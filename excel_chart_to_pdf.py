# -*- coding: utf-8 -*-
"""
Excel 图表批量导出 PDF 工具
================================
功能：
  * 选择一个 .xlsx / .xlsm 文件
  * 直接读取文件中所有 Sheet 的图表（柱状图、折线图等）
  * 只导出图表本身，不要数据表、不要单元格
  * 每个图表单独占 PDF 一页，按 Sheet 顺序排列
  * A4 横向，图表居中放大，打印清晰
  * 调用本机 Microsoft Excel 渲染图表（标题、图例、坐标轴、数据标签、颜色、字体与原文件一致）
  * 不修改原始 Excel 文件，不联网

做法：
  * 不用 Excel 自带的 PDF 导出（实测 Workbook.ExportAsFixedFormat 在 COM
    调用下反复报“无法保存该文档”；Chart.ExportAsFixedFormat 则会把整个
    工作表含单元格都导出来）。
  * 导出前先把图表自带的数据表关掉（HasDataTable=False），再用
    Chart.Export 把图表单独存为高清 PNG 图片（不含数据表、不含单元格），
    最后用 reportlab 把图片拼成一个 PDF，每张图一页、A4 横向、居中放大。

运行环境：Windows + Python 3 + Microsoft Excel
依赖：pywin32、reportlab（只需装一次：pip install pywin32 reportlab）
界面：tkinter（Python 自带）

用法：双击运行本文件，或 python excel_chart_to_pdf.py
"""

import os
import sys
import queue
import tempfile
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# A4 横向可用区域（单位：点 pt；1cm = 28.35pt）
# A4 横向 29.7cm x 21cm，上下左右各留 1cm 边距
# 可用宽 = 27.7cm = 785pt，可用高 = 19cm = 539pt
PAGE_W = 785.0
PAGE_H = 539.0
MARGIN_PT = 28.35  # 1cm

# 导出 PNG 时把图表放大的倍数（相对填满 A4 的尺寸），保证图片够清晰
EXPORT_SCALE = 2.0


# ------------------------------------------------------------------
# 核心导出逻辑（调用本机 Excel 导出图片，再拼成 PDF）
# ------------------------------------------------------------------
def export_charts_to_pdf(excel_path, progress_cb=None, log_cb=None):
    """把文件中所有 Sheet 的图表导出为一个 PDF（按 Sheet 顺序，只含图表）。

    参数：
        excel_path : Excel 文件完整路径
        progress_cb: 进度回调 progress_cb(done, total, text)
        log_cb     : 日志回调 log_cb(text)
    返回：
        (pdf_path, warnings)  warnings 为跳过说明列表
    异常：
        出错时抛异常，异常信息即为具体原因。
    """
    try:
        import win32com.client as win32
    except ImportError:
        raise RuntimeError(
            "缺少 pywin32，请先安装一次：\n"
            "    pip install pywin32\n"
            "装好后重新运行本工具。"
        )
    try:
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.lib.utils import ImageReader
        from reportlab.pdfgen import canvas as pdf_canvas
    except ImportError:
        raise RuntimeError(
            "缺少 reportlab，请先安装一次：\n"
            "    pip install reportlab\n"
            "装好后重新运行本工具。"
        )

    if not os.path.isfile(excel_path):
        raise FileNotFoundError("找不到文件：\n" + excel_path)

    ext = os.path.splitext(excel_path)[1].lower()
    if ext not in (".xlsx", ".xlsm", ".xls"):
        raise ValueError(
            "不支持的文件格式：%s\n仅支持 .xlsx / .xlsm / .xls" % ext
        )

    pdf_path = os.path.splitext(excel_path)[0] + "Charts.pdf"

    excel = None
    wb = None
    tmpdir = tempfile.mkdtemp(prefix="excel_charts_")
    pngs = []
    try:
        # 用一个独立的 Excel 进程，不干扰用户已打开的 Excel
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False

        # 只读打开，不改原文件
        try:
            wb = excel.Workbooks.Open(excel_path, ReadOnly=True)
        except Exception as e:
            raise RuntimeError(
                "无法打开 Excel 文件。\n"
                "可能文件正被其他程序占用，或文件已损坏。\n"
                "请关闭占用该文件的程序后重试。\n"
                "原始错误：%s" % e
            )

        # ---- 按 Sheet 顺序收集所有图表 ----
        jobs = []       # (sheet名, ChartObject)
        for ws in wb.Worksheets:
            cos = ws.ChartObjects()
            for idx in range(1, cos.Count + 1):
                jobs.append((ws.Name, cos.Item(idx)))

        if not jobs:
            raise ValueError("文件中没有找到任何图表，无法导出。")

        total = len(jobs)
        for n, (sheet_name, co) in enumerate(jobs, 1):
            title = ""
            try:
                if co.Chart.HasTitle:
                    title = co.Chart.ChartTitle.Text
            except Exception:
                title = ""
            label = "[%d/%d] %s%s" % (
                n, total, sheet_name,
                ("：" + title) if title else "",
            )
            if log_cb:
                log_cb(label)
            if progress_cb:
                progress_cb(n - 1, total, "正在处理：%s" % label)

            cht = co.Chart

            # 关掉图表自带的数据表（只要图表，不要表格）
            hadTable = False
            try:
                hadTable = cht.HasDataTable
            except Exception:
                hadTable = False
            try:
                cht.HasDataTable = False
            except Exception:
                pass

            # 临时放大图表，保证导出的 PNG 够清晰，导出完恢复
            ow = float(co.Width)
            oh = float(co.Height)
            scaled = False
            if ow > 0 and oh > 0:
                fill = min(PAGE_W / ow, PAGE_H / oh)
                nw = ow * fill * EXPORT_SCALE
                nh = oh * fill * EXPORT_SCALE
                co.Width = nw
                co.Height = nh
                scaled = True

            # 把图表单独存为 PNG（只含图表，不含数据表、不含单元格）
            png_path = os.path.join(tmpdir, "chart_%02d.png" % n)
            try:
                cht.Export(png_path, "PNG")
            except Exception as e:
                raise RuntimeError(
                    "导出图表图片失败（%s）：\n%s" % (label, e)
                )
            pngs.append(png_path)

            # 恢复图表原始大小和数据表开关（源文件是只读打开的，本来也不会保存）
            if scaled:
                try:
                    co.Width = ow
                    co.Height = oh
                except Exception:
                    pass
            try:
                cht.HasDataTable = hadTable
            except Exception:
                pass

            if progress_cb:
                progress_cb(n, total, "已完成：%s" % label)

        # ---- 用 reportlab 把图片拼成一个 PDF（每张图一页，A4 横向）----
        if os.path.exists(pdf_path):
            try:
                os.remove(pdf_path)
            except OSError:
                raise RuntimeError(
                    "无法覆盖旧的 PDF：\n%s\n"
                    "可能它正被 PDF 阅读器打开着，请关闭后重试。" % pdf_path
                )

        page_w, page_h = landscape(A4)  # 841.89 x 595.28 pt
        c = pdf_canvas.Canvas(pdf_path, pagesize=landscape(A4))
        c.setTitle(os.path.basename(pdf_path))
        for png_path in pngs:
            img = ImageReader(png_path)
            iw, ih = img.getSize()
            if iw <= 0 or ih <= 0:
                continue
            # 按比例放大到填满可用区域并居中
            scale = min(PAGE_W / iw, PAGE_H / ih)
            dw = iw * scale
            dh = ih * scale
            x = MARGIN_PT + (PAGE_W - dw) / 2.0
            y = MARGIN_PT + (PAGE_H - dh) / 2.0
            c.drawImage(img, x, y, width=dw, height=dh,
                        preserveAspectRatio=True, mask="auto")
            c.showPage()
        c.save()

        if not os.path.isfile(pdf_path):
            raise RuntimeError("PDF 未生成，未知原因。请检查 Excel 是否正常。")

        return pdf_path, []

    finally:
        # 清理临时图片
        for png_path in pngs:
            try:
                if os.path.exists(png_path):
                    os.remove(png_path)
            except Exception:
                pass
        try:
            if tmpdir and os.path.isdir(tmpdir):
                os.rmdir(tmpdir)
        except Exception:
            pass
        # 关闭 Excel（源文件只读打开，不保存）
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if excel is not None:
                excel.Quit()
        except Exception:
            pass


def open_folder(path):
    """在资源管理器中打开 path 所在文件夹（Windows）。"""
    if sys.platform.startswith("win"):
        try:
            os.startfile(os.path.dirname(path))
        except Exception:
            os.startfile(path)


# ------------------------------------------------------------------
# GUI
# ------------------------------------------------------------------
class App:
    def __init__(self, root):
        self.root = root
        root.title("Excel 图表批量导出 PDF")
        root.geometry("560x340")
        root.resizable(False, False)

        self.file_var = tk.StringVar()
        self.status_var = tk.StringVar(value="请选择 Excel 文件。")
        self.q = queue.Queue()
        self.exporting = False
        self.last_pdf = None

        # --- 文件选择 ---
        frm_file = ttk.LabelFrame(root, text="选择 Excel 文件", padding=10)
        frm_file.pack(fill="x", padx=12, pady=(12, 6))
        self.entry = ttk.Entry(frm_file, textvariable=self.file_var)
        self.entry.pack(side="left", fill="x", expand=True)
        ttk.Button(frm_file, text="浏览…", command=self.browse).pack(
            side="left", padx=(8, 0))

        # --- 说明 ---
        frm_info = ttk.Frame(root, padding=(12, 2))
        frm_info.pack(fill="x")
        ttk.Label(
            frm_info,
            text="只导出文件中所有 Sheet 的图表（不要数据表、不要单元格），按 Sheet 顺序排列，每个图表单独占一页。",
            foreground="#666666",
            wraplength=520,
            justify="left",
        ).pack(anchor="w")

        # --- 导出按钮 + 进度 ---
        frm_run = ttk.Frame(root, padding=10)
        frm_run.pack(fill="x", padx=12, pady=6)
        self.btn_export = ttk.Button(
            frm_run, text="导出图表 PDF", command=self.start_export)
        self.btn_export.pack(fill="x")
        self.progress = ttk.Progressbar(
            frm_run, orient="horizontal", mode="determinate", maximum=100)
        self.progress.pack(fill="x", pady=(10, 4))
        ttk.Label(frm_run, textvariable=self.status_var).pack(anchor="w")

        # --- 日志 ---
        frm_log = ttk.LabelFrame(root, text="日志", padding=10)
        frm_log.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        self.log = tk.Text(frm_log, height=5, state="disabled")
        self.log.pack(fill="both", expand=True)

    # -- 小工具 --
    def append_log(self, text):
        self.log.configure(state="normal")
        self.log.insert("end", text + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def browse(self):
        path = filedialog.askopenfilename(
            title="选择 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm *.xls"),
                       ("所有文件", "*.*")],
        )
        if path:
            self.file_var.set(path)
            self.status_var.set("已选择：%s" % os.path.basename(path))

    # -- 导出流程（后台线程 + 队列回主线程更新界面）--
    def start_export(self):
        if self.exporting:
            return
        excel_path = self.file_var.get().strip()
        if not excel_path:
            messagebox.showwarning("未选择文件", "请先选择一个 Excel 文件。")
            return
        if not os.path.isfile(excel_path):
            messagebox.showerror("找不到文件", "找不到文件：\n" + excel_path)
            return

        self.exporting = True
        self.btn_export.configure(state="disabled")
        self.progress["value"] = 0
        self.status_var.set("正在启动 Excel…")
        self.append_log("开始导出：%s" % excel_path)

        def worker():
            def progress_cb(done, total, text):
                self.q.put(("progress", done, total, text))

            def log_cb(text):
                self.q.put(("log", text))

            try:
                pdf_path, warnings = export_charts_to_pdf(
                    excel_path, progress_cb=progress_cb, log_cb=log_cb)
                self.q.put(("done", pdf_path, warnings))
            except Exception as e:
                self.q.put(("error", str(e)))

        threading.Thread(target=worker, daemon=True).start()
        self.root.after(100, self.poll_queue)

    def poll_queue(self):
        try:
            while True:
                item = self.q.get_nowait()
                kind = item[0]
                if kind == "progress":
                    _, done, total, text = item
                    if total > 0:
                        self.progress["value"] = done * 100.0 / total
                    self.status_var.set(text)
                elif kind == "log":
                    self.append_log(item[1])
                elif kind == "done":
                    _, pdf_path, warnings = item
                    self.finish_success(pdf_path, warnings)
                    return
                elif kind == "error":
                    self.finish_error(item[1])
                    return
        except queue.Empty:
            pass
        if self.exporting:
            self.root.after(100, self.poll_queue)

    def finish_success(self, pdf_path, warnings):
        self.exporting = False
        self.btn_export.configure(state="normal")
        self.progress["value"] = 100
        self.last_pdf = pdf_path
        self.status_var.set("导出成功：%s" % os.path.basename(pdf_path))
        self.append_log("导出成功：%s" % pdf_path)
        for w in warnings:
            self.append_log("提示：" + w)

        msg = "导出成功！\n\nPDF 已保存到：\n%s" % pdf_path
        if warnings:
            msg += "\n\n提示：\n" + "\n".join(warnings)
        msg += "\n\n是否打开 PDF 所在文件夹？"
        if messagebox.askyesno("导出成功", msg):
            open_folder(pdf_path)

    def finish_error(self, err_text):
        self.exporting = False
        self.btn_export.configure(state="normal")
        self.progress["value"] = 0
        self.status_var.set("导出失败。")
        self.append_log("导出失败：" + err_text)
        messagebox.showerror("导出失败", err_text)


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
