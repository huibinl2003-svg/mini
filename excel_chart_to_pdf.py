# -*- coding: utf-8 -*-
"""
Excel 图表批量导出 PDF 工具
================================
功能：
  * 选择一个 .xlsx / .xlsm 文件
  * 直接读取文件中所有 Sheet 的图表（柱状图、折线图等）
  * 导出整个图表：图 + 图表自带的数据表 + 图例，一个都不少；
    但不要工作表的单元格（不要表格）
  * 每个图表单独占 PDF 一页，按 Sheet 顺序排列
  * A4 横向，图表按原比例放大居中，绝不拉伸变形，效果和 Excel 里看到的一模一样
  * 调用本机 Microsoft Excel 渲染图表（标题、图例、坐标轴、数据标签、颜色、字体与原文件一致）
  * 不修改原始 Excel 文件，不联网

做法：
  * 不用 Excel 自带的 PDF 导出（实测 Workbook.ExportAsFixedFormat 在 COM
    调用下反复报“无法保存该文档”；Chart.ExportAsFixedFormat 则会把整个
    工作表含单元格都导出来）。
  * 用 Chart.Export 把图表单独存为高清 PNG 图片（图表 + 自带的数据表，
    不含工作表单元格），再用 reportlab 把图片拼成 PDF，每张图一页、
    A4 横向、按原比例放大居中。Chart.Export 是按屏幕显示渲染的，数据表里
    的数字按原格式显示（比如 21.0，不会变成 21.0072212）。
    导出前先把图表对象放大以提高清晰度：先锁定纵横比再改宽度，高度自动
    按比例变，导完读回尺寸确认比例没变形才用；一旦比例变了就恢复原尺寸
    导出（再用 PIL 放大像素兜底）。直接改宽高不锁定纵横比会把图拉变形。

运行环境：Windows + Python 3 + Microsoft Excel
依赖：pywin32、reportlab、pillow（只需装一次：pip install pywin32 reportlab pillow）
界面：tkinter（Python 自带）

用法：双击运行本文件，或 python excel_chart_to_pdf.py
选文件方式：点「浏览…」选一个 Excel 文件，再点「导出图表 PDF」。
"""

import os
import sys
import queue
import tempfile
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

EXCEL_EXTS = (".xlsx", ".xlsm", ".xls")

VERSION = "v12"

# A4 横向可用区域（单位：点 pt；1cm = 28.35pt）
# A4 横向 29.7cm x 21cm，上下左右各留 1cm 边距
# 可用宽 = 27.7cm = 785pt，可用高 = 19cm = 539pt
PAGE_W = 785.0
PAGE_H = 539.0
MARGIN_PT = 28.35  # 1cm

# 临时放大图表对象的倍数（相对填满 A4 的尺寸），导出高清 PNG 用；
# 放大时先锁定纵横比，只改宽度，高度自动按比例变，保证不变形
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
    if ext not in EXCEL_EXTS:
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

            # 数据表是图表的一部分，要一起导出（不要关掉）
            # 注意：这里不用 CopyPicture，Chart.Export 是按屏幕显示渲染的，
            # 数据表里的数字会按原格式显示（比如 21.0，不会变成 21.0072212）

            # 临时放大图表以提高清晰度：先锁定纵横比，只改宽度，高度自动按比例变，
            # 保证不变形；导完恢复原尺寸。v11 不改尺寸所以比例对但图糊，
            # 这里用锁定纵横比的方式放大，既清晰又不变形。
            ow = float(co.Width)
            oh = float(co.Height)
            scaled = False
            if ow > 0 and oh > 0:
                fill = min(PAGE_W / ow, PAGE_H / oh)
                target_w = ow * fill * EXPORT_SCALE
                try:
                    co.ShapeRange.LockAspectRatio = 1  # msoTrue，锁定纵横比
                except Exception:
                    pass
                try:
                    co.Width = target_w
                except Exception:
                    pass
                # 读回实际尺寸，确认比例没被拉变形
                try:
                    nw2 = float(co.Width)
                    nh2 = float(co.Height)
                    if nw2 > 0 and nh2 > 0:
                        r0 = ow / oh
                        r1 = nw2 / nh2
                        if abs(r1 - r0) / r0 < 0.01:
                            scaled = True
                        else:
                            # 比例变了，恢复原尺寸，按原尺寸导出
                            try:
                                co.Width = ow
                                co.Height = oh
                            except Exception:
                                pass
                            scaled = False
                except Exception:
                    pass

            # 去掉图表外框的黑边：ChartArea 默认带边框，导出后会变成黑框
            try:
                cht.ChartArea.Format.Line.Visible = 0  # msoFalse，隐藏边框
            except Exception:
                try:
                    cht.ChartArea.Border.LineStyle = -4142  # xlLineStyleNone
                except Exception:
                    pass

            # 按当前尺寸导出 PNG（图表 + 自带的数据表，不含工作表单元格）
            png_path = os.path.join(tmpdir, "chart_%02d.png" % n)
            try:
                cht.Export(png_path, "PNG")
            except Exception as e:
                raise RuntimeError(
                    "导出图表图片失败（%s）：\n%s" % (label, e)
                )

            # 恢复图表原始大小和纵横比锁定（源文件是只读打开的，本来也不会保存）
            if scaled:
                try:
                    co.ShapeRange.LockAspectRatio = 0  # msoFalse，解除锁定
                except Exception:
                    pass
                try:
                    co.Width = ow
                    co.Height = oh
                except Exception:
                    pass

            # 如果没能放大（回落到原尺寸），原图分辨率低，用 PIL 放大 2 倍
            # （LANCZOS），只改像素、不改比例，让图清楚一点。
            try:
                from PIL import Image
                im = Image.open(png_path)
                w0, h0 = im.size
                if w0 > 0 and h0 > 0 and w0 < 1600:
                    im.resize((w0 * 2, h0 * 2), Image.LANCZOS).save(png_path)
            except Exception:
                pass
            pngs.append(png_path)

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
            # 按原比例放大到填满可用区域并居中，绝不拉伸变形
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


# ------------------------------------------------------------------
# GUI
# ------------------------------------------------------------------
class App:
    def __init__(self, root):
        self.root = root
        root.title("Excel 图表批量导出 PDF  " + VERSION)
        root.geometry("500x210")
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

    # -- 小工具 --
    def append_log(self, text):
        pass  # 日志显示已去掉，只留状态栏

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
        # 不再弹“是否打开文件夹”的确认框，只在状态栏和日志里给路径

    def finish_error(self, err_text):
        self.exporting = False
        self.btn_export.configure(state="normal")
        self.progress["value"] = 0
        self.status_var.set("导出失败。")
        self.append_log("导出失败：" + err_text)
        messagebox.showerror("导出失败", err_text)


def main():
    root = tk.Tk()
    app = App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
