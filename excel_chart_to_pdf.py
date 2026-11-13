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
  * 调用本机 Microsoft Excel 渲染（标题、图例、坐标轴、数据标签、颜色、字体与原文件一致）
  * 不修改原始 Excel 文件，不联网

做法：
  * 导出前先把图表自带的数据表关掉（HasDataTable=False），再用 CopyPicture
    （xlPrinter，高清）把图表复制为图片，粘贴到临时工作簿的一张表上，
    临时表设为 A4 横向、居中、一页，最后整本导出为一个 PDF。
  * 这样 PDF 里只有图表，没有数据表，也没有单元格内容；页是 A4 横向。
  * 不用 Chart.ExportAsFixedFormat 直接导出图表——实测它会把整个工作表
    （含单元格）都导出来，页数和纸张都不对。

运行环境：Windows + Python 3 + Microsoft Excel
依赖：pywin32（只需装一次：pip install pywin32）
界面：tkinter（Python 自带）

用法：双击运行本文件，或 python excel_chart_to_pdf.py
"""

import os
import sys
import queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Excel 常量（用数字，避免依赖 win32com.client.constants）
XL_PAPER_A4 = 9            # A4
XL_LANDSCAPE = 2           # 横向
XL_TYPE_PDF = 0            # xlTypePDF
XL_QUALITY_STANDARD = 0    # 标准质量
XL_PRINTER = 2             # CopyPicture Appearance:=xlPrinter（高清）
XL_PICTURE = -4147         # CopyPicture Format:=xlPicture

# A4 横向可用区域（单位：点 pt；1cm = 28.35pt）
# A4 横向 29.7cm x 21cm，上下左右各留 1cm 边距
# 可用宽 = 27.7cm = 785pt，可用高 = 19cm = 539pt
PAGE_W = 785.0
PAGE_H = 539.0
MARGIN_PT = 28.35  # 1cm


# ------------------------------------------------------------------
# 核心导出逻辑（调用本机 Excel）
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
    temp_wb = None
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

        # ---- 创建临时工作簿（只放图表图片，每张表一个图表）----
        temp_wb = excel.Workbooks.Add()
        while temp_wb.Worksheets.Count > 1:
            temp_wb.Worksheets(temp_wb.Worksheets.Count).Delete()

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

            if n == 1:
                tws = temp_wb.Worksheets(1)
            else:
                tws = temp_wb.Worksheets.Add(
                    After=temp_wb.Worksheets(temp_wb.Worksheets.Count))
            tws.Name = "Chart_%d" % n

            # 页面设置：A4 横向，1cm 边距，居中，一页
            ps = tws.PageSetup
            ps.PaperSize = XL_PAPER_A4
            ps.Orientation = XL_LANDSCAPE
            ps.LeftMargin = MARGIN_PT
            ps.RightMargin = MARGIN_PT
            ps.TopMargin = MARGIN_PT
            ps.BottomMargin = MARGIN_PT
            ps.HeaderMargin = 0
            ps.FooterMargin = 0
            ps.CenterHorizontally = True
            ps.CenterVertically = True
            ps.Zoom = False
            ps.FitToPagesWide = 1
            ps.FitToPagesTall = 1

            # 关掉图表自带的数据表（只要图表，不要表格），再高清复制为图片
            cht = co.Chart
            hadTable = False
            try:
                hadTable = cht.HasDataTable
            except Exception:
                hadTable = False
            try:
                cht.HasDataTable = False
            except Exception:
                pass
            co.CopyPicture(Appearance=XL_PRINTER, Format=XL_PICTURE)
            # 立刻恢复数据表开关（源文件是只读打开的，本来也不会保存）
            try:
                cht.HasDataTable = hadTable
            except Exception:
                pass

            # 粘贴图片到临时表，放大到填满可用区域并居中
            tws.Activate()
            tws.Paste()
            shapes = tws.Shapes
            pic = shapes.Item(shapes.Count)
            ow = float(pic.Width)
            oh = float(pic.Height)
            if ow > 0 and oh > 0:
                scale = min(PAGE_W / ow, PAGE_H / oh)
                nw = ow * scale
                nh = oh * scale
                try:
                    pic.LockAspectRatio = 0  # msoFalse
                except Exception:
                    pass
                pic.Width = nw
                pic.Height = nh
                pic.Left = MARGIN_PT + (PAGE_W - nw) / 2.0
                pic.Top = MARGIN_PT + (PAGE_H - nh) / 2.0

            # 打印区域只框图片所在范围，保证一页一个图表
            try:
                tl = pic.TopLeftCell
                br = pic.BottomRightCell
                ps.PrintArea = "%s:%s" % (tl.Address, br.Address)
            except Exception:
                pass

            if progress_cb:
                progress_cb(n, total, "已完成：%s" % label)

        # ---- 导出整个临时工作簿为一个 PDF（每张临时表 = 一页）----
        if os.path.exists(pdf_path):
            try:
                os.remove(pdf_path)
            except OSError:
                raise RuntimeError(
                    "无法覆盖旧的 PDF：\n%s\n"
                    "可能它正被 PDF 阅读器打开着，请关闭后重试。" % pdf_path
                )

        temp_wb.ExportAsFixedFormat(
            Type=XL_TYPE_PDF,
            Filename=pdf_path,
            Quality=XL_QUALITY_STANDARD,
            IncludeDocProperties=True,
            IgnorePrintAreas=False,
            OpenAfterPublish=False,
        )

        if not os.path.isfile(pdf_path):
            raise RuntimeError("PDF 未生成，未知原因。请检查 Excel 是否正常。")

        return pdf_path, []

    finally:
        # 清理：先关临时工作簿，再关源文件，最后退出 Excel
        try:
            if temp_wb is not None:
                temp_wb.Close(SaveChanges=False)
        except Exception:
            pass
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
