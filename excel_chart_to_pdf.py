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

做法（为什么原生打印清楚、之前的方法糊）：
  * Excel 原生打印 / 另存为 PDF 清楚，是因为它是矢量图：柱子、文字、坐标轴
    都是画出来的线条和文字，放大多少倍都不会糊。
  * 之前 v5–v13 用的是 Chart.Export 存 PNG 图片，那是点阵图，只有屏幕
    分辨率（约 96 DPI），放大到 A4 就糊；事后用 PIL 放大像素也补不回
    真正的细节，只能让边缘稍微平滑一点。
  * v15 的主办法（直接矢量导出）：每个图表直接调 Chart.ExportAsFixedFormat
    导成一个矢量 PDF，再用 pypdf 把它们拼到 A4 横向页面上、合成一个 PDF。
    拼的时候还是矢量，不会糊，和原生打印一样清楚。不用临时工作簿、
    不用剪贴板、没有外部链接。关键：全程不改图表对象的尺寸（v5–v12
    试过了，改尺寸都会变形），放大靠合并时的缩放来做，比例一定是对的。
  * v14 的临时工作簿办法留着当第二选择；万一两条矢量路都失败，自动退回
    v13 的 PNG 办法，至少还能出一份 PDF，不会报错退出。
  * 走的是哪条路会明明白白显示在状态栏里：如果退回图片方式，会弹框提醒，
    不会让用户拿着糊图还以为是矢量。

运行环境：Windows + Python 3 + Microsoft Excel
依赖：pywin32、reportlab、pillow、pypdf（只需装一次：
  pip install pywin32 reportlab pillow pypdf）
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

VERSION = "v15"

# A4 横向可用区域（单位：点 pt；1cm = 28.35pt）
# A4 横向 29.7cm x 21cm，上下左右各留 1cm 边距
# 可用宽 = 27.7cm = 785pt，可用高 = 19cm = 539pt
PAGE_W = 785.0
PAGE_H = 539.0
MARGIN_PT = 28.35  # 1cm

# PNG 兜底时图片像素放大倍数（PIL 用，只改像素、不改比例）
EXPORT_SCALE = 3.0


# ------------------------------------------------------------------
# 核心导出逻辑
# ------------------------------------------------------------------
def _collect_charts(wb):
    """按 Sheet 顺序收集所有图表，返回 [(sheet名, ChartObject)]。"""
    jobs = []
    for ws in wb.Worksheets:
        cos = ws.ChartObjects()
        for idx in range(1, cos.Count + 1):
            jobs.append((ws.Name, cos.Item(idx)))
    return jobs


def _chart_label(n, total, sheet_name, co):
    title = ""
    try:
        if co.Chart.HasTitle:
            title = co.Chart.ChartTitle.Text
    except Exception:
        title = ""
    return "[%d/%d] %s%s" % (
        n, total, sheet_name,
        ("：" + title) if title else "",
    )


def _hide_chart_border(cht):
    """去掉图表外框的黑边：ChartArea 默认带边框，导出后会变成黑框。"""
    try:
        cht.ChartArea.Format.Line.Visible = 0  # msoFalse，隐藏边框
    except Exception:
        try:
            cht.ChartArea.Border.LineStyle = -4142  # xlLineStyleNone
        except Exception:
            pass


def _remove_old_pdf(pdf_path):
    """旧 PDF 先删掉，避免覆盖失败。"""
    if os.path.exists(pdf_path):
        try:
            os.remove(pdf_path)
        except OSError:
            raise RuntimeError(
                "无法覆盖旧的 PDF：\n%s\n"
                "可能它正被 PDF 阅读器打开着，请关闭后重试。" % pdf_path
            )


def _export_vector_direct(excel, wb_src, jobs, pdf_path, tmpdir,
                          progress_cb=None, log_cb=None):
    """直接矢量导出（主办法）：每个图表直接导成矢量 PDF，
    再用 pypdf 把它们拼到 A4 横向页面上、合成一个 PDF。

    不用临时工作簿、不用剪贴板、没有外部链接。
    全程不改图表对象尺寸，放大靠合并时的缩放，比例不会变形。
    成功返回 None；失败抛异常。
    """
    try:
        from pypdf import PdfReader, PdfWriter, Transformation
        from reportlab.lib.pagesizes import A4, landscape
    except ImportError as e:
        raise RuntimeError("缺少 pypdf 或 reportlab：%s" % e)

    total = len(jobs)
    chart_pdfs = []
    try:
        for n, (sheet_name, co) in enumerate(jobs, 1):
            label = _chart_label(n, total, sheet_name, co)
            if log_cb:
                log_cb(label)
            if progress_cb:
                progress_cb(n - 1, total, "正在处理：%s" % label)

            cht = co.Chart
            # 数据表是图表的一部分，跟着图表一起导出；
            # 数据表数字按原格式显示（比如 21.0，不会变成 21.0072212）
            _hide_chart_border(cht)

            chart_pdf = os.path.join(tmpdir, "chart_%02d.pdf" % n)
            try:
                cht.ExportAsFixedFormat(0, chart_pdf, 0, True, False)
            except Exception as e:
                raise RuntimeError(
                    "导出图表 PDF 失败（%s）：\n%s" % (label, e)
                )
            if not os.path.isfile(chart_pdf):
                raise RuntimeError("图表 PDF 未生成（%s）" % label)
            chart_pdfs.append(chart_pdf)

            if progress_cb:
                progress_cb(n, total, "已完成：%s" % label)

        _remove_old_pdf(pdf_path)

        # 用 pypdf 把各图表 PDF 拼到 A4 横向页面上（保持矢量）
        page_w, page_h = landscape(A4)  # 841.89 x 595.28 pt
        writer = PdfWriter()
        for chart_pdf in chart_pdfs:
            reader = PdfReader(chart_pdf)
            if len(reader.pages) == 0:
                continue
            chart_page = reader.pages[0]
            cw = float(chart_page.mediabox.width)
            ch = float(chart_page.mediabox.height)
            if cw <= 0 or ch <= 0:
                continue
            # 按原比例放大到填满可用区域并居中，绝不拉伸变形
            scale = min(PAGE_W / cw, PAGE_H / ch)
            new_page = writer.add_blank_page(width=page_w, height=page_h)
            tx = MARGIN_PT + (PAGE_W - cw * scale) / 2.0
            ty = MARGIN_PT + (PAGE_H - ch * scale) / 2.0
            t = Transformation().scale(scale).translate(tx, ty)
            new_page.merge_transformed_page(chart_page, t)
        writer.add_metadata({"/Title": os.path.basename(pdf_path)})
        with open(pdf_path, "wb") as f:
            writer.write(f)

        if not os.path.isfile(pdf_path):
            raise RuntimeError("PDF 未生成，未知原因。请检查 Excel 是否正常。")
    finally:
        for p in chart_pdfs:
            try:
                if os.path.exists(p):
                    os.remove(p)
            except Exception:
                pass


def _export_vector_tempwb(excel, wb_src, jobs, pdf_path,
                          progress_cb=None, log_cb=None):
    """矢量导出（第二选择，v14 办法）：把活图表复制到临时工作簿，
    用页面缩放填满 A4，再用 Excel 自带的 ExportAsFixedFormat 导出矢量 PDF。

    全程不改图表对象尺寸，放大靠 PageSetup.Zoom，比例不会变形。
    成功返回 None；失败抛异常。
    """
    total = len(jobs)
    wb_tmp = None
    try:
        wb_tmp = excel.Workbooks.Add()
        # 临时工作簿正好需要 len(jobs) 个 Sheet，多了删、少了补
        while wb_tmp.Worksheets.Count < total:
            wb_tmp.Worksheets.Add(
                After=wb_tmp.Worksheets(wb_tmp.Worksheets.Count))
        while wb_tmp.Worksheets.Count > total:
            wb_tmp.Worksheets(wb_tmp.Worksheets.Count).Delete()

        for n, (sheet_name, co) in enumerate(jobs, 1):
            label = _chart_label(n, total, sheet_name, co)
            if log_cb:
                log_cb(label)
            if progress_cb:
                progress_cb(n - 1, total, "正在处理：%s" % label)

            ws_tmp = wb_tmp.Worksheets(n)
            _hide_chart_border(co.Chart)

            # 复制活图表（不是图片），粘贴到临时表里。
            co.Copy()
            ws_tmp.Activate()
            ws_tmp.Paste()
            pasted = ws_tmp.ChartObjects(ws_tmp.ChartObjects().Count)
            # 放到左上角，不改尺寸
            try:
                pasted.Left = 0
                pasted.Top = 0
            except Exception:
                pass

            # 页面缩放：按原比例放大到填满 A4 可用区域。
            ow = float(pasted.Width)
            oh = float(pasted.Height)
            if ow > 0 and oh > 0:
                zoom = int(round(min(PAGE_W / ow, PAGE_H / oh) * 100))
                zoom = max(10, min(400, zoom))
            else:
                zoom = 100

            ps = ws_tmp.PageSetup
            ps.PaperSize = 9      # xlPaperA4
            ps.Orientation = 2    # xlLandscape 横向
            ps.LeftMargin = MARGIN_PT
            ps.RightMargin = MARGIN_PT
            ps.TopMargin = MARGIN_PT
            ps.BottomMargin = MARGIN_PT
            ps.Zoom = zoom
            try:
                ps.FitToPages = False
            except Exception:
                pass
            # 打印区域只含图表所在的单元格
            try:
                tl = pasted.TopLeftCell.Address
                br = pasted.BottomRightCell.Address
                ps.PrintArea = "%s:%s" % (tl, br)
            except Exception:
                pass

            if progress_cb:
                progress_cb(n, total, "已完成：%s" % label)

        _remove_old_pdf(pdf_path)

        # 导出整个临时工作簿为 PDF（矢量，和原生打印一样清楚）
        wb_tmp.ExportAsFixedFormat(0, pdf_path, 0, True, False)

        if not os.path.isfile(pdf_path):
            raise RuntimeError("PDF 未生成，未知原因。请检查 Excel 是否正常。")
    finally:
        try:
            if wb_tmp is not None:
                wb_tmp.Close(SaveChanges=False)
        except Exception:
            pass


def _export_png(excel, wb_src, jobs, pdf_path, tmpdir,
                progress_cb=None, log_cb=None):
    """PNG 兜底（v13 办法）：Chart.Export 存 PNG，再用 reportlab 拼成 PDF。
    比例是对的，但图片是点阵的，清晰度不如矢量导出。
    """
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

    total = len(jobs)
    pngs = []
    try:
        for n, (sheet_name, co) in enumerate(jobs, 1):
            label = _chart_label(n, total, sheet_name, co)
            if log_cb:
                log_cb(label)
            if progress_cb:
                progress_cb(n - 1, total, "正在处理：%s" % label)

            cht = co.Chart
            # 不改图表对象尺寸，按原尺寸导出，比例一定对
            _hide_chart_border(cht)

            png_path = os.path.join(tmpdir, "chart_%02d.png" % n)
            try:
                cht.Export(png_path, "PNG")
            except Exception as e:
                raise RuntimeError(
                    "导出图表图片失败（%s）：\n%s" % (label, e)
                )

            # 原图分辨率低，用 PIL 放大再锐化，只改像素、不改比例
            try:
                from PIL import Image, ImageFilter
                im = Image.open(png_path)
                w0, h0 = im.size
                if w0 > 0 and h0 > 0:
                    im = im.resize((int(w0 * EXPORT_SCALE),
                                    int(h0 * EXPORT_SCALE)), Image.LANCZOS)
                    im = im.filter(ImageFilter.UnsharpMask(
                        radius=2, percent=120, threshold=2))
                    im.save(png_path)
            except Exception:
                pass
            pngs.append(png_path)

            if progress_cb:
                progress_cb(n, total, "已完成：%s" % label)

        _remove_old_pdf(pdf_path)

        c = pdf_canvas.Canvas(pdf_path, pagesize=landscape(A4))
        c.setTitle(os.path.basename(pdf_path))
        for png_path in pngs:
            img = ImageReader(png_path)
            iw, ih = img.getSize()
            if iw <= 0 or ih <= 0:
                continue
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
    finally:
        for png_path in pngs:
            try:
                if os.path.exists(png_path):
                    os.remove(png_path)
            except Exception:
                pass


def export_charts_to_pdf(excel_path, progress_cb=None, log_cb=None):
    """把文件中所有 Sheet 的图表导出为一个 PDF（按 Sheet 顺序，只含图表）。

    依次尝试：直接矢量导出 → 临时工作簿矢量导出 → PNG 办法。
    返回 (pdf_path, warnings, method)，method 为 "矢量" 或 "图片"。
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
    if ext not in EXCEL_EXTS:
        raise ValueError(
            "不支持的文件格式：%s\n仅支持 .xlsx / .xlsm / .xls" % ext
        )

    pdf_path = os.path.splitext(excel_path)[0] + "Charts.pdf"

    excel = None
    wb = None
    tmpdir = tempfile.mkdtemp(prefix="excel_charts_")
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False

        try:
            wb = excel.Workbooks.Open(excel_path, ReadOnly=True)
        except Exception as e:
            raise RuntimeError(
                "无法打开 Excel 文件。\n"
                "可能文件正被其他程序占用，或文件已损坏。\n"
                "请关闭占用该文件的程序后重试。\n"
                "原始错误：%s" % e
            )

        jobs = _collect_charts(wb)
        if not jobs:
            raise ValueError("文件中没有找到任何图表，无法导出。")

        warnings = []
        # 1) 直接矢量导出（主办法）
        try:
            _export_vector_direct(excel, wb, jobs, pdf_path, tmpdir,
                                  progress_cb=progress_cb, log_cb=log_cb)
            return pdf_path, warnings, "矢量"
        except Exception as e:
            warnings.append("直接矢量导出失败：%s" % e)
            if log_cb:
                log_cb("直接矢量导出失败：%s" % e)

        # 2) 临时工作簿矢量导出（v14 办法）
        try:
            _export_vector_tempwb(excel, wb, jobs, pdf_path,
                                  progress_cb=progress_cb, log_cb=log_cb)
            return pdf_path, warnings, "矢量"
        except Exception as e:
            warnings.append("临时工作簿矢量导出失败：%s" % e)
            if log_cb:
                log_cb("临时工作簿矢量导出失败：%s" % e)

        # 3) PNG 办法兜底（比例对、但没矢量清楚）
        _export_png(excel, wb, jobs, pdf_path, tmpdir,
                    progress_cb=progress_cb, log_cb=log_cb)
        warnings.append("两条矢量路都失败了，已改用图片方式导出，"
                        "清晰度会比原生打印差。")
        return pdf_path, warnings, "图片"

    finally:
        try:
            if tmpdir and os.path.isdir(tmpdir):
                os.rmdir(tmpdir)
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
                pdf_path, warnings, method = export_charts_to_pdf(
                    excel_path, progress_cb=progress_cb, log_cb=log_cb)
                self.q.put(("done", pdf_path, warnings, method))
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
                    _, pdf_path, warnings, method = item
                    self.finish_success(pdf_path, warnings, method)
                    return
                elif kind == "error":
                    self.finish_error(item[1])
                    return
        except queue.Empty:
            pass
        if self.exporting:
            self.root.after(100, self.poll_queue)

    def finish_success(self, pdf_path, warnings, method):
        self.exporting = False
        self.btn_export.configure(state="normal")
        self.progress["value"] = 100
        self.last_pdf = pdf_path
        # 让用户一眼看出走的是哪条路：矢量才和原生打印一样清楚；
        # 如果退回图片方式，必须明说，不然用户拿到糊图还以为矢量没起作用。
        if method == "矢量":
            self.status_var.set("导出成功（矢量，和原生打印一样清楚）：%s"
                                % os.path.basename(pdf_path))
        else:
            self.status_var.set("导出成功（图片方式，清晰度一般）：%s"
                                % os.path.basename(pdf_path))
            messagebox.showwarning("已用图片方式导出",
                                   "两条矢量导出都失败了，已改用图片方式导出，"
                                   "清晰度会比原生打印差。\n\n" + "\n".join(warnings))
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
