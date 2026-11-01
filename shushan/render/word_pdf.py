# -*- coding: utf-8 -*-
"""Word 真渲染：docx → PDF（后台 Word COM）→ 逐页 PNG。

页面视图的数据管线（只读渲染，不承担编辑——写盘仍由 python-docx 整稿引擎
负责）：
  1. export_pdf_via_word：用独立 COM 实例（DispatchEx，不影响用户开着的
     Word）只读打开文档 → 先刷新域（目录页码真实生成）→ 导出 PDF →
     finally 关闭并退出，防后台残留；
  2. render_pdf_pages：pypdfium2 逐页渲染 PNG 字节；
  3. render_pages_from_docx：组合入口（工作台线程与测试共用）。

依赖：Windows + 本机 Word；pypdfium2（requirements.txt）。
"""
import tempfile
from pathlib import Path

PDF_EXPORT_FORMAT = 17          # wdExportFormatPDF
_PDF_MAGIC = b"%PDF-"


def export_pdf_via_word(docx_path, out_pdf) -> Path:
    """后台 Word 把 docx 导出为 PDF（返回 out_pdf 路径）。

    路径一律转绝对——Word COM 以自身工作目录解析相对路径（会落到 system32）。
    """
    docx_path = str(Path(docx_path).resolve())
    out_pdf = str(Path(out_pdf).resolve())
    if not Path(docx_path).exists():
        raise FileNotFoundError(docx_path)
    try:
        import pythoncom
        import win32com.client
    except ImportError as e:
        raise RuntimeError("页面视图需要 pywin32（Word COM 渲染）：" + str(e)) from e
    pythoncom.CoInitialize()
    word = None
    doc = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(docx_path, ReadOnly=True, AddToRecentFiles=False)
        # 刷新域：目录页码/引用编号在 PDF 里真实生成
        try:
            for toc in doc.TablesOfContents:
                toc.Update()
        except Exception:
            pass
        try:
            doc.Fields.Update()
        except Exception:
            pass
        doc.ExportAsFixedFormat(OutputFileName=out_pdf, ExportFormat=PDF_EXPORT_FORMAT)
    except Exception as e:
        raise RuntimeError(
            f"Word 渲染失败（请确认已安装 Word；若 Word 卡在对话框，"
            f"可用工作台「清理Word后台」后重试）：{e}") from e
    finally:
        try:
            if doc is not None:
                doc.Close(SaveChanges=0)
        except Exception:
            pass
        try:
            if word is not None:
                word.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()
    out = Path(out_pdf)
    if not out.exists() or out.read_bytes()[:5] != _PDF_MAGIC:
        raise RuntimeError("Word 未生成有效 PDF（导出中断）")
    return out


def render_pdf_pages(pdf_path, dpi: int = 110) -> list[bytes]:
    """PDF → 逐页 PNG 字节列表（pypdfium2）。"""
    import pypdfium2 as pdfium

    scale = dpi / 72.0
    pdf = pdfium.PdfDocument(str(pdf_path))
    out = []
    try:
        for page in pdf:
            bitmap = page.render(scale=scale)
            pil = bitmap.to_pil()
            import io
            buf = io.BytesIO()
            pil.save(buf, format="PNG")
            out.append(buf.getvalue())
    finally:
        pdf.close()
    if not out:
        raise RuntimeError("PDF 渲染结果为空")
    return out


def render_pages_from_docx(docx_path, dpi: int = 110,
                           workdir: str | None = None) -> list[bytes]:
    """docx → 逐页 PNG（工作台线程与测试的共用入口）。

    workdir：PDF 中转目录（默认系统临时目录；工作台传入会话级临时目录便于清理）。
    路径一律转绝对（Word COM 以自身工作目录解析相对路径）。
    """
    docx_path = str(Path(docx_path).resolve())
    tmp = None
    if workdir is None:
        tmp = tempfile.TemporaryDirectory(prefix="shushan_pdf_")
        workdir = tmp.name
    else:
        workdir = str(Path(workdir).resolve())
    pdf_path = Path(workdir) / (Path(docx_path).stem + ".pdf")
    try:
        export_pdf_via_word(docx_path, pdf_path)
        return render_pdf_pages(pdf_path, dpi)
    finally:
        if tmp is not None:
            tmp.cleanup()
