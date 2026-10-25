# -*- coding: utf-8 -*-
"""格式转换：.doc → .docx（借本机 Word 完成，格式零损失）。

原则：
- 原文件只读打开，绝不回写；
- 转换产物放在原文件旁（<原名>_书山转换.docx）或指定目录；
- Word 未运行时静默启动隐藏实例，用完即退，不影响已打开的文档。
"""
from pathlib import Path

WD_FORMAT_XML_DOCUMENT = 12   # wdFormatXMLDocument（.docx）

SUPPORTED_NOTE = "当前支持 .docx；.doc 会经本机 Word 自动转换；PDF 提取属于二期计划"


class ConvertError(Exception):
    pass


def ensure_docx(path, work_dir=None) -> tuple[Path, Path | None]:
    """保证拿到一个 .docx：.docx 原样返回；.doc 转换后返回新路径。

    返回 (可分析的 docx 路径, 转换产物路径或 None)。
    """
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(p)
    suffix = p.suffix.lower()
    if suffix == ".docx":
        return p, None
    if suffix == ".doc":
        out_dir = Path(work_dir) if work_dir else p.parent
        dest = out_dir / f"{p.stem}_书山转换.docx"
        convert_doc_to_docx(p, dest)
        return dest, dest
    raise ConvertError(f"暂不支持 {suffix} 格式。{SUPPORTED_NOTE}。"
                       f"建议先在 Word 中另存为 .docx 再用书山处理。")


def convert_doc_to_docx(src: Path, dest: Path):
    """用本机 Word 把 .doc 转存为 .docx。"""
    try:
        import win32com.client
        import pythoncom
    except ImportError as e:
        raise ConvertError(f"转换 .doc 需要 pywin32（{e}）")
    pythoncom.CoInitialize()
    word = None
    own_instance = False
    try:
        try:
            word = win32com.client.GetActiveObject("Word.Application")
        except Exception:
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            own_instance = True
        doc = word.Documents.Open(str(src.resolve()), ReadOnly=True)
        try:
            doc.SaveAs2(str(dest.resolve()), FileFormat=WD_FORMAT_XML_DOCUMENT)
        finally:
            doc.Close(SaveChanges=0)
    except Exception as e:
        raise ConvertError(f".doc 转换失败：{e}") from e
    finally:
        if own_instance and word is not None:
            word.Quit()
