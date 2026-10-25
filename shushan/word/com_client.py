# -*- coding: utf-8 -*-
"""与正在运行的 Microsoft Word 实时对接（COM）。

三种能力：
1. 探测当前打开的 Word 与活动文档；
2. 把 .dotx 方案模板挂接到活动文档并刷新样式（附加模板 + UpdateStyles，
   必要时用 OrganizerCopy 逐个补齐样式）；
3. 在活动文档上按归类结果逐段套样式并清除直接格式。

说明：COM 段落索引与 python-docx 的 body 段落索引不一致（前者含表格内段落），
所以按"正文段落序号"对齐：遍历 COM 段落时跳过表格内段落，计数器与
python-docx doc.paragraphs 的顺序天然一致。
"""
from pathlib import Path

from ..types import BUILTIN_STYLE_MAP, DEFAULT_CN_NAMES, IGNORE_TYPE, style_name_of

WD_ORGANIZER_STYLES = 1     # wdOrganizerObjectStyles
WD_WITH_IN_TABLE = 12       # WdInformation.wdWithInTable


class WordNotRunning(Exception):
    pass


def get_word_app():
    """连接正在运行的 Word；未运行返回 None（绝不新起 Word 进程）。"""
    import win32com.client
    import pythoncom
    pythoncom.CoInitialize()
    try:
        return win32com.client.GetActiveObject("Word.Application")
    except Exception:
        return None


def active_document():
    """返回 (word, doc)；拿不到活动文档时抛 WordNotRunning。"""
    word = get_word_app()
    if word is None:
        raise WordNotRunning("未检测到正在运行的 Word")
    try:
        doc = word.ActiveDocument
        name = doc.Name
    except Exception as e:
        raise WordNotRunning(f"Word 已运行但没有打开的文档（{e}）")
    return word, doc


def document_file_path(doc) -> Path:
    """活动文档的磁盘路径；未保存过的文档没有路径。"""
    try:
        full = doc.FullName
    except Exception:
        full = ""
    if not full or str(full).startswith("文档") or "Document" in str(full) and ":" not in str(full):
        raise WordNotRunning("当前文档尚未保存到磁盘，请先保存（Ctrl+S）再使用书山")
    return Path(str(full))


def _try_style(doc, name):
    try:
        st = doc.Styles(name)
        return st
    except Exception:
        return None


def style_ref(doc, type_id: str):
    """解析某内容类型在 Word 里的样式对象：内置样式用语言无关常量，其余按名尝试。"""
    const = BUILTIN_STYLE_MAP.get(type_id, (None, None, None))[2]
    if const is not None:
        try:
            return doc.Styles(const)
        except Exception:
            pass
    candidates = []
    if type_id in BUILTIN_STYLE_MAP:
        canonical = BUILTIN_STYLE_MAP[type_id][1]
        candidates += [DEFAULT_CN_NAMES.get(type_id, canonical),
                       canonical.replace("heading", "Heading").title(),
                       canonical.title()]
    candidates.append(style_name_of(type_id))
    for name in candidates:
        st = _try_style(doc, name)
        if st is not None:
            return st
    return None


def inject_styles_via_com(word, doc, dotx_path: Path, style_names: list) -> dict:
    """附加模板 + UpdateStyles；对仍未就位的样式用 OrganizerCopy 逐个补。"""
    report = {"attached": False, "organizer_copied": [], "missing": []}
    try:
        doc.AttachedTemplate = str(Path(dotx_path).resolve())
        doc.UpdateStyles()
        report["attached"] = True
    except Exception as e:
        report["missing"].append(f"附加模板失败：{e}")
        return report
    for name in style_names:
        if _try_style(doc, name) is None:
            try:
                word.OrganizerCopy(Source=str(Path(dotx_path).resolve()),
                                   Destination=doc.FullName, Name=name,
                                   Object=WD_ORGANIZER_STYLES)
                report["organizer_copied"].append(name)
            except Exception:
                report["missing"].append(name)
    return report


WD_LINE_SPACE_EXACTLY = 4
WD_LINE_SPACE_SINGLE = 0


def resize_inline_shapes_com(doc, objects_cfg: dict) -> dict:
    """COM 路径的图片宽度统一（objects.figure.resize: true 时生效）。

    用 InlineShapes 集合直接改宽（LockAspectRatio 保持纵横比），
    目标宽度不超过版心宽。"""
    fig = (objects_cfg or {}).get("figure") or {}
    if not fig.get("resize"):
        return {"resized": 0}
    median = (fig.get("width_cm") or {}).get("median")
    if not median:
        return {"resized": 0}
    try:
        ps = doc.PageSetup
        content_pt = ps.PageWidth - ps.LeftMargin - ps.RightMargin
    except Exception:
        content_pt = None
    target_pt = min(float(median) * 28.35, content_pt) if content_pt else float(median) * 28.35
    resized = 0
    for i in range(1, doc.InlineShapes.Count + 1):
        try:
            shp = doc.InlineShapes.Item(i)
            if abs(shp.Width - target_pt) <= target_pt * 0.03:
                continue
            shp.LockAspectRatio = True
            shp.Width = target_pt
            resized += 1
        except Exception:
            continue
    return {"resized": resized}


def apply_assignments_com(word, doc, assignments: dict, profile: dict | None = None,
                          numbered_map: dict | None = None) -> dict:
    """assignments: {正文段落序号(0基): 内容类型}。返回报告。

    含嵌入式图片/公式的段落（InlineShapes>0）如果行距是"固定值"会被 Word
    裁剪掉图片，这里统一改为单倍行距（防裁剪保护）。"""
    counts, failed, protected = {}, [], 0
    outline_only = 0
    body_idx = -1
    keep_heading_numbering = ((profile or {}).get("numbering") or {}).get(
        "mode") == "keep_target"
    for i in range(1, doc.Paragraphs.Count + 1):
        para = doc.Paragraphs.Item(i)
        rng = para.Range
        try:
            in_table = rng.Information(WD_WITH_IN_TABLE)
        except Exception:
            in_table = False
        if in_table:
            continue
        body_idx += 1
        if body_idx in assignments and assignments[body_idx] != IGNORE_TYPE:
            tid = assignments[body_idx]
            is_heading = tid.startswith("heading_")
            wants_number = (not is_heading or numbered_map is None
                            or numbered_map.get(body_idx, True))
            if is_heading and not wants_number:
                # 未勾选：只设定大纲级别，样式与外观保持原样，去编号
                try:
                    para.OutlineLevel = int(tid[8:])
                    rng.ListFormat.RemoveNumbers()
                    outline_only += 1
                    counts[tid] = counts.get(tid, 0) + 1
                except Exception as e:
                    failed.append(f"段落{body_idx}: {e}")
            else:
                st = style_ref(doc, tid)
                if st is None:
                    failed.append(f"段落{body_idx}: 找不到样式（{tid}）")
                    continue
                try:
                    rng.Style = st
                    rng.Font.Reset()
                    # keep_target 模式且标题参与编号：保留目标编号（Reset 会清掉直接编号）
                    if keep_heading_numbering and is_heading:
                        pass
                    else:
                        rng.ParagraphFormat.Reset()
                    counts[tid] = counts.get(tid, 0) + 1
                except Exception as e:
                    failed.append(f"段落{body_idx}: {e}")
        try:
            if rng.InlineShapes.Count > 0 and para.LineSpacingRule == WD_LINE_SPACE_EXACTLY:
                para.LineSpacingRule = WD_LINE_SPACE_SINGLE
                protected += 1
        except Exception:
            pass
    rep = {"counts": counts, "failed": failed, "protected_objects": protected,
           "headings_outline_only": outline_only}
    try:
        rep.update(resize_inline_shapes_com(doc, (profile or {}).get("objects")))
    except Exception:
        pass
    return rep


def refresh_status() -> dict:
    """面板状态行用：返回 {running, doc_name, path, saved}。"""
    word = get_word_app()
    if word is None:
        return {"running": False, "doc_name": "", "path": "", "saved": True}
    try:
        doc = word.ActiveDocument
        return {"running": True, "doc_name": doc.Name,
                "path": str(doc.FullName or ""), "saved": bool(doc.Saved)}
    except Exception:
        return {"running": True, "doc_name": "", "path": "", "saved": True}
