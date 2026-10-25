# -*- coding: utf-8 -*-
"""文档级分析入口：产出段落特征列表 + 页面设置 + 图片尺寸统计。"""
from dataclasses import dataclass, field
from pathlib import Path

import docx

from .effective import (StylesDB, theme_fonts, effective_para_format,
                        format_signature, read_page_setup,
                        paragraph_has_object, object_kind, figure_extents_cm,
                        parse_citations, RE_CITATION)
from docx.oxml.ns import qn


@dataclass
class ParaInfo:
    index: int
    text: str
    fmt: dict | None
    signature: tuple
    suggestion: str = "body_text"
    confidence: float = 0.0
    reason: str = ""
    has_object: bool = False   # 含图片/OLE/公式，整稿时保留防裁剪保护
    object_kind: str | None = None   # 'figure' | 'formula'
    citations: tuple = ()            # 本段引用的文献编号（展开后）


@dataclass
class AnalysisResult:
    path: str
    paragraphs: list = field(default_factory=list)   # 全部段落（空段 fmt=None）
    page: dict = field(default_factory=dict)
    source_path: str = ""
    figure_sizes_cm: list = field(default_factory=list)   # [(宽cm, 高cm), ...]
    table_settings: dict = field(default_factory=dict)    # 跨页规范（多数派）
    citation_count: int = 0                               # 正文引用标记 [n] 处数
    ref_entries: dict = field(default_factory=dict)       # {段落号: 文献序号}
    citation_positions: dict = field(default_factory=dict)  # {文献序号: [引用段落号]}
    dangling_citations: list = field(default_factory=list)  # 悬空引用（无对应条目）
    uncited_entries: list = field(default_factory=list)     # 未被引用的文献序号

    def nonempty(self):
        """可归类段落：有文字，或承载图片/公式（对象段也有版式可提取）。"""
        return [p for p in self.paragraphs if p.fmt is not None]

    @property
    def object_indices(self):
        """含图片/OLE/公式的段落号（防裁剪保护对象）。"""
        return [p.index for p in self.paragraphs if p.has_object]


def _extract_table_settings(doc) -> dict:
    """从参考文档的表格多数派提取跨页规范：标题行重复、行不跨页拆分。"""
    tables = doc.tables
    if not tables:
        return {}
    header_n = sum(
        1 for t in tables if t.rows and t.rows[0]._tr.find(
            qn("w:trPr") + "/" + qn("w:tblHeader")) is not None)
    cant_n = 0
    total_rows = 0
    for t in tables:
        for row in t.rows:
            total_rows += 1
            if row._tr.find(qn("w:trPr") + "/" + qn("w:cantSplit")) is not None:
                cant_n += 1
    return {
        "repeat_header": bool(tables) and header_n * 2 > len(tables),
        "cant_split_rows": bool(total_rows) and cant_n * 2 > total_rows,
    }


def _count_citations(paragraphs) -> int:
    """统计正文段里的引用标记 [1]、[2-4]、[1,3] 等出现次数。"""
    import re as _re
    pat = _re.compile(r"\[\d+(?:\s*[-,，~]\s*\d+)*\]")
    return sum(len(pat.findall(p.text)) for p in paragraphs
               if p.fmt is not None and p.suggestion not in ("reference", "ignore"))


def analyze_document(path) -> AnalysisResult:
    path = Path(path)
    doc = docx.Document(str(path))
    sdb = StylesDB(doc.styles.element, theme_fonts(doc))
    result = AnalysisResult(path=str(path), page=read_page_setup(doc), source_path=str(path))

    from .cluster import suggest_type   # 延迟导入避免环
    for i, para in enumerate(doc.paragraphs):
        text = para.text
        has_obj = paragraph_has_object(para._p)
        kind = object_kind(para._p) if has_obj else None
        if not text.strip() and not has_obj:
            result.paragraphs.append(ParaInfo(index=i, text=text, fmt=None,
                                              signature=("",) * 19))
            continue
        # 对象段（即使无文字）也解析段落版式：对齐/间距/行距可提取
        fmt = effective_para_format(para, sdb)
        sig = format_signature(fmt)
        sug, conf, reason = suggest_type(text, fmt, result)
        if has_obj:
            if kind == "formula":
                sug, conf, reason = "formula_para", 0.9, "公式段（版式提取）"
            else:
                sug, conf, reason = "figure_para", 0.9, "图片段（版式提取）"
            if kind == "figure":
                result.figure_sizes_cm.extend(figure_extents_cm(para._p))
        cits = ()
        if sug not in ("reference", "ignore"):
            cits = tuple(sorted(parse_citations(text)))
        result.paragraphs.append(ParaInfo(index=i, text=text, fmt=fmt, signature=sig,
                                          suggestion=sug, confidence=conf, reason=reason,
                                          has_object=has_obj, object_kind=kind,
                                          citations=cits))
    result.table_settings = _extract_table_settings(doc)
    result.citation_count = _count_citations(result.nonempty())

    # ---- 引用与文献的对应关系、完整性检查 ----
    import re as _re
    seq = 0
    for p in result.nonempty():
        if p.suggestion == "reference":
            seq += 1
            m2 = _re.match(r'^\s*\[(\d{1,3})\]', p.text)
            result.ref_entries[p.index] = int(m2.group(1)) if m2 else seq
    for p in result.nonempty():
        for n in p.citations:
            result.citation_positions.setdefault(n, []).append(p.index)
    entry_nums = set(result.ref_entries.values())
    result.dangling_citations = sorted(n for n in result.citation_positions
                                       if n not in entry_nums)
    result.uncited_entries = sorted(n for n in entry_nums
                                    if n not in result.citation_positions)
    return result
