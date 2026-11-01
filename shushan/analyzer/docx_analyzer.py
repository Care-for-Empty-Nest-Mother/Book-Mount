# -*- coding: utf-8 -*-
"""文档级分析入口：产出段落特征列表 + 页面设置 + 图片尺寸统计。"""
from dataclasses import dataclass, field
from pathlib import Path

import docx

from docx.oxml.ns import qn

from .effective import (StylesDB, theme_fonts, effective_para_format,
                        format_signature, read_page_setup,
                        paragraph_has_object, object_kind, figure_extents_cm,
                        parse_citations, RE_LIT_NUM, parse_ppr, parse_rpr)


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
    header_footer: dict = field(default_factory=dict)     # 页眉页脚配置（M2 提取）
    table_cell_style: dict = field(default_factory=dict)  # 表格单元格文字样式多数派（M2）

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


# ---- M2：页眉页脚 与 表格单元格样式 提取 ----

def _field_token(instr: str) -> str | None:
    """域代码 → {page}/{pages} 记号（其余域不识别）。"""
    s = (instr or "").upper()
    if "NUMPAGES" in s:
        return "{pages}"
    if "PAGE" in s:
        return "{page}"
    return None


def _story_paragraph_template(p_el) -> str:
    """把页脚段落里的文字与 PAGE/NUMPAGES 域还原成模板串。

    复杂域（fldChar begin…separate…缓存结果…end）只取 instrText 记号、
    跳过缓存结果文字；fldSimple 直接按 instr 映射。"""
    parts = []
    state = {"in_result": False}

    def walk(el):
        for child in el:
            tag = child.tag
            if tag == qn("w:fldSimple"):
                tok = _field_token(child.get(qn("w:instr")) or "")
                if tok:
                    parts.append(tok)
            elif tag == qn("w:r"):
                fld = child.find(qn("w:fldChar"))
                if fld is not None:
                    ftype = fld.get(qn("w:fldCharType"))
                    if ftype == "separate":
                        state["in_result"] = True
                    elif ftype == "end":
                        state["in_result"] = False
                    continue
                instr = child.find(qn("w:instrText"))
                if instr is not None and (instr.text or "").strip():
                    tok = _field_token(instr.text)
                    if tok:
                        parts.append(tok)
                    continue
                if not state["in_result"]:
                    for t_el in child.findall(qn("w:t")):
                        parts.append(t_el.text or "")
            elif tag in (qn("w:hyperlink"), qn("w:smartTag")):
                walk(child)

    walk(p_el)
    return "".join(parts)


def _extract_header_footer(doc, sdb: StylesDB) -> dict:
    """读参考文档页眉页脚 → 方案 header_footer 键（§5.5）。

    内容取第一个有内容的节（封面节常为空，取正文节才有意义）；
    PAGE/NUMPAGES 域还原为 {page}/{pages} 模板；首页不同/奇偶不同读
    sectPr 与文档级标志。参考文档没有任何页眉页脚时返回 {}。
    """
    try:
        sections = list(doc.sections)
    except Exception:
        return {}
    if not sections:
        return {}
    header_cfg = None
    footer_cfg = None
    for sec in sections:
        if header_cfg is None:
            try:
                if not sec.header.is_linked_to_previous:
                    # 遍历全部后代段落：页眉内容可能位于 sdt 内容控件、
                    # 表格或文本框内（python-docx 的 .paragraphs 只看直接子级）
                    paras = [p for p in sec.header._element.iter(qn("w:p"))
                             if "".join(t.text or "" for t in p.findall(".//" + qn("w:t"))).strip()]
                    if paras:
                        p0 = paras[0]
                        align = parse_ppr(p0).get("align") or "center"
                        # 裸 w:p 元素：样式链解析 + 段落直接格式 + 首个非空 run
                        ppr0 = p0.find(qn("w:pPr"))
                        sid = None
                        if ppr0 is not None:
                            ps = ppr0.find(qn("w:pStyle"))
                            if ps is not None:
                                sid = ps.get(qn("w:val"))
                        fmt = {**sdb.resolve(sid), **parse_ppr(ppr0)}
                        for r in p0.findall(qn("w:r")):
                            txt = "".join(t.text or "" for t in r.findall(qn("w:t")))
                            if txt.strip():
                                fmt.update({k: v for k, v in
                                            parse_rpr(r.find(qn("w:rPr"))).items()
                                            if v is not None})
                                break
                        font = {k: fmt[k] for k in ("font_cn", "font_en", "size_pt")
                                if fmt.get(k)}
                        header_cfg = {
                            "enabled": True,
                            "content": "\n".join(
                                "".join(t.text or "" for t in p.findall(".//" + qn("w:t"))).strip()
                                for p in paras),
                            "align": align,
                        }
                        if font:
                            header_cfg["font"] = font
            except Exception:
                header_cfg = None
        if footer_cfg is None:
            try:
                if not sec.footer.is_linked_to_previous:
                    best = None
                    for p in sec.footer._element.iter(qn("w:p")):
                        tpl = _story_paragraph_template(p)
                        if not tpl.strip():
                            continue
                        has_field = "{page}" in tpl or "{pages}" in tpl
                        pos = parse_ppr(p).get("align") or "center"
                        if has_field:
                            best = (pos, tpl)
                            break
                        if best is None:
                            best = (pos, tpl)
                    if best:
                        footer_cfg = {"enabled": True, "page_number": {
                            "position": best[0], "format": best[1], "start_at": 1}}
            except Exception:
                footer_cfg = None
        if header_cfg is not None and footer_cfg is not None:
            break
    # 正文起编标志：参考文档在非首节写入 pgNumType start → 方案记为正文起编
    start_at = None
    start_idx = None
    for i, s in enumerate(sections):
        try:
            pg = s._sectPr.find(qn("w:pgNumType"))
        except Exception:
            pg = None
        if pg is not None and pg.get(qn("w:start")):
            try:
                start_at = int(pg.get(qn("w:start")))
                start_idx = i
            except (TypeError, ValueError):
                pass
            break
    body_starts = start_idx is not None and start_idx > 0
    if header_cfg is None and footer_cfg is None and not body_starts:
        return {}
    diff_first = False
    try:
        diff_first = any(s.different_first_page_header_footer for s in sections)
    except Exception:
        pass
    diff_oe = False
    try:
        diff_oe = bool(doc.settings.odd_and_even_pages_header_footer)
    except Exception:
        pass
    header_cfg = header_cfg or {"enabled": False}
    if diff_first:
        header_cfg["different_first"] = True
    if diff_oe:
        header_cfg["different_odd_even"] = True
    if footer_cfg is None:
        footer_cfg = {"enabled": False}
    elif body_starts:
        footer_cfg.setdefault("page_number", {})["start_at"] = start_at
    return {"header": header_cfg,
            "footer": footer_cfg,
            "body_starts_new_section": bool(body_starts)}


def _extract_table_cell_style(doc, sdb: StylesDB) -> dict:
    """表格单元格段落格式多数派 → tables.cell_style（§5.3/§5.5）。

    统计全部非空单元格段落的字体/字号/对齐，取严格多数派
    （>50%，平票不写），避免把噪声写成方案。"""
    from collections import Counter
    votes = {"font_cn": Counter(), "font_en": Counter(),
             "size_pt": Counter(), "align": Counter()}
    total = 0
    try:
        tables = doc.tables
    except Exception:
        return {}
    for tbl in tables:
        # 合并单元格去重：持有元素强引用按身份比较（id() 会因 lxml 代理
        # 地址复用产生误判，见 page_furniture.restyle_table_cells 注释）
        seen_els: list = []
        for row in tbl.rows:
            for cell in row.cells:
                tc_el = cell._tc
                if any(tc_el is e for e in seen_els):
                    continue
                seen_els.append(tc_el)
                for para in cell.paragraphs:
                    if not para.text.strip():
                        continue
                    try:
                        fmt = effective_para_format(para, sdb)
                    except Exception:
                        continue
                    for k in ("font_cn", "font_en"):
                        if fmt.get(k):
                            votes[k][fmt[k]] += 1
                    if fmt.get("size_pt"):
                        votes["size_pt"][fmt["size_pt"]] += 1
                    if fmt.get("align"):
                        votes["align"][fmt["align"]] += 1
                    total += 1
    if not total:
        return {}
    out = {}
    for k, cnt in votes.items():
        if not cnt:
            continue
        v, c = cnt.most_common(1)[0]
        if c * 2 > sum(cnt.values()):         # 严格多数派
            out["alignment" if k == "align" else k] = v
    return out


def analyze_document(path) -> AnalysisResult:
    path = Path(path)
    doc = docx.Document(str(path))
    sdb = StylesDB(doc.styles.element, theme_fonts(doc))
    result = AnalysisResult(path=str(path), page=read_page_setup(doc), source_path=str(path))

    from .cluster import suggest_type   # 延迟导入避免环

    # ---- 识别修正：多节文档的首节一般是封面，封面标题不设大纲级别 ----
    # 首节 = 第一个内嵌分节符之前的区域；若首节里已出现「第X章/1.1」这类
    # 章节编号，说明它不是纯封面（如封面与正文未分节的期刊稿），不启用该规则。
    sec0_end = None
    for i, para in enumerate(doc.paragraphs):
        if para._p.find(qn("w:pPr") + "/" + qn("w:sectPr")) is not None:
            sec0_end = i
            break
    first_section_is_cover = False
    if sec0_end is not None:
        from .cluster import RE_CHAPTER, RE_NUM_HEAD
        first_section_is_cover = not any(
            RE_CHAPTER.match(p.text.strip()) or RE_NUM_HEAD.match(p.text.strip())
            for p in doc.paragraphs[:sec0_end + 1] if p.text.strip())

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
        if (first_section_is_cover and i <= sec0_end
                and sug.startswith("heading_")):
            sug, conf, reason = "cover", 0.75, "首节视为封面：标题不设大纲级别"
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
    # M2：页眉页脚配置 与 表格单元格文字样式多数派
    try:
        result.header_footer = _extract_header_footer(doc, sdb)
    except Exception:
        result.header_footer = {}
    try:
        result.table_cell_style = _extract_table_cell_style(doc, sdb)
    except Exception:
        result.table_cell_style = {}

    # ---- 引用与文献的对应关系、完整性检查 ----
    seq = 0
    for p in result.nonempty():
        if p.suggestion == "reference":
            seq += 1
            m2 = RE_LIT_NUM.match(p.text)
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
