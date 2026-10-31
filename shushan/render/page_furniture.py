# -*- coding: utf-8 -*-
"""页面附件（M2）：页眉页脚页码、目录域、表格单元格文字重排。

设计约定：
- 全部函数对输入做防御、对执行做 try/except——页眉页脚/目录属于"锦上添花"，
  任何失败只记警告，不阻断整稿主流程；
- 页码使用 Word 原生域（w:fldSimple 内嵌占位 run），由 settings.xml 的
  updateFields 让 Word 打开时自动刷新；
- 分节处理：封面/摘要等前置节的页眉页脚清空并断开继承，正文起编通过
  在正文节 sectPr 写入 w:pgNumType/@w:start 实现——不做全局一刀切。

明确不做（写入交付说明）：
- 题注 SEQ 域（会改动文字内容，推迟到 M3 新文档骨架）；
- 奇偶页差异化内容（v0.2 仅做 Word 开关，奇偶内容镜像相同）；
- 页眉页脚的 COM 实时路径（实时路径已整体下线）。
"""
import re

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from .styles_xml import JC_MAP

# {page} / {pages} 记号 → PAGE / NUMPAGES 域
RE_PAGE_TOKEN = re.compile(r"(\{page\}|\{pages\})")

# settings.xml 中位于 w:updateFields 之后的元素（CT_Settings 顺序，取第一个作插入锚点）
_SETTINGS_TAIL_TAGS = [
    "w:hdrShapeDefaults", "w:footnotePr", "w:endnotePr", "w:compat",
    "w:rsids", "w:themeFontLang", "w:clrSchemeMapping", "w:shapeDefaults",
    "w:decimalSymbol", "w:listSeparator",
]


# ---------------------------------------------------------------- 分节定位

def paragraph_section_map(doc) -> dict:
    """段落号 → 节号 映射。

    节以段内 sectPr（分节符）为界：带 sectPr 的段落是其所在节的最后一段，
    最后一个节由 body 级 sectPr 承接。表格内段落不参与分节，无需考虑。
    """
    sec_of = {}
    cur = 0
    try:
        for i, para in enumerate(doc.paragraphs):
            sec_of[i] = cur
            ppr = para._p.find(qn("w:pPr"))
            if ppr is not None and ppr.find(qn("w:sectPr")) is not None:
                cur += 1
    except Exception:
        return {}
    return sec_of


# ---------------------------------------------------------------- 基础写入

def _style_run(r, font: dict | None):
    """给一个 w:r 写入字体直接格式（rFonts / 字号），可安全重复调用。"""
    if not isinstance(font, dict) or not font:
        return
    rpr = r.find(qn("w:rPr"))
    if rpr is None:
        rpr = OxmlElement("w:rPr")
        r.insert(0, rpr)
    cn, en = font.get("font_cn"), font.get("font_en")
    if cn or en:
        for old in rpr.findall(qn("w:rFonts")):
            rpr.remove(old)
        rf = OxmlElement("w:rFonts")
        if en:
            rf.set(qn("w:ascii"), en)
            rf.set(qn("w:hAnsi"), en)
        if cn:
            rf.set(qn("w:eastAsia"), cn)
        rpr.insert(0, rf)
    size = font.get("size_pt")
    if size:
        for tag in ("w:sz", "w:szCs"):
            for old in rpr.findall(qn(tag)):
                rpr.remove(old)
            el = OxmlElement(tag)
            el.set(qn("w:val"), str(int(round(size * 2))))
            rpr.append(el)


def _add_text_run(p_el, text: str, font: dict | None = None):
    r = OxmlElement("w:r")
    _style_run(r, font)
    t = OxmlElement("w:t")
    if text != text.strip():
        t.set(qn("xml:space"), "preserve")
    t.text = text
    r.append(t)
    p_el.append(r)


def _add_simple_field(p_el, instr: str, placeholder: str, font: dict | None = None):
    """w:fldSimple 域：内嵌占位 run，Word 打开/更新域时刷新为实际值。"""
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), instr)
    r = OxmlElement("w:r")
    _style_run(r, font)
    t = OxmlElement("w:t")
    t.text = placeholder
    r.append(t)
    fld.append(r)
    p_el.append(fld)


def _add_fld_char(p_el, fld_type: str):
    r = OxmlElement("w:r")
    fld = OxmlElement("w:fldChar")
    fld.set(qn("w:fldCharType"), fld_type)
    r.append(fld)
    p_el.append(r)


def _set_paragraph_align(p_el, align: str):
    jc = p_el.get_or_add_pPr().get_or_add_jc()
    jc.set(qn("w:val"), JC_MAP.get(align, "left"))


def _clear_story(story):
    """清空页眉/页脚内容并断开与上一节的链接，返回可写入的空段落。"""
    try:
        story.is_linked_to_previous = False
    except Exception:
        pass
    el = story._element
    for child in list(el):
        if child.tag in (qn("w:p"), qn("w:tbl")):
            el.remove(child)
    el.append(OxmlElement("w:p"))
    return story.paragraphs[0]


def _set_pgnum_start(section, start: int):
    """在节 sectPr 写入 w:pgNumType/@w:start（页码重启）。"""
    sect_pr = section._sectPr
    for old in sect_pr.findall(qn("w:pgNumType")):
        sect_pr.remove(old)
    el = OxmlElement("w:pgNumType")
    el.set(qn("w:start"), str(int(start)))
    # CT_SectPr 顺序：pgNumType 须位于 cols/titlePg/docGrid 等元素之前
    after = {qn(f"w:{t}") for t in ("cols", "formProt", "vAlign", "noEndnote",
                                    "titlePg", "textDirection", "bidi",
                                    "rtlGutter", "docGrid", "printerSettings")}
    anchor = next((c for c in sect_pr if c.tag in after), None)
    if anchor is not None:
        anchor.addprevious(el)
    else:
        sect_pr.append(el)


# ---------------------------------------------------------------- 页眉页脚

def _apply_section_header(sec, header_cfg: dict, odd_even_mirror: bool) -> list:
    warns = []
    if not header_cfg.get("enabled"):
        return warns
    try:
        p = _clear_story(sec.header)
        content = str(header_cfg.get("content") or "")
        if content.strip():
            _add_text_run(p._p, content, header_cfg.get("font"))
        if header_cfg.get("align"):
            _set_paragraph_align(p._p, header_cfg["align"])
    except Exception as e:
        warns.append(f"页眉写入失败：{e}")
        return warns
    if header_cfg.get("different_odd_even") and odd_even_mirror:
        # v0.2 奇偶内容相同：把同样的页眉镜像到偶数页（Word 开关已打开）
        try:
            p_even = _clear_story(sec.even_page_header)
            if content.strip():
                _add_text_run(p_even._p, content, header_cfg.get("font"))
            if header_cfg.get("align"):
                _set_paragraph_align(p_even._p, header_cfg["align"])
        except Exception as e:
            warns.append(f"偶数页页眉镜像失败：{e}")
    return warns


def _apply_section_footer(sec, footer_cfg: dict, odd_even_mirror: bool) -> list:
    warns = []
    if not footer_cfg.get("enabled"):
        return warns
    pn = footer_cfg.get("page_number")
    if not isinstance(pn, dict) or not (pn.get("format") or pn.get("position")):
        return warns
    fmt = str(pn.get("format") or "{page}")
    font = pn.get("font") if isinstance(pn.get("font"), dict) else None
    try:
        p = _clear_story(sec.footer)
        for part in RE_PAGE_TOKEN.split(fmt):
            if part == "{page}":
                _add_simple_field(p._p, r" PAGE \* MERGEFORMAT ", "1", font)
            elif part == "{pages}":
                _add_simple_field(p._p, r" NUMPAGES \* MERGEFORMAT ", "1", font)
            elif part:
                _add_text_run(p._p, part, font)
        if pn.get("position"):
            _set_paragraph_align(p._p, pn["position"])
    except Exception as e:
        warns.append(f"页脚写入失败：{e}")
        return warns
    if pn.get("different_odd_even") and odd_even_mirror:
        try:
            p_even = _clear_story(sec.even_page_footer)
            for part in RE_PAGE_TOKEN.split(fmt):
                if part == "{page}":
                    _add_simple_field(p_even._p, r" PAGE \* MERGEFORMAT ", "1", font)
                elif part == "{pages}":
                    _add_simple_field(p_even._p, r" NUMPAGES \* MERGEFORMAT ", "1", font)
                elif part:
                    _add_text_run(p_even._p, part, font)
            if pn.get("position"):
                _set_paragraph_align(p_even._p, pn["position"])
        except Exception as e:
            warns.append(f"偶数页页脚镜像失败：{e}")
    return warns


def apply_headers_footers(doc, hf_cfg: dict, body_section_idx: int | None = None) -> dict:
    """按方案 header_footer 配置给整份文档写入页眉/页脚/页码（§5.1）。

    body_section_idx: 第一个正文级标题所在节（apply.py 用 assignment 定位传入）。
    body_starts_new_section=true 时：
      - 该节之前的节（封面/摘要）页眉页脚清空并断开链接；
      - 页码在该节用 w:pgNumType 重启（start_at）；
      - 该节起的各节写入页眉页脚内容。
    传 None（未找到正文标题）时全部节统一写入、不重启页码。
    返回报告 {"applied", "sections", "warnings"}。
    """
    report = {"applied": False, "sections": 0, "warnings": []}
    if not isinstance(hf_cfg, dict) or not hf_cfg:
        return report
    header_cfg = hf_cfg.get("header") if isinstance(hf_cfg.get("header"), dict) else {}
    footer_cfg = hf_cfg.get("footer") if isinstance(hf_cfg.get("footer"), dict) else {}
    body_first = bool(hf_cfg.get("body_starts_new_section"))
    odd_even = bool(header_cfg.get("different_odd_even"))
    if not (header_cfg.get("enabled") or footer_cfg.get("enabled") or body_first):
        return report

    if odd_even:
        try:
            doc.settings.odd_and_even_pages_header_footer = True
        except Exception as e:
            report["warnings"].append(f"奇偶页开关设置失败：{e}")

    pn = footer_cfg.get("page_number") if isinstance(footer_cfg.get("page_number"), dict) else {}
    start_at = pn.get("start_at") if isinstance(pn.get("start_at"), int) and not isinstance(
        pn.get("start_at"), bool) else 1

    try:
        sections = list(doc.sections)
    except Exception as e:
        report["warnings"].append(f"读取分节失败：{e}")
        return report
    for i, sec in enumerate(sections):
        try:
            front = body_first and body_section_idx is not None and i < body_section_idx
            if front:
                # 前置节（封面/摘要）：页眉页脚清空且断开继承，不编页码；
                # 首页不同/奇偶不同的残余变体（自带部件的）也一并清空，
                # 避免封面首页从旧部件里冒出页眉页脚
                _clear_story(sec.header)
                _clear_story(sec.footer)
                for variant in (sec.first_page_header, sec.first_page_footer,
                                sec.even_page_header, sec.even_page_footer):
                    try:
                        if not variant.is_linked_to_previous:
                            _clear_story(variant)
                    except Exception:
                        continue
                continue
            report["warnings"] += _apply_section_header(sec, header_cfg, odd_even)
            report["warnings"] += _apply_section_footer(sec, footer_cfg, odd_even)
            if body_first and body_section_idx is not None and i == body_section_idx:
                _set_pgnum_start(sec, start_at)
            # 首页不同：只作用于页眉页脚区的第一节（正文首节，找不到正文节
            # 时为文档首节）。多节论文若每节都设 titlePg，每章首页都会丢
            # 页眉页脚页码——「首页不同」语义上是全文首页，不是每章首页。
            first_furnished = body_section_idx if (body_first and
                                                   body_section_idx is not None) else 0
            if (header_cfg.get("different_first") or footer_cfg.get("different_first")) \
                    and i == first_furnished:
                sec.different_first_page_header_footer = True
                _clear_story(sec.first_page_header)
                _clear_story(sec.first_page_footer)
            report["sections"] += 1
        except Exception as e:
            report["warnings"].append(f"第{i + 1}节页眉页脚应用失败：{e}")
    report["applied"] = report["sections"] > 0 or bool(report["warnings"])
    return report


# ---------------------------------------------------------------- 目录域

def document_has_toc(doc) -> bool:
    """文档正文是否已含 TOC 域（instrText 或 fldSimple 形式）。

    以 TOC 起头才算目录域——PAGEREF _Toc… 这类书签引用（旧目录删除后
    残留的交叉引用）不能误判成"已有目录"。"""
    try:
        body = doc.element.body
    except Exception:
        return False
    try:
        for it in body.iter(qn("w:instrText")):
            if (it.text or "").strip().upper().startswith("TOC"):
                return True
        for fld in body.iter(qn("w:fldSimple")):
            if (fld.get(qn("w:instr")) or "").strip().upper().startswith("TOC"):
                return True
    except Exception:
        return False
    return False


def insert_toc_field(doc, toc_cfg: dict, anchor_para_idx: int | None) -> dict:
    """在第一个标题段之前插入 TOC 域段落（已有目录或不满足条件时不插）。"""
    report = {"inserted": False, "reason": ""}
    if not isinstance(toc_cfg, dict) or not toc_cfg.get("insert"):
        report["reason"] = "方案未开启目录插入"
        return report
    if document_has_toc(doc):
        report["reason"] = "文档已有目录域"
        return report
    if anchor_para_idx is None:
        report["reason"] = "未找到标题段，无法定位目录插入点"
        return report
    try:
        paragraphs = doc.paragraphs
        if anchor_para_idx < 0 or anchor_para_idx >= len(paragraphs):
            report["reason"] = "目录插入点越界"
            return report
        anchor = paragraphs[anchor_para_idx]._p
    except Exception as e:
        report["reason"] = f"定位目录插入点失败：{e}"
        return report
    levels = toc_cfg.get("levels", 3)
    if not isinstance(levels, int) or isinstance(levels, bool) or not (1 <= levels <= 3):
        levels = 3
    try:
        p = OxmlElement("w:p")
        _add_fld_char(p, "begin")
        r_instr = OxmlElement("w:r")
        instr = OxmlElement("w:instrText")
        instr.set(qn("xml:space"), "preserve")
        instr.text = f' TOC \\o "1-{levels}" \\h \\z \\u '
        r_instr.append(instr)
        p.append(r_instr)
        _add_fld_char(p, "separate")
        _add_text_run(p, "（目录：右键→更新域）")
        _add_fld_char(p, "end")
        anchor.addprevious(p)
        report["inserted"] = True
    except Exception as e:
        report["reason"] = f"目录域插入失败：{e}"
    return report


def set_update_fields(doc) -> bool:
    """settings.xml 写入 <w:updateFields w:val="true"/>，Word 打开时自动刷新域。"""
    try:
        settings = doc.settings.element
        existing = settings.find(qn("w:updateFields"))
        if existing is not None:
            existing.set(qn("w:val"), "true")
            return True
        el = OxmlElement("w:updateFields")
        el.set(qn("w:val"), "true")
        anchor = None
        for tag in _SETTINGS_TAIL_TAGS:
            anchor = settings.find(qn(tag))
            if anchor is not None:
                break
        if anchor is not None:
            anchor.addprevious(el)
        else:
            settings.append(el)
        return True
    except Exception:
        return False


def apply_toc_updates(doc, toc_cfg: dict, anchor_para_idx: int | None = None) -> dict:
    """目录域处理入口：按需插入 TOC 域 + 无论插入与否都写 updateFields。

    updateFields 无条件写入：整稿后页码/引用域需要 Word 打开时自动刷新。
    """
    report = {"inserted": False, "update_fields": False, "reason": ""}
    try:
        ins = insert_toc_field(doc, toc_cfg or {}, anchor_para_idx)
        report["inserted"] = ins.get("inserted", False)
        report["reason"] = ins.get("reason", "")
    except Exception as e:
        report["reason"] = f"目录域处理失败：{e}"
    try:
        report["update_fields"] = set_update_fields(doc)
    except Exception:
        pass
    return report


# ---------------------------------------------------------------- 表格重排

def _apply_cell_format(paragraph, spec: dict):
    """把 cell_style 以直接格式写入单元格段落（字体/字号/加粗/对齐）。

    只动直接子级 run（不碰单元格内图形/文本框），不改文字内容，
    保留 vertAlign（上标等语义格式）。"""
    align = spec.get("alignment")
    if align:
        _set_paragraph_align(paragraph._p, align)
    for r in paragraph._p.findall(qn("w:r")):
        rpr = r.find(qn("w:rPr"))
        if rpr is None:
            rpr = OxmlElement("w:rPr")
            r.insert(0, rpr)
        cn, en = spec.get("font_cn"), spec.get("font_en")
        if cn or en:
            for old in rpr.findall(qn("w:rFonts")):
                rpr.remove(old)
            rf = OxmlElement("w:rFonts")
            if en:
                rf.set(qn("w:ascii"), en)
                rf.set(qn("w:hAnsi"), en)
            if cn:
                rf.set(qn("w:eastAsia"), cn)
            rpr.insert(0, rf)
        size = spec.get("size_pt")
        if size:
            for tag in ("w:sz", "w:szCs"):
                for old in rpr.findall(qn(tag)):
                    rpr.remove(old)
                el = OxmlElement(tag)
                el.set(qn("w:val"), str(int(round(size * 2))))
                rpr.append(el)
        bold = spec.get("bold")
        if bold is not None:
            for old in rpr.findall(qn("w:b")):
                rpr.remove(old)
            b = OxmlElement("w:b")
            if not bold:
                b.set(qn("w:val"), "0")
            rpr.append(b)


def restyle_table_cells(doc, cell_style: dict | None) -> int:
    """表格内部文字重排（§5.3）：全部表格全部单元格按 cell_style 统一格式。

    不改单元格文字内容，不动表格结构（列宽/合并/跨页设置保持）。
    返回处理的单元格段落数。cell_style 为空或非 dict 时不动（返回 0）。
    """
    if not isinstance(cell_style, dict) or not cell_style:
        return 0
    n = 0
    try:
        tables = doc.tables
    except Exception:
        return 0
    for tbl in tables:
        # 合并单元格去重：必须持有元素强引用后按身份比较。
        # 不能用 id() 去重——lxml 代理对象回收后地址复用，会把未处理的
        # 单元格误判为已处理而整格跳过。
        seen_els: list = []
        try:
            rows = tbl.rows
        except Exception:
            continue
        for row in rows:
            for cell in row.cells:
                tc_el = cell._tc
                if any(tc_el is e for e in seen_els):
                    continue
                seen_els.append(tc_el)
                for para in cell.paragraphs:
                    try:
                        _apply_cell_format(para, cell_style)
                        n += 1
                    except Exception:
                        continue
    return n
