# -*- coding: utf-8 -*-
"""离线整稿：把方案注入 docx 文件并按归类结果重排段落格式。

安全边界：
- 写入前自动备份原文件；
- 只动段落/字符格式属性与样式引用，不增删文字内容；
- 段落里的分节符(sectPr)、域代码、书签一律保留；
- 检测到标题带手工编号（第X章/1.1）时不挂自动编号，避免"第一章 第一章"重复。
"""
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from ..analyzer.docx_analyzer import analyze_document
from ..analyzer.cluster import RE_CHAPTER, RE_NUM_HEAD
from ..analyzer.effective import StylesDB, parse_ppr, RE_CITATION
from ..backup import backup_file
from ..profile_builder import sig_str
from ..types import IGNORE_TYPE, style_name_of
from .dotx_renderer import render_styles_into, apply_page_setup
from .page_furniture import (apply_headers_footers, apply_toc_updates,
                             restyle_table_cells, paragraph_section_map)
from .styles_xml import _spacing_xml, _indent_xml, _set_font

# 段落直接格式里允许被清除的属性（其余如 sectPr 保留）。
# w:pageBreakBefore 不在清除列表：段前分页是版面结构（决定分页位置），
# 清掉会改变原文分页布局（封面/摘要页空白、章节起始页漂移）；方案 styles
# 的 page_break_before 字段经样式生效，不受影响。
CLEAR_PPR_TAGS = ["w:rPr", "w:numPr", "w:outlineLvl", "w:jc", "w:spacing", "w:ind",
                  "w:keepNext", "w:keepLines", "w:textAlignment",
                  "w:pBdr", "w:shd", "w:tabs"]


def resolve_assignment(analysis, profile, para_override: dict | None = None) -> dict:
    """决定每个段落的类型，优先级：用户改判（逐段）> 对象段按内容判定 >
    方案 mapping（签名多数派，仅文字段）> 自动建议。

    para_override: {段落号: 类型}——批注界面确认后的结果，可含 "ignore"。
    返回 {段落号: 类型}。
    """
    mapping = profile.get("mapping") or {}
    assign = {}
    for p in analysis.nonempty():
        if para_override and p.index in para_override:
            assign[p.index] = para_override[p.index]
            continue
        if p.has_object:
            # 图片/公式段的类型由内容决定（图片段/公式段），不依赖签名映射
            assign[p.index] = p.suggestion
            continue
        t = mapping.get(sig_str(p.signature))
        if t and t.startswith("heading_") and p.suggestion == "cover":
            # 封面标题不设大纲级别（识别修正优先于旧映射，防止旧方案把
            # 封面大标题带进编号体系）
            t = None
        assign[p.index] = t if t else p.suggestion
    return assign


def has_manual_numbering(analysis, assignment: dict) -> bool:
    for p in analysis.nonempty():
        if assignment.get(p.index, "").startswith("heading_"):
            if RE_CHAPTER.match(p.text.strip()) or RE_NUM_HEAD.match(p.text.strip()):
                return True
    return False


def reattach_numbering(paragraph, num_id: str, num_ilvl: int = 0):
    """把原有的自动编号结构重新挂回段落（整稿不应打掉参考文献/列表的编号）。"""
    numpr = paragraph._p.get_or_add_pPr().get_or_add_numPr()
    ilvl_el = numpr.get_or_add_ilvl()
    ilvl_el.set(qn("w:val"), str(int(num_ilvl)))
    nid_el = numpr.get_or_add_numId()
    nid_el.set(qn("w:val"), str(num_id))


def suppress_numbering(paragraph):
    """抑制段落编号（numId=0）：样式里继承来的编号也被压掉，
    该段不编号、不占号，后续章节编号照常接续。"""
    numpr = paragraph._p.get_or_add_pPr().get_or_add_numPr()
    numpr.get_or_add_ilvl().set(qn("w:val"), "0")
    numpr.get_or_add_numId().set(qn("w:val"), "0")


def set_outline_level(paragraph, level0: int):
    """直写段落大纲级别（0 基）。只改大纲属性，其余格式一概不动——
    用于"仅设定大纲级别"的标题：导航窗格与目录可见，外观保持原样。"""
    ppr = paragraph._p.get_or_add_pPr()
    ol = ppr.find(qn("w:outlineLvl"))
    if ol is None:
        ol = OxmlElement("w:outlineLvl")
        rpr = ppr.find(qn("w:rPr"))
        if rpr is not None:
            rpr.addprevious(ol)
        else:
            ppr.append(ol)
    ol.set(qn("w:val"), str(int(level0)))


def strip_direct_formatting(paragraph, keep_numpr: bool = False):
    ppr = paragraph._p.find(qn("w:pPr"))
    if ppr is not None:
        for tag in CLEAR_PPR_TAGS:
            if keep_numpr and tag == "w:numPr":
                continue          # 保留目标标题自己的编号（keep_target 模式）
            for el in ppr.findall(qn(tag)):
                ppr.remove(el)
    # 只清直接子级 run 的字符格式：图形内部的文本框 run 不能动，
    # 否则文本框里的文字格式会被抹掉（超级链接里的 run 也一并保留）。
    # 上标/下标（w:vertAlign，如正文引用标记 [1]）属于内容语义，一并保留，
    # 避免"引用和正文彼此覆盖"。
    for r in paragraph._p.findall(qn("w:r")):
        rpr = r.find(qn("w:rPr"))
        if rpr is None:
            continue
        va = rpr.find(qn("w:vertAlign"))
        r.remove(rpr)
        if va is not None:
            new_rpr = OxmlElement("w:rPr")
            new_rpr.append(va)
            r.insert(0, new_rpr)


WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
EMU_PER_CM = 360000


def resize_figures_to_reference(doc, analysis, objects_cfg, content_cm: float) -> int:
    """按方案开关统一图片宽度（objects.figure.resize: true 时生效）。

    把内嵌图片（DrawingML）宽度等比缩放到参考中位数，且不超过版心宽；
    与参考中位数相差 ≤3% 的视为已一致，不动。浮动图与 VML 旧图不处理。
    """
    fig = (objects_cfg or {}).get("figure") or {}
    if not fig.get("resize"):
        return 0
    median = (fig.get("width_cm") or {}).get("median")
    if not median:
        return 0
    target_cm = min(float(median), content_cm)
    target_emu = int(target_cm * EMU_PER_CM)

    resized = 0
    for p in analysis.paragraphs:
        if not p.has_object:
            continue
        p_el = doc.paragraphs[p.index]._p
        for drawing in p_el.findall(".//" + qn("w:drawing")):
            inline = drawing.find(f"{{{WP_NS}}}inline")
            if inline is None:
                continue                     # 浮动图不动
            ext = inline.find(f"{{{WP_NS}}}extent")
            if ext is None:
                continue
            try:
                cx = int(ext.get("cx"))
            except (TypeError, ValueError):
                continue
            if abs(cx - target_emu) <= target_emu * 0.03:
                continue                     # 已与参考一致
            factor = target_emu / cx
            ext.set("cx", str(target_emu))
            ext.set("cy", str(int(int(ext.get("cy", "0")) * factor)))
            for a_ext in drawing.findall(f".//{{{A_NS}}}ext"):
                try:
                    a_ext.set("cx", str(int(int(a_ext.get("cx", "0")) * factor)))
                    a_ext.set("cy", str(int(int(a_ext.get("cy", "0")) * factor)))
                except (TypeError, ValueError):
                    pass
            resized += 1
    return resized


def fit_oversized_figures(doc, analysis, page: dict) -> dict:
    """安全缩放：宽度超过版心的嵌入式图片等比缩到版心宽以内。

    只动 DrawingML 内联图片（wp:extent + a:xfrm/a:ext 同步缩放，保持纵横比）；
    浮动图片（wp:anchor）与 VML 旧图不动，仅计数上报——盲目缩放浮动图会影响
    版面锚定，交给用户在 Word 里处理更安全。
    """
    if not isinstance(page, dict):
        return {"resized": 0, "skipped_float": 0}
    margins = page.get("margins", {}) or {}
    paper_w = {"A4": 21.0, "A3": 29.7, "B5": 17.6, "Letter": 21.59}.get(
        page.get("paper", "A4"), 21.0)
    if page.get("orientation") == "landscape":
        paper_w = {"A4": 29.7, "A3": 42.0, "B5": 25.0, "Letter": 27.94}.get(
            page.get("paper", "A4"), 29.7)
    content_cm = paper_w - margins.get("left_cm", 3.17) - margins.get("right_cm", 3.17)
    limit_emu = int(content_cm * EMU_PER_CM)
    if limit_emu <= 0:
        return {"resized": 0, "skipped_float": 0}

    resized = skipped = 0
    for p in analysis.paragraphs:
        if not p.has_object:
            continue
        p_el = doc.paragraphs[p.index]._p
        for drawing in p_el.findall(".//" + qn("w:drawing")):
            inline = drawing.find(f"{{{WP_NS}}}inline")
            anchor = drawing.find(f"{{{WP_NS}}}anchor")
            container = inline if inline is not None else anchor
            if container is None:
                continue
            ext = container.find(f"{{{WP_NS}}}extent")   # extent 在 inline/anchor 之下
            if ext is None:
                continue
            try:
                cx = int(ext.get("cx"))
            except (TypeError, ValueError):
                continue
            if cx <= limit_emu:
                continue
            if inline is None:      # 浮动图：不动，只上报
                skipped += 1
                continue
            factor = limit_emu / cx
            ext.set("cx", str(limit_emu))
            ext.set("cy", str(int(int(ext.get("cy", "0")) * factor)))
            # 同步图形变换尺寸，避免显示尺寸与内部变换不一致
            for a_ext in drawing.findall(f".//{{{A_NS}}}ext"):
                try:
                    a_ext.set("cx", str(int(int(a_ext.get("cx", "0")) * factor)))
                    a_ext.set("cy", str(int(int(a_ext.get("cy", "0")) * factor)))
                except (TypeError, ValueError):
                    pass
            resized += 1
    return {"resized": resized, "skipped_float": skipped}


def protect_object_paragraphs(doc, analysis) -> int:
    """图形/公式段防裁剪保护。

    嵌入式图片/公式的高度超过"固定值行距"时会被 Word 裁剪到一行行高，
    视觉上等同于图片消失。即使该段落本身未被重排，其引用的样式（如 Normal）
    也可能被方案重定义成固定值行距——因此这里对每个含对象段落解析生效行距，
    凡是 exact 一律覆盖为单倍行距（直接格式，样式不动）。
    """
    sdb = StylesDB(doc.styles.element, {})
    protected = 0
    for p in analysis.paragraphs:
        if not p.has_object:
            continue
        p_new = doc.paragraphs[p.index]
        ppr_el = p_new._p.find(qn("w:pPr"))
        sid = None
        if ppr_el is not None:
            ps = ppr_el.find(qn("w:pStyle"))
            if ps is not None:
                sid = ps.get(qn("w:val"))
        if sid is None and p_new.style is not None:
            sid = p_new.style.style_id
        fmt = sdb.resolve(sid)
        if ppr_el is not None:
            fmt = {**fmt, **parse_ppr(ppr_el)}
        if fmt.get("line_mode") == "exact":
            spacing = p_new._p.get_or_add_pPr().get_or_add_spacing()
            spacing.set(qn("w:line"), "240")
            spacing.set(qn("w:lineRule"), "auto")
            protected += 1
    return protected


def normalize_citations(doc, analysis, assignment, enabled: bool) -> int:
    """标准引用格式（GB/T 7714 顺序编码制）：把正文中独立成 run 的 [n] 标记
    统一为上标。只动格式不改文字；混合在长句 run 里的标记不处理（需拆 run）。
    """
    if not enabled:
        return 0
    n = 0
    for p in analysis.nonempty():
        if assignment.get(p.index) not in ("body_text", "figure_para", "formula_para"):
            continue
        for r in doc.paragraphs[p.index]._p.findall(qn("w:r")):
            text = "".join(t.text or "" for t in r.findall(qn("w:t")))
            if not text or not RE_CITATION.fullmatch(text.strip()):
                continue
            rpr = r.find(qn("w:rPr"))
            if rpr is None:
                rpr = OxmlElement("w:rPr")
                r.insert(0, rpr)
            va = rpr.find(qn("w:vertAlign"))
            if va is None:
                va = OxmlElement("w:vertAlign")
                rpr.append(va)
            va.set(qn("w:val"), "superscript")
            n += 1
    return n


def normalize_tables(doc, tables_cfg: dict) -> dict:
    """跨页表格规范：标题行跨页重复（tblHeader）+ 行不跨页拆分（cantSplit）。

    配置来自方案 tables 字段（默认都开启）；单个 trPr 内按
    cantSplit → tblHeader 的 schema 顺序插入。
    """
    repeat = bool((tables_cfg or {}).get("repeat_header", True))
    cant = bool((tables_cfg or {}).get("cant_split_rows", True))
    n_tables = n_header = n_cant = 0
    for tbl in doc.tables:
        n_tables += 1
        for ri, row in enumerate(tbl.rows):
            trpr = row._tr.get_or_add_trPr()
            if cant and trpr.find(qn("w:cantSplit")) is None:
                trpr.insert(0, OxmlElement("w:cantSplit"))
                n_cant += 1
            if repeat and ri == 0 and trpr.find(qn("w:tblHeader")) is None:
                el = OxmlElement("w:tblHeader")
                cs = trpr.find(qn("w:cantSplit"))
                if cs is not None:
                    cs.addnext(el)
                else:
                    trpr.insert(0, el)
                n_header += 1
    return {"tables": n_tables, "header_repeat": n_header, "rows_cant_split": n_cant}


def detect_numbering_mode(analysis, assignment) -> str:
    """智能判断标题编号来源：
    - 标题段自带自动编号（直接 numPr）→ keep_target（保留目标编号体系）
    - 标题文字里已有手工编号（第X章 / 1.1）→ manual（文字即编号，不另挂）
    - 都没有 → reference（套用参考方案的编号体系）
    """
    for p in analysis.nonempty():
        t = assignment.get(p.index) or p.suggestion
        if t.startswith("heading_") and p.fmt and p.fmt.get("num_id"):
            return "keep_target"
    if has_manual_numbering(analysis, assignment):
        return "manual"
    return "reference"


MODE_LABELS = {
    "auto": "智能判断", "keep_target": "保留目标文档编号",
    "reference": "套用参考方案编号", "manual": "保留手工编号（文字即编号）",
    "custom": "自定义编号样式",
}

# 前置区类型（摘要/关键词等）：正文起编与目录插入点要跳过它们之前的标题段
FRONT_MATTER_TYPES = ("abstract", "en_abstract", "keywords", "en_keywords",
                      "cover", "statement")


def find_body_heading_index(analysis, assignment: dict) -> int | None:
    """定位"正文级标题"段落（M2：页眉页脚正文起编 与 目录域插入点共用）。

    封面大标题常被自动建议为 heading_1，直接取第一个 heading 段会把封面
    当成正文（页眉写入封面节、目录插到封面标题前）。优先取前置区
    （摘要/关键词段与目录条目行）之后的第一个标题段；没有前置区标记时
    回落为第一个标题段。无标题段返回 None。
    """
    headings = [p.index for p in analysis.nonempty()
                if str(assignment.get(p.index, "")).startswith("heading_")]
    if not headings:
        return None
    last_front = None
    for p in analysis.nonempty():
        t = assignment.get(p.index)
        if t in FRONT_MATTER_TYPES or (t == IGNORE_TYPE and "目录条目" in (p.reason or "")):
            last_front = p.index
    if last_front is not None:
        after = [h for h in headings if h > last_front]
        if after:
            return after[0]
    return headings[0]


def apply_format_override(paragraph, spec: dict):
    """把手动修改的格式以直接格式写到段落上（优先于样式）。

    只写 spec 中出现的字段；字体/字号/加粗/斜体作用于直接子级 run
    （不碰图形内部文本框）。"""
    ppr = paragraph._p.get_or_add_pPr()
    spacing_attrs = _spacing_xml(spec)
    if spacing_attrs:
        sp = ppr.get_or_add_spacing()
        for attr, val in spacing_attrs:
            sp.set(qn(attr), str(val))
    indent_attrs = _indent_xml(spec)
    if indent_attrs:
        ind = ppr.get_or_add_ind()
        for attr, val in indent_attrs:
            ind.set(qn(attr), str(val))
    align = spec.get("alignment")
    if align:
        from .styles_xml import JC_MAP
        jc = ppr.get_or_add_jc()
        jc.set(qn("w:val"), JC_MAP.get(align, "left"))
    rpr_targets = paragraph._p.findall(qn("w:r"))
    for r in rpr_targets:
        rpr = r.find(qn("w:rPr"))
        if rpr is None:
            rpr = OxmlElement("w:rPr")
            r.insert(0, rpr)
        _set_font(rpr, spec) if (spec.get("font_cn") or spec.get("font_en")) else None
        size = spec.get("size_pt")
        if size:
            for tag in ("w:sz", "w:szCs"):
                for old in rpr.findall(qn(tag)):
                    rpr.remove(old)
                el = OxmlElement(tag)
                el.set(qn("w:val"), str(int(round(size * 2))))
                rpr.append(el)
        for tag, val in (("w:b", spec.get("bold")), ("w:i", spec.get("italic"))):
            if val is None:
                continue
            for old in rpr.findall(qn(tag)):
                rpr.remove(old)
            el = OxmlElement(tag)
            if not val:
                el.set(qn("w:val"), "0")
            rpr.append(el)


def apply_profile_offline(profile: dict, docx_path, para_override: dict | None = None,
                          make_backup: bool = True,
                          numbering_override: tuple | None = None,
                          numbered_map: dict | None = None,
                          fmt_overrides: dict | None = None) -> dict:
    """对磁盘上的 docx 执行整稿。

    numbering_override: (mode, scheme_levels)——界面选择的标题编号模式；
    mode ∈ auto/keep_target/reference/manual/custom，custom 时 scheme_levels
    为自定义编号定义（reference 语义）。缺省 None 时按方案与既有规则处理。
    返回报告 dict。"""
    docx_path = Path(docx_path)
    if not docx_path.exists() or docx_path.suffix.lower() != ".docx":
        raise FileNotFoundError(docx_path)

    analysis = analyze_document(docx_path)
    assignment = resolve_assignment(analysis, profile, para_override)

    # ---- 解析标题编号模式 ----
    numbering_cfg = profile.get("numbering") or {}
    scheme_levels = (numbering_cfg.get("heading_scheme") or {}).get("levels") or []
    if numbering_override is not None:
        mode, override_scheme = numbering_override
        if mode == "auto":
            mode = detect_numbering_mode(analysis, assignment)
        elif mode == "custom":
            mode = "reference"
            scheme_levels = override_scheme or scheme_levels
    else:
        mode = numbering_cfg.get("mode", "reference")
        if mode == "auto":
            mode = detect_numbering_mode(analysis, assignment)
        if mode == "manual" or (mode == "reference" and has_manual_numbering(analysis, assignment)
                                and not numbering_override):
            mode = "manual"

    backup_path = backup_file(docx_path) if make_backup else None

    doc = Document(str(docx_path))
    # 按解析出的模式构造注入配置
    profile_for_inject = dict(profile)
    if mode == "keep_target":
        profile_for_inject["numbering"] = {"mode": "keep_target"}
    elif mode == "manual":
        profile_for_inject["numbering"] = {}
    elif mode == "reference":
        profile_for_inject["numbering"] = {"mode": "reference",
                                           "heading_scheme": {"levels": scheme_levels}}
    inject_report = render_styles_into(doc, profile_for_inject,
                                       remove_other_paragraph_styles=False)
    page = profile.get("page")
    if page:
        apply_page_setup(doc, page)

    styles = doc.styles
    counts, failed, skipped = {}, [], 0
    renumbered = 0
    unnumbered = 0
    outline_only = 0
    suppressed_num = 0
    for p in analysis.nonempty():
        tid = assignment.get(p.index, "body_text")
        if tid == IGNORE_TYPE:
            skipped += 1                     # 归类为忽略的段落保持原样
            continue
        sname = style_name_of(tid)
        try:
            is_heading = tid.startswith("heading_")
            # 标题对号（缺省=勾选）：勾选=完整标题处理（套样式+按模式编号）；
            # 不勾=仅设定大纲级别——不套标题样式、不参与编号，其余原样保留
            wants_full = (not is_heading or numbered_map is None
                          or numbered_map.get(p.index, True))
            if is_heading and not wants_full:
                p_new = doc.paragraphs[p.index]      # 不换样式、不清格式
                suppress_numbering(p_new)            # 编号抑制（不编号不占号）
                set_outline_level(p_new, int(tid[8:]) - 1)
                if fmt_overrides and p.index in fmt_overrides:
                    apply_format_override(p_new, fmt_overrides[p.index])
                outline_only += 1
                counts[tid] = counts.get(tid, 0) + 1
                continue
            p_new = doc.paragraphs[p.index]
            p_new.style = styles[sname]
            keep_num = mode == "keep_target" and is_heading
            strip_direct_formatting(p_new, keep_numpr=keep_num)
            # 编号归属：
            # - keep_target 模式：标题段自己的编号原样保留（只改级别与视觉样式）
            # - reference 模式：标题段交给参考方案的编号体系（直接 numPr 清除）；
            #   非标题段的原有自动编号是内容结构（参考文献 [1][2]、项目列表），必须保留
            old_num = p.fmt.get("num_id") if p.fmt else None
            if old_num and not is_heading:
                if tid == "cover":
                    # 封面段落不参与任何编号（旧稿残留的列表编号一并压掉）
                    suppress_numbering(p_new)
                    suppressed_num += 1
                else:
                    reattach_numbering(p_new, old_num, p.fmt.get("num_ilvl", 0))
                    renumbered += 1
            if fmt_overrides and p.index in fmt_overrides:
                apply_format_override(p_new, fmt_overrides[p.index])
            counts[tid] = counts.get(tid, 0) + 1
        except KeyError:
            failed.append(f"段落{p.index}: 样式 {sname} 不存在")
        except Exception as e:                     # 单段失败不拖垮整稿
            failed.append(f"段落{p.index}: {e}")
    # ---- M2：页眉页脚 / 目录域 / 表格单元格文字重排 ----
    # 顺序：段落重排（上方循环）→ 页眉页脚 → 表格重排 →（收尾）目录域；
    # 三者全部 try/except 包裹，失败记警告不阻断整稿。
    sec_map = {}
    try:
        sec_map = paragraph_section_map(doc)
    except Exception:
        sec_map = {}
    body_heading_idx = find_body_heading_index(analysis, assignment)
    hf_cfg = profile.get("header_footer") if isinstance(profile.get("header_footer"), dict) else {}
    hf_report = {"applied": False, "sections": 0, "warnings": []}
    if hf_cfg:
        body_idx = None
        if hf_cfg.get("body_starts_new_section") and body_heading_idx is not None:
            body_idx = sec_map.get(body_heading_idx, 0)
        try:
            hf_report = apply_headers_footers(doc, hf_cfg, body_idx)
        except Exception as e:
            hf_report = {"applied": False, "sections": 0,
                         "warnings": [f"页眉页脚应用失败：{e}"]}
        inject_report["warnings"] += hf_report.get("warnings", [])
    tables_cfg = profile.get("tables") if isinstance(profile.get("tables"), dict) else {}
    cell_style = tables_cfg.get("cell_style") if isinstance(tables_cfg.get("cell_style"), dict) else None
    table_cells_n = 0
    if cell_style:
        try:
            table_cells_n = restyle_table_cells(doc, cell_style)
        except Exception as e:
            inject_report["warnings"].append(f"表格单元格重排失败：{e}")

    protected = protect_object_paragraphs(doc, analysis)
    tables_report = normalize_tables(doc, tables_cfg)
    citations_cfg = profile.get("citations") or {}
    sup_n = normalize_citations(doc, analysis, assignment,
                                citations_cfg.get("superscript", True))
    # 图片宽度统一（仅当方案 objects.figure.resize=true）→ 超版心安全缩放兜底
    resized_n = 0
    objects_cfg = profile.get("objects") or {}
    if (objects_cfg.get("figure") or {}).get("resize"):
        margins = (profile.get("page") or {}).get("margins", {}) or {}
        paper_w = {"A4": 21.0, "A3": 29.7, "B5": 17.6, "Letter": 21.59}.get(
            (profile.get("page") or {}).get("paper", "A4"), 21.0)
        content_cm = paper_w - margins.get("left_cm", 3.17) - margins.get("right_cm", 3.17)
        resized_n = resize_figures_to_reference(doc, analysis, objects_cfg, content_cm)
    fits = fit_oversized_figures(doc, analysis, profile.get("page") or {})
    # ---- M2：目录域收尾处理。插入 TOC 段落会使其后段落号整体 +1，
    # 因此必须在所有按 analysis 段落号定位的操作（防裁剪保护/引用上标/
    # 图宽缩放）之后执行；updateFields 无论插入与否都写。
    toc_report = {"inserted": False, "update_fields": False, "reason": ""}
    try:
        toc_report = apply_toc_updates(doc, profile.get("toc") or {}, body_heading_idx)
    except Exception as e:
        toc_report = {"inserted": False, "update_fields": False,
                      "reason": f"目录域处理失败：{e}"}
    doc.save(str(docx_path))

    return {
        "backup": str(backup_path) if backup_path else None,
        "paragraphs": sum(counts.values()),
        "ignored": skipped,
        "protected_objects": protected,
        "numbering_preserved": renumbered,
        "numbering_suppressed": suppressed_num,
        "headings_unnumbered": unnumbered,
        "headings_outline_only": outline_only,
        "numbering_mode": mode,
        "numbering_mode_label": MODE_LABELS.get(mode, mode),
        "tables": tables_report,
        "citations_found": analysis.citation_count,
        "citations_superscripted": sup_n,
        "dangling_citations": analysis.dangling_citations,
        "uncited_entries": analysis.uncited_entries,
        "figures_resized": resized_n + fits["resized"],
        "figures_float_skipped": fits["skipped_float"],
        "header_footer": hf_report,
        "toc": toc_report,
        "table_cells": table_cells_n,
        "counts": counts,
        "styles": inject_report["styles"],
        "numbering": inject_report["numbering"],
        "warnings": inject_report["warnings"] + failed,
        "manual_numbering_kept": mode == "manual",
    }
