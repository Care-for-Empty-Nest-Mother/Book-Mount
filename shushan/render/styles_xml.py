# -*- coding: utf-8 -*-
"""由方案 JSON 生成 w:style XML 元素。

覆盖策略：内置样式（标题1~4、正文、题注、目录1~3、页眉、页脚）用 Word 内置
styleId/w:name，Word 会按名匹配并本地化显示，大纲/目录/交叉引用等机制原生可用；
其余类型生成书山自定义样式（styleId 加 Shushan 前缀防撞）。
"""
from docx.oxml.ns import qn
from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
TW_PER_CM = 567.0

JC_MAP = {"left": "left", "center": "center", "right": "right",
          "justify": "both", "distributed": "distribute"}


def E(tag, **attrs):
    el = etree.SubElement(etree.Element("root"), qn(tag))  # 占位父级，稍后摘取
    for k, v in attrs.items():
        el.set(qn("w:" + k), str(v))
    return el


def _finish(el, parent_holder):
    parent_holder.append(el)
    return el


def _set_font(rpr, spec):
    cn, en = spec.get("font_cn"), spec.get("font_en")
    if cn or en:
        rf = E("w:rFonts")
        if en:
            rf.set(qn("w:ascii"), en)
            rf.set(qn("w:hAnsi"), en)
        if cn:
            rf.set(qn("w:eastAsia"), cn)
        rpr.append(rf)


def _spacing_xml(spec):
    parts = []
    if spec.get("space_before_pt") is not None:
        parts.append(("w:before", int(round(spec["space_before_pt"] * 20))))
    if spec.get("space_after_pt") is not None:
        parts.append(("w:after", int(round(spec["space_after_pt"] * 20))))
    ls = spec.get("line_spacing")
    line_rule, line_val = None, None
    if isinstance(ls, dict):
        mode, value = ls.get("mode"), ls.get("value")
        if mode == "multiple" and value:
            line_rule, line_val = "auto", int(round(value * 240))
        elif mode in ("exact", "at_least") and value:
            line_rule, line_val = ("exact" if mode == "exact" else "atLeast"), int(round(value * 20))
    if line_val is not None:
        parts.append(("w:line", line_val))
        parts.append(("w:lineRule", line_rule))
    return parts


def _indent_xml(spec):
    parts = []
    fli = spec.get("first_line_indent")
    if isinstance(fli, dict) and fli.get("value"):
        if fli.get("unit") == "chars":
            parts.append(("w:firstLineChars", int(round(fli["value"] * 100))))
        else:
            parts.append(("w:firstLine", int(round(fli["value"] * TW_PER_CM))))
    for key, prefix in (("hanging_indent", "hanging"), ("left_indent", "left"),
                        ("right_indent", "right")):
        v = spec.get(key)
        if isinstance(v, dict) and v.get("value"):
            if v.get("unit") == "chars":
                parts.append((f"w:{prefix}Chars", int(round(v["value"] * 100))))
            else:
                parts.append((f"w:{prefix}", int(round(v["value"] * TW_PER_CM))))
    return parts


def build_style_element(type_id: str, spec: dict, num_ref=None, default=False):
    """构造一个 w:style 元素。num_ref=(ilvl, numId) 时挂接多级编号。"""
    from ..types import style_id_of, style_name_of, DEFAULT_CN_NAMES

    style_id = style_id_of(type_id)
    holder = []
    root = etree.Element(qn("w:style"))
    root.set(qn("w:type"), "paragraph")
    root.set(qn("w:styleId"), style_id)
    if default:
        root.set(qn("w:default"), "1")

    name = etree.SubElement(root, qn("w:name"))
    name.set(qn("w:val"), style_name_of(type_id))
    based = etree.SubElement(root, qn("w:basedOn"))
    based.set(qn("w:val"), spec.get("base_on", "Normal"))
    if spec.get("display_name") and DEFAULT_CN_NAMES.get(type_id) != spec["display_name"]:
        aliases = etree.SubElement(root, qn("w:aliases"))
        aliases.set(qn("w:val"), spec["display_name"])
    etree.SubElement(root, qn("w:qFormat"))

    # ---- pPr ----
    ppr = etree.SubElement(root, qn("w:pPr"))
    if spec.get("keep_with_next"):
        etree.SubElement(ppr, qn("w:keepNext"))
    if spec.get("page_break_before"):
        etree.SubElement(ppr, qn("w:pageBreakBefore"))
    if num_ref is not None:
        ilvl, num_id = num_ref
        numpr = etree.SubElement(ppr, qn("w:numPr"))
        il = etree.SubElement(numpr, qn("w:ilvl"))
        il.set(qn("w:val"), str(ilvl))
        ni = etree.SubElement(numpr, qn("w:numId"))
        ni.set(qn("w:val"), str(num_id))
    spacing = _spacing_xml(spec)
    if spacing:
        sp = etree.SubElement(ppr, qn("w:spacing"))
        for attr, val in spacing:
            sp.set(qn(attr), str(val))
    indents = _indent_xml(spec)
    if indents:
        ind = etree.SubElement(ppr, qn("w:ind"))
        for attr, val in indents:
            ind.set(qn(attr), str(val))
    align = spec.get("alignment")
    if align:
        jc = etree.SubElement(ppr, qn("w:jc"))
        jc.set(qn("w:val"), JC_MAP.get(align, "left"))
    ol = spec.get("outline_level")
    if ol is not None:
        ole = etree.SubElement(ppr, qn("w:outlineLvl"))
        ole.set(qn("w:val"), str(int(ol)))

    # ---- rPr ----
    rpr = etree.SubElement(root, qn("w:rPr"))
    _set_font(rpr, spec)
    if spec.get("bold") is not None:
        b = etree.SubElement(rpr, qn("w:b"))
        if not spec["bold"]:
            b.set(qn("w:val"), "0")
    if spec.get("italic") is not None:
        i_el = etree.SubElement(rpr, qn("w:i"))
        if not spec["italic"]:
            i_el.set(qn("w:val"), "0")
    color = spec.get("color")
    if color and color != "auto":
        c = etree.SubElement(rpr, qn("w:color"))
        c.set(qn("w:val"), color)
    size = spec.get("size_pt")
    if size:
        sz = etree.SubElement(rpr, qn("w:sz"))
        sz.set(qn("w:val"), str(int(round(size * 2))))
        szcs = etree.SubElement(rpr, qn("w:szCs"))
        szcs.set(qn("w:val"), str(int(round(size * 2))))
    return root


def update_doc_defaults(styles_el, body_spec: dict):
    """把正文字体写进 docDefaults，让表格等未显式设置格式的对象也继承方案。"""
    if not isinstance(body_spec, dict):
        return
    dd = styles_el.find(qn("w:docDefaults"))
    if dd is None:
        dd = etree.SubElement(styles_el, qn("w:docDefaults"))
    rprd = dd.find(qn("w:rPrDefault"))
    if rprd is None:
        rprd = etree.SubElement(dd, qn("w:rPrDefault"))
    rpr = rprd.find(qn("w:rPr"))
    if rpr is None:
        rpr = etree.SubElement(rprd, qn("w:rPr"))
    for old in rpr.findall(qn("w:rFonts")):
        rpr.remove(old)
    _set_font(rpr, body_spec)
    size = body_spec.get("size_pt")
    if size:
        for tag in ("w:sz", "w:szCs"):
            for old in rpr.findall(qn(tag)):
                rpr.remove(old)
            sz = etree.SubElement(rpr, qn(tag))
            sz.set(qn("w:val"), str(int(round(size * 2))))


def find_style_by_id_or_name(styles_el, style_id, name):
    """在 styles.xml 中查找同 styleId 或同 w:name 的段落样式元素。"""
    for st in styles_el.findall(qn("w:style")):
        stype = st.get(qn("w:type"), "paragraph")
        if stype != "paragraph":
            continue
        if st.get(qn("w:styleId")) == style_id:
            return st
        name_el = st.find(qn("w:name"))
        if name_el is not None and name_el.get(qn("w:val")) == name:
            return st
    return None
