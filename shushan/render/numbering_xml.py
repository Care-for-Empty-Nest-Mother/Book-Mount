# -*- coding: utf-8 -*-
"""由方案 JSON 生成多级编号（abstractNum + num）XML。"""
from docx.oxml.ns import qn
from lxml import etree
from ..types import NUM_FMT_XML, style_id_of


def existing_ids(numbering_root) -> tuple[set, set]:
    aids, nids = set(), set()
    if numbering_root is None:
        return aids, nids
    for el in numbering_root.findall(qn("w:abstractNum")):
        v = el.get(qn("w:abstractNumId"))
        if v is not None:
            aids.add(int(v))
    for el in numbering_root.findall(qn("w:num")):
        v = el.get(qn("w:numId"))
        if v is not None:
            nids.add(int(v))
    return aids, nids


def build_numbering(scheme_levels: list, abstract_num_id: int, num_id: int):
    """构造 (abstractNum 元素, num 元素)。scheme_levels 见规范第 5 节。"""
    absnum = etree.Element(qn("w:abstractNum"))
    absnum.set(qn("w:abstractNumId"), str(abstract_num_id))
    mlt = etree.SubElement(absnum, qn("w:multiLevelType"))
    mlt.set(qn("w:val"), "multilevel")

    for lv in sorted(scheme_levels, key=lambda x: x.get("level", 1)):
        ilvl = int(lv.get("level", 1)) - 1
        lvl = etree.SubElement(absnum, qn("w:lvl"))
        lvl.set(qn("w:ilvl"), str(ilvl))
        start = etree.SubElement(lvl, qn("w:start"))
        start.set(qn("w:val"), "1")
        fmt = etree.SubElement(lvl, qn("w:numFmt"))
        fmt.set(qn("w:val"), NUM_FMT_XML.get(lv.get("num_fmt", "decimal"), "decimal"))
        pstyle = etree.SubElement(lvl, qn("w:pStyle"))
        pstyle.set(qn("w:val"), style_id_of(lv.get("style", f"heading_{ilvl + 1}")))
        suff = etree.SubElement(lvl, qn("w:suff"))
        suff.set(qn("w:val"), lv.get("suffix", "space"))
        text = etree.SubElement(lvl, qn("w:lvlText"))
        text.set(qn("w:val"), lv.get("text", f"%{ilvl + 1}"))
        jc = etree.SubElement(lvl, qn("w:lvlJc"))
        jc.set(qn("w:val"), "left")
        ppr = etree.SubElement(lvl, qn("w:pPr"))
        ind = etree.SubElement(ppr, qn("w:ind"))
        ind.set(qn("w:left"), "0")
        ind.set(qn("w:firstLine"), "0")

    num = etree.Element(qn("w:num"))
    num.set(qn("w:numId"), str(num_id))
    ref = etree.SubElement(num, qn("w:abstractNumId"))
    ref.set(qn("w:val"), str(abstract_num_id))
    return absnum, num


def num_ref_for_styles(scheme_levels: list, num_id: int) -> dict:
    """返回 {type_id: (ilvl, numId)}，供样式挂接 numPr。"""
    out = {}
    for lv in scheme_levels:
        tid = lv.get("style")
        if tid:
            out[tid] = (int(lv.get("level", 1)) - 1, num_id)
    return out
