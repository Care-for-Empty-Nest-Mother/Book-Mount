# -*- coding: utf-8 -*-
"""渲染器：格式方案 JSON → Word 原生 .dotx 模板。

做法：以 python-docx 默认模板为底座，重建段落样式表、写入 docDefaults、
追加多级编号定义、设置页面参数，最后把包内容类型改写为 template。
"""
import tempfile
import zipfile
from pathlib import Path

from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml.ns import qn
from docx.shared import Cm

from .styles_xml import build_style_element, update_doc_defaults, find_style_by_id_or_name
from .numbering_xml import existing_ids, build_numbering, num_ref_for_styles

PAPER_SIZE_CM = {
    "A4": (21.0, 29.7),
    "A3": (29.7, 42.0),
    "B5": (17.6, 25.0),
    "Letter": (21.59, 27.94),
}


def numbering_root_of(document):
    """取文档的 numbering part 根元素；没有则返回 None。"""
    try:
        part = document.part.part_related_by(RT.NUMBERING)
        return part.element
    except KeyError:
        return None


def apply_page_setup(document, page: dict):
    if not isinstance(page, dict):
        return
    paper = page.get("paper", "A4")
    w, h = PAPER_SIZE_CM.get(paper, PAPER_SIZE_CM["A4"])
    margins = page.get("margins", {}) or {}
    for sec in document.sections:
        pw, ph = (w, h)
        if page.get("orientation") == "landscape":
            pw, ph = h, w
        sec.page_width = Cm(pw)
        sec.page_height = Cm(ph)
        if "top_cm" in margins:
            sec.top_margin = Cm(margins["top_cm"])
        if "bottom_cm" in margins:
            sec.bottom_margin = Cm(margins["bottom_cm"])
        if "left_cm" in margins:
            sec.left_margin = Cm(margins["left_cm"])
        if "right_cm" in margins:
            sec.right_margin = Cm(margins["right_cm"])
        if "header_cm" in margins:
            sec.header_distance = Cm(margins["header_cm"])
        if "footer_cm" in margins:
            sec.footer_distance = Cm(margins["footer_cm"])


def _scheme_of(profile):
    numbering = profile.get("numbering") or {}
    scheme = (numbering.get("heading_scheme") or {}).get("levels") or []
    return [lv for lv in scheme if isinstance(lv, dict)]


def render_styles_into(document, profile, remove_other_paragraph_styles=False) -> dict:
    """把方案样式注入 document。返回报告 dict。

    dotx 渲染：remove_other_paragraph_styles=True（模板里只留方案样式）；
    注入已有文档：False（只替换同 id/同名的样式，其余一律不动）。
    numbering.mode="keep_target" 时：不把方案编号挂到标题样式，并且替换标题
    样式定义时保留目标文档样式里已有的编号（只改视觉样式，编号体系不变）。
    """
    styles_el = document.styles.element
    report = {"styles": [], "numbering": None, "warnings": []}

    numbering_cfg = profile.get("numbering") or {}
    numbering_mode = numbering_cfg.get("mode", "reference")
    scheme = [] if numbering_mode == "keep_target" else _scheme_of(profile)
    num_id = None
    if scheme:
        nroot = numbering_root_of(document)
        if nroot is None:
            report["warnings"].append("该文档没有编号部件，跳过多级编号（样式仍会生成）")
        else:
            aids, nids = existing_ids(nroot)
            aid = (max(aids) + 1) if aids else 1
            num_id = (max(nids) + 1) if nids else 1
            absnum, num = build_numbering(scheme, aid, num_id)
            # abstractNum 必须位于所有 w:num 之前
            first_num = nroot.find(qn("w:num"))
            if first_num is not None:
                first_num.addprevious(absnum)
            else:
                nroot.append(absnum)
            nroot.append(num)
            report["numbering"] = {"abstractNumId": aid, "numId": num_id,
                                   "levels": len(scheme)}

    num_refs = num_ref_for_styles(scheme, num_id) if num_id else {}

    if remove_other_paragraph_styles:
        for st in list(styles_el.findall(qn("w:style"))):
            if st.get(qn("w:type"), "paragraph") == "paragraph":
                styles_el.remove(st)

    style_items = dict(profile.get("styles") or {})
    style_items.setdefault("body_text", {})
    for tid, spec in style_items.items():
        num_ref = num_refs.get(tid)
        el = build_style_element(tid, spec or {}, num_ref=num_ref,
                                 default=(tid == "body_text"))
        old = find_style_by_id_or_name(styles_el, el.get(qn("w:styleId")),
                                       (el.find(qn("w:name"))).get(qn("w:val")))
        if old is not None:
            if numbering_mode == "keep_target":
                # 保留目标样式自带的编号（自定义编号样式模式：只换视觉样式）
                old_ppr = old.find(qn("w:pPr"))
                if old_ppr is not None:
                    old_numpr = old_ppr.find(qn("w:numPr"))
                    if old_numpr is not None:
                        new_ppr = el.find(qn("w:pPr"))
                        if new_ppr is None:
                            from lxml import etree as _etree
                            new_ppr = _etree.SubElement(el, qn("w:pPr"))
                        import copy as _copy
                        new_ppr.insert(0, _copy.deepcopy(old_numpr))
            styles_el.remove(old)
        styles_el.append(el)
        report["styles"].append(tid)

    body_spec = style_items.get("body_text") or {}
    update_doc_defaults(styles_el, body_spec)
    return report


def save_as_dotx(document, out_path):
    """保存为真正的 .dotx：改写包内容类型后重新打包。"""
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
        tmp_path = Path(tmp.name)
    document.save(str(tmp_path))

    doc_ct = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
    tpl_ct = "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml"
    with zipfile.ZipFile(str(tmp_path), "r") as zin, \
            zipfile.ZipFile(str(out_path), "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "[Content_Types].xml":
                data = data.replace(doc_ct.encode(), tpl_ct.encode())
            zout.writestr(item, data)
    tmp_path.unlink(missing_ok=True)
    return out_path


def render_dotx(profile: dict, out_path) -> Path:
    """方案 JSON → .dotx 模板文件。"""
    doc = Document()          # python-docx 默认模板为底座
    report = render_styles_into(doc, profile, remove_other_paragraph_styles=True)
    apply_page_setup(doc, profile.get("page") or {})
    save_as_dotx(doc, out_path)
    return Path(out_path)
