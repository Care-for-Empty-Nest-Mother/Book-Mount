# -*- coding: utf-8 -*-
"""实际生效格式解析：docDefaults ← 样式继承链 ← 段落直接格式 ← 字符直接格式。

"实际生效格式"指 Word 打开文档时真正渲染出来的格式。一篇手动排版的文献可能
格式写在样式里、写在段落属性里、写在字符属性里，三者任意叠加；本模块把它们
合成出最终值，供聚类与方案生成使用。
"""
import re
from copy import deepcopy
from docx.oxml.ns import qn

TW_PER_CM = 567.0        # 1 厘米 ≈ 567 缇
LINE_UNITS_PER_MULT = 240.0  # lineRule=auto 时，line/240 = 行距倍数

# 含这些元素的段落承载图片/OLE对象/公式——整稿时必须保护：
# 1) 不重排（内容载体）；
# 2) 生效行距若是"固定值"，嵌入式图片会被 Word 裁剪到行高（视觉上"消失"），
#    必须覆盖为倍数行距。
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
O_NS = "urn:schemas-microsoft-com:office:office"
OBJECT_TAGS = tuple(f"{{{ns}}}{tag}" for ns, tag in (
    (W_NS, "drawing"), (W_NS, "pict"), (W_NS, "object"), (M_NS, "oMath")))


def paragraph_has_object(p_el) -> bool:
    """段落是否包含图片 / VML / OLE 对象 / 公式。"""
    return any(p_el.find(".//" + tag) is not None for tag in OBJECT_TAGS)


def object_kind(p_el) -> str | None:
    """判定对象种类：'formula'（公式）或 'figure'（图片/其他对象）。"""
    if p_el.find(".//" + f"{{{M_NS}}}oMath") is not None:
        return "formula"
    has_drawing = (p_el.find(".//" + f"{{{W_NS}}}drawing") is not None
                   or p_el.find(".//" + f"{{{W_NS}}}pict") is not None)
    ole = p_el.find(".//" + f"{{{O_NS}}}OLEObject")
    if ole is not None and not has_drawing:
        proid = (ole.get("ProID") or "") + (ole.get("ProgID") or "")
        if "equation" in proid.lower() or "math" in proid.lower():
            return "formula"
        return "figure"
    if has_drawing:
        return "figure"
    return "formula" if ole is not None else None


def figure_extents_cm(p_el) -> list:
    """提取段内图片的显示尺寸 [(宽cm, 高cm), ...]。

    兼容两种存储：DrawingML（wp:extent，EMU 单位）与旧版 VML
    （v:shape 的 style="width:Xpt;height:Ypt"，.doc 转换件常见）。
    """
    out = []
    for ext in p_el.findall(".//" + f"{{{WP_NS}}}extent"):
        try:
            cx, cy = int(ext.get("cx")), int(ext.get("cy"))
            out.append((round(cx / 360000, 2), round(cy / 360000, 2)))
        except (TypeError, ValueError):
            continue
    V_NS = "urn:schemas-microsoft-com:vml"
    for shape in p_el.findall(".//" + f"{{{V_NS}}}shape"):
        style = shape.get("style") or ""
        wm = re.search(r"(?<![-\w])width\s*:\s*([\d.]+)pt", style)
        hm = re.search(r"(?<![-\w])height\s*:\s*([\d.]+)pt", style)
        if wm and hm:
            out.append((round(float(wm.group(1)) * 2.54 / 72, 2),
                        round(float(hm.group(1)) * 2.54 / 72, 2)))
    return out


def _bool_val(el):
    """w:b 等布尔元素：无 val 属性视为 True，val 为 0/false/off 视为 False。"""
    if el is None:
        return None
    v = el.get(qn("w:val"))
    if v is None:
        return True
    return str(v).lower() not in ("0", "false", "off")


def _tw_to_pt(v):
    return round(int(v) / 20.0, 2) if v is not None else None


def _merge(base: dict, overlay: dict) -> dict:
    out = deepcopy(base)
    for k, v in overlay.items():
        if v is not None:
            out[k] = v
    return out


def parse_rpr(rpr) -> dict:
    """解析一个 w:rPr 元素为格式字段字典（缺失字段为 None）。"""
    if rpr is None:
        return {}
    f = {}
    fonts = rpr.find(qn("w:rFonts"))
    if fonts is not None:
        ea = fonts.get(qn("w:eastAsia"))
        ascii_ = fonts.get(qn("w:ascii")) or fonts.get(qn("w:hAnsi"))
        f["font_cn"] = ea
        f["font_en"] = ascii_
    sz = rpr.find(qn("w:sz"))
    if sz is not None and sz.get(qn("w:val")):
        f["size_pt"] = int(sz.get(qn("w:val"))) / 2.0
    f["bold"] = _bool_val(rpr.find(qn("w:b")))
    f["italic"] = _bool_val(rpr.find(qn("w:i")))
    color = rpr.find(qn("w:color"))
    if color is not None:
        f["color"] = color.get(qn("w:val")) or None
    return f


def parse_ppr(ppr) -> dict:
    """解析一个 w:pPr 元素为段落格式字段字典。"""
    f = {}
    if ppr is None:
        return f
    jc = ppr.find(qn("w:jc"))
    if jc is not None:
        v = jc.get(qn("w:val")) or "left"
        f["align"] = {"left": "left", "start": "left", "center": "center",
                      "right": "right", "end": "right", "both": "justify",
                      "justify": "justify", "distribute": "distributed"}.get(v, "left")
    sp = ppr.find(qn("w:spacing"))
    if sp is not None:
        f["before_pt"] = _tw_to_pt(sp.get(qn("w:before")))
        f["after_pt"] = _tw_to_pt(sp.get(qn("w:after")))
        line, rule = sp.get(qn("w:line")), sp.get(qn("w:lineRule"), "auto")
        if line is not None:
            n = int(line)
            if rule == "auto":
                f["line_mode"], f["line_value"] = "multiple", round(n / LINE_UNITS_PER_MULT, 2)
            else:
                f["line_mode"] = "exact" if rule == "exact" else "at_least"
                f["line_value"] = round(n / 20.0, 2)
    ind = ppr.find(qn("w:ind"))
    if ind is not None:
        flc = ind.get(qn("w:firstLineChars"))
        fl = ind.get(qn("w:firstLine"))
        if flc:
            f["first_chars"] = int(flc) / 100.0
        elif fl:
            f["first_cm"] = round(int(fl) / TW_PER_CM, 2)
        hgc = ind.get(qn("w:hangingChars"))
        hg = ind.get(qn("w:hanging"))
        if hgc:
            f["hanging_chars"] = int(hgc) / 100.0
        elif hg:
            f["hanging_cm"] = round(int(hg) / TW_PER_CM, 2)
        lc = ind.get(qn("w:leftChars"))
        li = ind.get(qn("w:left"))
        if lc:
            f["left_chars"] = int(lc) / 100.0
        elif li:
            f["left_cm"] = round(int(li) / TW_PER_CM, 2)
    ol = ppr.find(qn("w:outlineLvl"))
    if ol is not None:
        f["outline"] = int(ol.get(qn("w:val")))
    numpr = ppr.find(qn("w:numPr"))
    if numpr is not None:
        nid = numpr.find(qn("w:numId"))
        if nid is not None and nid.get(qn("w:val")) not in (None, "0"):
            f["numbering"] = True
            f["num_id"] = nid.get(qn("w:val"))
            ilvl_el = numpr.find(qn("w:ilvl"))
            f["num_ilvl"] = int(ilvl_el.get(qn("w:val"))) if (
                ilvl_el is not None and ilvl_el.get(qn("w:val"))) else 0
    return f


class StylesDB:
    """styles.xml 的内存数据库：按 styleId 索引，支持继承链解析。"""

    def __init__(self, styles_el, theme: dict):
        self.theme = theme or {}
        self.doc_defaults = {}
        self.by_id = {}
        if styles_el is None:
            return
        dd = styles_el.find(qn("w:docDefaults"))
        if dd is not None:
            rpr_d = dd.find(qn("w:rPrDefault") + "/" + qn("w:rPr"))
            ppr_d = dd.find(qn("w:pPrDefault") + "/" + qn("w:pPr"))
            self.doc_defaults = {**parse_rpr(rpr_d), **parse_ppr(ppr_d)}
            for k in ("font_cn", "font_en"):
                if not self.doc_defaults.get(k):
                    self.doc_defaults[k] = self.theme.get(
                        "ea" if k == "font_cn" else "latin")
        for st in styles_el.findall(qn("w:style")):
            sid = st.get(qn("w:styleId"))
            if not sid:
                continue
            based = st.find(qn("w:basedOn"))
            self.by_id[sid] = {
                "name": (st.find(qn("w:name")).get(qn("w:val"))
                         if st.find(qn("w:name")) is not None else sid),
                "basedOn": based.get(qn("w:val")) if based is not None else None,
                "rPr": parse_rpr(st.find(qn("w:rPr"))),
                "pPr": parse_ppr(st.find(qn("w:pPr"))),
            }

    def resolve(self, style_id) -> dict:
        """沿 basedOn 链自根向叶合并，返回该样式的完整格式字典。"""
        merged = deepcopy(self.doc_defaults)
        chain, seen, cur = [], set(), style_id
        while cur and cur in self.by_id and cur not in seen:
            chain.append(cur)
            seen.add(cur)
            cur = self.by_id[cur]["basedOn"]
        for sid in reversed(chain):          # 根 → 叶，后者覆盖前者
            st = self.by_id[sid]
            merged = _merge(merged, {**st["rPr"], **st["pPr"]})
        return merged

    def name_of(self, style_id) -> str:
        st = self.by_id.get(style_id)
        return st["name"] if st else (style_id or "")


def theme_fonts(document) -> dict:
    """从 theme1.xml 读出默认西文/中文字体（minorFont）。"""
    try:
        theme_part = document.part.package.part_related_by(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme")
        root = theme_part._element
    except (KeyError, AttributeError):
        try:
            from docx.opc.constants import RELATIONSHIP_TYPE as RT
            theme_part = document.part.part_related_by(RT.THEME)
            from lxml import etree
            root = etree.fromstring(theme_part.blob)
        except Exception:
            return {}
    ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
    out = {}
    minor = root.find(".//a:minorFont", ns)
    if minor is not None:
        latin = minor.find("a:latin", ns)
        ea = minor.find("a:ea", ns)
        out["latin"] = latin.get("typeface") if latin is not None else None
        out["ea"] = ea.get("typeface") if ea is not None else None
    out = {k: (v if v not in ("", None) else None) for k, v in out.items()}
    return out


RUN_KEYS = ["font_cn", "font_en", "size_pt", "bold", "italic", "color"]


def effective_para_format(paragraph, sdb: StylesDB) -> dict:
    """合成一个段落的实际生效格式。"""
    p_el = paragraph._p
    ppr = p_el.find(qn("w:pPr"))
    sid = None
    if ppr is not None:
        pstyle = ppr.find(qn("w:pStyle"))
        if pstyle is not None:
            sid = pstyle.get(qn("w:val"))
    if sid is None and paragraph.style is not None:
        sid = paragraph.style.style_id

    fmt = sdb.resolve(sid)
    fmt = _merge(fmt, parse_ppr(ppr))

    # 字符级：按各 run 的文字长度加权投票（混排段落取占字数最多的那套）
    acc = {}
    for r in p_el.findall(".//" + qn("w:r")):
        text = "".join(t.text or "" for t in r.findall(qn("w:t")))
        if not text.strip():
            continue
        rf = _merge({k: fmt.get(k) for k in RUN_KEYS}, parse_rpr(r.find(qn("w:rPr"))))
        for k, v in rf.items():
            if v is None:
                continue
            w = acc.setdefault(k, {})
            w[v] = w.get(v, 0) + len(text)
    for k, w in acc.items():
        fmt[k] = max(w.items(), key=lambda kv: kv[1])[0]

    fmt["style_id"] = sid or ""
    fmt["style_name"] = sdb.name_of(sid)
    return fmt


def format_signature(fmt: dict) -> tuple:
    """把格式字典压成可哈希签名：同一排版方式的段落共享同一签名。"""
    def r2(v):
        return round(v, 2) if isinstance(v, float) else v
    return (
        fmt.get("font_cn"), fmt.get("font_en"),
        r2(fmt.get("size_pt")), fmt.get("bold"), fmt.get("italic"), fmt.get("color"),
        fmt.get("align"), fmt.get("line_mode"), r2(fmt.get("line_value")),
        r2(fmt.get("before_pt")), r2(fmt.get("after_pt")),
        r2(fmt.get("first_chars")), r2(fmt.get("first_cm")),
        r2(fmt.get("hanging_chars")), r2(fmt.get("hanging_cm")),
        r2(fmt.get("left_chars")), r2(fmt.get("left_cm")),
        fmt.get("outline"), fmt.get("numbering", False),
    )


def read_page_setup(document) -> dict:
    """读取第一节页面设置，返回方案 page 键结构。"""
    sec = document.sections[0]
    w_cm = round(sec.page_width.cm, 2) if sec.page_width else 21.0
    h_cm = round(sec.page_height.cm, 2) if sec.page_height else 29.7
    paper = "A4" if (abs(w_cm - 21.0) < 0.3 and abs(h_cm - 29.7) < 0.3) else \
            "B5" if (abs(w_cm - 17.6) < 0.3 and abs(h_cm - 25.0) < 0.3) else \
            "Letter" if (abs(w_cm - 21.59) < 0.5 and abs(h_cm - 27.94) < 0.5) else "A4"
    return {
        "paper": paper,
        "orientation": "portrait",
        "margins": {
            "top_cm": round(sec.top_margin.cm, 2),
            "bottom_cm": round(sec.bottom_margin.cm, 2),
            "left_cm": round(sec.left_margin.cm, 2),
            "right_cm": round(sec.right_margin.cm, 2),
            "header_cm": round(sec.header_distance.cm, 2),
            "footer_cm": round(sec.footer_distance.cm, 2),
        },
    }


# ---- 引用标记解析（顺序编码制 [1]、[2-4]、[1,3] 等） ----
import re as _re

RE_CITATION = _re.compile(r"\[(\d{1,3}(?:\s*[-,，~]\s*\d{1,3})*)\]")


def parse_citations(text: str, max_expand: int = 60) -> set:
    """提取正文里的引用编号集合；范围 [2-4] 展开为 {2,3,4}。"""
    nums = set()
    for m in RE_CITATION.finditer(text or ""):
        expr = m.group(1)
        for part in _re.split(r"[,，]", expr):
            part = part.strip()
            if "-" in part or "~" in part:
                mm = _re.match(r"(\d{1,3})\s*[-~]\s*(\d{1,3})", part)
                if mm:
                    lo, hi = int(mm.group(1)), int(mm.group(2))
                    if 0 < lo <= hi <= lo + max_expand:
                        nums.update(range(lo, hi + 1))
            elif part.isdigit():
                nums.add(int(part))
    return nums
