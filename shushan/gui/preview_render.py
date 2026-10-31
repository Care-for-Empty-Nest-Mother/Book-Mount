# -*- coding: utf-8 -*-
"""docx → QTextDocument HTML 预览渲染器（工作台右栏的渲染层）。

纯函数（para_css / para_html / table_html / images_of）便于离屏单测；
build_preview 用光标逐元素插入并返回 段落号↔块号 双向映射——
双栏联动、ExtraSelection 底色、点击定位都以此映射为地基。
（实测：insertHtml 的 <p> 片段会并入当前块，必须逐段 insertBlock 分隔；
 表格按行列占块；图片经 addResource + <img src="img://..."> 渲染。）
"""
import html as _html

from docx.oxml.ns import qn

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
EMU_PER_PT = 12700

ALIGN_CSS = {"left": "left", "center": "center", "right": "right",
             "justify": "justify", "distributed": "justify"}


def para_css(fmt: dict | None) -> str:
    """实际生效格式 → 段落内联 CSS（Qt 富文本支持的子集）。"""
    fmt = fmt or {}
    css = []
    fams = []
    if fmt.get("font_en"):
        fams.append(f"'{fmt['font_en']}'")
    if fmt.get("font_cn"):
        fams.append(f"'{fmt['font_cn']}'")
    fams += ["'Times New Roman'", "'宋体'"]
    css.append("font-family:" + ",".join(fams))
    if fmt.get("size_pt"):
        css.append(f"font-size:{fmt['size_pt']:g}pt")
    if fmt.get("bold"):
        css.append("font-weight:bold")
    if fmt.get("italic"):
        css.append("font-style:italic")
    color = fmt.get("color")
    if color and color != "auto":
        css.append(f"color:#{color}")
    if fmt.get("align"):
        css.append(f"text-align:{ALIGN_CSS.get(fmt['align'], 'left')}")
    lm, lv = fmt.get("line_mode"), fmt.get("line_value")
    if lm == "multiple" and lv:
        css.append(f"line-height:{lv:g}")
    elif lm in ("exact", "at_least") and lv:
        css.append(f"line-height:{lv:g}pt")
    css.append(f"margin:{(fmt.get('before_pt') or 0):g}pt 0 {(fmt.get('after_pt') or 0):g}pt")
    if fmt.get("first_chars"):
        css.append(f"text-indent:{fmt['first_chars']:g}em")
    elif fmt.get("first_cm"):
        css.append(f"text-indent:{fmt['first_cm']:g}cm")
    pad = None
    if fmt.get("hanging_chars"):
        pad = f"{fmt['hanging_chars']:g}em"
    elif fmt.get("hanging_cm"):
        pad = f"{fmt['hanging_cm']:g}cm"
    elif fmt.get("left_chars"):
        pad = f"{fmt['left_chars']:g}em"
    elif fmt.get("left_cm"):
        pad = f"{fmt['left_cm']:g}cm"
    if pad:
        css.append(f"padding-left:{pad}")
    return ";".join(css)


def para_html(p, fmt: dict | None, badge: str = "", badge_color: str = "#868e96",
              extra: str = "") -> str:
    """单段 HTML 片段（<p> 包裹）。badge 为形如 '☑ 标题1' 的标记文字；
    extra 为图片/公式占位等附加内联内容。"""
    css = para_css(fmt)
    if p is not None and p.text.strip():
        body = _html.escape(p.text.strip())
    elif extra:
        body = extra
    else:
        hint = "【公式】" if (p is not None and p.object_kind == "formula") else "【图片】"
        body = f'<span style="color:#74a3b3">{hint}</span>'
    b = ""
    if badge:
        b = (f'<span style="background-color:{badge_color}; color:#ffffff; '
             f'font-size:8pt">&nbsp;{badge}&nbsp;</span>&nbsp;')
    return f'<p style="{css}">{b}{body}{extra}</p>'


def empty_para_html() -> str:
    """空段的极薄占位块（减少视觉噪声，但仍占一个块以对齐段落号）。"""
    return '<p style="font-size:4pt; margin:1pt 0">&nbsp;</p>'


def table_html(tbl, cell_css=None, max_cell_text: int = 60) -> str:
    """表格 → HTML 表。水平合并单元格按连续重复分组计 colspan；
    跨行合并（竖向）以重复文本近似（预览精度取舍，结构不受影响）。"""
    rows_html = []
    for row in tbl.rows:
        cells = row.cells
        tds = []
        i = 0
        while i < len(cells):
            j = i
            while j + 1 < len(cells) and cells[j + 1]._tc is cells[i]._tc:
                j += 1
            tc_el = cells[i]
            span = j - i + 1
            texts = []
            for cp in tc_el.paragraphs:
                t = cp.text.strip()
                if t:
                    texts.append(t)
            txt = "<br>".join(_html.escape(t[:max_cell_text]) for t in texts) or "&nbsp;"
            css = cell_css(tc_el) if cell_css else ""
            style = (f"padding:3pt 5pt; font-size:9pt; {css}") if css else \
                    "padding:3pt 5pt; font-size:9pt;"
            open_tag = f'<td colspan="{span}" style="{style}">' if span > 1 \
                else f'<td style="{style}">'
            tds.append(open_tag + txt + "</td>")
            i = j + 1
        rows_html.append("<tr>" + "".join(tds) + "</tr>")
    return (f'<table border="0.5" cellpadding="0" cellspacing="0" '
            f'style="margin:4pt 0; border-collapse:collapse">'
            + "".join(rows_html) + "</table>")


def images_of(p_el, part) -> list[tuple[bytes, str, float]]:
    """段内 DrawingML 内嵌图 → [(blob, content_type, 宽度pt), ...]。
    VML 旧图不处理（由调用方占位）。"""
    out = []
    for d in p_el.findall(".//" + qn("w:drawing")):
        ext = d.find(f"{{{WP_NS}}}extent")
        cx = int(ext.get("cx")) if (ext is not None and ext.get("cx")) else None
        blip = d.find(".//" + f"{{{A_NS}}}blip")
        rid = blip.get(f"{{{R_NS}}}embed") if blip is not None else None
        if not rid:
            continue
        try:
            prt = part.related_parts[rid]
            blob = prt.blob
            ct = prt.content_type or "image/png"
        except Exception:
            continue
        wpt = round(cx / EMU_PER_PT) if cx else 160
        out.append((blob, ct, max(24.0, min(float(wpt), 460.0))))
    return out


def insert_blocks(qdoc, frags: list[str]) -> list[int]:
    """把 HTML 片段逐块插入 qdoc（左栏样式卡用），返回各片段块号。"""
    from PySide6.QtGui import QTextCursor

    qdoc.clear()
    cur = QTextCursor(qdoc)
    out = []
    for frag in frags:
        out.append(cur.blockNumber())
        cur.insertHtml(frag)
        cur.insertBlock()
    return out


def build_preview(qdoc, analysis, docx, fmt_of, badge_of, cell_css=None,
                  img_width_pt: float = 420.0) -> tuple[dict, dict]:
    """按文档顺序把整篇插入 qdoc：段落（带徽标）+ 表格 + 图片。

    fmt_of(idx) → 段落生效格式；badge_of(idx) → (标记文字, 徽标色)，
    fmt 为 None 的空段自动渲染为薄占位块。
    返回 (para_to_block, block_to_para)；表格/单元格块不在映射内（点击忽略）。
    """
    from PySide6.QtCore import QUrl
    from PySide6.QtGui import QImage, QTextCursor, QTextDocument

    qdoc.clear()
    cur = QTextCursor(qdoc)
    paras = analysis.paragraphs
    tables = list(docx.tables)
    p_idx = 0
    t_idx = 0
    p2b: dict[int, int] = {}
    for child in docx.element.body:
        if child.tag == qn("w:p"):
            para = paras[p_idx] if p_idx < len(paras) else None
            fmt = fmt_of(p_idx)
            if fmt is None:
                p2b[p_idx] = cur.blockNumber()
                cur.insertHtml(empty_para_html())
                cur.insertBlock()
                p_idx += 1
                continue
            badge, bcolor = badge_of(p_idx)
            extra = ""
            if para is not None and para.has_object:
                if para.object_kind == "figure":
                    for k_i, (blob, ct, wpt) in enumerate(
                            images_of(child, docx.part)):
                        url = f"img://p{p_idx}_{k_i}"
                        qimg = QImage.fromData(blob)
                        if not qimg.isNull():
                            qdoc.addResource(QTextDocument.ImageResource,
                                             QUrl(url), qimg)
                            w = min(wpt, img_width_pt)
                            extra += f' <img src="{url}" width="{w:g}">'
                        else:
                            extra += ' <span style="color:#74a3b3">【图片】</span>'
                if para.object_kind == "formula" or (
                        para.has_object and not extra):
                    extra += ' <span style="color:#5c6bc0">【公式】</span>'
            p2b[p_idx] = cur.blockNumber()
            cur.insertHtml(para_html(para, fmt, badge, bcolor, extra))
            cur.insertBlock()
            p_idx += 1
        elif child.tag == qn("w:tbl"):
            if t_idx < len(tables):
                cur.insertHtml(table_html(tables[t_idx], cell_css))
                cur.insertBlock()
            t_idx += 1
    b2p = {b: p for p, b in p2b.items()}
    return p2b, b2p
