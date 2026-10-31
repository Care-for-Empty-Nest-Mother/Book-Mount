# -*- coding: utf-8 -*-
"""Word 同款功能图标：QPainter 程序化绘制，不依赖图标资源文件。

对齐四钮 = Word"文档行"线条图标；行距/缩进 = 行线+箭头；字体颜色 =
大写 A + 底部色条（色条颜色实时反映当前所选颜色）；清除格式 = A + 红斜杠。
纯函数（kind → QIcon），便于离屏单测。
"""
from PySide6.QtCore import QPointF, QRectF, Qt
from PySide6.QtGui import QColor, QFont, QIcon, QPainter, QPen, QPixmap

INK = "#343a40"          # 默认线条色（近 Word 的深灰）


def _base_pixmap(size: int) -> tuple[QPixmap, QPainter]:
    pix = QPixmap(size, size)
    pix.fill(Qt.transparent)
    p = QPainter(pix)
    p.setRenderHint(QPainter.Antialiasing)
    return pix, p


def _pen(p: QPainter, color: str, width: float = 2.0):
    p.setPen(QPen(QColor(color), width, Qt.SolidLine, Qt.RoundCap))


def _draw_text_lines(p: QPainter, size: int, anchor: str, justify: bool = False,
                     color: str = INK):
    """画四条"文档行"：anchor ∈ left/center/right；justify=两端（全等宽）。"""
    m = size * 0.18
    w = size - 2 * m
    ys = [size * (0.22 + 0.19 * i) for i in range(4)]
    fracs = (1.0, 0.62, 1.0, 1.0) if justify else (1.0, 0.62, 1.0, 0.72)
    _pen(p, color, max(1.8, size / 14))
    for y, frac in zip(ys, fracs):
        lw = w if justify else w * frac
        if anchor == "left":
            x0 = m
        elif anchor == "center":
            x0 = m + (w - lw) / 2
        else:
            x0 = m + w - lw
        p.drawLine(QPointF(x0, y), QPointF(x0 + lw, y))


def _draw_arrow(p: QPainter, x0, y0, x1, y1, color: str, head: float = 3.2):
    """带双端箭头的线段（行距/缩进图标）。"""
    _pen(p, color, 1.8)
    p.drawLine(QPointF(x0, y0), QPointF(x1, y1))
    import math
    ang = math.atan2(y1 - y0, x1 - x0)
    for tx, ty, back in ((x1, y1, True), (x0, y0, False)):
        a = ang + math.pi if back else ang
        p.drawLine(QPointF(tx, ty),
                   QPointF(tx - head * math.cos(a - 0.45),
                           ty - head * math.sin(a - 0.45)))
        p.drawLine(QPointF(tx, ty),
                   QPointF(tx - head * math.cos(a + 0.45),
                           ty - head * math.sin(a + 0.45)))


def make_icon(kind: str, color: str = INK, size: int = 28) -> QIcon:
    """kind → Word 同款 QIcon。"""
    pix, p = _base_pixmap(size)
    m = size * 0.16
    w = size - 2 * m
    if kind in ("align_left", "align_center", "align_right", "align_justify"):
        _draw_text_lines(p, size, {"align_left": "left", "align_center": "center",
                                   "align_right": "right"}.get(kind, "left"),
                         justify=(kind == "align_justify"), color=color)
    elif kind == "line_spacing":
        # 左侧上下双箭头 + 右侧三行文档行（行距渐宽，同 Word）
        x_arrow = m + 2
        _draw_arrow(p, x_arrow, size * 0.20, x_arrow, size * 0.80, color)
        _pen(p, color, max(1.8, size / 14))
        for i, frac in enumerate((0.85, 0.85, 0.85)):
            y = size * (0.26 + 0.24 * i)
            p.drawLine(QPointF(m + 7, y), QPointF(m + 7 + (w - 7) * frac, y))
    elif kind in ("indent_inc", "indent_dec"):
        # 三行文档行 + 水平箭头（增=向右，减=向左），同 Word 缩进按钮
        _draw_text_lines(p, size, "left", color=color)
        y = size * 0.5
        if kind == "indent_inc":
            _draw_arrow(p, size * 0.30, y, size * 0.86, y, color)
        else:
            _draw_arrow(p, size * 0.86, y, size * 0.30, y, color)
    elif kind == "font_color":
        p.setPen(QPen(QColor(color), 1))
        f = QFont("Times New Roman", int(size * 0.52), QFont.Bold)
        p.setFont(f)
        p.drawText(QRectF(0, -size * 0.10, size, size * 0.82),
                   Qt.AlignCenter, "A")
        p.setPen(Qt.NoPen)
        p.setBrush(QColor(color))
        p.drawRect(QRectF(m, size * 0.80, w, size * 0.11))
    elif kind == "clear_format":
        # 灰色 A + 红色斜杠（清除格式）
        p.setPen(QPen(QColor("#868e96"), 1))
        p.setFont(QFont("Times New Roman", int(size * 0.5), QFont.Bold))
        p.drawText(QRectF(0, -size * 0.08, size, size * 0.8), Qt.AlignCenter, "A")
        _pen(p, "#c92a2a", max(2.0, size / 11))
        p.drawLine(QPointF(size * 0.16, size * 0.84), QPointF(size * 0.84, size * 0.16))
    else:                                    # 未知种类：画个占位圆点，不崩
        p.setPen(Qt.NoPen)
        p.setBrush(QColor(color))
        p.drawEllipse(QRectF(m, m, w, w))
    p.end()
    return QIcon(pix)


def glyph_button_text(kind: str) -> str:
    """B/I 等衬线字形按钮的文字（由调用方配 serif 加粗字体渲染）。"""
    return {"bold": "B", "italic": "I"}.get(kind, kind)
