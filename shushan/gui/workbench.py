# -*- coding: utf-8 -*-
"""镜像 Word · 双栏对照格式工作台。

布局：顶部 = Word"开始"选项卡式功能区（字体/段落/样式三组，按钮外观与
快捷键都与 Word 一致）；左栏 = 方案样式库（模板，样式卡按其格式渲染）；
右栏 = 目标文档逐段预览（类型样式 + 逐段覆盖合并后的实际效果，保留
批注式彩色标记）。点选段落即改格式（落到 fmt_overrides 引擎），
"更新到类型样式"把所选格式提升为类型样式（Word"更新样式以匹配所选内容"
的心智，仅影响本次整稿，不写回方案文件）。

Accepted 后暴露四个属性（与 ClassifyDialog 契约兼容 + 新增一项）：
    assignment      {段落号: 类型}
    numbered        {标题段落号: bool}
    fmt_overrides   {段落号: 格式 spec}
    style_edits     {类型id: 格式 spec}
"""
import copy
import html as _html
from collections import Counter

from PySide6.QtCore import Qt
from PySide6.QtGui import (QColor, QFont, QKeySequence, QShortcut,
                           QTextCharFormat, QTextCursor)
from PySide6.QtWidgets import (QColorDialog, QComboBox, QDialog,
                               QDialogButtonBox, QDoubleSpinBox, QHBoxLayout,
                               QLabel, QPushButton, QSplitter, QToolButton,
                               QVBoxLayout, QWidget)

from ..analyzer.docx_analyzer import AnalysisResult
from ..profile_builder import spec_to_fmt_patch
from ..types import CN_FONT_SIZES, IGNORE_TYPE
from .classify_dialog import (SELECT_COLOR, TYPE_ITEMS, _ManuscriptView,
                              colors_of, fmt_detail, label_of)
from .font_picker import fill_font_combo, fill_size_combo, select_size
from .format_edit_dialog import FormatEditDialog
from .preview_render import build_preview, insert_blocks, para_css
from .word_icons import make_icon

RIBBON_CSS = """
QGroupBox { font-size: 11px; color: #666; border: 0; margin-top: 0; }
QGroupBox::title { subcontrol-origin: margin; subcontrol-position: bottom left;
                   left: 6px; }
QToolButton { border: 1px solid transparent; border-radius: 3px;
              padding: 2px 4px; font-size: 12px; }
QToolButton:hover { background: #e7f5ff; border-color: #74c0fc; }
QToolButton:checked { background: #d0ebff; border-color: #339af0; }
QComboBox { font-size: 12px; }
QDoubleSpinBox { font-size: 12px; }
"""

# 左栏样式卡的示例文字
SAMPLE_TEXT = {
    "heading_1": "第一章 绪论", "heading_2": "1.1 研究背景",
    "heading_3": "1.1.1 研究现状", "heading_4": "（四级标题）",
    "body_text": "正文示例：学术论文的排版格式应当规范统一，便于阅读与存档。",
    "abstract": "摘  要", "keywords": "关键词：格式；排版；复用",
    "caption_figure": "图1-1 系统总体架构", "caption_table": "表2-1 方案对比",
    "reference": "[1] 张三. 学术论文格式研究[J]. 期刊, 2021.",
    "appendix": "附录A", "cover": "论文题目", "statement": "原创性声明",
    "toc_1": "第一章 绪论", "toc_2": "1.1 研究背景", "toc_3": "1.1.1 现状",
    "author_info": "作者信息", "affiliation": "（作者单位）",
    "en_title": "Title of the Paper", "en_abstract": "Abstract",
    "en_keywords": "Keywords: format; layout",
    "figure_para": "【图片段】", "formula_para": "【公式段】",
    "header": "页眉示例", "footer": "页脚示例",
    "quote": "引用段落示例", "note": "注释示例",
}

ALIGN_ORDER = ["left", "center", "right", "justify"]  # 顺序仅供展示参考
LINE_ITEMS = [("multiple", 1.0), ("multiple", 1.15), ("multiple", 1.25),
              ("multiple", 1.5), ("multiple", 2.0), ("multiple", 2.5),
              ("multiple", 3.0)]


def _fmt_to_spec(fmt: dict) -> dict:
    """analyzer fmt 字典 → 格式 spec（方案 styles 同构），fmt_to_spec 的桥。"""
    spec = {}
    if fmt.get("font_cn"):
        spec["font_cn"] = fmt["font_cn"]
    if fmt.get("font_en"):
        spec["font_en"] = fmt["font_en"]
    if fmt.get("size_pt"):
        spec["size_pt"] = round(float(fmt["size_pt"]), 1)
    if fmt.get("bold") is not None:
        spec["bold"] = bool(fmt["bold"])
    if fmt.get("italic"):
        spec["italic"] = True
    if fmt.get("color") and fmt["color"] != "auto":
        spec["color"] = fmt["color"]
    if fmt.get("align"):
        spec["alignment"] = fmt["align"]
    if fmt.get("line_mode") and fmt.get("line_value"):
        spec["line_spacing"] = {"mode": fmt["line_mode"],
                                "value": round(float(fmt["line_value"]), 2)}
    if fmt.get("before_pt") is not None:
        spec["space_before_pt"] = round(float(fmt["before_pt"]), 1)
    if fmt.get("after_pt") is not None:
        spec["space_after_pt"] = round(float(fmt["after_pt"]), 1)
    for chars_key, cm_key, spec_key in (
            ("first_chars", "first_cm", "first_line_indent"),
            ("hanging_chars", "hanging_cm", "hanging_indent"),
            ("left_chars", "left_cm", "left_indent")):
        if fmt.get(chars_key):
            spec[spec_key] = {"unit": "chars", "value": round(float(fmt[chars_key]), 2)}
        elif fmt.get(cm_key):
            spec[spec_key] = {"unit": "cm", "value": round(float(fmt[cm_key]), 2)}
    return spec


class WorkbenchDialog(QDialog):
    """双栏对照格式工作台。用法：

        dlg = WorkbenchDialog(analysis, profile, 文档名, docx路径, parent)
        if dlg.exec() == QDialog.Accepted:
            apply_profile_offline(profile, docx, para_override=dlg.assignment,
                                  numbered_map=dlg.numbered,
                                  fmt_overrides=dlg.fmt_overrides)
            profile["styles"].update(dlg.style_edits)   # 类型样式修改（可选）
    """

    def __init__(self, analysis: AnalysisResult, profile: dict, doc_name: str,
                 docx_path, parent=None):
        super().__init__(parent)
        self.analysis = analysis
        self.profile = profile
        self._docx_path = str(docx_path)
        self.setWindowTitle(f"书山 · 格式工作台 —— {doc_name}")
        self.resize(1400, 800)
        self.paras = analysis.paragraphs
        self._docx = None                       # 惰性打开（渲染表格/图片用）
        self._styles: dict = dict(profile.get("styles") or {})
        self.suggestion = {p.index: p.suggestion for p in analysis.nonempty()}
        self.assignment: dict = dict(self.suggestion)
        self.numbered: dict = {p.index: True for p in analysis.nonempty()
                               if p.suggestion.startswith("heading_")}
        self.fmt_overrides: dict = {}
        self.style_edits: dict = {}
        self.selected: set = set()
        self._anchor: int | None = None
        self._nonempty_idx = {p.index for p in analysis.nonempty()}
        self._hover_block = None
        self._echo = False                      # 功能区回显中（抑制信号）
        self._undo_stack: list = []
        self._redo_stack: list = []
        self._p2b: dict = {}
        self._b2p: dict = {}
        self._card_tid: list = []               # 左栏块号 → 类型
        self._card_b2t: dict = {}

        self._build_ui()
        self._bind_shortcuts()
        self._rebuild_cards()
        self._rebuild_preview()
        self._refresh_highlights()
        self._sync_selection_label()
        self._update_citation_label()
        self._sync_ribbon()

    # ================= UI 构建 =================
    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setSpacing(4)
        root.addWidget(QLabel(
            "左=模板样式库（点卡片定位右栏同类段落）｜右=目标文档（点选段落，"
            "顶部按键即改格式，Shift 连选 / Ctrl 加选 / Alt+点标题切对号）"))

        root.addWidget(self._build_ribbon())

        split = QSplitter(Qt.Horizontal)
        self.card_view = _ManuscriptView(self._on_card_click, self._on_card_hover,
                                         lambda b: None, self._card_context,
                                         lambda b, pos: False)
        self.card_view.setFont(QFont("Microsoft YaHei", 10))
        self.card_view.setPlaceholderText("（模板样式库）")
        self.view = _ManuscriptView(self._on_click, self._on_hover,
                                    self._on_alt_click, self._context_menu,
                                    self._on_double)
        self.view.setFont(QFont("Microsoft YaHei", 10))
        self.view.setPlaceholderText("（目标文档预览）")
        split.addWidget(self.card_view)
        split.addWidget(self.view)
        split.setStretchFactor(0, 2)
        split.setStretchFactor(1, 3)
        root.addWidget(split, stretch=1)

        foot = QHBoxLayout()
        self.citation_label = QLabel("")
        self.citation_label.setWordWrap(True)
        self.citation_label.setStyleSheet("color:#8a6d1a; font-size:11px;")
        foot.addWidget(self.citation_label, stretch=1)
        self.sel_label = QLabel("已选中 0 段")
        foot.addWidget(self.sel_label)
        root.addLayout(foot)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("确认整稿")
        buttons.button(QDialogButtonBox.Cancel).setText("取消")
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        btn_reset = QPushButton("全部重置")
        btn_reset.clicked.connect(self._on_reset)
        buttons.addButton(btn_reset, QDialogButtonBox.ActionRole)
        root.addWidget(buttons)

    # ---------- 功能区（Word 开始选项卡分布） ----------
    def _build_ribbon(self) -> QWidget:
        from PySide6.QtWidgets import QFrame, QGroupBox
        bar = QWidget()
        bar.setStyleSheet(RIBBON_CSS)
        lay = QHBoxLayout(bar)
        lay.setContentsMargins(6, 4, 6, 2)
        lay.setSpacing(6)

        # ---- 字体组 ----
        g_font = QGroupBox("字体")
        fl = QHBoxLayout(g_font)
        fl.setContentsMargins(4, 2, 4, 8)
        self.font_combo = QComboBox()
        self.font_combo.setMinimumWidth(150)
        self.font_combo.setToolTip("字体（同设中/西文，与 Word 一致）；"
                                   "中西文分别设置请右键→修改格式")
        fill_font_combo(self.font_combo)
        self.font_combo.currentTextChanged.connect(self._on_font_changed)
        fl.addWidget(self.font_combo, stretch=3)
        self.size_combo = QComboBox()
        self.size_combo.setMinimumWidth(80)
        fill_size_combo(self.size_combo)
        self.size_combo.currentIndexChanged.connect(self._on_size_index)
        fl.addWidget(self.size_combo, stretch=1)
        self.btn_bold = self._glyph_button("B", True, False, "加粗 (Ctrl+B)",
                                           self._toggle_bold)
        self.btn_italic = self._glyph_button("I", False, True, "斜体 (Ctrl+I)",
                                             self._toggle_italic)
        fl.addWidget(self.btn_bold)
        fl.addWidget(self.btn_italic)
        self.color = "#000000"
        self.btn_color = QToolButton()
        self.btn_color.setIcon(make_icon("font_color", self.color))
        self.btn_color.setToolTip("字体颜色")
        self.btn_color.clicked.connect(self._pick_color)
        fl.addWidget(self.btn_color)
        lay.addWidget(g_font)

        sep1 = QFrame(); sep1.setFrameShape(QFrame.VLine)
        sep1.setStyleSheet("color:#d0d0d0")
        lay.addWidget(sep1)

        # ---- 段落组 ----
        g_para = QGroupBox("段落")
        pl = QHBoxLayout(g_para)
        pl.setContentsMargins(4, 2, 4, 8)
        self.align_btns = {}
        for kind, tip in (("left", "左对齐 (Ctrl+L)"), ("center", "居中 (Ctrl+E)"),
                          ("right", "右对齐 (Ctrl+R)"), ("justify", "两端对齐 (Ctrl+J)")):
            b = QToolButton()
            b.setIcon(make_icon("align_" + kind))
            b.setToolTip(tip)
            b.setCheckable(True)
            b.clicked.connect(lambda _=False, k=kind: self._set_align(k))
            self.align_btns[kind] = b
            pl.addWidget(b)
        self.btn_indent_inc = QToolButton()
        self.btn_indent_inc.setIcon(make_icon("indent_inc"))
        self.btn_indent_inc.setToolTip("增加首行缩进 0.5 字符 (Ctrl+M)")
        self.btn_indent_inc.clicked.connect(lambda: self._indent_step(0.5))
        self.btn_indent_dec = QToolButton()
        self.btn_indent_dec.setIcon(make_icon("indent_dec"))
        self.btn_indent_dec.setToolTip("减少首行缩进 0.5 字符 (Ctrl+Shift+M)")
        self.btn_indent_dec.clicked.connect(lambda: self._indent_step(-0.5))
        pl.addWidget(self.btn_indent_inc)
        pl.addWidget(self.btn_indent_dec)
        self.line_combo = QComboBox()
        self.line_combo.setToolTip("行距 (Ctrl+1 / Ctrl+5 / Ctrl+2)")
        for mode, v in LINE_ITEMS:
            self.line_combo.addItem(f"{v:g} 倍", (mode, v))
        self.line_combo.currentIndexChanged.connect(self._on_line_index)
        pl.addWidget(self.line_combo)
        self.before_spin = QDoubleSpinBox()
        self.before_spin.setRange(0, 200)
        self.before_spin.setDecimals(1)
        self.before_spin.setSingleStep(0.5)
        self.before_spin.setPrefix("段前 ")
        self.before_spin.setSuffix("磅")
        self.before_spin.setToolTip("段前间距（磅）")
        self.before_spin.valueChanged.connect(self._on_before_changed)
        self.after_spin = QDoubleSpinBox()
        self.after_spin.setRange(0, 200)
        self.after_spin.setDecimals(1)
        self.after_spin.setSingleStep(0.5)
        self.after_spin.setPrefix("段后 ")
        self.after_spin.setSuffix("磅")
        self.after_spin.setToolTip("段后间距（磅）")
        self.after_spin.valueChanged.connect(self._on_after_changed)
        pl.addWidget(self.before_spin)
        pl.addWidget(self.after_spin)
        lay.addWidget(g_para)

        sep2 = QFrame(); sep2.setFrameShape(QFrame.VLine)
        sep2.setStyleSheet("color:#d0d0d0")
        lay.addWidget(sep2)

        # ---- 样式组 ----
        g_style = QGroupBox("样式")
        sl = QHBoxLayout(g_style)
        sl.setContentsMargins(4, 2, 4, 8)
        self.type_combo = QComboBox()
        for tid, label in TYPE_ITEMS:
            self.type_combo.addItem(label, tid)
        self.type_combo.setToolTip("选中段的类型（可改判）")
        sl.addWidget(self.type_combo, stretch=1)
        btn_reassign = QPushButton("改判")
        btn_reassign.setToolTip("把选中段落改判为所选类型 (Ctrl+Shift+N=改判为正文)")
        btn_reassign.clicked.connect(self._on_reassign)
        sl.addWidget(btn_reassign)
        btn_upd = QPushButton("更新到类型样式")
        btn_upd.setToolTip("把所选段落的格式提升为该类型的样式\n"
                           "（Word「更新样式以匹配所选内容」；仅本次整稿生效）")
        btn_upd.clicked.connect(self._update_style)
        sl.addWidget(btn_upd)
        lay.addWidget(g_style)

        # ---- 通用组 ----
        g_ops = QGroupBox("")
        ol = QHBoxLayout(g_ops)
        ol.setContentsMargins(4, 2, 4, 8)
        self.btn_clear = QToolButton()
        self.btn_clear.setIcon(make_icon("clear_format"))
        self.btn_clear.setToolTip("清除本段的格式修改 (Ctrl+Q)")
        self.btn_clear.clicked.connect(self._clear_overrides)
        ol.addWidget(self.btn_clear)
        self.btn_undo = QPushButton("↺ 撤销")
        self.btn_undo.setToolTip("撤销 (Ctrl+Z)")
        self.btn_undo.clicked.connect(self._undo)
        self.btn_redo = QPushButton("↻ 重做")
        self.btn_redo.setToolTip("重做 (Ctrl+Y)")
        self.btn_redo.clicked.connect(self._redo)
        ol.addWidget(self.btn_undo)
        ol.addWidget(self.btn_redo)
        lay.addWidget(g_ops)
        lay.addStretch(1)
        return bar

    def _glyph_button(self, text, bold, italic, tip, slot) -> QToolButton:
        b = QToolButton()
        b.setText(text)
        f = QFont("Times New Roman", 13)
        f.setBold(bold)
        f.setItalic(italic)
        b.setFont(f)
        b.setToolTip(tip)
        b.setCheckable(True)
        b.clicked.connect(slot)
        return b

    def _bind_shortcuts(self):
        pairs = [
            ("Ctrl+B", self._toggle_bold), ("Ctrl+I", self._toggle_italic),
            ("Ctrl+E", lambda: self._set_align("center")),
            ("Ctrl+L", lambda: self._set_align("left")),
            ("Ctrl+R", lambda: self._set_align("right")),
            ("Ctrl+J", lambda: self._set_align("justify")),
            ("Ctrl+1", lambda: self._set_line("multiple", 1.0)),
            ("Ctrl+5", lambda: self._set_line("multiple", 1.5)),
            ("Ctrl+2", lambda: self._set_line("multiple", 2.0)),
            ("Ctrl+M", lambda: self._indent_step(0.5)),
            ("Ctrl+Shift+M", lambda: self._indent_step(-0.5)),
            ("Ctrl+D", self._open_format_dialog),
            ("Ctrl+Q", self._clear_overrides),
            ("Ctrl+Shift+N", lambda: self._reassign_to("body_text")),
            ("Ctrl+A", self._select_all),
            ("Escape", self._select_none),
            ("Ctrl+Z", self._undo), ("Ctrl+Y", self._redo),
            ("Ctrl+Shift+>", lambda: self._size_jump(1)),
            ("Ctrl+Shift+<", lambda: self._size_jump(-1)),
            ("Ctrl+]", lambda: self._size_step(1)),
            ("Ctrl+[", lambda: self._size_step(-1)),
        ]
        for seq, slot in pairs:
            sc = QShortcut(QKeySequence(seq), self)
            sc.activated.connect(slot)

    # ================= 数据视图 =================
    def _merged_styles(self, tid: str) -> dict | None:
        """类型最终样式 spec = 方案 styles + 本次 style_edits。无定义返回 None。"""
        spec = dict(self._styles.get(tid) or {})
        if tid in self.style_edits:
            spec = {**spec, **self.style_edits[tid]}
        return spec or None

    def _fmt_of(self, idx: int) -> dict | None:
        """右栏渲染用"整稿后实际生效格式"。

        ignore 段与"仅大纲级别"标题保持原格式；其余 = 类型样式 patch +
        逐段覆盖 patch（整稿会清直接格式，原 fmt 不参与合成）。
        """
        p = self.paras[idx]
        tid = self.assignment.get(idx, p.suggestion)
        if p.fmt is None:
            return None
        if tid == IGNORE_TYPE:
            return dict(p.fmt)
        if tid.startswith("heading_") and not self.numbered.get(idx, True):
            return dict(p.fmt)              # 仅大纲级别：样式与外观保持原样
        spec = self._merged_styles(tid)
        if spec is None:
            base = dict(p.fmt)              # 方案无此类型样式 → 整稿时保持原样
        else:
            base = spec_to_fmt_patch(spec)
        ov = self.fmt_overrides.get(idx)
        if ov:
            base = {**base, **spec_to_fmt_patch(ov)}
        return base

    def _badge_of(self, idx: int) -> tuple[str, str]:
        p = self.paras[idx]
        tid = self.assignment.get(idx, p.suggestion)
        _bg, badge = colors_of(tid)
        mark = label_of(tid)
        if tid.startswith("heading_"):
            mark = (f"☑ {label_of(tid)}" if self.numbered.get(idx, True)
                    else f"☐ 仅大纲{tid[8:]}")
        marks = ""
        if self.assignment.get(idx) != p.suggestion:
            marks += "✎"
        if idx in self.fmt_overrides:
            marks += "✍"
        if marks:
            mark = f"{mark} {marks}"
        return mark, badge

    def _cell_css(self, cell) -> str:
        """表格单元格统一 CSS = 正文样式 + tables.cell_style（整稿行为一致）。"""
        cs = dict(self._merged_styles("body_text") or {})
        cs = {**cs, **(self.profile.get("tables") or {}).get("cell_style", {})}
        return para_css(spec_to_fmt_patch(cs)) if cs else ""

    def _rebuild_preview(self):
        from docx import Document
        if self._docx is None:
            self._docx = Document(self._docx_path)
        self._p2b, self._b2p = build_preview(
            self.view.document(), self.analysis, self._docx,
            self._fmt_of, self._badge_of, cell_css=self._cell_css)

    def _refresh_highlights(self):
        doc = self.view.document()
        sels = []
        for idx, block_no in self._p2b.items():
            p = self.paras[idx]
            tid = self.assignment.get(idx, p.suggestion)
            bg = SELECT_COLOR if idx in self.selected else colors_of(tid)[0]
            if not bg:
                continue
            block = doc.findBlockByNumber(block_no)
            if not block.isValid():
                continue
            cur = QTextCursor(block)
            cur.movePosition(QTextCursor.EndOfBlock, QTextCursor.KeepAnchor)
            sel = self.view.ExtraSelection()
            sel.cursor = cur
            fmt = QTextCharFormat()
            fmt.setBackground(QColor(bg))
            sel.format = fmt
            sels.append(sel)
        self.view.setExtraSelections(sels)
        # 左栏：当前类型卡高亮
        csels = []
        maj = self._majority_type()
        cdoc = self.card_view.document()
        for block_no, tid in self._card_b2t.items():
            block = cdoc.findBlockByNumber(block_no)
            if not block.isValid():
                continue
            if maj and tid == maj:
                cur = QTextCursor(block)
                cur.movePosition(QTextCursor.EndOfBlock, QTextCursor.KeepAnchor)
                sel = self.card_view.ExtraSelection()
                sel.cursor = cur
                f = QTextCharFormat()
                f.setBackground(QColor(SELECT_COLOR))
                sel.format = f
                csels.append(sel)
        self.card_view.setExtraSelections(csels)

    def _rebuild_cards(self):
        frags = []
        self._card_tid = []
        order = [tid for tid, _label in TYPE_ITEMS]
        shown = [tid for tid in order if self._merged_styles(tid)]
        # 目标文档用到但方案未定义的类型，也列出（提示保持原样）
        used = {self.assignment.get(p.index, p.suggestion)
                for p in self.analysis.nonempty()}
        extra = [tid for tid in order if tid in used and tid not in shown]
        for tid in shown + extra:
            spec = self._merged_styles(tid) or {}
            label = label_of(tid)
            badge, badge_color = colors_of(tid)
            sample = SAMPLE_TEXT.get(tid, label)
            card = (f'<p style="{para_css(spec_to_fmt_patch(spec))}">'
                    f'<span style="background-color:{badge_color}; color:#fff; '
                    f'font-size:8pt">&nbsp;{label}&nbsp;</span>&nbsp;'
                    f'{_html.escape(sample)}</p>'
                    f'<p style="font-size:8pt; color:#868e96; margin:0 0 6pt 0">'
                    f'&nbsp;&nbsp;{fmt_detail(spec_to_fmt_patch(spec)) or "（保持 Word 默认）"}'
                    f'</p>')
            frags.append(card)
            self._card_tid.append(tid)
        self._card_b2t = {b: tid for b, tid in zip(
            insert_blocks(self.card_view.document(), frags), self._card_tid)}

    # ================= 选择与联动 =================
    def _scope(self) -> set:
        return {i for i in self.selected if self.paras[i].fmt is not None}

    def _majority_type(self) -> str | None:
        if not self.selected:
            return None
        return Counter(self.assignment.get(i, self.paras[i].suggestion)
                       for i in self.selected).most_common(1)[0][0]

    def _on_click(self, block_no, modifiers):
        idx = self._b2p.get(block_no)
        if idx is None:
            return
        if modifiers & Qt.ControlModifier:
            self.selected.symmetric_difference_update({idx})
            self._anchor = idx
        elif modifiers & Qt.ShiftModifier and self._anchor is not None:
            lo, hi = sorted((self._anchor, idx))
            self.selected = {i for i in range(lo, hi + 1) if i in self._nonempty_idx}
        else:
            self.selected = {idx}
            self._anchor = idx
        self._after_selection_change()

    def _on_card_click(self, block_no, modifiers):
        tid = self._card_b2t.get(block_no)
        if tid is None:
            return
        targets = sorted(i for i in self._nonempty_idx
                         if self.assignment.get(i, self.paras[i].suggestion) == tid)
        if not targets:
            return
        self.selected = set(targets)
        self._anchor = targets[0]
        block = self.view.document().findBlockByNumber(self._p2b.get(targets[0], -1))
        if block.isValid():
            self._scroll_block_to_middle(block)
        self._after_selection_change()

    def _after_selection_change(self):
        self._refresh_highlights()
        self._sync_selection_label()
        self._sync_ribbon()

    def _sync_selection_label(self):
        n = len(self.selected)
        maj = self._majority_type()
        self.sel_label.setText(f"已选中 {n} 段" + (f" · 类型 {label_of(maj)}" if maj else ""))

    def _select_all(self):
        self.selected = set(self._nonempty_idx)
        self._anchor = min(self.selected) if self.selected else None
        self._after_selection_change()

    def _select_none(self):
        self.selected = set()
        self._anchor = None
        self._after_selection_change()

    def _scroll_block_to_middle(self, block):
        cursor = QTextCursor(block)
        self.view.setTextCursor(cursor)
        rect = self.view.document().documentLayout().blockBoundingRect(block)
        viewport_h = self.view.viewport().height()
        target = max(0, int(rect.center().y() - viewport_h / 2))
        self.view.verticalScrollBar().setValue(target)

    def _on_hover(self, block_no, global_pos):
        from PySide6.QtWidgets import QToolTip
        if block_no == self._hover_block:
            return
        self._hover_block = block_no
        idx = self._b2p.get(block_no)
        if idx is None:
            QToolTip.hideText()
            return
        p = self.paras[idx]
        tid = self.assignment.get(idx, p.suggestion)
        changed = "" if tid == p.suggestion else "（已改判）"
        if idx in self.fmt_overrides:
            changed += " ✍已修改格式"
        text = (f"P{idx} · 判定 {label_of(tid)}{changed}"
                f"（建议 {label_of(p.suggestion)}）\n{fmt_detail(self._fmt_of(idx))}\n"
                f"依据：{p.reason}")
        QToolTip.showText(global_pos, text, self.view)

    def _on_card_hover(self, block_no, global_pos):
        from PySide6.QtWidgets import QToolTip
        tid = self._card_b2t.get(block_no)
        if tid is None:
            return
        spec = self._merged_styles(tid) or {}
        QToolTip.showText(global_pos, f"{label_of(tid)}\n{fmt_detail(spec_to_fmt_patch(spec))}",
                          self.card_view)

    def _card_context(self, block_no, global_pos):
        pass

    def _on_alt_click(self, block_no):
        idx = self._b2p.get(block_no)
        if idx is None:
            return
        tid = self.assignment.get(idx, self.paras[idx].suggestion)
        if not tid.startswith("heading_"):
            return
        self._push_undo()
        self.numbered[idx] = not self.numbered.get(idx, True)
        self._refresh_after_edit()

    def _on_double(self, block_no, global_pos):
        return False                       # 引用互跳沿用右键（v1 精简）

    def _context_menu(self, block_no, global_pos):
        from PySide6.QtWidgets import QMenu
        idx = self._b2p.get(block_no)
        if idx is None:
            return
        menu = QMenu(self)
        if idx in self.selected and self.selected:
            scope = set(self._scope())
            scope_label = f"修改格式…（{len(scope)} 段）"
        else:
            scope = {idx}
            self.selected = scope
            self._anchor = idx
            self._after_selection_change()
            scope_label = "修改格式…（本段）"
        act_edit = menu.addAction(scope_label)
        act_clear = menu.addAction("清除本段的格式修改 (Ctrl+Q)")
        act_clear.setEnabled(any(i in self.fmt_overrides for i in scope))
        chosen = menu.exec(global_pos)
        if chosen is act_edit:
            self._open_format_dialog()
        elif chosen is act_clear:
            self._clear_overrides()

    # ================= 功能区动作（按键与快捷键共用） =================
    def _push_undo(self):
        self._undo_stack.append(copy.deepcopy(
            (self.assignment, self.fmt_overrides, self.numbered, self.style_edits)))
        if len(self._undo_stack) > 100:
            self._undo_stack.pop(0)
        self._redo_stack.clear()

    def _undo(self):
        if not self._undo_stack:
            return
        self._redo_stack.append(copy.deepcopy(
            (self.assignment, self.fmt_overrides, self.numbered, self.style_edits)))
        self.assignment, self.fmt_overrides, self.numbered, self.style_edits = \
            self._undo_stack.pop()
        self._refresh_after_edit(cards=True)

    def _redo(self):
        if not self._redo_stack:
            return
        self._undo_stack.append(copy.deepcopy(
            (self.assignment, self.fmt_overrides, self.numbered, self.style_edits)))
        self.assignment, self.fmt_overrides, self.numbered, self.style_edits = \
            self._redo_stack.pop()
        self._refresh_after_edit(cards=True)

    def _majority_fmt(self) -> dict:
        """选中段（可归类）生效格式的多数派——功能区回显与快捷键基准。"""
        fmts = [self._fmt_of(i) for i in sorted(self._scope())]
        if not fmts:
            return {}
        keys = ("font_cn", "font_en", "size_pt", "bold", "italic", "color",
                "align", "line_mode", "line_value", "before_pt", "after_pt",
                "first_chars", "first_cm")
        out = {}
        for k in keys:
            vals = [f.get(k) for f in fmts if f.get(k) is not None]
            if not vals:
                continue
            v, n = Counter(vals).most_common(1)[0]
            out[k] = v if n * 2 > len(fmts) else None
        return out

    def _apply_spec(self, patch: dict):
        """把 patch 并入选中段每段的 fmt_overrides（一次撤销单位）。"""
        scope = self._scope()
        if not scope or not patch:
            return
        self._push_undo()
        for i in scope:
            ov = dict(self.fmt_overrides.get(i) or {})
            ov.update(copy.deepcopy(patch))
            self.fmt_overrides[i] = ov
        self._refresh_after_edit()

    def _toggle_bold(self):
        cur = self._majority_fmt().get("bold")
        self._apply_spec({"bold": not cur if cur is not None else True})

    def _toggle_italic(self):
        cur = self._majority_fmt().get("italic")
        self._apply_spec({"italic": not cur if cur is not None else True})

    def _on_font_changed(self, text):
        if self._echo or not self.selected:
            return
        name = text.strip()
        if name:
            self._apply_spec({"font_cn": name, "font_en": name})

    def _on_size_index(self, _i):
        if self._echo or not self.selected:
            return
        pt = self.size_combo.currentData()
        if pt:
            self._apply_spec({"size_pt": float(pt)})

    def _size_step(self, delta: int):
        cur = self._majority_fmt().get("size_pt") or 12
        self._apply_spec({"size_pt": round(max(1.0, min(200.0, cur + delta)), 1)})

    def _size_jump(self, direction: int):
        sizes = sorted(set(CN_FONT_SIZES.values()) |
                       {8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 26, 32, 42, 48})
        cur = self._majority_fmt().get("size_pt") or 12
        if direction > 0:
            nxt = next((s for s in sizes if s > cur + 0.01), sizes[-1])
        else:
            nxt = next((s for s in reversed(sizes) if s < cur - 0.01), sizes[0])
        self._apply_spec({"size_pt": round(float(nxt), 1)})

    def _set_align(self, kind: str):
        self._apply_spec({"alignment": kind})

    def _set_line(self, mode: str, value: float):
        self._apply_spec({"line_spacing": {"mode": mode, "value": value}})

    def _on_line_index(self, _i):
        if self._echo or not self.selected:
            return
        pair = self.line_combo.currentData()
        if pair:
            self._set_line(pair[0], pair[1])

    def _indent_step(self, delta: float):
        fmt = self._majority_fmt()
        cur = fmt.get("first_chars") or 0.0
        new = round(max(-10.0, min(20.0, cur + delta)), 2)
        patch = {}
        if new == 0:
            patch["first_line_indent"] = {"unit": "chars", "value": 0}
        else:
            patch["first_line_indent"] = {"unit": "chars", "value": new}
        self._apply_spec(patch)

    def _on_before_changed(self, v):
        if self._echo or not self.selected:
            return
        self._apply_spec({"space_before_pt": round(v, 1)})

    def _on_after_changed(self, v):
        if self._echo or not self.selected:
            return
        self._apply_spec({"space_after_pt": round(v, 1)})

    def _pick_color(self):
        from PySide6.QtWidgets import QColorDialog
        c = QColorDialog.getColor(QColor(self.color), self, "字体颜色")
        if not c.isValid():
            return
        self.color = c.name().lstrip("#").upper()
        self.btn_color.setIcon(make_icon("font_color", "#" + self.color))
        self._apply_spec({"color": self.color})

    def _open_format_dialog(self):
        scope = self._scope()
        if not scope:
            return
        pre = self._majority_fmt()
        dlg = FormatEditDialog(pre, len(scope), self)
        if dlg.exec() == QDialog.Accepted:
            self._apply_spec(dlg.spec())

    def _clear_overrides(self):
        scope = self._scope()
        if not any(i in self.fmt_overrides for i in scope):
            return
        self._push_undo()
        for i in scope:
            self.fmt_overrides.pop(i, None)
        self._refresh_after_edit()

    def _reassign_to(self, tid: str):
        if not self.selected:
            return
        self._push_undo()
        for i in self.selected:
            self.assignment[i] = tid
            if tid.startswith("heading_"):
                self.numbered.setdefault(i, True)
            else:
                self.numbered.pop(i, None)
        self._refresh_after_edit(cards=True)

    def _on_reassign(self):
        tid = self.type_combo.currentData()
        if tid:
            self._reassign_to(tid)

    def _update_style(self):
        """把选中段多数派格式提升为该类型的样式（style_edits）。"""
        scope = self._scope()
        tid = self._majority_type()
        if not scope or not tid or tid == IGNORE_TYPE:
            return
        spec = _fmt_to_spec(self._majority_fmt())
        if not spec:
            return
        self._push_undo()
        self.style_edits[tid] = {**(self.style_edits.get(tid) or {}), **spec}
        self._refresh_after_edit(cards=True)

    def _on_reset(self):
        self._push_undo()
        self.assignment = dict(self.suggestion)
        self.fmt_overrides.clear()
        self.numbered = {p.index: True for p in self.analysis.nonempty()
                         if p.suggestion.startswith("heading_")}
        self.style_edits.clear()
        self._refresh_after_edit(cards=True)

    def _refresh_after_edit(self, cards: bool = False):
        if cards:
            self._rebuild_cards()
        self._rebuild_preview()
        self._refresh_highlights()
        self._sync_selection_label()
        self._update_citation_label()
        self._sync_ribbon()

    # ================= 功能区回显 =================
    def _ribbon_set_enabled(self, on: bool):
        for w in (self.font_combo, self.size_combo, self.btn_bold, self.btn_italic,
                  self.btn_color, self.line_combo, self.before_spin, self.after_spin,
                  self.btn_clear, self.btn_indent_inc, self.btn_indent_dec):
            w.setEnabled(on)
        for b in self.align_btns.values():
            b.setEnabled(on)

    def _sync_ribbon(self):
        scope = self._scope()
        self._ribbon_set_enabled(bool(scope))
        if not scope:
            self.btn_undo.setEnabled(bool(self._undo_stack))
            self.btn_redo.setEnabled(bool(self._redo_stack))
            return
        fmt = self._majority_fmt()
        self._echo = True
        try:
            if fmt.get("font_cn"):
                self.font_combo.setEditText(fmt["font_cn"])
            elif fmt.get("font_en"):
                self.font_combo.setEditText(fmt["font_en"])
            if fmt.get("size_pt"):
                select_size(self.size_combo, float(fmt["size_pt"]))
            self.btn_bold.setChecked(bool(fmt.get("bold")))
            self.btn_italic.setChecked(bool(fmt.get("italic")))
            for kind, b in self.align_btns.items():
                b.setChecked(fmt.get("align") == kind)
            if fmt.get("line_mode") == "multiple" and fmt.get("line_value"):
                pos = self.line_combo.findData(
                    ("multiple", round(float(fmt["line_value"]), 2)))
                self.line_combo.setCurrentIndex(pos if pos >= 0 else -1)
            self.before_spin.setValue(float(fmt.get("before_pt") or 0))
            self.after_spin.setValue(float(fmt.get("after_pt") or 0))
            maj = self._majority_type()
            if maj:
                pos = next((i for i, (t, _) in enumerate(TYPE_ITEMS) if t == maj), 0)
                self.type_combo.setCurrentIndex(pos)
        finally:
            self._echo = False
        self.btn_undo.setEnabled(bool(self._undo_stack))
        self.btn_redo.setEnabled(bool(self._redo_stack))

    def _update_citation_label(self):
        a = self.analysis
        if not a.ref_entries and not a.citation_positions:
            self.citation_label.setText("")
            return
        parts = []
        if a.dangling_citations:
            nums = "、".join(f"[{n}]" for n in a.dangling_citations)
            parts.append(f"⚠ 悬空引用 {nums}")
        if a.uncited_entries:
            nums = "、".join(f"[{n}]" for n in a.uncited_entries[:8])
            parts.append(f"⚠ {len(a.uncited_entries)} 条文献未被引用：{nums}")
        if parts:
            self.citation_label.setText("引用检查：" + "；".join(parts))
        else:
            self.citation_label.setText(
                f"引用检查 ✓ {len(a.ref_entries)} 条文献与正文引用一一对应")
            self.citation_label.setStyleSheet("color:#1a6b45; font-size:11px;")
