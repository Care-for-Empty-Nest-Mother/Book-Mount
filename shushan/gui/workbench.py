# -*- coding: utf-8 -*-
"""书山主窗口 = 镜像 Word 双栏格式工作台（主界面与整稿台归一）。

布局：顶栏（方案库 + 提取/导入/渲染/整稿应用/换文档/清理Word后台）→
Word"开始"式功能区（字体/段落/样式/通用，按钮外观与快捷键同 Word）→
双栏（左=模板样式库，随所选方案即时刷新；右=目标文档预览，彩色标记保留）→
状态行（确认整稿/全部重置）→ 可折叠日志。

流程：启动即工作台（空态：左栏渲染所选方案样式卡，右栏引导）→
「整稿应用…」选文档 → 双栏填充 → 点选段落即改格式 →「确认整稿」→
报告写日志（窗口不关，可继续调整再次整稿或换文档）。

确认整稿后窗口仍持有四态（与旧 WorkbenchDialog 契约一致）：
    assignment / numbered / fmt_overrides / style_edits
"""
import copy
import html as _html
import os
import shutil
import tempfile
from collections import Counter
from pathlib import Path

from PySide6.QtCore import QEvent, Qt, QThread, QTimer, QUrl, Signal
from PySide6.QtGui import (QColor, QFont, QIcon, QKeySequence, QPixmap,
                           QShortcut, QTextCharFormat, QTextCursor)
from PySide6.QtWidgets import (QColorDialog, QComboBox, QDialog,
                               QDoubleSpinBox, QFileDialog, QFrame, QGroupBox,
                               QHBoxLayout, QLabel, QMainWindow, QMessageBox,
                               QPlainTextEdit, QPushButton, QScrollArea,
                               QSplitter, QStackedWidget, QToolButton,
                               QVBoxLayout, QWidget)

from ..analyzer.docx_analyzer import AnalysisResult, analyze_document
from ..convert import ensure_docx
from ..profile import ProfileError, load_profile, save_profile
from ..profile_builder import build_profile, spec_to_fmt_patch
from ..regions import (MARKABLE_TYPES, REGION_COLORS, REGION_LABELS, Region,
                       apply_region_rules, normalize_regions, region_type_at,
                       restore_suggestions, seed_regions)
from ..render import apply_profile_offline
from ..render.dotx_renderer import render_dotx
from ..types import CN_FONT_SIZES, IGNORE_TYPE
from ..word import cleanup as word_cleanup
from .classify_dialog import (SELECT_COLOR, TYPE_ITEMS, ClassifyDialog,
                              _ManuscriptView, colors_of, fmt_detail, label_of)
from .font_picker import fill_font_combo, fill_size_combo, select_size
from .format_edit_dialog import FormatEditDialog
from .numbering_dialog import NumberingModeDialog
from .preview_render import build_preview, insert_blocks, para_css
from .word_icons import make_icon

APP_TITLE = "书山 · 文献格式工坊"

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
QPushButton { font-size: 12px; }
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

LINE_ITEMS = [("multiple", 1.0), ("multiple", 1.15), ("multiple", 1.25),
              ("multiple", 1.5), ("multiple", 2.0), ("multiple", 2.5),
              ("multiple", 3.0)]

EMPTY_GUIDE = """
<p style="font-size:14pt; color:#1a3c34; margin:24pt 12pt 6pt 12pt">
<b>书山 · 把中意文献的格式一键套用到你的文档</b></p>
<p style="font-size:11pt; color:#495057; margin:2pt 12pt; line-height:1.7">
① 顶部选择一个<b>格式方案</b>（左侧即它的样式库预览）<br>
② 点 <b>「整稿应用…」</b> 选择你的 docx 文档<br>
③ 点选段落（Shift 连选 / Ctrl 加选），用上方 Word 同款按键调整格式，
   或点左栏样式卡定位同类段落<br>
④ 点 <b>「确认整稿」</b> ——原文件自动备份，报告见下方日志</p>
<p style="font-size:10pt; color:#868e96; margin:10pt 12pt">
提示：没有方案？点「提取格式…」从一篇排版规范的文献学起。</p>
"""


def _fmt_to_spec(fmt: dict) -> dict:
    """analyzer fmt 字典 → 格式 spec（方案 styles 同构）。"""
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


class ShushanWorkbench(QMainWindow):
    """书山主窗口：方案库 + 双栏对照工作台 + 整稿，一窗完成。"""

    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.resize(1400, 820)
        try:
            from ..paths import ICON_FILE
            if ICON_FILE.exists():
                self.setWindowIcon(QIcon(str(ICON_FILE)))
        except Exception:
            pass

        self.profile: dict | None = None
        self.profile_name: str = ""
        self.analysis: AnalysisResult | None = None
        self.work_path: str | None = None
        self._docx = None

        self.paras = []
        self._styles: dict = {}
        self.suggestion: dict = {}
        self.assignment: dict = {}
        self.numbered: dict = {}
        self.fmt_overrides: dict = {}
        self.style_edits: dict = {}
        self.selected: set = set()
        self._anchor: int | None = None
        self._nonempty_idx: set = set()
        self._hover_block = None
        self._echo = False
        self._undo_stack: list = []
        self._redo_stack: list = []
        self._p2b: dict = {}
        self._b2p: dict = {}
        self._card_tid: list = []
        self._card_b2t: dict = {}
        self._card_region_jump: dict = {}
        # 区域标记（M2.6）
        self.regions: list = []
        self._base_suggestion: dict = {}
        self._base_reasons: dict = {}
        self._touched_numbered: set = set()
        # 页面视图（Word 真渲染）
        self._page_mode = False
        self._page_tmp: str | None = None
        self._page_images: list = []
        self._page_zoom = 1.0
        self._pages_dirty = False
        self._render_thread = None
        self._render_busy = False

        self._build_ui()
        self._bind_shortcuts()
        self.refresh_profiles()
        self._show_empty_state()

    # ================= UI 构建 =================
    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setSpacing(3)
        root.setContentsMargins(6, 4, 6, 4)

        root.addWidget(self._build_topbar())
        root.addWidget(self._build_ribbon())

        split = QSplitter(Qt.Horizontal)
        self.card_view = _ManuscriptView(self._on_card_click, self._on_card_hover,
                                         lambda b: None, self._card_context,
                                         lambda b, pos: False)
        self.card_view.setFont(QFont("Microsoft YaHei", 10))
        self.view = _ManuscriptView(self._on_click, self._on_hover,
                                    self._on_alt_click, self._context_menu,
                                    self._on_double)
        self.view.setFont(QFont("Microsoft YaHei", 10))
        self.view_stack = QStackedWidget()
        self.view_stack.addWidget(self.view)              # index 0 连续流预览
        self.view_stack.addWidget(self._build_page_canvas())  # index 1 页面视图
        split.addWidget(self.card_view)
        split.addWidget(self.view_stack)
        split.setStretchFactor(0, 2)
        split.setStretchFactor(1, 3)
        root.addWidget(split, stretch=1)

        foot = QHBoxLayout()
        self.citation_label = QLabel("")
        self.citation_label.setWordWrap(True)
        self.citation_label.setStyleSheet("color:#8a6d1a; font-size:11px;")
        foot.addWidget(self.citation_label, stretch=1)
        self.sel_label = QLabel("未加载文档")
        foot.addWidget(self.sel_label)
        self.btn_confirm = QPushButton("确认整稿")
        self.btn_confirm.setToolTip("按当前对照结果整稿（原文件自动备份）；窗口不关闭，可继续调整")
        self.btn_confirm.setStyleSheet(
            "font-weight:600; padding:4px 14px; background:#d3f9d8;")
        self.btn_confirm.clicked.connect(self._apply_now)
        foot.addWidget(self.btn_confirm)
        self.btn_reset = QPushButton("全部重置")
        self.btn_reset.clicked.connect(self._on_reset)
        foot.addWidget(self.btn_reset)
        root.addLayout(foot)

        # 可折叠日志
        log_row = QHBoxLayout()
        self.btn_log = QPushButton("日志 ▾")
        self.btn_log.setFlat(True)
        self.btn_log.setStyleSheet("color:#495057; font-size:11px;")
        self.btn_log.clicked.connect(self._toggle_log)
        log_row.addWidget(self.btn_log)
        log_row.addStretch(1)
        root.addLayout(log_row)
        self.log = _LogView()
        root.addWidget(self.log)

    def _build_topbar(self) -> QWidget:
        bar = QWidget()
        lay = QHBoxLayout(bar)
        lay.setContentsMargins(4, 2, 4, 2)
        lay.setSpacing(5)
        lay.addWidget(QLabel("方案:"))
        self.profile_combo = QComboBox()
        self.profile_combo.setMinimumWidth(210)
        self.profile_combo.setToolTip("格式方案库（profiles/）")
        self.profile_combo.currentIndexChanged.connect(self._on_profile_combo)
        lay.addWidget(self.profile_combo, stretch=2)
        for text, tip, slot in (
                ("提取格式…", "从一篇排版规范的文献提取格式方案（弹批注视图）",
                 self.on_extract),
                ("导入…", "导入 .fprof.json 方案文件", self.on_import),
                ("渲染模板…", "方案 → .dotx Word 模板", self.on_render),
                ("删除", "删除选中的方案（不可恢复）", self.on_delete_profile)):
            b = QPushButton(text)
            b.setToolTip(tip)
            b.clicked.connect(slot)
            lay.addWidget(b)
        sep = QFrame(); sep.setFrameShape(QFrame.VLine)
        sep.setStyleSheet("color:#d0d0d0")
        lay.addWidget(sep)
        self.btn_pick_doc = QPushButton("整稿应用…")
        self.btn_pick_doc.setToolTip("选择目标文档并载入双栏预览")
        self.btn_pick_doc.setStyleSheet("font-weight:600; background:#d0ebff;")
        self.btn_pick_doc.clicked.connect(self._pick_and_load_document)
        lay.addWidget(self.btn_pick_doc)
        self.btn_page_view = QPushButton("页面视图")
        self.btn_page_view.setCheckable(True)
        self.btn_page_view.setToolTip("右侧切换为 Word 真渲染的逐页预览（需本机 Word，"
                                      "渲染约几秒）——分页/页眉页脚/页码/目录 100% 真实；\n"
                                      "再点一次返回连续流预览")
        self.btn_page_view.setEnabled(False)
        self.btn_page_view.clicked.connect(self._toggle_page_view)
        lay.addWidget(self.btn_page_view)
        self.btn_refresh_pages = QPushButton("刷新页面视图")
        self.btn_refresh_pages.setToolTip("按当前调整重新渲染页面视图（手动刷新）")
        self.btn_refresh_pages.setEnabled(False)
        self.btn_refresh_pages.clicked.connect(self._refresh_page_view)
        lay.addWidget(self.btn_refresh_pages)
        self.btn_cleanup_word = QPushButton("清理Word后台")
        self.btn_cleanup_word.setToolTip(
            "结束后台残留的 Word 进程（自动化残留）；\n"
            "你正在使用的、有窗口的 Word 不受影响")
        self.btn_cleanup_word.clicked.connect(self.on_cleanup_word)
        lay.addWidget(self.btn_cleanup_word)
        lay.addStretch(1)
        return bar

    def _build_ribbon(self) -> QWidget:
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

        # ---- 区域组（M2.6）----
        g_region = QGroupBox("区域")
        rl = QHBoxLayout(g_region)
        rl.setContentsMargins(4, 2, 4, 8)
        self.region_combo = QComboBox()
        for tid, label in MARKABLE_TYPES:
            self.region_combo.addItem(label, tid)
        self.region_combo.setToolTip("先点选/框选段落，再标记为所选区域类型；"
                                     "区间内的表格与图片随区间归属")
        rl.addWidget(self.region_combo, stretch=1)
        self.btn_region_mark = QPushButton("标记选中区域")
        self.btn_region_mark.setToolTip("把选中段落范围标记为所选区域，识别规则按区域重跑")
        self.btn_region_mark.clicked.connect(self._mark_region)
        rl.addWidget(self.btn_region_mark)
        self.btn_region_clear = QPushButton("清除标记")
        self.btn_region_clear.setToolTip("清除与选中范围重叠的区域标记（建议还原到分析基线）")
        self.btn_region_clear.clicked.connect(self._clear_region_marks)
        rl.addWidget(self.btn_region_clear)
        lay.addWidget(g_region)

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

    # ================= 日志 =================
    def log_line(self, text: str):
        self.log.append_line(text)

    def _toggle_log(self):
        self.log.toggle()

    # ================= 方案库（顶栏动作，自 panel.py 迁移） =================
    def refresh_profiles(self):
        from ..paths import PROFILES_DIR
        self.profile_combo.blockSignals(True)
        self.profile_combo.clear()
        names = sorted(p.name for p in PROFILES_DIR.glob("*.fprof.json"))
        self.profile_combo.addItem("（未选择方案）", None)
        for name in names:
            self.profile_combo.addItem(name, name)
        self.profile_combo.blockSignals(False)
        self.profile_combo.setCurrentIndex(1 if names else 0)   # 触发 _on_profile_combo

    def _on_profile_combo(self, _i=None):
        name = self.profile_combo.currentData()
        if not name:
            self.use_profile(None, "")
            return
        from ..paths import PROFILES_DIR
        try:
            data = load_profile(PROFILES_DIR / name)
        except ProfileError as e:
            QMessageBox.critical(self, APP_TITLE, str(e))
            return
        self.use_profile(data, name)

    def use_profile(self, profile: dict | None, name: str = ""):
        """切换当前方案（顶栏下拉与测试共用）。文档已加载时按新方案重渲染。"""
        self.profile = profile
        self.profile_name = name
        self._styles = dict((profile or {}).get("styles") or {})
        self.style_edits = {}               # 样式修改属于旧方案，随切换清空
        if self.analysis is None:
            self._rebuild_cards()
            return
        self.log_line(f"[方案] 已切换为「{name}」——双栏按新方案重渲染")
        self._rebuild_cards()
        self._rebuild_preview()
        self._refresh_highlights()
        self._sync_ribbon()

    # ---- 提取 / 导入 / 渲染 / 删除（原 panel.py 逻辑） ----
    def on_extract(self):
        src, _ = QFileDialog.getOpenFileName(self, "选择要提取格式的参考文献",
                                             "", "Word 文档 (*.docx *.doc)")
        if not src:
            return
        try:
            analysis_path, converted = ensure_docx(src)
            if converted:
                self.log_line(f"[转换] {Path(src).name} → {Path(converted).name}"
                              f"（原文件未动）")
            analysis = analyze_document(analysis_path)
        except Exception as e:
            QMessageBox.critical(self, APP_TITLE, f"分析失败：{e}")
            return
        analysis.source_path = str(src)
        dlg = ClassifyDialog(analysis, Path(src).name, self)
        if dlg.exec() != ClassifyDialog.Accepted:
            return
        from PySide6.QtWidgets import QInputDialog
        default_name = Path(src).stem + " 的格式"
        name, ok = QInputDialog.getText(self, "方案名称", "名称：",
                                        text=default_name)
        name = (name or "").strip()
        if not (ok and name):
            return
        profile = build_profile(analysis, dlg.assignment, name=name, source="docx",
                                fmt_overrides=dlg.fmt_overrides)
        from ..paths import PROFILES_DIR
        out = PROFILES_DIR / f"{name}.fprof.json"
        try:
            save_profile(profile, out)
        except ProfileError as e:
            QMessageBox.critical(self, APP_TITLE, str(e))
            return
        self.refresh_profiles()
        self.log_line(f"[提取] {src} → {out.name}"
                      f"（{len(profile.get('styles', {}))} 类格式）")

    def on_import(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "选择格式方案 JSON", "", "书山格式方案 (*.fprof.json *.fprof *.json)")
        if not path:
            return
        try:
            data = load_profile(path)
        except ProfileError as e:
            QMessageBox.critical(self, APP_TITLE, str(e))
            return
        from ..paths import PROFILES_DIR
        dest = PROFILES_DIR / Path(path).name
        shutil.copy2(path, dest)
        self.refresh_profiles()
        self.log_line(f"[导入] {Path(path).name} —— {data.get('meta', {}).get('name', '')}")

    def on_render(self):
        if not self._require_profile():
            return
        from ..paths import PROFILES_DIR
        stem = Path(self.profile_name).stem if self.profile_name else \
            (self.profile.get("meta", {}).get("name") or "方案")
        default_out = PROFILES_DIR / f"{stem}.dotx"
        out, _ = QFileDialog.getSaveFileName(self, "保存 Word 模板",
                                             str(default_out), "Word 模板 (*.dotx)")
        if not out:
            return
        try:
            render_dotx(self.profile, out)
        except Exception as e:
            QMessageBox.critical(self, APP_TITLE, f"渲染失败：{e}")
            return
        self.log_line(f"[渲染] {self.profile_name} → {out}\n        "
                      f"在 Word 中：文件→选项→加载项→管理:模板→选用 该模板并勾选"
                      f"「自动更新文档样式」即可换装。")

    def on_delete_profile(self):
        name = self.profile_combo.currentData()
        if not name:
            QMessageBox.information(self, APP_TITLE, "请先在方案下拉中选择一个方案。")
            return
        from ..paths import PROFILES_DIR
        ppath = PROFILES_DIR / name
        if not ppath.exists():
            QMessageBox.warning(self, APP_TITLE, f"方案文件不存在：{name}")
            return
        box = QMessageBox(self)
        box.setWindowTitle(APP_TITLE)
        box.setIcon(QMessageBox.Warning)
        box.setText(f"确定删除方案「{name}」吗？\n删除后不可恢复"
                    f"（渲染出的 .dotx 模板不受影响）。")
        yes = box.addButton("删除", QMessageBox.DestructiveRole)
        box.addButton("取消", QMessageBox.RejectRole)
        box.exec()
        if box.clickedButton() is not yes:
            self.log_line(f"[删除] 已取消：{name}")
            return
        try:
            ppath.unlink()
        except Exception as e:
            QMessageBox.critical(self, APP_TITLE, f"删除失败：{e}")
            return
        self.refresh_profiles()
        self.log_line(f"[删除] {name} 已删除")

    def on_cleanup_word(self):
        try:
            rep = word_cleanup.cleanup_word_background()
        except ImportError:
            QMessageBox.warning(self, APP_TITLE, "清理 Word 后台需要 pywin32 组件。")
            return
        except Exception as e:
            QMessageBox.critical(self, APP_TITLE, f"清理失败：{e}")
            return
        if rep["found"] == 0:
            self.log_line("[Word后台] 未发现 Word 进程")
        else:
            msg = (f"发现 {rep['found']} 个 Word 进程：结束后台 {len(rep['killed'])} 个")
            if rep["kept"]:
                msg += f"，保留可见窗口 {len(rep['kept'])} 个（你正在用的）"
            if rep["failed"]:
                msg += f"，{len(rep['failed'])} 个结束失败"
            self.log_line(f"[Word后台] {msg}")

    def _require_profile(self) -> bool:
        if self.profile is None:
            QMessageBox.information(self, APP_TITLE,
                                    "请先在顶部方案下拉中选择一个格式方案。")
            return False
        return True

    # ================= 文档载入 / 空态 =================
    def _pick_and_load_document(self):
        if not self._require_profile():
            return
        src, _ = QFileDialog.getOpenFileName(self, "选择待整稿的文档", "",
                                             "Word 文档 (*.docx *.doc)")
        if not src:
            return
        self.load_document(src)

    def load_document(self, src: str):
        """选定的目标文档载入双栏（顶栏与测试共用）。"""
        if not self._require_profile():
            return
        try:
            work_path, converted = ensure_docx(src)
            if converted:
                self.log_line(f"[转换] {Path(src).name} → {Path(converted).name}"
                              f"（整稿写入转换件，原 .doc 保持只读）")
            analysis = analyze_document(work_path)
        except Exception as e:
            QMessageBox.critical(self, APP_TITLE, f"分析失败：{e}")
            return
        self.work_path = str(work_path)
        self.analysis = analysis
        self.paras = analysis.paragraphs
        self._docx = None
        self._nonempty_idx = {p.index for p in analysis.nonempty()}
        self._base_suggestion = {p.index: p.suggestion for p in analysis.nonempty()}
        self._base_reasons = {p.index: p.reason for p in analysis.nonempty()}
        self._touched_numbered = set()
        self.regions = seed_regions(analysis)
        self.suggestion = {p.index: p.suggestion for p in analysis.nonempty()}
        self.assignment = dict(self.suggestion)
        self.numbered = {p.index: True for p in analysis.nonempty()
                         if p.suggestion.startswith("heading_")}
        self.fmt_overrides = {}
        self.style_edits = {}
        self.selected = set()
        self._anchor = None
        self._undo_stack.clear()
        self._redo_stack.clear()
        self.btn_pick_doc.setText("换文档…")
        self.btn_confirm.setEnabled(True)
        self.btn_reset.setEnabled(True)
        self._ribbon_set_enabled(True)
        for w in (self.region_combo, self.btn_region_mark, self.btn_region_clear):
            w.setEnabled(True)
        self.btn_page_view.setEnabled(True)
        self._page_images = []
        self._pages_dirty = True
        self._populate_pages()
        if self._page_mode:
            QTimer.singleShot(0, self._refresh_page_view)
        self._rebuild_cards()
        self._rebuild_preview()
        self._refresh_highlights()
        self._sync_selection_label()
        self._update_citation_label()
        self._sync_ribbon()
        self.log_line(f"[载入] {Path(src).name}（{len(list(analysis.nonempty()))} 段待归类）"
                      f"——对照调整后点「确认整稿」")

    def _show_empty_state(self):
        self.btn_confirm.setEnabled(False)
        self.btn_reset.setEnabled(False)
        self._ribbon_set_enabled(False)
        for w in (self.region_combo, self.btn_region_mark, self.btn_region_clear):
            w.setEnabled(False)
        self.btn_page_view.setEnabled(False)
        self.btn_refresh_pages.setEnabled(False)
        if self._page_mode:
            self.btn_page_view.setChecked(False)
            self._page_mode = False
        self._page_images = []
        self.view_stack.setCurrentIndex(0)
        self.sel_label.setText("未加载文档")
        self.view.setHtml(EMPTY_GUIDE.replace("its", "它的"))
        self._p2b, self._b2p = {}, {}

    # ================= 数据视图 =================
    def _merged_styles(self, tid: str) -> dict | None:
        spec = dict(self._styles.get(tid) or {})
        if tid in self.style_edits:
            spec = {**spec, **self.style_edits[tid]}
        return spec or None

    def _fmt_of(self, idx: int) -> dict | None:
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
        # 区域横幅：区域首段的徽标前置【区域名】并换区域色（不插行，零位移）
        if self.regions:
            rt = region_type_at(self.regions, idx)
            if rt != "auto" and any(r.type == rt and r.start == idx
                                    for r in self.regions):
                mark = f"【{REGION_LABELS.get(rt, rt)}】 {mark}"
                badge = REGION_COLORS.get(rt, badge)
        return mark, badge

    def _cell_css(self, cell) -> str:
        cs = dict(self._merged_styles("body_text") or {})
        cs = {**cs, **((self.profile or {}).get("tables") or {}).get("cell_style", {})}
        return para_css(spec_to_fmt_patch(cs)) if cs else ""

    def _rebuild_preview(self):
        from docx import Document
        if self.analysis is None:
            return
        if self._docx is None:
            self._docx = Document(self.work_path)
        self._p2b, self._b2p = build_preview(
            self.view.document(), self.analysis, self._docx,
            self._fmt_of, self._badge_of, cell_css=self._cell_css)

    def _refresh_highlights(self):
        if self.analysis is None:
            return
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
        actions = []                       # ("type", tid) / ("region", start) / ("none", None)
        self._card_tid = []
        # 区域总览卡（点击定位区域首段）
        if self.regions:
            frags.append('<p style="font-size:9pt; color:#1a3c34; margin:2pt 0">'
                         "<b>区域</b>（点击定位）</p>")
            actions.append(("none", None))
            for r in self.regions:
                color = REGION_COLORS.get(r.type, "#868e96")
                label = REGION_LABELS.get(r.type, r.type)
                frags.append(
                    f'<p style="font-size:9pt; margin:1pt 0 1pt 4pt">'
                    f'<span style="background-color:{color}; color:#fff; '
                    f'font-size:8pt">&nbsp;{label}&nbsp;</span>'
                    f' <span style="color:#495057">P{r.start}–P{r.end}</span></p>')
                actions.append(("region", r.start))
        if self.profile is None:
            frags.append('<p style="font-size:11pt; color:#868e96; margin:10pt">'
                         "（请先在顶部选择一个格式方案）</p>")
            actions.append(("none", None))
        else:
            order = [tid for tid, _label in TYPE_ITEMS]
            shown = [tid for tid in order if self._merged_styles(tid)]
            used = set()
            if self.analysis is not None:
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
                actions.append(("type", tid))
                self._card_tid.append(tid)
        blocks = insert_blocks(self.card_view.document(), frags)
        self._card_b2t = {}
        self._card_region_jump = {}
        for blk, (kind, val) in zip(blocks, actions):
            if kind == "type":
                self._card_b2t[blk] = val
            elif kind == "region":
                self._card_region_jump[blk] = val

    # ================= 选择与联动 =================
    def _scope(self) -> set:
        return {i for i in self.selected
                if self.analysis is not None and self.paras[i].fmt is not None}

    def _majority_type(self) -> str | None:
        if not self.selected or self.analysis is None:
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
        jump = self._card_region_jump.get(block_no)
        if jump is not None:
            self.selected = {jump}
            self._anchor = jump
            block = self.view.document().findBlockByNumber(self._p2b.get(jump, -1))
            if block.isValid():
                self._scroll_block_to_middle(block)
            self._after_selection_change()
            return
        tid = self._card_b2t.get(block_no)
        if tid is None or self.analysis is None:
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
        if self.analysis is None:
            self.sel_label.setText("未加载文档")
            return
        maj = self._majority_type()
        self.sel_label.setText(f"已选中 {n} 段" + (f" · 类型 {label_of(maj)}" if maj else ""))

    def _select_all(self):
        if self.analysis is None:
            return
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
        if idx is None or self.analysis is None:
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
        QToolTip.showText(global_pos,
                          f"{label_of(tid)}\n{fmt_detail(spec_to_fmt_patch(spec))}",
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
        self._touched_numbered.add(idx)             # 用户手动动过对号，区域规则不再覆盖
        self.numbered[idx] = not self.numbered.get(idx, True)
        self._refresh_after_edit()

    def _on_double(self, block_no, global_pos):
        return False

    def _context_menu(self, block_no, global_pos):
        from PySide6.QtWidgets import QMenu
        idx = self._b2p.get(block_no)
        if idx is None or self.analysis is None:
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
        self._apply_spec({"first_line_indent": {"unit": "chars", "value": new}})

    def _on_before_changed(self, v):
        if self._echo or not self.selected:
            return
        self._apply_spec({"space_before_pt": round(v, 1)})

    def _on_after_changed(self, v):
        if self._echo or not self.selected:
            return
        self._apply_spec({"space_after_pt": round(v, 1)})

    def _pick_color(self):
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
        if not self.selected or self.analysis is None:
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
        scope = self._scope()
        tid = self._majority_type()
        if not scope or not tid or tid == IGNORE_TYPE:
            return
        spec = _fmt_to_spec(self._majority_fmt())
        if not spec:
            return
        self._push_undo()
        self.style_edits[tid] = {**(self.style_edits.get(tid) or {}), **spec}
        self.log_line(f"[样式] 「{label_of(tid)}」已按所选内容更新"
                      f"（仅本次整稿生效）")
        self._refresh_after_edit(cards=True)

    # ================= 区域标记（M2.6） =================
    def _mark_region(self):
        if self.analysis is None or not self.selected:
            QMessageBox.information(self, APP_TITLE, "请先点选/框选要标记的段落。")
            return
        rtype = self.region_combo.currentData()
        if not rtype:
            return
        lo, hi = min(self.selected), max(self.selected)
        self._push_undo()
        self.regions = normalize_regions(
            self.regions + [Region(rtype, lo, hi)], len(self.paras))
        rep = apply_region_rules(self.analysis, self.regions)
        for idx in rep["numbered_off"]:
            if idx not in self._touched_numbered:
                self.numbered[idx] = False          # 非正文区标题：仅大纲不占号
        for p in self.analysis.nonempty():          # 区内建议重置为区域规则结果
            if rep["region_of"].get(p.index, "auto") != "auto":
                self.assignment[p.index] = p.suggestion
        self.log_line(f"[区域] 已标记 {REGION_LABELS.get(rtype, rtype)}：P{lo}–P{hi}"
                      f"（规则重跑改写 {rep['changed']} 段，"
                      f"{len(rep['numbered_off'])} 个标题默认仅大纲）")
        self._refresh_after_edit(cards=True)

    def _clear_region_marks(self):
        if self.analysis is None or not self.selected or not self.regions:
            return
        lo, hi = min(self.selected), max(self.selected)
        cleared = [r for r in self.regions
                   if r.type != "auto" and not (r.end < lo or r.start > hi)]
        if not cleared:
            self.log_line("[区域] 选中范围内没有可清除的标记")
            return
        self._push_undo()
        self.regions = normalize_regions(
            [r for r in self.regions if r not in cleared], len(self.paras))
        restore_suggestions(self.analysis, self._base_suggestion,
                            self._base_reasons, cleared,
                            self._touched_numbered, self.numbered)
        for p in self.analysis.nonempty():
            if any(r.start <= p.index <= r.end for r in cleared):
                self.assignment[p.index] = p.suggestion
        self.log_line("[区域] 已清除选中范围的区域标记（建议还原到分析基线）")
        self._refresh_after_edit(cards=True)

    # ================= 页面视图（Word 真渲染，M2.7） =================
    def _build_page_canvas(self) -> QWidget:
        wrap = QWidget()
        v = QVBoxLayout(wrap)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(0)
        self.page_banner = QLabel("")
        self.page_banner.setWordWrap(True)
        self.page_banner.setStyleSheet(
            "background:#fff3bf; color:#8a6d1a; padding:3px 6px; font-size:11px;")
        self.page_banner.hide()
        v.addWidget(self.page_banner)
        self.page_scroll = QScrollArea()
        self.page_scroll.setWidgetResizable(True)
        self.page_scroll.setStyleSheet("background:#525659;")
        self.page_container = QWidget()
        self.page_container.setStyleSheet("background:#525659;")
        self.page_lay = QVBoxLayout(self.page_container)
        self.page_lay.setAlignment(Qt.AlignHCenter)
        self.page_lay.setSpacing(8)
        self.page_scroll.setWidget(self.page_container)
        self.page_scroll.viewport().installEventFilter(self)
        v.addWidget(self.page_scroll, stretch=1)
        return wrap

    def eventFilter(self, obj, ev):
        if (obj is self.page_scroll.viewport() and ev.type() == QEvent.Wheel
                and ev.modifiers() & Qt.ControlModifier):
            factor = 1.1 if ev.angleDelta().y() > 0 else 1 / 1.1
            self._page_zoom = max(0.4, min(3.0, self._page_zoom * factor))
            self._populate_pages()
            return True
        return super().eventFilter(obj, ev)

    def _populate_pages(self):
        """按当前缩放把基准页图填充到画布。"""
        while self.page_lay.count():
            item = self.page_lay.takeAt(0)
            w = item.widget()
            if w is not None:
                w.deleteLater()
        for png in self._page_images:
            pix = QPixmap()
            pix.loadFromData(png)
            if self._page_zoom != 1.0:
                pix = pix.scaled(int(pix.width() * self._page_zoom),
                                 int(pix.height() * self._page_zoom),
                                 Qt.KeepAspectRatio, Qt.SmoothTransformation)
            lab = QLabel()
            lab.setAlignment(Qt.AlignCenter)
            lab.setPixmap(pix)
            self.page_lay.addWidget(lab)
        self.page_lay.addStretch(1)

    def _toggle_page_view(self):
        self._page_mode = self.btn_page_view.isChecked()
        if self._page_mode:
            if self.analysis is None:
                self.btn_page_view.setChecked(False)
                self._page_mode = False
                return
            self.view_stack.setCurrentIndex(1)
            self.btn_refresh_pages.setEnabled(not self._render_busy)
            if not self._page_images or self._pages_dirty:
                self._refresh_page_view()
        else:
            self.view_stack.setCurrentIndex(0)
            self.btn_refresh_pages.setEnabled(False)

    def _current_effective_profile(self) -> dict:
        """当前四态合并后的方案（style_edits + 缺失类型补空样式）。"""
        profile = copy.deepcopy(self.profile or {})
        if self.style_edits:
            profile.setdefault("styles", {}).update(copy.deepcopy(self.style_edits))
        styles = profile.setdefault("styles", {})
        for tid in set(self.assignment.values()):
            if tid and tid != IGNORE_TYPE and tid not in styles:
                styles[tid] = {}
        return profile

    def _refresh_page_view(self):
        if self.analysis is None or self._render_thread is not None:
            return
        if not self._require_profile():
            return
        if self._page_tmp is None:
            self._page_tmp = tempfile.mkdtemp(prefix="shushan_pages_")
        self._pages_dirty = False
        self._update_page_banner()
        self._set_page_busy(True)
        self._render_thread = _PageRenderWorker(
            self.work_path, self._page_tmp, self._current_effective_profile(),
            dict(self.assignment), dict(self.numbered), dict(self.fmt_overrides))
        self._render_thread.done_sig.connect(self._on_pages_ready)
        self._render_thread.error_sig.connect(self._on_pages_error)
        self._render_thread.start()

    def _on_pages_ready(self, pngs, elapsed):
        self._page_images = pngs
        self._populate_pages()
        self._set_page_busy(False)
        self._render_thread = None
        self._update_page_banner()
        self.log_line(f"[页面视图] 渲染完成：{len(pngs)} 页，耗时 {elapsed:.1f} 秒"
                      f"（Ctrl+滚轮缩放）")

    def _on_pages_error(self, msg):
        self._set_page_busy(False)
        self._render_thread = None
        self._pages_dirty = True
        self._update_page_banner()
        self.log_line(f"⚠ [页面视图] {msg}")

    def _set_page_busy(self, busy, text="渲染中…（Word 导出 PDF，约几秒）"):
        self._render_busy = busy
        self.btn_refresh_pages.setEnabled(not busy and self._page_mode)
        self.btn_page_view.setEnabled(not busy and self.analysis is not None)
        self.page_banner.setText(text if busy else "")
        self.page_banner.setVisible(busy)
        if busy:
            self.view_stack.setCurrentIndex(1)

    def _update_page_banner(self):
        if self._render_busy:
            return
        if self._page_mode and self._pages_dirty and self.analysis is not None:
            self.page_banner.setText("预览已过期（有新的调整）——点「刷新页面视图」更新")
            self.page_banner.setVisible(True)
        elif not self._render_busy:
            self.page_banner.setVisible(False)

    def closeEvent(self, ev):
        if self._page_tmp:
            shutil.rmtree(self._page_tmp, ignore_errors=True)
            self._page_tmp = None
        super().closeEvent(ev)

    def _on_reset(self):
        if self.analysis is None:
            return
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
        self._pages_dirty = True
        self._update_page_banner()

    # ================= 整稿（窗口内，不关窗） =================
    def _apply_now(self, numbering_override="__dialog__"):
        """确认整稿：style_edits 合并 → 编号模式 → apply → 日志。

        numbering_override：测试/高级用法直接传 (mode, scheme_levels) 跳过弹窗；
        缺省 "__dialog__" 时弹编号模式确认。
        """
        if self.analysis is None or self.work_path is None:
            QMessageBox.information(self, APP_TITLE, "请先「整稿应用…」选择目标文档。")
            return
        profile = self._current_effective_profile()
        assignment = dict(self.assignment)
        if numbering_override == "__dialog__":
            override = self._ask_numbering_mode(assignment)
            if override is None:
                self.log_line("[整稿] 已取消（编号模式未确认）")
                return
        else:
            override = numbering_override
        try:
            report = apply_profile_offline(
                profile, self.work_path, para_override=assignment,
                numbering_override=override,
                numbered_map=dict(self.numbered),
                fmt_overrides=dict(self.fmt_overrides))
        except Exception as e:
            QMessageBox.critical(self, APP_TITLE, f"整稿失败：{e}")
            self.log_line(f"[整稿] 失败：{e}")
            return
        self._log_report(report)
        # 整稿已写入磁盘（目录域插入等会使段落号位移）——按整稿后的真实状态
        # 重载分析，四态回到干净基线，用户可继续微调或换文档
        self.load_document(self.work_path)

    def _ask_numbering_mode(self, assignment):
        from ..render.apply import detect_numbering_mode
        detected = detect_numbering_mode(self.analysis, assignment)
        ndlg = NumberingModeDialog(detected, self)
        if ndlg.exec() != NumberingModeDialog.Accepted:
            return None
        override = ndlg.result_override()
        self.log_line(f"[编号模式] {override[0]}"
                      + ("（自定义编号样式）" if override[0] == "custom" else ""))
        return override

    def _log_report(self, report: dict):
        from ..types import DEFAULT_CN_NAMES
        self.log_line(f"[整稿] {Path(self.work_path).name}")
        msg = (f"已整稿 {report['paragraphs']} 段："
               + "、".join(f"{DEFAULT_CN_NAMES.get(k, k)}×{v}"
                           for k, v in sorted(report["counts"].items())))
        self.log_line("        " + msg)
        if report.get("backup"):
            self.log_line(f"        备份：{report['backup']}")
        if report.get("manual_numbering_kept"):
            self.log_line("        检测到手工编号，已保留文字编号、未启用自动编号")
        if report.get("numbering_suppressed"):
            self.log_line(f"        封面段落编号压制：{report['numbering_suppressed']} 段"
                          f"（封面不参与编号）")
        if report.get("protected_objects"):
            self.log_line(f"        图形/公式段防裁剪保护：{report['protected_objects']} 段")
        if report.get("headings_outline_only"):
            self.log_line(f"        {report['headings_outline_only']} 个标题按对号只设定大纲级别")
        if report.get("figures_resized"):
            self.log_line(f"        图片宽度调整：{report['figures_resized']} 张")
        if report.get("citations_superscripted"):
            self.log_line(f"        引用标记上标规范化：{report['citations_superscripted']} 处")
        if report.get("table_cells"):
            self.log_line(f"        表格单元格重排：{report['table_cells']} 段"
                          f"（tables.cell_style）")
        hf_rep = report.get("header_footer") or {}
        if hf_rep.get("applied"):
            self.log_line(f"        页眉页脚：已写入 {hf_rep.get('sections', 0)} 节"
                          f"（header_footer）")
        toc_rep = report.get("toc") or {}
        if toc_rep.get("inserted"):
            self.log_line("        目录域已插入（Word 中右键 → 更新域 生成条目）")
        elif toc_rep.get("reason"):
            self.log_line(f"        目录域：{toc_rep['reason']}")
        if report.get("dangling_citations"):
            nums = "、".join(f"[{n}]" for n in report["dangling_citations"])
            self.log_line(f"        ⚠ 悬空引用 {nums}：正文引用了不存在的文献条目")
        if report.get("uncited_entries"):
            nums = "、".join(f"[{n}]" for n in report["uncited_entries"][:8])
            self.log_line(f"        ⚠ {len(report['uncited_entries'])} 条文献未被正文引用：{nums}")
        for w in report.get("warnings", [])[:10]:
            self.log_line(f"        ⚠ {w}")
        self.log.expand()

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
        self.btn_undo.setEnabled(bool(self._undo_stack))
        self.btn_redo.setEnabled(bool(self._redo_stack))
        if not scope:
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

    def _update_citation_label(self):
        if self.analysis is None:
            self.citation_label.setText("")
            return
        a = self.analysis
        if not a.ref_entries and not a.citation_positions:
            self.citation_label.setText("")
            return
        parts = []
        if a.dangling_citations:
            nums = "、".join(f"[{n}]" for n in a.dangling_citations)
            parts.append(f"⚠ 悬空引用 {nums}（正文引用了不存在的条目）")
        if a.uncited_entries:
            nums = "、".join(f"[{n}]" for n in a.uncited_entries[:8])
            more = "…" if len(a.uncited_entries) > 8 else ""
            parts.append(f"⚠ {len(a.uncited_entries)} 条文献未被正文引用：{nums}{more}")
        if parts:
            self.citation_label.setText("引用检查：" + "；".join(parts))
        else:
            self.citation_label.setText(
                f"引用检查 ✓ {len(a.ref_entries)} 条文献与正文引用一一对应")
            self.citation_label.setStyleSheet("color:#1a6b45; font-size:11px;")


class _LogView(QPlainTextEdit):
    """可折叠日志：收起时显示最近一行，展开后滚动到底。"""

    def __init__(self):
        super().__init__()
        self.setReadOnly(True)
        self.setMaximumHeight(26)
        self.setPlaceholderText("操作日志…")
        self._expanded = False

    def append_line(self, text: str):
        self.appendPlainText(text)
        self.verticalScrollBar().setValue(self.verticalScrollBar().maximum())

    def expand(self):
        if not self._expanded:
            self.toggle()

    def toggle(self):
        self._expanded = not self._expanded
        self.setMaximumHeight(160 if self._expanded else 26)
        self.verticalScrollBar().setValue(self.verticalScrollBar().maximum())


# 兼容别名：旧调用仍可用 WorkbenchDialog 名字拿到主窗口类
WorkbenchDialog = ShushanWorkbench


class _PageRenderWorker(QThread):
    """后台渲染线程：四态落地临时副本 → Word 导出 PDF → 逐页 PNG。"""

    done_sig = Signal(list, float)          # (png 字节列表, 耗时秒)
    error_sig = Signal(str)

    def __init__(self, src_docx: str, workdir: str, profile: dict,
                 assignment: dict, numbered: dict, fmt_overrides: dict):
        super().__init__()
        self._src = src_docx
        self._workdir = workdir
        self._profile = profile
        self._assignment = assignment
        self._numbered = numbered
        self._fmt_overrides = fmt_overrides

    def run(self):
        import time
        t0 = time.time()
        try:
            dst = os.path.join(self._workdir, "预览.docx")
            shutil.copy2(self._src, dst)
            from ..render import apply_profile_offline
            apply_profile_offline(self._profile, dst,
                                  para_override=dict(self._assignment),
                                  numbered_map=dict(self._numbered),
                                  fmt_overrides=dict(self._fmt_overrides),
                                  make_backup=False)
            from ..render.word_pdf import render_pages_from_docx
            pngs = render_pages_from_docx(dst, dpi=110, workdir=self._workdir)
            self.done_sig.emit(pngs, time.time() - t0)
        except Exception as e:
            self.error_sig.emit(str(e))
