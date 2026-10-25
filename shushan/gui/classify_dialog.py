# -*- coding: utf-8 -*-
"""批注式归类确认界面（导师批注视图 v2）。

设计：
- 左侧"原文视图"逐段显示文档，每段开头带彩色标签【标题1】等，段底色即该段
  将要套用的格式颜色；鼠标悬浮在任意段落上，显示该段提取到的实际生效格式
  （字体、字号、行距、缩进、判定依据）；
- 右侧"类型分组树"：按内容类型分组（标题1/标题2/图注…），点击类型 = 整类
  选中，展开可见组内每个段落，点击段落 = 定位并单选；
- 鼠标点选段落（Shift 连选 / Ctrl 加选），选中段显示蓝色高亮；改判把选中段落
  改为任意类型（含忽略），树与视图即时同步；
- 确认后返回 {段落号: 类型} 逐段 assignment，供生成方案或整稿使用。
"""
import html as _html
from collections import Counter

from PySide6.QtGui import QColor, QFont, QTextCharFormat, QTextCursor
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (QComboBox, QDialog, QDialogButtonBox, QHBoxLayout,
                               QLabel, QPushButton, QTextEdit, QToolTip,
                               QTreeWidget, QTreeWidgetItem, QVBoxLayout, QWidget)

from ..analyzer.docx_analyzer import AnalysisResult
from ..types import CONTENT_TYPES, DEFAULT_CN_NAMES, IGNORE_TYPE
from .format_edit_dialog import FormatEditDialog, spec_to_fmt_patch

# 各类型 (段落底色, 标签底色)；底色 None = 不着色（正文保持纸面白）
TYPE_COLORS = {
    "heading_1":      ("#ffe3e3", "#c92a2a"),
    "heading_2":      ("#ffedd8", "#d9480f"),
    "heading_3":      ("#fff3bf", "#997a00"),
    "heading_4":      ("#f4fce3", "#5c940d"),
    "body_text":      (None,       "#868e96"),
    "abstract":       ("#ffdeeb", "#c2255c"),
    "keywords":       ("#ffdeeb", "#a61e4d"),
    "en_title":       ("#ffdeeb", "#c2255c"),
    "en_abstract":    ("#ffdeeb", "#c2255c"),
    "en_keywords":    ("#ffdeeb", "#a61e4d"),
    "caption_figure": ("#d0ebff", "#1971c2"),
    "caption_table":  ("#e5dbff", "#6741d9"),
    "reference":      ("#f3d9fa", "#9c36b5"),
    "ignore":         ("#e9ecef", "#495057"),
    "figure_para":    ("#c5f6fa", "#0b7285"),
    "formula_para":   ("#dbe4ff", "#3b5bdb"),
}
DEFAULT_COLOR = ("#f1f3f5", "#343a40")
SELECT_COLOR = "#a5d8ff"          # 被选中段落的蓝色高亮

# 忽略类型放在第一项：即使未来再出现初始值异常，失败模式也是"保持原样"而非误改
TYPE_ITEMS = [("ignore", "忽略（保持原样）")] + [
    (tid, DEFAULT_CN_NAMES.get(tid, tid)) for tid in CONTENT_TYPES]

ALIGN_CN = {"left": "左对齐", "center": "居中", "right": "右对齐",
            "justify": "两端对齐", "distributed": "分散对齐"}

# 默认收起的大组（段落太多，展开会淹没界面）
COLLAPSED_GROUPS = {"body_text", "ignore"}


def label_of(tid: str) -> str:
    if tid == IGNORE_TYPE:
        return "忽略"
    return DEFAULT_CN_NAMES.get(tid, tid)


def colors_of(tid: str):
    return TYPE_COLORS.get(tid, DEFAULT_COLOR)


def fmt_detail(fmt: dict | None) -> str:
    """把提取到的实际生效格式压成一行可读摘要（悬浮提示用）。"""
    if not fmt:
        return "（空段，无格式）"
    parts = []
    if fmt.get("font_cn"):
        parts.append(f"中文 {fmt['font_cn']}")
    if fmt.get("font_en"):
        parts.append(f"西文 {fmt['font_en']}")
    if fmt.get("size_pt"):
        parts.append(f"{fmt['size_pt']:g}磅")
    if fmt.get("bold") is not None:
        parts.append("加粗" if fmt["bold"] else "不加粗")
    if fmt.get("italic"):
        parts.append("斜体")
    if fmt.get("align"):
        parts.append(ALIGN_CN.get(fmt["align"], fmt["align"]))
    if fmt.get("line_mode"):
        v = fmt.get("line_value")
        name = {"multiple": f"{v:g}倍行距", "exact": f"固定值{v:g}磅",
                "at_least": f"最小值{v:g}磅"}.get(fmt["line_mode"], fmt["line_mode"])
        parts.append(name)
    if fmt.get("before_pt") is not None:
        parts.append(f"段前{fmt['before_pt']:g}磅")
    if fmt.get("after_pt") is not None:
        parts.append(f"段后{fmt['after_pt']:g}磅")
    if fmt.get("first_chars"):
        parts.append(f"首行缩进{fmt['first_chars']:g}字符")
    elif fmt.get("first_cm"):
        parts.append(f"首行缩进{fmt['first_cm']:g}cm")
    if fmt.get("hanging_chars"):
        parts.append(f"悬挂缩进{fmt['hanging_chars']:g}字符")
    elif fmt.get("hanging_cm"):
        parts.append(f"悬挂缩进{fmt['hanging_cm']:g}cm")
    if fmt.get("left_chars"):
        parts.append(f"左缩进{fmt['left_chars']:g}字符")
    elif fmt.get("left_cm"):
        parts.append(f"左缩进{fmt['left_cm']:g}cm")
    if fmt.get("outline") is not None:
        parts.append(f"大纲级别{fmt['outline'] + 1}")
    if fmt.get("numbering"):
        parts.append("带编号")
    return " ｜ ".join(parts) if parts else "（继承样式默认，无直接格式）"


class _ManuscriptView(QTextEdit):
    """只读原文视图：点击换算成段落号；悬浮显示该段提取到的格式；
    Alt+点击标题段切换"是否参与编号"。"""

    def __init__(self, on_click, on_hover, on_alt_click, on_context, on_double):
        super().__init__()
        self.setReadOnly(True)
        self.setMouseTracking(True)          # 悬浮即响应，无需按下
        self._on_click = on_click
        self._on_hover = on_hover
        self._on_alt_click = on_alt_click
        self._on_context = on_context
        self._on_double = on_double

    def mouseDoubleClickEvent(self, event):
        cursor = self.cursorForPosition(event.position().toPoint())
        if self._on_double(cursor.blockNumber(), event.globalPos()):
            return
        super().mouseDoubleClickEvent(event)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            cursor = self.cursorForPosition(event.position().toPoint())
            if event.modifiers() & Qt.AltModifier:
                self._on_alt_click(cursor.blockNumber())
            else:
                self._on_click(cursor.blockNumber(), event.modifiers())
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        cursor = self.cursorForPosition(event.position().toPoint())
        self._on_hover(cursor.blockNumber(), event.globalPosition().toPoint())
        super().mouseMoveEvent(event)

    def contextMenuEvent(self, event):
        cursor = self.cursorForPosition(event.pos())
        self._on_context(cursor.blockNumber(), event.globalPos())



class ClassifyDialog(QDialog):
    """批注式归类确认。用法：

        dlg = ClassifyDialog(analysis, 文档名, parent)
        if dlg.exec() == QDialog.Accepted:
            assignment = dlg.assignment     # {段落号: 类型}
    """

    def __init__(self, analysis: AnalysisResult, doc_name: str, parent=None):
        super().__init__(parent)
        self.analysis = analysis
        self.setWindowTitle(f"书山 · 归类批注 —— {doc_name}")
        self.resize(1150, 720)
        self.paras = analysis.paragraphs
        self.suggestion = {p.index: p.suggestion for p in analysis.nonempty()}
        self.assignment: dict = dict(self.suggestion)
        self.selected: set = set()
        self._anchor: int | None = None
        self._nonempty_idx = {p.index for p in analysis.nonempty()}
        self._hover_block = None
        self._tree_syncing = False
        # 标题逐段"是否参与编号"（对号），默认全勾；非标题段不出现在此表
        self.numbered: dict = {p.index: True for p in analysis.nonempty()
                               if p.suggestion.startswith("heading_")}
        # 段落级格式覆盖 {段落号: 格式 spec}——右键"修改格式"的编辑结果
        self.fmt_overrides: dict = {}

        self._build_ui()
        self._rebuild_html()
        self._rebuild_tree()
        self._refresh_highlights()
        self._sync_selection_label()
        self._update_citation_label()

    # ---------------- UI 构建 ----------------
    def _build_ui(self):
        root = QVBoxLayout(self)
        tip = QLabel("颜色即将要套用的格式，悬浮可查看该段提取到的实际格式；"
                     "左键点选段落（Shift 连选 / Ctrl 加选），右侧按类型分组："
                     "点类型整类选中，点段落定位，改判即时同步。")
        tip.setWordWrap(True)
        root.addWidget(tip)

        body = QHBoxLayout()
        self.view = _ManuscriptView(self._on_click, self._on_hover, self._on_alt_click,
                                    self._context_menu, self._on_double)
        self.view.setFont(QFont("Microsoft YaHei", 10))
        body.addWidget(self.view, stretch=3)

        side = QWidget()
        side_l = QVBoxLayout(side)
        side.setFixedWidth(340)

        side_l.addWidget(QLabel("类型分组（点类型=整类选中，点段落=定位）："))
        self.tree = QTreeWidget()
        self.tree.setHeaderHidden(True)
        self.tree.setUniformRowHeights(True)
        self.tree.itemClicked.connect(self._on_tree_click)
        self.tree.itemChanged.connect(self._on_tree_check)
        side_l.addWidget(self.tree, stretch=1)

        self.citation_label = QLabel("")
        self.citation_label.setWordWrap(True)
        self.citation_label.setStyleSheet("color:#8a6d1a; font-size:11px;")
        side_l.addWidget(self.citation_label)

        self.sel_label = QLabel("已选中 0 段")
        side_l.addWidget(self.sel_label)
        re_row = QHBoxLayout()
        self.reassign_combo = QComboBox()
        for tid, label in TYPE_ITEMS:
            self.reassign_combo.addItem(label, tid)
        re_row.addWidget(self.reassign_combo, stretch=1)
        btn_apply = QPushButton("改判")
        btn_apply.clicked.connect(self._on_reassign)
        re_row.addWidget(btn_apply)
        side_l.addLayout(re_row)
        body.addWidget(side)
        root.addLayout(body, stretch=1)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("确认归类")
        buttons.button(QDialogButtonBox.Cancel).setText("取消")
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        btn_reset = QPushButton("全部采纳建议")
        btn_reset.clicked.connect(self._on_reset)
        buttons.addButton(btn_reset, QDialogButtonBox.ActionRole)
        root.addWidget(buttons)

    # ---------------- 渲染 ----------------
    def _rebuild_html(self):
        parts = []
        for p in self.paras:
            if p.fmt is None:
                parts.append('<p style="color:#ced4da; font-size:8pt">⸱</p>')
                continue
            tid = self.assignment.get(p.index, p.suggestion)
            _bg, badge = colors_of(tid)
            mark = ""
            if tid.startswith("heading_"):
                if self.numbered.get(p.index, True):
                    mark = f"☑ {label_of(tid)}"
                else:
                    mark = f"☐ 仅大纲{tid[8:]}"
            badge_html = (f'<span style="background-color:{badge}; color:#ffffff; '
                          f'font-size:8pt">&nbsp;{mark}&nbsp;</span>&nbsp;')
            if not p.text.strip():
                hint = "（公式）" if p.object_kind == "formula" else "（图片）"
                text_html = f'<span style="color:#74a3b3">{hint}</span>'
            else:
                text_html = _html.escape(p.text.strip())
                if tid == IGNORE_TYPE:
                    text_html = f'<span style="color:#adb5bd">{text_html}</span>'
            parts.append(f"<p>{badge_html}{text_html}</p>")
        self.view.setHtml("".join(parts))

    def _refresh_highlights(self):
        doc = self.view.document()
        sels = []
        for p in self.analysis.nonempty():
            tid = self.assignment.get(p.index, p.suggestion)
            bg = SELECT_COLOR if p.index in self.selected else colors_of(tid)[0]
            if not bg:
                continue
            block = doc.findBlockByNumber(p.index)
            if not block.isValid():
                continue
            cur = QTextCursor(block)
            cur.movePosition(QTextCursor.EndOfBlock, QTextCursor.KeepAnchor)
            sel = QTextEdit.ExtraSelection()
            sel.cursor = cur
            fmt = QTextCharFormat()
            fmt.setBackground(QColor(bg))
            sel.format = fmt
            sels.append(sel)
        self.view.setExtraSelections(sels)

    def _rebuild_tree(self):
        """按内容类型分组建树：顶层=类型（含段数），子级=段落。"""
        self._tree_syncing = True
        self.tree.clear()
        groups = {}
        for p in self.analysis.nonempty():
            tid = self.assignment.get(p.index, p.suggestion)
            groups.setdefault(tid, []).append(p)
        for tid, label in TYPE_ITEMS:
            plist = groups.get(tid)
            if not plist:
                continue
            bg, badge = colors_of(tid)
            if tid.startswith("heading_"):
                top_text = f"大纲级别 {tid[8:]}   ×{len(plist)}"
            else:
                top_text = f"{label}   ×{len(plist)}"
            top = QTreeWidgetItem([top_text])
            top.setData(0, Qt.UserRole, tid)
            top.setData(0, Qt.UserRole + 1, "type")
            top.setBackground(0, QColor(bg))
            top.setForeground(0, QColor(badge))
            for p in plist:
                mark = "✎ " if self.assignment.get(p.index) != p.suggestion else ""
                if p.index in self.fmt_overrides:
                    mark += "✍ "
                preview = p.text.strip()[:30] or (
                    "（公式）" if p.object_kind == "formula" else "（图片）")
                child = QTreeWidgetItem([f"{mark}P{p.index} · {preview}"])
                child.setData(0, Qt.UserRole, tid)
                child.setData(0, Qt.UserRole + 1, "para")
                child.setData(0, Qt.UserRole + 2, p.index)
                child.setBackground(0, QColor(bg))
                numbered_tip = ""
                if tid.startswith("heading_"):
                    child.setFlags(child.flags() | Qt.ItemIsUserCheckable)
                    child.setCheckState(0, Qt.Checked if self.numbered.get(
                        p.index, True) else Qt.Unchecked)
                    numbered_tip = ("☑ 勾选：套用标题样式并参与编号"
                                    if self.numbered.get(p.index, True)
                                    else "☐ 未勾：只设定大纲级别（保留原样式与外观，无编号字样）")
                else:
                    # QTreeWidgetItem 默认带 UserCheckable，显式去掉避免出现空复选框
                    child.setFlags(child.flags() & ~Qt.ItemIsUserCheckable)
                child.setToolTip(0, f"P{p.index}\n判定：{label_of(tid)}"
                                     f"（建议：{label_of(p.suggestion)}）\n{numbered_tip}"
                                     f"\n{fmt_detail(p.fmt)}")
                top.addChild(child)
            self.tree.addTopLevelItem(top)
            top.setExpanded(tid not in COLLAPSED_GROUPS)
        self._tree_syncing = False

    def _effective_fmt(self, p):
        """段落当前的实际格式 = 提取格式 + 手动覆盖。"""
        base = dict(p.fmt or {})
        ov = self.fmt_overrides.get(p.index)
        if ov:
            base = {**base, **spec_to_fmt_patch(ov)}
        return base

    def _sync_selection_label(self):
        n = len(self.selected)
        self.sel_label.setText(f"已选中 {n} 段" + ("" if n else "（点击左侧正文选择）"))
        if n:
            majority = Counter(self.assignment.get(i) for i in self.selected).most_common(1)[0][0]
            pos = next((i for i, (tid, _) in enumerate(TYPE_ITEMS) if tid == majority), 0)
            self.reassign_combo.setCurrentIndex(pos)

    # ---------------- 交互 ----------------
    def _context_menu(self, block_no, global_pos):
        """原文视图右键：修改格式（作用于选中的一组段落）。"""
        if block_no < 0 or block_no >= len(self.paras):
            return
        p = self.paras[block_no]
        if p.fmt is None:
            return
        from PySide6.QtWidgets import QMenu
        menu = QMenu(self)
        if block_no in self.selected and self.selected:
            scope = set(self.selected)
            scope_label = f"修改格式…（{len(scope)} 段）"
        else:
            scope = {p.index}
            self.selected = scope
            self._anchor = p.index
            self._refresh_highlights()
            self._sync_selection_label()
            scope_label = "修改格式…（本段）"
        act_edit = menu.addAction(scope_label)
        act_jump = None
        if p.citations or p.index in self.analysis.ref_entries:
            act_jump = menu.addAction("定位引用的文献 / 引用位置…")
        menu.addSeparator()
        act_clear = menu.addAction("清除本段的手动格式修改")
        act_clear.setEnabled(any(i in self.fmt_overrides for i in scope))
        chosen = menu.exec(global_pos)
        if chosen is act_jump:
            self._citation_jump_menu(p, global_pos)
        elif chosen is act_edit:
            pre = self._effective_fmt(p)
            dlg = FormatEditDialog(pre, len(scope), self)
            if dlg.exec() == QDialog.Accepted:
                spec = dlg.spec()
                for i in scope:
                    self.fmt_overrides[i] = spec
                self._rebuild_html()
                self._rebuild_tree()
                self._refresh_highlights()
        elif chosen is act_clear:
            for i in scope:
                self.fmt_overrides.pop(i, None)
            self._rebuild_html()
            self._rebuild_tree()
            self._refresh_highlights()

    def _citation_targets(self, p):
        """段落 p 的引用跳转目标：[(文献序号, 条目段落号或 None), ...]。"""
        entry_by_num = {num: idx for idx, num in self.analysis.ref_entries.items()}
        return [(n, entry_by_num.get(n)) for n in p.citations]

    def _scroll_block_to_middle(self, block):
        """把目标段滚动到可视区中部（默认 setTextCursor 会顶到边缘，
        上下文看不到，不方便对照修改）。用文档坐标直接计算滚动位置，
        不依赖当前滚动状态，确定宜居中。"""
        cursor = QTextCursor(block)
        self.view.setTextCursor(cursor)
        rect_abs = self.view.document().documentLayout().blockBoundingRect(block)
        viewport_h = self.view.viewport().height()
        target_top = max(0, int(rect_abs.center().y() - viewport_h / 2))
        self.view.verticalScrollBar().setValue(target_top)

    def _jump_to(self, idx):
        self.selected = {idx}
        self._anchor = idx
        block = self.view.document().findBlockByNumber(idx)
        if block.isValid():
            self._scroll_block_to_middle(block)
        self._refresh_highlights()
        self._sync_selection_label()
        self._sync_tree_current()

    def _citation_jump_menu(self, p, global_pos):
        """双击/菜单触发：引用↔文献互跳。返回 True 表示已处理。"""
        from PySide6.QtWidgets import QMenu
        entry_by_num = {num: idx for idx, num in self.analysis.ref_entries.items()}
        if p.citations:
            menu = QMenu(self)
            menu.addAction("跳转到被引用的文献：").setEnabled(False)
            valid = []
            for n in p.citations:
                idx = entry_by_num.get(n)
                if idx is None:
                    menu.addAction(f"文献 [{n}]（⚠ 无对应条目）").setEnabled(False)
                else:
                    valid.append((n, idx))
                    target = self.paras[idx]
                    preview = target.text.strip()[:22]
                    menu.addAction(f"→ 文献 [{n}] {preview}",
                                   lambda idx=idx: self._jump_to(idx))
            menu.exec(global_pos)
            return True
        if p.index in self.analysis.ref_entries:
            num = self.analysis.ref_entries[p.index]
            positions = self.analysis.citation_positions.get(num, [])
            menu = QMenu(self)
            if not positions:
                menu.addAction("⚠ 该文献未被正文引用").setEnabled(False)
            else:
                menu.addAction(f"跳转到引用 [{num}] 的正文位置：").setEnabled(False)
                for idx in positions:
                    preview = self.paras[idx].text.strip()[:22]
                    menu.addAction(f"→ P{idx} {preview}",
                                   lambda idx=idx: self._jump_to(idx))
            menu.exec(global_pos)
            return True
        return False

    def _on_double(self, block_no, global_pos):
        if block_no < 0 or block_no >= len(self.paras):
            return False
        p = self.paras[block_no]
        if p.fmt is None:
            return False
        return self._citation_jump_menu(p, global_pos)

    def _update_citation_label(self):
        a = self.analysis
        n_entries = len(a.ref_entries)
        if not n_entries and not a.citation_positions:
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
            self.citation_label.setText(f"引用检查 ✓ {n_entries} 条文献与正文引用一一对应")
            self.citation_label.setStyleSheet("color:#1a6b45; font-size:11px;")

    def _on_click(self, block_no, modifiers):
        if block_no < 0 or block_no >= len(self.paras):
            return
        p = self.paras[block_no]
        if p.fmt is None:
            return
        idx = p.index
        if modifiers & Qt.ControlModifier:
            self.selected.symmetric_difference_update({idx})
            self._anchor = idx
        elif modifiers & Qt.ShiftModifier and self._anchor is not None:
            lo, hi = sorted((self._anchor, idx))
            self.selected = {i for i in range(lo, hi + 1) if i in self._nonempty_idx}
        else:
            self.selected = {idx}
            self._anchor = idx
        self._refresh_highlights()
        self._sync_selection_label()
        self._sync_tree_current()

    def _on_hover(self, block_no, global_pos):
        if block_no == self._hover_block:
            return
        self._hover_block = block_no
        if block_no < 0 or block_no >= len(self.paras):
            QToolTip.hideText()
            return
        p = self.paras[block_no]
        if p.fmt is None:
            QToolTip.hideText()
            return
        tid = self.assignment.get(p.index, p.suggestion)
        changed = "" if tid == p.suggestion else "（已改判）"
        merged = self._effective_fmt(p)
        if p.index in self.fmt_overrides:
            changed += " ✍已修改格式"
        text = (f"P{p.index} · 判定 {label_of(tid)}{changed}"
                f"（建议 {label_of(p.suggestion)}）\n{fmt_detail(p.fmt)}\n依据：{p.reason}")
        QToolTip.showText(global_pos, text, self.view)

    def _sync_tree_current(self):
        """单选时让树定位到对应段落（不触发点击递归）。"""
        if self._tree_syncing or len(self.selected) != 1:
            return
        idx = next(iter(self.selected))
        for i in range(self.tree.topLevelItemCount()):
            top = self.tree.topLevelItem(i)
            for j in range(top.childCount()):
                child = top.child(j)
                if child.data(0, Qt.UserRole + 2) == idx:
                    self._tree_syncing = True
                    self.tree.setCurrentItem(child)
                    self._tree_syncing = False
                    return

    def _on_tree_check(self, item, _col):
        """树中标题子节点的对号被点击 → 更新编号状态并同步原文视图。"""
        if self._tree_syncing or item.data(0, Qt.UserRole + 1) != "para":
            return
        tid = item.data(0, Qt.UserRole)
        if not tid.startswith("heading_"):
            return
        idx = item.data(0, Qt.UserRole + 2)
        self.numbered[idx] = item.checkState(0) == Qt.Checked
        self._rebuild_html()
        self._refresh_highlights()

    def _toggle_numbering(self, idx):
        self.numbered[idx] = not self.numbered.get(idx, True)
        self._rebuild_html()
        self._rebuild_tree()
        self._refresh_highlights()

    def _on_alt_click(self, block_no):
        if block_no < 0 or block_no >= len(self.paras):
            return
        p = self.paras[block_no]
        if p.fmt is None or not p.suggestion.startswith("heading_"):
            tid_now = self.assignment.get(block_no, "")
            if not tid_now.startswith("heading_"):
                return
        self._toggle_numbering(self.paras[block_no].index)

    def _on_tree_click(self, item, _col):
        kind = item.data(0, Qt.UserRole + 1)
        if kind == "type":
            tid = item.data(0, Qt.UserRole)
            self.selected = {p.index for p in self.analysis.nonempty()
                             if self.assignment.get(p.index, p.suggestion) == tid}
            self._anchor = min(self.selected) if self.selected else None
        else:                                   # para
            idx = item.data(0, Qt.UserRole + 2)
            self.selected = {idx}
            self._anchor = idx
            block = self.view.document().findBlockByNumber(idx)
            if block.isValid():
                self._scroll_block_to_middle(block)
        self._refresh_highlights()
        self._sync_selection_label()

    def _on_reassign(self):
        if not self.selected:
            return
        tid = self.reassign_combo.currentData()
        for i in self.selected:
            self.assignment[i] = tid
            if tid.startswith("heading_"):
                self.numbered.setdefault(i, True)   # 新改判进来的标题默认参与编号
            else:
                self.numbered.pop(i, None)          # 非标题段无编号概念
        self._rebuild_html()
        self._rebuild_tree()
        self._refresh_highlights()
        self._sync_selection_label()

    def _on_reset(self):
        self.assignment = dict(self.suggestion)
        self.fmt_overrides.clear()
        self.numbered = {p.index: True for p in self.analysis.nonempty()
                         if p.suggestion.startswith("heading_")}
        self._rebuild_html()
        self._rebuild_tree()
        self._refresh_highlights()
        self._sync_selection_label()
