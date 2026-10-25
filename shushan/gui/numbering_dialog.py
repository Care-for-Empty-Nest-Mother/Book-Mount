# -*- coding: utf-8 -*-
"""标题编号模式选择 + 自定义编号样式编辑器。

整稿流程中弹出，让用户决定标题编号的来源：
- 智能判断（推荐）：目标标题自带自动编号 → 保留；文字里已有手工编号 → 不另挂；
  都没有 → 套用参考方案；
- 保留目标文档编号：只改标题级别与视觉样式（keep_target）；
- 套用参考方案编号：把方案里的编号体系（如"第X章"）套给标题（reference）；
- 自定义编号样式：现场输入各级编号格式（如 %1.%2 十进制）。

返回 (mode, scheme_levels)，传给 apply_profile_offline 的 numbering_override。
"""
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (QComboBox, QDialog, QDialogButtonBox, QFormLayout,
                               QGroupBox, QHBoxLayout, QLabel, QLineEdit,
                               QPushButton, QRadioButton, QVBoxLayout, QWidget)

from ..render.apply import MODE_LABELS

NUM_FMT_ITEMS = [("decimal", "1,2,3…"), ("chinese_counting", "一,二,三…"),
                 ("upper_roman", "Ⅰ,Ⅱ,Ⅲ…"), ("lower_roman", "ⅰ,ⅱ,ⅲ…"),
                 ("upper_letter", "A,B,C…"), ("lower_letter", "a,b,c…")]
SUFFIX_ITEMS = [("space", "空格"), ("tab", "制表符"), ("nothing", "无")]

_DEFAULT_SCHEME = [
    {"level": 1, "style": "heading_1", "text": "第%1章", "num_fmt": "chinese_counting", "suffix": "space"},
    {"level": 2, "style": "heading_2", "text": "%1.%2", "num_fmt": "decimal", "suffix": "space"},
    {"level": 3, "style": "heading_3", "text": "%1.%2.%3", "num_fmt": "decimal", "suffix": "space"},
]
_FMT_SAMPLE = {"decimal": ["1", "2", "3"], "chinese_counting": ["一", "二", "三"],
               "upper_roman": ["Ⅰ", "Ⅱ", "Ⅲ"], "lower_roman": ["ⅰ", "ⅱ", "ⅲ"],
               "upper_letter": ["A", "B", "C"], "lower_letter": ["a", "b", "c"]}


def _sample_render(text: str, level: int, num_fmt: str) -> str:
    """把编号文本模板渲染成示例（%1→本级示例数字）。"""
    out = text
    nums = _FMT_SAMPLE.get(num_fmt, _FMT_SAMPLE["decimal"])
    for k in range(level, 0, -1):
        placeholder = f"%{k}"
        if placeholder in out:
            out = out.replace(placeholder, nums[0] if k == level else nums[min(k - 1, 2)])
    return out


class CustomNumberingDialog(QDialog):
    """自定义编号样式：级别 1~3 的编号文本、数字类型与间隔。"""

    def __init__(self, scheme_levels=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("书山 · 自定义编号样式")
        self.resize(560, 300)
        base = scheme_levels or _DEFAULT_SCHEME
        self._rows = []
        lay = QVBoxLayout(self)
        lay.addWidget(QLabel("各级标题的编号格式：%1 为本级编号，%2/%3 为上级编号"
                             "（必须包含 %1）。"))
        form = QFormLayout()
        for lv in (1, 2, 3):
            base_lv = next((x for x in base if x.get("level") == lv), None)
            txt = QLineEdit(base_lv.get("text") if base_lv else
                            _DEFAULT_SCHEME[lv - 1]["text"])
            fmt_box = QComboBox()
            for k, label in NUM_FMT_ITEMS:
                fmt_box.addItem(label, k)
            cur = (base_lv or _DEFAULT_SCHEME[lv - 1]).get("num_fmt", "decimal")
            fmt_box.setCurrentIndex(max(0, [k for k, _ in NUM_FMT_ITEMS].index(cur)))
            suf_box = QComboBox()
            for k, label in SUFFIX_ITEMS:
                suf_box.addItem(label, k)
            cur_suf = (base_lv or _DEFAULT_SCHEME[lv - 1]).get("suffix", "space")
            suf_box.setCurrentIndex(max(0, [k for k, _ in SUFFIX_ITEMS].index(cur_suf)))
            row = QHBoxLayout()
            row.addWidget(txt, stretch=3)
            row.addWidget(fmt_box, stretch=2)
            row.addWidget(suf_box, stretch=1)
            wrap = QWidget()
            wrap.setLayout(row)
            form.addRow(f"级别 {lv}：", wrap)
            self._rows.append((lv, txt, fmt_box, suf_box))
        lay.addLayout(form)
        self.preview = QLabel("")
        self.preview.setStyleSheet("color:#1a6b45;")
        lay.addWidget(self.preview)
        for _, txt, fmt_box, suf_box in self._rows:
            txt.textChanged.connect(self._update_preview)
            fmt_box.currentIndexChanged.connect(self._update_preview)
            suf_box.currentIndexChanged.connect(self._update_preview)
        self._update_preview()

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("确定")
        buttons.button(QDialogButtonBox.Cancel).setText("取消")
        buttons.accepted.connect(self._on_ok)
        buttons.rejected.connect(self.reject)
        lay.addWidget(buttons)

    def _update_preview(self):
        samples = []
        for lv, txt, fmt_box, _ in self._rows:
            t = txt.text().strip()
            if t and "%1" in t:
                samples.append(_sample_render(t, lv, fmt_box.currentData()))
        self.preview.setText("预览：" + "　→　".join(samples) if samples
                             else "预览：请填写包含 %1 的编号文本")

    def _on_ok(self):
        for _, txt, _, _ in self._rows:
            t = txt.text().strip()
            if t and "%1" not in t:
                self.preview.setText(f"⚠ 编号文本 {t!r} 缺少 %1（本级编号占位符）")
                return
        self.accept()

    def scheme(self):
        out = []
        for lv, txt, fmt_box, suf_box in self._rows:
            t = txt.text().strip()
            if t:
                out.append({"level": lv, "style": f"heading_{lv}", "text": t,
                            "num_fmt": fmt_box.currentData(),
                            "suffix": suf_box.currentData()})
        return out


class NumberingModeDialog(QDialog):
    """整稿前的标题编号模式选择。result_override() → (mode, scheme_levels)。"""

    def __init__(self, detected: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle("书山 · 标题编号模式")
        self.resize(560, 360)
        self._custom_scheme = None
        self._detected = detected

        lay = QVBoxLayout(self)
        intro = QLabel("标题的编号从哪里来？（只影响编号，标题的字体字号等视觉样式"
                       "一律按方案套用）")
        intro.setWordWrap(True)
        lay.addWidget(intro)

        detected_cn = MODE_LABELS.get(detected, detected)
        hint = {"keep_target": "检测到目标文档标题自带自动编号，建议保留",
                "manual": "检测到目标文档标题文字里已含编号（如「第一章」），无需另挂",
                "reference": "目标文档标题没有编号，建议套用参考方案的编号体系"}.get(
            detected, "")
        self.r_auto = QRadioButton(f"智能判断（推荐）—— 当前检测结果：{detected_cn}")
        if hint:
            self.r_auto.setToolTip(hint)
        self.r_keep = QRadioButton("保留目标文档编号（只改标题级别与视觉样式）")
        self.r_ref = QRadioButton("套用参考方案的编号体系")
        self.r_custom = QRadioButton("自定义编号样式…（现场输入各级编号格式）")
        self.r_auto.setChecked(True)
        for r in (self.r_auto, self.r_keep, self.r_ref, self.r_custom):
            lay.addWidget(r)
            r.toggled.connect(self._update_hint)

        box = QGroupBox("说明")
        box_l = QVBoxLayout(box)
        self.hint = QLabel(hint or "智能判断会根据目标文档的现状自动选择最合适的方式。")
        self.hint.setWordWrap(True)
        box_l.addWidget(self.hint)
        lay.addWidget(box)

        edit_row = QHBoxLayout()
        edit_row.addStretch(1)
        self.btn_edit = QPushButton("编辑自定义编号…")
        self.btn_edit.setEnabled(False)
        self.btn_edit.clicked.connect(self._edit_custom)
        edit_row.addWidget(self.btn_edit)
        lay.addLayout(edit_row)
        self.custom_preview = QLabel("")
        self.custom_preview.setStyleSheet("color:#1a6b45;")
        lay.addWidget(self.custom_preview)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("开始整稿")
        buttons.button(QDialogButtonBox.Cancel).setText("取消")
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        lay.addWidget(buttons)

    def _update_hint(self):
        self.btn_edit.setEnabled(self.r_custom.isChecked())
        texts = {
            "auto": "智能判断：目标标题自带自动编号则保留；文字里已有编号则不另挂；否则套用方案。",
            "keep_target": "目标文档的编号体系一个字不动（含编号格式与样式），标题只换字体字号等版式。",
            "reference": "清除目标标题原有编号，改用方案里的编号体系（如「第一章 / 1.1 / 1.1.1」）。",
            "custom": "现场定义各级编号格式，应用到标题（清除目标原有编号）。",
        }
        if self.r_auto.isChecked():
            texts["auto"] += f" 当前检测结果：{MODE_LABELS.get(self._detected, self._detected)}。"
            if self._detected == "manual":
                texts["auto"] += " 注意：保留手工编号时标题文字里的编号就是最终显示。"
        self.hint.setText(texts.get(
            "auto" if self.r_auto.isChecked() else
            "keep_target" if self.r_keep.isChecked() else
            "reference" if self.r_ref.isChecked() else "custom", ""))

    def _edit_custom(self):
        dlg = CustomNumberingDialog(self._custom_scheme, self)
        if dlg.exec() == QDialog.Accepted and dlg.scheme():
            self._custom_scheme = dlg.scheme()
            samples = []
            for lv in self._custom_scheme:
                samples.append(_sample_render(lv["text"], lv["level"], lv["num_fmt"]))
            self.custom_preview.setText("自定义编号：" + "　→　".join(samples))

    def result_override(self):
        if self.r_keep.isChecked():
            return ("keep_target", None)
        if self.r_ref.isChecked():
            return ("reference", None)
        if self.r_custom.isChecked():
            return ("custom", self._custom_scheme)
        return ("auto", None)
