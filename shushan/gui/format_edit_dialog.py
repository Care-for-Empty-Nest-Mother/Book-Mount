# -*- coding: utf-8 -*-
"""段落格式编辑器：把悬浮显示的"实际生效格式"变成可编辑表单。

确定后返回一份格式 spec（与方案 styles 里同构的字段子集），作为选中段落的
格式覆盖：整稿时以直接格式落地（优先于样式），提取时参与该类型代表格式的投票。
"""
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (QCheckBox, QComboBox, QDialog, QDialogButtonBox,
                               QDoubleSpinBox, QFormLayout, QGroupBox, QHBoxLayout,
                               QLabel, QLineEdit, QVBoxLayout)

CN_FONTS = ["宋体", "黑体", "楷体", "仿宋", "仿宋_GB2312", "微软雅黑", "方正小标宋简体"]
EN_FONTS = ["Times New Roman", "Arial", "Calibri", "Cambria", "SimSun"]
ALIGN_ITEMS = [("left", "左对齐"), ("center", "居中"), ("right", "右对齐"),
               ("justify", "两端对齐"), ("distributed", "分散对齐")]
LINE_MODE_ITEMS = [("multiple", "倍数行距"), ("exact", "固定值(磅)"), ("at_least", "最小值(磅)")]


def _set_combo(combo, data_key):
    idx = max(0, [k for k, _ in [(combo.itemData(i), 0) for i in range(combo.count())]]
              .index(data_key)) if data_key in [combo.itemData(i) for i in range(combo.count())] else 0
    combo.setCurrentIndex(idx)


class FormatEditDialog(QDialog):
    """编辑选中段落（组）的格式；pre_fmt 为当前实际生效格式（悬浮所显）。"""

    def __init__(self, pre_fmt: dict, count: int, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"书山 · 修改格式（{count} 段）")
        self.resize(480, 460)
        lay = QVBoxLayout(self)
        lay.addWidget(QLabel("以下为该区域当前的实际生效格式，修改后确定即按新格式套用"
                             "（整稿时以此直接格式优先于样式）。"))

        form = QFormLayout()
        self.font_cn = QComboBox()
        self.font_cn.setEditable(True)
        self.font_cn.addItems(CN_FONTS)
        self.font_cn.setCurrentText(pre_fmt.get("font_cn") or "宋体")
        form.addRow("中文字体：", self.font_cn)
        self.font_en = QComboBox()
        self.font_en.setEditable(True)
        self.font_en.addItems(EN_FONTS)
        self.font_en.setCurrentText(pre_fmt.get("font_en") or "Times New Roman")
        form.addRow("西文字体：", self.font_en)

        self.size = QDoubleSpinBox()
        self.size.setRange(1, 200)
        self.size.setDecimals(1)
        self.size.setValue(float(pre_fmt.get("size_pt") or 12))
        form.addRow("字号（磅）：", self.size)

        chk_row = QHBoxLayout()
        self.bold = QCheckBox("加粗")
        self.bold.setChecked(bool(pre_fmt.get("bold")))
        self.italic = QCheckBox("斜体")
        self.italic.setChecked(bool(pre_fmt.get("italic")))
        chk_row.addWidget(self.bold)
        chk_row.addWidget(self.italic)
        chk_row.addStretch(1)
        form.addRow("字形：", chk_row)

        self.align = QComboBox()
        for k, label in ALIGN_ITEMS:
            self.align.addItem(label, k)
        _set_combo(self.align, pre_fmt.get("align") or "left")
        form.addRow("对齐：", self.align)

        line_box = QGroupBox("行距")
        line_l = QHBoxLayout(line_box)
        self.line_mode = QComboBox()
        for k, label in LINE_MODE_ITEMS:
            self.line_mode.addItem(label, k)
        _set_combo(self.line_mode, pre_fmt.get("line_mode") or "multiple")
        self.line_value = QDoubleSpinBox()
        self.line_value.setRange(0.5, 200)
        self.line_value.setDecimals(2)
        self.line_value.setValue(float(pre_fmt.get("line_value") or 1.5))
        line_l.addWidget(self.line_mode, stretch=2)
        line_l.addWidget(self.line_value, stretch=1)
        form.addRow("行距：", line_box)

        self.before = QDoubleSpinBox()
        self.before.setRange(0, 200)
        self.before.setDecimals(1)
        self.before.setValue(float(pre_fmt.get("before_pt") or 0))
        self.after = QDoubleSpinBox()
        self.after.setRange(0, 200)
        self.after.setDecimals(1)
        self.after.setValue(float(pre_fmt.get("after_pt") or 0))
        form.addRow("段前（磅）：", self.before)
        form.addRow("段后（磅）：", self.after)

        ind_box = QGroupBox("缩进（字符）")
        ind_l = QFormLayout(ind_box)
        self.first_indent = QDoubleSpinBox()
        self.first_indent.setRange(-10, 20)
        self.first_indent.setDecimals(1)
        self.first_indent.setSingleStep(0.5)
        self.first_indent.setValue(float(pre_fmt.get("first_chars") or 0))
        self.hanging = QDoubleSpinBox()
        self.hanging.setRange(0, 20)
        self.hanging.setDecimals(1)
        self.hanging.setSingleStep(0.5)
        self.hanging.setValue(float(pre_fmt.get("hanging_chars") or 0))
        self.left_indent = QDoubleSpinBox()
        self.left_indent.setRange(0, 30)
        self.left_indent.setDecimals(1)
        self.left_indent.setSingleStep(0.5)
        self.left_indent.setValue(float(pre_fmt.get("left_chars") or 0))
        ind_l.addRow("首行缩进：", self.first_indent)
        ind_l.addRow("悬挂缩进：", self.hanging)
        ind_l.addRow("左缩进：", self.left_indent)
        form.addRow("缩进：", ind_box)
        lay.addLayout(form)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("按此格式套用")
        buttons.button(QDialogButtonBox.Cancel).setText("取消")
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        lay.addWidget(buttons)

    def spec(self) -> dict:
        """编辑结果 → 格式 spec（与方案 styles 字段同构）。"""
        ls_mode = self.line_mode.currentData()
        ls_value = round(self.line_value.value(), 2)
        first = round(self.first_indent.value(), 2)
        hanging = round(self.hanging.value(), 2)
        left = round(self.left_indent.value(), 2)
        spec = {
            "font_cn": self.font_cn.currentText().strip(),
            "font_en": self.font_en.currentText().strip(),
            "size_pt": round(self.size.value(), 1),
            "bold": self.bold.isChecked(),
            "italic": self.italic.isChecked(),
            "alignment": self.align.currentData(),
            "line_spacing": {"mode": ls_mode,
                             "value": ls_value if ls_mode == "multiple" else ls_value},
            "space_before_pt": round(self.before.value(), 1),
            "space_after_pt": round(self.after.value(), 1),
        }
        if first:
            spec["first_line_indent"] = {"unit": "chars", "value": first}
        if hanging:
            spec["hanging_indent"] = {"unit": "chars", "value": hanging}
        if left:
            spec["left_indent"] = {"unit": "chars", "value": left}
        return spec
