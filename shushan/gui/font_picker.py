# -*- coding: utf-8 -*-
"""系统字体枚举：Word 的字体列表 = 系统字体列表。

fill_font_combo 把"论文常用"字体置顶（且每项用该字体自身字形渲染，
同 Word 字体下拉的标志性细节），分隔线之后是其余全部系统字体。
供工作台功能区与 FormatEditDialog 共用。
"""
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont, QFontDatabase

# 论文常用字体（按顺序置顶；实际安装的才显示）
PAPER_FONTS = ["宋体", "黑体", "楷体", "仿宋", "仿宋_GB2312", "微软雅黑",
               "方正小标宋简体", "Times New Roman", "Cambria", "Arial"]

# 中文字号 → 磅（来自 types.CN_FONT_SIZES 的展示顺序），追加常用数值磅值
SIZE_ITEMS = [("初号", 42), ("小初", 36), ("一号", 26), ("小一", 24),
              ("二号", 22), ("小二", 18), ("三号", 16), ("小三", 15),
              ("四号", 14), ("小四", 12), ("五号", 10.5), ("小五", 9),
              ("六号", 7.5), ("小六", 6.5)]
EXTRA_SIZES = [8, 10, 11, 13, 20, 24, 28, 36, 48, 72]


def all_families() -> list[str]:
    """全部系统字体族（Qt6 静态方法；旧式实例法兜底）。"""
    try:
        return list(QFontDatabase.families())
    except TypeError:                       # 兼容需要实例化的旧绑定
        return list(QFontDatabase().families())


def fill_font_combo(combo, editable: bool = True):
    """字体下拉：论文常用置顶（自身字形渲染）+ 分隔线 + 其余系统字体。"""
    combo.clear()
    combo.setEditable(editable)
    installed = set(all_families())
    common = [f for f in PAPER_FONTS if f in installed]
    for name in common:
        combo.addItem(name)
        combo.setItemData(combo.count() - 1, QFont(name, 11), Qt.FontRole)
    if installed - set(common):
        combo.insertSeparator(combo.count())
        for name in sorted(installed - set(common)):
            combo.addItem(name)
            combo.setItemData(combo.count() - 1, QFont(name, 11), Qt.FontRole)
    if not combo.count():                   # 极端环境无字体：给可编辑空项
        combo.addItem("宋体")
    # 默认选常用字体（当前文本由调用方按 pre_fmt setCurrentText 覆盖）


def fill_size_combo(combo):
    """字号下拉：中文字号 + 常用数值磅值，data=磅值(float)。"""
    combo.clear()
    combo.setEditable(True)
    seen = set()
    for name, pt in SIZE_ITEMS:
        combo.addItem(f"{name}　{pt:g}", float(pt))
        seen.add(pt)
    combo.insertSeparator(combo.count())
    for pt in EXTRA_SIZES:
        if pt in seen:
            continue
        combo.addItem(f"{pt:g}", float(pt))


def select_size(combo, pt: float | None):
    """按磅值选中字号项；无匹配时写入可编辑文本。"""
    if pt is None:
        return
    for i in range(combo.count()):
        if combo.itemData(i) is not None and abs(float(combo.itemData(i)) - pt) < 0.26:
            combo.setCurrentIndex(i)
            return
    combo.setEditText(f"{pt:g}")
