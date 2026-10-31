# -*- coding: utf-8 -*-
"""书山入口薄壳：主界面已与工作台归一为 ShushanWorkbench（见 workbench.py）。

main.py 仍 `from shushan.gui.panel import run`——此文件只保留启动入口。
"""
import sys

from .workbench import APP_TITLE, ShushanWorkbench

__all__ = ["APP_TITLE", "ShushanWorkbench", "run"]


def run():
    from PySide6.QtWidgets import QApplication
    app = QApplication(sys.argv)
    app.setApplicationName("书山")
    win = ShushanWorkbench()
    win.show()
    sys.exit(app.exec())
