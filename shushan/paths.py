# -*- coding: utf-8 -*-
"""路径约定。冻结为 exe 时，方案库/示例随 exe 所在目录走。"""
import sys
from pathlib import Path

if getattr(sys, "frozen", False):
    APP_DIR = Path(sys.executable).resolve().parent
else:
    APP_DIR = Path(__file__).resolve().parent.parent   # 书山/
PROFILES_DIR = APP_DIR / "profiles"
EXAMPLES_DIR = APP_DIR / "examples"

PROFILES_DIR.mkdir(exist_ok=True)
