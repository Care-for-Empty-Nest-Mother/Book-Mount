# -*- coding: utf-8 -*-
"""路径约定。冻结为 exe 时，方案库/示例随 exe 所在目录走。"""
import os
import sys
from pathlib import Path

if getattr(sys, "frozen", False):
    APP_DIR = Path(sys.executable).resolve().parent
else:
    APP_DIR = Path(__file__).resolve().parent.parent   # 书山/
PROFILES_DIR = APP_DIR / "profiles"
EXAMPLES_DIR = APP_DIR / "examples"


def _ensure_profiles_dir():
    """确保方案库目录可用。

    历史打包缺陷曾把 profiles 写成"同名文件"而非目录，会让启动直接崩溃
    （WinError 183）。这里自动修复：把异常文件挪到备份名，再建真目录。
    任何失败都不阻断启动（方案库退化为 exe 所在目录）。"""
    global PROFILES_DIR
    try:
        if PROFILES_DIR.exists() and not PROFILES_DIR.is_dir():
            os.rename(PROFILES_DIR, APP_DIR / "profiles.异常备份")
        PROFILES_DIR.mkdir(parents=True, exist_ok=True)
    except Exception:
        PROFILES_DIR = APP_DIR


_ensure_profiles_dir()
