# -*- coding: utf-8 -*-
"""写操作前的自动备份。"""
import shutil
from datetime import datetime
from pathlib import Path

BACKUP_DIR_NAME = "书山备份"


def backup_file(docx_path) -> Path:
    """把文档复制到同目录 书山备份/ 下，带时间戳。返回备份路径。"""
    src = Path(docx_path)
    if not src.exists():
        raise FileNotFoundError(src)
    dest_dir = src.parent / BACKUP_DIR_NAME
    dest_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    dest = dest_dir / f"{src.stem}.{stamp}{src.suffix}"
    shutil.copy2(src, dest)
    return dest
