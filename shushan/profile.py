# -*- coding: utf-8 -*-
"""格式方案（Profile）的读写与轻量数据模型。"""
import json
from pathlib import Path
from .validator import validate_profile
from .types import SCHEMA_VERSION


class ProfileError(Exception):
    pass


def load_profile(path) -> dict:
    """读取并校验方案文件；不合法时抛 ProfileError，错误信息已排版。"""
    path = Path(path)
    try:
        data = json.loads(path.read_text(encoding="utf-8-sig"))
    except json.JSONDecodeError as e:
        raise ProfileError(f"不是有效的 JSON 文件：{e}\n（文件：{path.name}）")
    except UnicodeDecodeError:
        raise ProfileError(f"文件编码不是 UTF-8：{path.name}")
    errors, warnings = validate_profile(data)
    if errors:
        raise ProfileError("方案文件校验未通过：\n  - " + "\n  - ".join(errors))
    return data


def save_profile(data: dict, path):
    path = Path(path)
    data.setdefault("schema_version", SCHEMA_VERSION)
    errors, _ = validate_profile(data)
    if errors:
        raise ProfileError("方案不合法，拒绝保存：\n  - " + "\n  - ".join(errors))
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def validate_or_report(path) -> str:
    """命令行用：返回人类可读的校验报告。"""
    path = Path(path)
    try:
        data = json.loads(path.read_text(encoding="utf-8-sig"))
    except json.JSONDecodeError as e:
        return f"[错误] 不是有效的 JSON：{e}"
    errors, warnings = validate_profile(data)
    lines = []
    for e in errors:
        lines.append(f"  [错误] {e}")
    for w in warnings:
        lines.append(f"  [警告] {w}")
    name = data.get("meta", {}).get("name", "?") if isinstance(data, dict) else "?"
    head = f"[通过] {path.name} —— {name}"
    if errors:
        head = f"[未通过] {path.name}（{len(errors)} 处错误）"
    elif warnings:
        head += f"（{len(warnings)} 条警告）"
    return head + ("\n" + "\n".join(lines) if lines else "")
