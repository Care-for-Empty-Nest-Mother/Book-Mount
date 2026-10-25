# -*- coding: utf-8 -*-
"""把分析结果 + 逐段落归类结果，组装成格式方案 JSON（Profile）。

归类结果以「段落级 assignment」为准：{段落号: 内容类型}，其中 "ignore" 表示
该段保持原样（不进方案、整稿时跳过）。同一签名内允许个别段落改判——
例如封面行与章标题共享格式时，可以把封面行单独改判为忽略。
"""
from collections import Counter, defaultdict

from .analyzer.cluster import RE_CHAPTER
from .analyzer.effective import format_signature
from .types import IGNORE_TYPE, SCHEMA_VERSION


def auto_assignment(analysis) -> dict:
    """不做人工改判：全部采纳建议（CLI 与界面初始值）。"""
    return {p.index: p.suggestion for p in analysis.nonempty()}


def spec_to_fmt_patch(spec: dict) -> dict:
    """格式 spec（方案字段）→ 实际生效格式字典的补丁（analyzer fmt 形状）。"""
    f = {}
    if spec.get("font_cn"):
        f["font_cn"] = spec["font_cn"]
    if spec.get("font_en"):
        f["font_en"] = spec["font_en"]
    if spec.get("size_pt"):
        f["size_pt"] = spec["size_pt"]
    if spec.get("bold") is not None:
        f["bold"] = spec["bold"]
    if spec.get("italic"):
        f["italic"] = True
    if spec.get("alignment"):
        f["align"] = spec["alignment"]
    ls = spec.get("line_spacing")
    if isinstance(ls, dict) and ls.get("value"):
        f["line_mode"] = ls.get("mode")
        f["line_value"] = ls["value"]
    if spec.get("space_before_pt") is not None:
        f["before_pt"] = spec["space_before_pt"]
    if spec.get("space_after_pt") is not None:
        f["after_pt"] = spec["space_after_pt"]
    for spec_key, chars_key, cm_key in (("first_line_indent", "first_chars", "first_cm"),
                                        ("hanging_indent", "hanging_chars", "hanging_cm"),
                                        ("left_indent", "left_chars", "left_cm")):
        v = spec.get(spec_key)
        if isinstance(v, dict) and v.get("value"):
            f[chars_key if v.get("unit") == "chars" else cm_key] = v["value"]
    return f


def signature_to_spec(sig: tuple) -> dict:
    """签名（analyzer.effective.format_signature 的输出）→ styles 里的一种格式定义。"""
    (font_cn, font_en, size_pt, bold, italic, color, align,
     line_mode, line_value, before_pt, after_pt,
     first_chars, first_cm, hanging_chars, hanging_cm,
     left_chars, left_cm, outline, _numbering) = sig

    def clean(v):
        if isinstance(v, float) and v == int(v):
            return int(v)
        return v

    spec = {}
    if font_cn:
        spec["font_cn"] = font_cn
    if font_en:
        spec["font_en"] = font_en
    if size_pt:
        spec["size_pt"] = clean(size_pt)
    if bold is not None:
        spec["bold"] = bool(bold)
    if italic:
        spec["italic"] = True
    if color and color != "auto":
        spec["color"] = color
    if align and align != "left":
        spec["alignment"] = align
    if line_mode and line_value:
        spec["line_spacing"] = {"mode": line_mode, "value": clean(line_value)}
    if before_pt:
        spec["space_before_pt"] = clean(before_pt)
    if after_pt:
        spec["space_after_pt"] = clean(after_pt)
    if first_chars:
        spec["first_line_indent"] = {"unit": "chars", "value": clean(first_chars)}
    elif first_cm:
        spec["first_line_indent"] = {"unit": "cm", "value": clean(first_cm)}
    if hanging_chars:
        spec["hanging_indent"] = {"unit": "chars", "value": clean(hanging_chars)}
    elif hanging_cm:
        spec["hanging_indent"] = {"unit": "cm", "value": clean(hanging_cm)}
    if left_chars:
        spec["left_indent"] = {"unit": "chars", "value": clean(left_chars)}
    elif left_cm:
        spec["left_indent"] = {"unit": "cm", "value": clean(left_cm)}
    if outline is not None:
        spec["outline_level"] = outline
    return spec


def sig_str(sig: tuple) -> str:
    return "|".join("·" if v is None else str(v) for v in sig)


def infer_numbering(analysis, assignment: dict) -> dict | None:
    """从标题文字推断多级编号方案。"""
    h1_texts = [p.text.strip() for p in analysis.nonempty()
                if assignment.get(p.index) == "heading_1"]
    if not h1_texts:
        return None
    chinese = any(RE_CHAPTER.match(t) for t in h1_texts)
    levels = [{
        "level": 1, "style": "heading_1",
        "text": "第%1章" if chinese else "%1",
        "num_fmt": "chinese_counting" if chinese else "decimal",
        "suffix": "space",
    }]
    assigned = {t for t in assignment.values() if t != IGNORE_TYPE}
    if "heading_2" in assigned:
        levels.append({"level": 2, "style": "heading_2", "text": "%1.%2",
                       "num_fmt": "decimal", "suffix": "space"})
    if "heading_3" in assigned:
        levels.append({"level": 3, "style": "heading_3", "text": "%1.%2.%3",
                       "num_fmt": "decimal", "suffix": "space"})
    return {"heading_scheme": {"levels": levels}}


def build_profile(analysis, assignment: dict, name: str, source: str = "docx",
                  fmt_overrides: dict | None = None) -> dict:
    """assignment: {段落号: 类型} → 方案 dict。

    同一类型出现多个签名时，取段落数最多的签名作为该类型的代表格式；
    标记为 ignore 的段落不产生样式、不进 mapping。
    图片/公式段的类型由内容决定（不按签名进 mapping，避免图/公式段互相误判）。
    fmt_overrides: {段落号: 格式 spec}——批注视图"修改格式"的结果，
    参与代表格式投票（改的段落多，类型样式就按改后的来）。
    """
    fmt_overrides = fmt_overrides or {}
    type_votes = defaultdict(Counter)
    sig_majority = defaultdict(Counter)
    for p in analysis.nonempty():
        t = assignment.get(p.index)
        if not t or t == IGNORE_TYPE:
            continue
        sig = p.signature
        if p.index in fmt_overrides and p.fmt:
            merged = {**p.fmt, **spec_to_fmt_patch(fmt_overrides[p.index])}
            sig = format_signature(merged)
        type_votes[t][sig] += 1
        if not p.has_object:          # 对象段类型不进签名映射
            sig_majority[sig][t] += 1

    styles = {}
    for t, votes in type_votes.items():
        best_sig = votes.most_common(1)[0][0]
        spec = signature_to_spec(best_sig)
        if t.startswith("heading_") and t[8:].isdigit():
            spec.setdefault("outline_level", int(t[8:]) - 1)
            spec["keep_with_next"] = True
        styles[t] = spec

    styles.setdefault("body_text", {"font_cn": "宋体", "size_pt": 12})

    profile = {
        "schema_version": SCHEMA_VERSION,
        "meta": {
            "name": name,
            "source": source,
            "notes": f"提取自：{analysis.source_path}",
        },
        "page": analysis.page,
        "styles": styles,
    }
    numbering = infer_numbering(analysis, assignment)
    if numbering:
        numbering["mode"] = "reference"      # reference=套用参考方案编号；keep_target=保留目标文档编号
        profile["numbering"] = numbering

    # 表格跨页规范（从参考文档多数派提取；无表格时用规范默认值）
    profile["tables"] = dict(analysis.table_settings) if analysis.table_settings else \
        {"repeat_header": True, "cant_split_rows": True}

    # 图片尺寸参考统计（信息性：应用时不盲目改图宽，仅做超版心安全缩放）
    sizes = analysis.figure_sizes_cm
    if sizes:
        widths = sorted(w for w, _ in sizes)
        n = len(widths)
        median = widths[n // 2] if n % 2 else (widths[n // 2 - 1] + widths[n // 2]) / 2
        profile["objects"] = {
            "figure": {
                "count": len(sizes),
                "width_cm": {"median": median, "min": widths[0], "max": widths[-1]},
            }
        }
    # mapping 按签名记录"多数派类型"，供相似文档套用时跳过人工确认；ignore 不入表
    profile["mapping"] = {sig_str(sig): votes.most_common(1)[0][0]
                          for sig, votes in sig_majority.items()}
    return profile
