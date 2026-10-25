# -*- coding: utf-8 -*-
"""格式方案 JSON 校验器（M0）。

不依赖 jsonschema，手工校验以给出中文、可定位到字段的报错。
用法：
    from shushan.validator import validate_profile
    errors, warnings = validate_profile(data)
"""
from .types import CONTENT_TYPES, ALIGNMENTS, NUM_FMTS, SCHEMA_VERSION

ALLOWED_STYLE_KEYS = {
    "display_name", "font_cn", "font_en", "size_pt", "bold", "italic",
    "color", "alignment", "line_spacing", "space_before_pt", "space_after_pt",
    "first_line_indent", "hanging_indent", "left_indent", "right_indent",
    "outline_level", "keep_with_next", "page_break_before", "base_on",
}
ALLOWED_TOP_KEYS = {"schema_version", "meta", "page", "styles", "numbering",
                    "header_footer", "toc", "caption", "skeleton", "mapping",
                    "objects", "tables", "citations"}


def _err(errors, path, msg):
    errors.append(f"{path}: {msg}" if path else msg)


def _warn(warnings, path, msg):
    warnings.append(f"{path}: {msg}" if path else msg)


def _check_indent(value, path, errors):
    if not isinstance(value, dict):
        _err(errors, path, "应为对象 {\"unit\": \"chars\"|\"cm\", \"value\": 数值}")
        return
    unit = value.get("unit")
    if unit not in ("chars", "cm"):
        _err(errors, f"{path}.unit", f'应为 "chars" 或 "cm"，当前为 {unit!r}')
    v = value.get("value")
    if not isinstance(v, (int, float)) or isinstance(v, bool) or v < -20 or v > 50:
        _err(errors, f"{path}.value", "应为数值（chars: 通常 0~10；cm: 通常 0~10）")


def _check_line_spacing(value, path, errors):
    if not isinstance(value, dict):
        _err(errors, path, '应为对象 {"mode": "multiple"|"exact"|"at_least", "value": 数值}')
        return
    mode = value.get("mode")
    if mode not in ("multiple", "exact", "at_least"):
        _err(errors, f"{path}.mode", f'应为 "multiple"/"exact"/"at_least"，当前为 {mode!r}')
        return
    v = value.get("value")
    if not isinstance(v, (int, float)) or isinstance(v, bool) or v <= 0:
        _err(errors, f"{path}.value", "应为正数")
        return
    if mode == "multiple" and not (0.5 <= v <= 5):
        _err(errors, f"{path}.value", f"倍数行距应在 0.5~5 之间，当前为 {v}")


def _check_style_spec(type_id, spec, errors, warnings):
    path = f"styles.{type_id}"
    if not isinstance(spec, dict):
        _err(errors, path, "应为对象")
        return
    for k in spec:
        if k not in ALLOWED_STYLE_KEYS:
            _warn(warnings, f"{path}.{k}", "未知字段，将被忽略")
    size = spec.get("size_pt")
    if size is not None and (not isinstance(size, (int, float)) or isinstance(size, bool)
                             or not (1 <= size <= 200)):
        _err(errors, f"{path}.size_pt", f"应为 1~200 之间的磅值，当前为 {size!r}")
    align = spec.get("alignment")
    if align is not None and align not in ALIGNMENTS:
        _err(errors, f"{path}.alignment",
             f'应为 {"/".join(ALIGNMENTS)} 之一，当前为 {align!r}')
    ls = spec.get("line_spacing")
    if ls is not None:
        _check_line_spacing(ls, f"{path}.line_spacing", errors)
    color = spec.get("color")
    if color is not None and color != "auto" and not (
            isinstance(color, str) and len(color) == 6
            and all(c in "0123456789abcdefABCDEF" for c in color)):
        _err(errors, f"{path}.color", '应为 6 位十六进制（如 "000000"）或 "auto"')
    for k in ("bold", "italic", "keep_with_next", "page_break_before"):
        v = spec.get(k)
        if v is not None and not isinstance(v, bool):
            _err(errors, f"{path}.{k}", "应为布尔值 true/false")
    for k in ("space_before_pt", "space_after_pt"):
        v = spec.get(k)
        if v is not None and (not isinstance(v, (int, float)) or isinstance(v, bool)
                              or not (-100 <= v <= 200)):
            _err(errors, f"{path}.{k}", "应为 -100~200 之间的磅值")
    for k in ("first_line_indent", "hanging_indent", "left_indent", "right_indent"):
        v = spec.get(k)
        if v is not None:
            _check_indent(v, f"{path}.{k}", errors)
    ol = spec.get("outline_level")
    if ol is not None and (not isinstance(ol, int) or isinstance(ol, bool)
                           or not (-1 <= ol <= 8)):
        _err(errors, f"{path}.outline_level", "应为 -1（正文级）或 0~8")
    for k in ("display_name", "font_cn", "font_en", "base_on"):
        v = spec.get(k)
        if v is not None and (not isinstance(v, str) or not v.strip()):
            _err(errors, f"{path}.{k}", "应为非空字符串")


def validate_profile(data):
    """返回 (errors, warnings)，两个字符串列表。errors 非空即方案不合法。"""
    errors, warnings = [], []
    if not isinstance(data, dict):
        _err(errors, "", "方案文件根节点应为 JSON 对象")
        return errors, warnings

    for k in data:
        if k not in ALLOWED_TOP_KEYS:
            _warn(warnings, k, "未知顶层字段，将被忽略")

    sv = data.get("schema_version", SCHEMA_VERSION)
    if sv != SCHEMA_VERSION:
        _err(errors, "schema_version", f"当前程序仅支持版本 {SCHEMA_VERSION}，收到 {sv!r}")

    meta = data.get("meta")
    if not isinstance(meta, dict) or not str(meta.get("name", "")).strip():
        _err(errors, "meta.name", "缺少方案名称（meta.name 必填）")

    # ---- styles ----
    styles = data.get("styles")
    if not isinstance(styles, dict) or not styles:
        _err(errors, "styles", "缺少 styles（至少需要 body_text）")
    else:
        for tid, spec in styles.items():
            if tid not in CONTENT_TYPES:
                _err(errors, f"styles.{tid}",
                     f"未知内容类型，可选值：{'、'.join(CONTENT_TYPES)}")
                continue
            _check_style_spec(tid, spec, errors, warnings)
        if "body_text" not in styles:
            _warn(warnings, "styles", "建议提供 body_text（正文），整稿时未识别段落将回落到它")
        for tid in ("heading_1", "heading_2", "heading_3"):
            if tid in styles and "heading_1" in styles:
                a, b = styles.get(tid), styles.get("heading_1")
                if tid != "heading_1" and isinstance(a, dict) and isinstance(b, dict) \
                        and a.get("size_pt") and b.get("size_pt") and a["size_pt"] >= b["size_pt"]:
                    _warn(warnings, f"styles.{tid}.size_pt",
                          f"{tid} 字号不小于 heading_1，通常标题层级字号应递减")

    # ---- page ----
    page = data.get("page")
    if page is not None:
        if not isinstance(page, dict):
            _err(errors, "page", "应为对象")
        else:
            paper = page.get("paper")
            if paper is not None and paper not in ("A4", "A3", "B5", "Letter"):
                _err(errors, "page.paper", f"应为 A4/A3/B5/Letter，当前为 {paper!r}")
            margins = page.get("margins")
            if margins is not None:
                if not isinstance(margins, dict):
                    _err(errors, "page.margins", "应为对象")
                else:
                    for k, v in margins.items():
                        if k not in ("top_cm", "bottom_cm", "left_cm", "right_cm",
                                     "gutter_cm", "header_cm", "footer_cm"):
                            _warn(warnings, f"page.margins.{k}", "未知字段")
                        elif not isinstance(v, (int, float)) or isinstance(v, bool) \
                                or not (0 <= v <= 10):
                            _err(errors, f"page.margins.{k}", "应为 0~10 之间的厘米值")

    # ---- numbering ----
    numbering = data.get("numbering")
    if numbering is not None:
        if not isinstance(numbering, dict):
            _err(errors, "numbering", "应为对象")
        else:
            mode = numbering.get("mode")
            if mode is not None and mode not in ("reference", "keep_target"):
                _err(errors, "numbering.mode",
                     '应为 "reference"（套用参考方案编号）或 "keep_target"（保留目标文档编号）')
        scheme = numbering.get("heading_scheme")
        if not isinstance(scheme, dict):
            _err(errors, "numbering.heading_scheme", "应为对象，含 levels 数组")
        else:
            levels = scheme.get("levels")
            if not isinstance(levels, list) or not levels:
                _err(errors, "numbering.heading_scheme.levels", "应为非空数组")
            else:
                for i, lv in enumerate(levels):
                    p = f"numbering.heading_scheme.levels[{i}]"
                    if not isinstance(lv, dict):
                        _err(errors, p, "应为对象")
                        continue
                    if lv.get("level") not in (1, 2, 3, 4):
                        _err(errors, f"{p}.level", "应为 1~4")
                    if lv.get("style") not in CONTENT_TYPES:
                        _err(errors, f"{p}.style", f"未知内容类型 {lv.get('style')!r}")
                    if "%1" not in str(lv.get("text", "")):
                        _err(errors, f"{p}.text", '编号文本必须包含 %1（本级编号占位符）')
                    fmt = lv.get("num_fmt", "decimal")
                    if fmt not in NUM_FMTS:
                        _err(errors, f"{p}.num_fmt",
                             f"应为 {'/'.join(NUM_FMTS)}，当前为 {fmt!r}")
                    if lv.get("suffix") not in (None, "space", "tab", "nothing"):
                        _err(errors, f"{p}.suffix", '应为 "space"/"tab"/"nothing"')

    # ---- objects（图片尺寸参考统计与可选的统一宽度开关） ----
    objects = data.get("objects")
    if objects is not None:
        if not isinstance(objects, dict):
            _err(errors, "objects", "应为对象")
        else:
            fig = objects.get("figure")
            if fig is not None:
                if not isinstance(fig, dict):
                    _err(errors, "objects.figure", "应为对象")
                else:
                    r = fig.get("resize")
                    if r is not None and not isinstance(r, bool):
                        _err(errors, "objects.figure.resize",
                             "应为布尔值 true/false（true 时整稿会把图片宽度统一到参考中位数）")
                    wcm = fig.get("width_cm")
                    if wcm is not None:
                        if not isinstance(wcm, dict):
                            _err(errors, "objects.figure.width_cm", "应为对象")
                        else:
                            for k in ("median", "min", "max"):
                                v = wcm.get(k)
                                if v is not None and (not isinstance(v, (int, float))
                                                      or isinstance(v, bool)
                                                      or not (0.5 <= v <= 40)):
                                    _err(errors, f"objects.figure.width_cm.{k}",
                                         "应为 0.5~40 之间的厘米值")

    # ---- tables（跨页表格规范开关） ----
    tables = data.get("tables")
    if tables is not None:
        if not isinstance(tables, dict):
            _err(errors, "tables", "应为对象")
        else:
            for k in ("repeat_header", "cant_split_rows"):
                v = tables.get(k)
                if v is not None and not isinstance(v, bool):
                    _err(errors, f"tables.{k}", "应为布尔值 true/false")

    # ---- citations（正文引用格式开关） ----
    citations = data.get("citations")
    if citations is not None:
        if not isinstance(citations, dict):
            _err(errors, "citations", "应为对象")
        else:
            v = citations.get("superscript")
            if v is not None and not isinstance(v, bool):
                _err(errors, "citations.superscript",
                     "应为布尔值（true 时整稿把正文独立 [n] 标记统一为上标）")

    # ---- caption / toc 简查 ----
    caption = data.get("caption")
    if caption is not None and not isinstance(caption, dict):
        _err(errors, "caption", "应为对象")
    toc = data.get("toc")
    if toc is not None:
        if not isinstance(toc, dict):
            _err(errors, "toc", "应为对象")
        elif "levels" in toc and (not isinstance(toc["levels"], int)
                                  or not (1 <= toc["levels"] <= 3)):
            _err(errors, "toc.levels", "应为 1~3 的整数")

    # ---- skeleton ----
    skeleton = data.get("skeleton")
    if skeleton is not None:
        if not isinstance(skeleton, list):
            _err(errors, "skeleton", "应为数组")
        else:
            for i, item in enumerate(skeleton):
                if not isinstance(item, dict) or "type" not in item:
                    _err(errors, f"skeleton[{i}]", '应为含 "type" 的对象')
                elif item["type"] not in CONTENT_TYPES and item["type"] not in ("toc_field", "blank"):
                    _err(errors, f"skeleton[{i}].type",
                         f"未知类型 {item['type']!r}（可用内容类型、toc_field 或 blank）")

    # ---- mapping ----
    mapping = data.get("mapping")
    if mapping is not None and not isinstance(mapping, dict):
        _err(errors, "mapping", "应为对象（原文档样式名 → 内容类型）")

    return errors, warnings
