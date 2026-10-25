# -*- coding: utf-8 -*-
"""书山 · 内容类型与内置样式映射常量。"""

# 内容类型全集（超集）。格式方案 JSON 的 styles 键只能取自这里。
CONTENT_TYPES = [
    "heading_1", "heading_2", "heading_3", "heading_4",
    "body_text", "abstract", "keywords",
    "caption_figure", "caption_table",
    "reference", "appendix",
    "cover", "statement",
    "toc_1", "toc_2", "toc_3",
    "author_info", "affiliation", "corresponding_author", "funding",
    "en_title", "en_abstract", "en_keywords",
    "header", "footer", "quote", "note",
    "figure_para", "formula_para",     # 图片段/公式段（段落版式：对齐、间距、行距）
]

# 「忽略」是归类界面专用的伪类型：标记为忽略的段落不进方案、整稿时保持原样
IGNORE_TYPE = "ignore"

# 各类型的默认中文名称（Word 样式显示名，custom 样式用）
DEFAULT_CN_NAMES = {
    "heading_1": "标题 1", "heading_2": "标题 2", "heading_3": "标题 3",
    "heading_4": "标题 4", "body_text": "正文", "abstract": "摘要",
    "keywords": "关键词", "caption_figure": "题注", "caption_table": "表注",
    "reference": "参考文献条目", "appendix": "附录标题", "cover": "封面标题",
    "statement": "原创性声明", "toc_1": "目录 1", "toc_2": "目录 2",
    "toc_3": "目录 3", "author_info": "作者信息", "affiliation": "作者单位",
    "corresponding_author": "通讯作者", "funding": "基金标注",
    "en_title": "英文标题", "en_abstract": "英文摘要", "en_keywords": "英文关键词",
    "header": "页眉", "footer": "页脚", "quote": "引用段落", "note": "注释",
    "figure_para": "图片段落", "formula_para": "公式段落",
}

# 映射到 Word 内置样式的类型：(styleId, 内置 w:name, COM 内置样式常量)
# 注意 w:name 用 Word 的内部名（题注/页眉/页脚为小写 caption/header/footer），
# COM 常量只列可靠的；其余类型在 COM 中按名称解析。
BUILTIN_STYLE_MAP = {
    "body_text": ("Normal", "Normal", -1),
    "heading_1": ("Heading1", "heading 1", -2),
    "heading_2": ("Heading2", "heading 2", -3),
    "heading_3": ("Heading3", "heading 3", -4),
    "heading_4": ("Heading4", "heading 4", -5),
    "caption_figure": ("Caption", "caption", None),
    "header": ("Header", "header", None),
    "footer": ("Footer", "footer", None),
    "toc_1": ("TOC1", "toc 1", None),
    "toc_2": ("TOC2", "toc 2", None),
    "toc_3": ("TOC3", "toc 3", None),
    "quote": ("Quote", "Quote", None),
}

def is_builtin(type_id: str) -> bool:
    return type_id in BUILTIN_STYLE_MAP

def style_id_of(type_id: str) -> str:
    if type_id in BUILTIN_STYLE_MAP:
        return BUILTIN_STYLE_MAP[type_id][0]
    return "Shushan" + "".join(w.capitalize() for w in type_id.split("_"))

def style_name_of(type_id: str) -> str:
    """样式的 w:name：内置样式用内置名（Word 按名匹配并本地化显示），自定义样式用中文名。"""
    if type_id in BUILTIN_STYLE_MAP:
        return BUILTIN_STYLE_MAP[type_id][1]
    return DEFAULT_CN_NAMES.get(type_id, type_id)

# 中文字号 → 磅值（供分析器与文档参考）
CN_FONT_SIZES = {"初号": 42, "小初": 36, "一号": 26, "小一": 24, "二号": 22, "小二": 18,
                 "三号": 16, "小三": 15, "四号": 14, "小四": 12, "五号": 10.5,
                 "小五": 9, "六号": 7.5, "小六": 6.5}

ALIGNMENTS = ["left", "center", "right", "justify", "distributed"]

NUM_FMTS = ["decimal", "chinese_counting", "upper_roman", "lower_roman",
            "upper_letter", "lower_letter"]

# w:numFmt 对应表
NUM_FMT_XML = {
    "decimal": "decimal", "chinese_counting": "chineseCounting",
    "upper_roman": "upperRoman", "lower_roman": "lowerRoman",
    "upper_letter": "upperLetter", "lower_letter": "lowerLetter",
}

SCHEMA_VERSION = 1
