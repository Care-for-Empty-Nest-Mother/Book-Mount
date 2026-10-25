# -*- coding: utf-8 -*-
"""段落类型建议（半自动归类的"建议"半边）。

依据：文字模式（第X章/1.1/一、/图x-x/[n]…）、大纲级别、字号与正文基准字号的
相对关系、对齐与加粗。只给建议不给结论，最终由用户在面板上确认或改判。
"""
import re
from collections import Counter

RE_CHAPTER = re.compile(r"^第\s*[0-9一二三四五六七八九十百千]+\s*[章节篇部分讲]")
RE_NUM_HEAD = re.compile(r"^(\d+(?:\.\d+)*)[\s.、．]")
RE_CN_HEAD = re.compile(r"^[一二三四五六七八九十百]+\s*[、．.]")
RE_PAREN_HEAD = re.compile(r"^[（(][一二三四五六七八九十]+[)）]")
RE_FIG = re.compile(r"^\s*图\s*\d")
RE_TAB = re.compile(r"^\s*表\s*\d")
RE_FIG_EN = re.compile(r"^\s*(Figure|Fig\.?)\s*\d", re.I)
RE_TAB_EN = re.compile(r"^\s*Table\s*\d", re.I)
RE_REF = re.compile(r"^\s*\[\d+\]")
RE_ABS_CN = re.compile(r"^摘\s*要$")
RE_KW_CN = re.compile(r"^关\s*键\s*词")
RE_ABS_EN = re.compile(r"^Abstract\b", re.I)
RE_KW_EN = re.compile(r"^Key\s?words?\b", re.I)
# 目录条目：正文 + 制表符/点线引导 + 页码结尾（Word 目录的标准形态）
RE_TOC_TAB = re.compile(r"\t\s*\d+\s*$")
RE_TOC_DOTS = re.compile(r"[·.…]{4,}\s*\d+\s*$")
# 文献著录标识符：[J]期刊 [D]学位论文 [C]会议 [M]专著 [R]报告 [S]标准 [P]专利 [A] 等
RE_BIB_MARKER = re.compile(r"\[(?:J|D|C|M|A|R|N|P|S|G|EB/OL|DB/OL|J/OL|M/OL|N/OL)\]", re.I)

MAX_HEADING = 4


def _body_size(paras) -> float:
    """正文基准字号：取长段落（≥20字）中最常见字号。"""
    sizes = Counter(p.fmt.get("size_pt") for p in paras
                    if p.fmt and len(p.text.strip()) >= 20 and p.fmt.get("size_pt"))
    return sizes.most_common(1)[0][0] if sizes else 12.0


def suggest_type(text, fmt, analysis):
    """返回 (建议类型, 置信度0~1, 理由)。analysis 用于跨段统计（正文基准字号）。"""
    text = text.strip()
    size = fmt.get("size_pt") or 0.0
    body = _body_size(analysis.nonempty()) if analysis else 12.0

    # 1) 大纲级别是最硬的证据
    ol = fmt.get("outline")
    if ol is not None and 0 <= ol <= MAX_HEADING - 1 and len(text) < 80:
        return f"heading_{ol + 1}", 0.95, f"大纲级别{ol + 1}级"

    # 1.5) 目录条目（带页码的行绝不能当标题，否则会污染标题格式）
    if RE_TOC_TAB.search(text) or RE_TOC_DOTS.search(text):
        return "ignore", 0.85, "目录条目（页码行）"

    # 1.6) 自动编号 + 著录标识符 → 参考文献条目（编号在列表机制里，正文无 [n] 前缀）
    if fmt.get("numbering") and RE_BIB_MARKER.search(text):
        return "reference", 0.9, "自动编号 + 文献标识符"

    # 2) 结构性文字模式
    if RE_ABS_CN.match(text):
        return "abstract", 0.9, "「摘要」"
    if RE_ABS_EN.match(text):
        return "en_abstract", 0.85, "Abstract"
    if RE_KW_CN.match(text):
        return "keywords", 0.9, "「关键词」"
    if RE_KW_EN.match(text):
        return "en_keywords", 0.85, "Key words"
    # 文献条目须有著录要素：只有 [n] 开头不足以判定（防止正文引用句被误判）
    if RE_REF.match(text) and RE_BIB_MARKER.search(text):
        return "reference", 0.9, "[n] 编号 + 文献标识符"
    if RE_REF.match(text) and len(text) > 60 and re.search(r"\d{4}", text):
        return "reference", 0.6, "[n] 编号 + 长条目形态（弱证据）"

    looks_caption = len(text) <= 60
    if looks_caption and (RE_FIG.match(text) or RE_FIG_EN.match(text)):
        return "caption_figure", 0.8, "图注编号模式"
    if looks_caption and (RE_TAB.match(text) or RE_TAB_EN.match(text)):
        return "caption_table", 0.8, "表注编号模式"

    # 3) 标题启发式：文字模式 + 字号/加粗/居中等版面证据
    bold = bool(fmt.get("bold"))
    centered = fmt.get("align") == "center"
    short = len(text) <= 40 and "。" not in text and "；" not in text
    bigger = size > body + 0.3

    if RE_CHAPTER.match(text) and short:
        return "heading_1", 0.9, "「第X章」模式"
    m = RE_NUM_HEAD.match(text)
    if m and short:
        depth = min(m.group(1).count(".") + 1, MAX_HEADING)
        if bigger or bold:
            return f"heading_{depth}", 0.85, f"{m.group(1)} 编号 + 版面证据"
        return f"heading_{depth}", 0.55, f"{m.group(1)} 编号（版面证据弱）"
    if RE_PAREN_HEAD.match(text) and short:
        return "heading_3", 0.6, "「（一）」模式"
    if RE_CN_HEAD.match(text) and short:
        return "heading_2", 0.6, "「一、」模式"

    if short and centered and bigger:
        return "heading_1", 0.6, "居中 + 大字号 + 短文本"
    if short and bold and bigger:
        return "heading_2", 0.5, "加粗 + 大字号 + 短文本"

    # 4) 其余一律视为正文
    return "body_text", 0.4, "默认正文"
