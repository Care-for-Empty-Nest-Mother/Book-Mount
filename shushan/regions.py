# -*- coding: utf-8 -*-
"""区域标记与分区域识别规则（M2.6）。

用户在工作台里把论文划成区域（封面页/目录页/中文摘要页/…），识别按区域
规则执行，避免跨区域误判（封面无标题、目录区整页保持原样、摘要区用摘要
类型、参考文献区条目化等）。区域 = 段落号闭区间（连续覆盖全篇，未标记段
落自动归入 auto 区走现有自动识别）；区间内的表格/图片随区间归属。

纯函数模块：seed/normalize/rules 均可离屏单测。区域标记只影响本次整稿，
不写方案文件；CLI 无 UI 不带区域，行为不变。
"""
import re
from dataclasses import dataclass

from .analyzer.cluster import RE_ABS_CN, RE_ABS_EN, RE_KW_CN, RE_KW_EN

REGION_TYPES = [
    ("auto", "自动（未标记）"),
    ("cover", "封面页"),
    ("toc", "目录页"),
    ("abstract_cn", "中文摘要页"),
    ("abstract_en", "英文摘要页"),
    ("body", "正文页"),
    ("acknowledgement", "致谢页"),
    ("reference", "参考文献页"),
    ("appendix", "附录"),
    ("statement", "声明页"),
]
REGION_LABELS = dict(REGION_TYPES)

# 区域横幅徽标底色（预览里区域首段的【区域名】用）
REGION_COLORS = {
    "auto": "#adb5bd",
    "cover": "#495057",
    "toc": "#1971c2",
    "abstract_cn": "#c2255c",
    "abstract_en": "#9c36b5",
    "body": "#868e96",
    "acknowledgement": "#0b7285",
    "reference": "#5f3dc4",
    "appendix": "#5c940d",
    "statement": "#a61e4d",
}

# 参与预填/标记的可选类型（auto 只作填充，不给用户主动标记）
MARKABLE_TYPES = [t for t in REGION_TYPES if t[0] != "auto"]

_TITLE_MAX = 30          # 区域首段视为"标题"的最大长度
_PAGENUM = re.compile(r"[\divxIVX]+\s*$")   # 目录行页码尾巴（辅助判断，不强依赖）


@dataclass
class Region:
    type: str
    start: int              # 段落号（闭区间，含端点）
    end: int


def seed_regions(analysis) -> list[Region]:
    """按分节预填区域：首节=封面（沿用首节封面判定），其余=auto。"""
    ends = list(getattr(analysis, "section_ends", None) or [])
    if not ends:
        ends = [max(len(analysis.paragraphs) - 1, 0)]
    regions, start = [], 0
    for si, end in enumerate(ends):
        rtype = "cover" if (si == 0 and getattr(analysis, "first_section_is_cover",
                                                False)) else "auto"
        regions.append(Region(rtype, start, end))
        start = end + 1
    return regions


def normalize_regions(regions: list[Region], n_paras: int) -> list[Region]:
    """排序、越界裁剪、auto 填洞、相邻同类型合并——保证连续覆盖 [0, n-1]。

    auto 是"未标记"兜底，优先级最低：任何显式标记（包括后标的）都覆盖
    auto 区；显式标记之间按 先标先得（后来者被截到前一区域之后）。
    """
    marked = sorted((r for r in regions
                     if r.type != "auto" and r.end >= r.start and r.start < n_paras),
                    key=lambda r: r.start)
    out: list[Region] = []
    cursor = 0
    for r in marked:
        s = max(cursor, max(0, r.start))
        e = min(max(r.start, r.end), n_paras - 1)
        if s > e:
            continue
        if s > cursor:
            out.append(Region("auto", cursor, s - 1))
        if out and out[-1].type == r.type and out[-1].end == s - 1:
            out[-1].end = e                       # 相邻同类型合并
        else:
            out.append(Region(r.type, s, e))
        cursor = e + 1
    if cursor < n_paras:
        out.append(Region("auto", cursor, n_paras - 1))
    return [r for r in out if r.end >= r.start]


def region_type_at(regions: list[Region], idx: int) -> str:
    for r in regions:
        if r.start <= idx <= r.end:
            return r.type
    return "auto"


def apply_region_rules(analysis, regions: list[Region]) -> dict:
    """按区域改写建议（就地）。返回 {"changed", "numbered_off", "region_of"}。

    - numbered_off：建议为标题的段落号——工作台据此把"用户没手动动过对号"
      的标题默认取消对号（仅大纲级别：进目录、不占编号号）；
    - auto 区域一律不动（自动兜底）。
    """
    n = len(analysis.paragraphs)
    changed = 0
    numbered_off: list[int] = []
    region_of: dict[int, str] = {}

    # 各"标题型"区域的首个非空段视为区域标题（进目录、不占编号）
    title_idx: dict[str, int] = {}
    for p in analysis.nonempty():
        rt = region_type_at(regions, p.index)
        region_of[p.index] = rt
        if rt in ("abstract_cn", "abstract_en", "acknowledgement", "reference",
                  "appendix") and rt not in title_idx:
            if len(p.text.strip()) <= _TITLE_MAX:
                title_idx[rt] = p.index

    for p in analysis.nonempty():
        rt = region_of[p.index]
        if rt == "auto" or rt == "body":
            continue
        text = p.text.strip()
        new, reason = None, ""
        if rt == "cover":
            if p.suggestion.startswith("heading_"):
                new, reason = "cover", "封面页区域：标题不设大纲级别"
            elif p.suggestion in ("caption_figure", "caption_table"):
                new, reason = "ignore", "封面页区域：题注不适用"
        elif rt == "toc":
            if p.suggestion != "ignore":
                new, reason = "ignore", "目录页区域：整页保持原样"
        elif rt == "abstract_cn":
            if title_idx.get(rt) == p.index:
                new, reason = "heading_1", "中文摘要页区域：标题仅大纲（不占编号）"
            elif RE_KW_CN.search(text):
                new, reason = "keywords", "中文摘要页区域：关键词行"
            else:
                new, reason = "abstract", "中文摘要页区域：摘要正文"
        elif rt == "abstract_en":
            if title_idx.get(rt) == p.index:
                new, reason = "heading_1", "英文摘要页区域：标题仅大纲（不占编号）"
            elif RE_KW_EN.search(text):
                new, reason = "en_keywords", "英文摘要页区域：Keywords 行"
            else:
                new, reason = "en_abstract", "英文摘要页区域：Abstract 正文"
        elif rt == "acknowledgement":
            if title_idx.get(rt) == p.index:
                new, reason = "heading_1", "致谢页区域：标题仅大纲（不占编号）"
            else:
                new, reason = "body_text", "致谢页区域：致谢正文"
        elif rt == "reference":
            if title_idx.get(rt) == p.index:
                new, reason = "heading_1", "参考文献页区域：标题仅大纲（不占编号）"
            else:
                new, reason = "reference", "参考文献页区域：文献条目"
        elif rt == "appendix":
            if title_idx.get(rt) == p.index:
                new, reason = "heading_1", "附录区域：标题仅大纲（不占编号）"
            else:
                new, reason = "body_text", "附录区域：附录正文"
        elif rt == "statement":
            new, reason = "statement", "声明页区域：声明文字"
        if new and new != p.suggestion:
            p.suggestion = new
            p.reason = reason
            changed += 1
        if new and new.startswith("heading_"):
            numbered_off.append(p.index)
    return {"changed": changed,
            "numbered_off": sorted(set(numbered_off)),
            "region_of": region_of}


def restore_suggestions(analysis, base_suggestion: dict, base_reason: dict,
                        cleared: list[Region], touched_numbered: set,
                        numbered: dict):
    """清除区域标记后，把被清区域的建议/理由还原到分析基线。"""
    for p in analysis.nonempty():
        if not any(r.start <= p.index <= r.end for r in cleared):
            continue
        p.suggestion = base_suggestion.get(p.index, p.suggestion)
        p.reason = base_reason.get(p.index, p.reason)
        if p.index in numbered and p.index not in touched_numbered:
            numbered[p.index] = True
