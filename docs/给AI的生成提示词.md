# 给 AI 的生成提示词模板

使用方法：把下面整段提示词复制给任意 AI（书山内置 AI 对话、网页版均可），
在末尾【格式要求】处贴入期刊/学校发布的**文字版格式要求**即可。
AI 输出的 JSON 保存为 `方案名.fprof.json`，导入书山时会自动校验。

---

```text
你是一个 Word 排版专家。请把我提供的格式要求整理成"书山格式方案" JSON 文件，
只输出 JSON，不要任何解释文字或代码块标记。

【JSON 规则】
1. 顶层结构：
{
  "schema_version": 1,
  "meta": {"name": "方案名", "source": "ai", "notes": "要求出处"},
  "page": {"paper": "A4", "margins": {"top_cm": 数, "bottom_cm": 数, "left_cm": 数, "right_cm": 数}},
  "styles": { 内容类型: 格式定义 },
  "numbering": {"mode": "reference", "heading_scheme": {"levels": [{"level": 1, "style": "heading_1", "text": "第%1章", "num_fmt": "chinese_counting", "suffix": "space"}]}},
  "tables": {"repeat_header": true, "cant_split_rows": true},
  "header_footer": {"header": {"enabled": 布尔, "content": "页眉文字", "align": "center"},
                    "footer": {"enabled": 布尔, "page_number": {"position": "center", "format": "第 {page} 页"}}},
  "caption": {"figure_pattern": "图{chap}-{seq}", "table_pattern": "表{chap}-{seq}",
              "detect_patterns": ["^图\\s*\\d", "^表\\s*\\d"]}
}
2. "styles" 的键只能是：heading_1, heading_2, heading_3, heading_4, body_text, abstract,
   keywords, caption_figure, caption_table, figure_para, formula_para, reference,
   appendix, cover, statement, toc_1, toc_2, toc_3, author_info, affiliation,
   corresponding_author, funding, en_title, en_abstract, en_keywords, header, footer,
   quote, note。
   按格式要求里出现的类型填，没提到的不要编造。
   figure_para（图片段）/formula_para（公式段）只填对齐、段前后间距、行距；
   这两类的 line_spacing 严禁用 {"mode":"exact"}（会把图片/公式裁剪掉），用 {"mode":"multiple"}。
3. 每个类型可用的字段：
   "font_cn"（中文字体，如"宋体""黑体""楷体"）、"font_en"（西文字体，如"Times New Roman"）、
   "size_pt"（磅值）、"bold"/"italic"（布尔）、"alignment"（left/center/right/justify）、
   "line_spacing"（{"mode":"multiple","value":1.5} 倍数行距；或 {"mode":"exact","value":20} 固定值磅）、
   "space_before_pt"/"space_after_pt"（段前后磅）、
   "hanging_indent"（悬挂缩进，参考文献条目常用 {"unit":"chars","value":2}）、
   "first_line_indent"（{"unit":"chars","value":2} 首行缩进字符）、
   "outline_level"（heading_1 是 0，heading_2 是 1，依次加一；其他类型不写）、
   "keep_with_next"（标题类写 true）。
4. 中文字号换算：初号=42磅 小初=36 一号=26 小一=24 二号=22 小二=18 三号=16 小三=15
   四号=14 小四=12 五号=10.5 小五=9。格式要求若写"小四"，size_pt 就填 12。
5. 磅与厘米：1 磅 ≈ 0.035 厘米；页边距若给的是厘米直接填厘米，若给磅则换算。
6. 多级编号：格式要求说"一级标题：第X章（X为中文数字）"，text 用 "第%1章" 且
   num_fmt 用 "chinese_counting"；说"1.1"式编号，text 用 "%1.%2" 且 num_fmt 用 "decimal"。
7. 页边距若未给出，用 A4 默认：上下 2.54cm、左右 3.17cm。
8. 不要输出 "objects" 字段（那是程序从文献实测写入的统计）；
   "citations" 字段也保持缺省（默认把正文 [n] 标记统一为上标）。

【格式要求】
（在这里粘贴期刊投稿须知 / 学校学位论文格式要求原文）
```
