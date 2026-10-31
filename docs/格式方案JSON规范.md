# 书山格式方案 JSON 规范（schema v1）

格式方案文件是"书山"的唯一中间格式：由 docx 分析器产出、或由 AI 按本规范生成，
再由程序校验、渲染成 Word 模板（.dotx）或直接整稿应用。

- 扩展名：`.fprof.json`（自定义后缀 `.fprof` 亦兼容，内容相同）
- 编码：UTF-8
- 校验：`python main.py validate 方案.fprof.json`

---

## 1. 顶层结构

```jsonc
{
  "schema_version": 1,          // 固定为 1
  "meta":    { ... },           // 必填：方案名称等
  "page":    { ... },           // 可选：页面设置
  "styles":  { ... },           // 必填：各内容类型的格式定义
  "numbering": { ... },         // 可选：标题多级编号
  "header_footer": { ... },     // 可选：页眉页脚页码（M2 已生效）
  "toc":     { ... },           // 可选：目录域（M2 已生效）
  "caption": { ... },           // 可选（M3 生效）：图表题注规则
  "skeleton": [ ... ],          // 可选（M3 生效）：新文档骨架
  "mapping": { ... }            // 可选：原文献样式名/特征 → 内容类型，供精确复用
}
```

## 2. meta（必填）

| 字段 | 类型 | 说明 |
|---|---|---|
| name | string | 方案名称，必填，如 "XX大学硕士学位论文" |
| author | string | 方案作者/来源说明 |
| source | string | `"docx"`（从成稿提取）/ `"ai"`（AI 按文字要求生成）/ `"manual"` |
| notes | string | 备注 |

## 3. styles（必填）

键为**内容类型**，只能取以下值（超集，用多少写多少，但 `body_text` 强烈建议提供）：

```
heading_1  heading_2  heading_3  heading_4      章标题/节/小节/四级标题
body_text  abstract   keywords                  正文/摘要/关键词
caption_figure  caption_table                  图注/表注
figure_para  formula_para                      图片段/公式段（段落版式：对齐、间距、行距）
reference  appendix   cover   statement         参考文献条目/附录/封面/声明
toc_1 toc_2 toc_3                              目录 1~3 级样式
author_info affiliation corresponding_author funding   期刊投稿场景
en_title en_abstract en_keywords                英文标题/摘要/关键词
header footer                                    页眉/页脚样式
quote note                                       引用段落/注释
```

> **图片段/公式段**：图片与公式所在段落的版式（对齐、段前后间距、行距）。
> 应用时图片尺寸、位置不会被改动；若此类段落生效行距为"固定值"，
> 书山会自动覆盖为单倍行距，防止嵌入式图片被 Word 裁剪。
> **重要约束：这两类样式的行距不要写 `exact`（固定值）**。

每个类型的格式定义（所有字段均可省略，省略=保持 Word 该样式默认）：

| 字段 | 类型 | 说明 | 示例 |
|---|---|---|---|
| font_cn | string | 中文字体（eastAsia） | `"黑体"` |
| font_en | string | 西文字体（ascii/hAnsi） | `"Times New Roman"` |
| size_pt | number | 字号磅值 | `12`（小四） |
| bold / italic | bool | 加粗/倾斜 | `true` |
| color | string | `"RRGGBB"` 或 `"auto"` | `"000000"` |
| alignment | string | `left` `center` `right` `justify` `distributed` | `"center"` |
| line_spacing | object | `{"mode": "multiple"|"exact"|"at_least", "value": 数值}`；multiple 单位为倍，其余为磅 | `{"mode":"multiple","value":1.5}` |
| space_before_pt / space_after_pt | number | 段前/段后间距（磅） | `12` |
| first_line_indent | object | 首行缩进 `{"unit": "chars"|"cm", "value": 数值}`，中文论文惯用 2 字符 | `{"unit":"chars","value":2}` |
| hanging_indent | object | 悬挂缩进（参考文献条目规范排版用） | `{"unit":"chars","value":2}` |
| left_indent / right_indent | object | 左/右缩进，同上 | |
| outline_level | int | 大纲级别：`0`~`8` 对应一级~九级标题；正文省略 | `0` |
| keep_with_next | bool | 与下段同页（标题建议 true） | `true` |
| page_break_before | bool | 段前分页（章标题常用） | `false` |
| display_name | string | 覆盖默认的 Word 样式显示名 | `"论文正文"` |

> 内置样式映射：heading_1~4 → 标题 1~4，body_text → 正文(Normal)，
> caption_figure → 题注(Caption)，toc_1~3 → 目录 1~3，header/footer → 页眉/页脚。
> 其余类型生成同名自定义样式。映射到内置样式可让 Word 的大纲级别、目录、
> 交叉引用等机制原生生效，**请勿用自定义样式替代标题**。

常用中文字号对照：初号42 / 小初36 / 一号26 / 小一24 / 二号22 / 小二18 / 三号16 /
小三15 / 四号14 / 小四12 / 五号10.5 / 小五9 磅。

## 4. page（可选）

| 字段 | 说明 |
|---|---|
| paper | `"A4"` / `"A3"` / `"B5"` / `"Letter"` |
| orientation | `"portrait"` / `"landscape"` |
| margins | `{"top_cm","bottom_cm","left_cm","right_cm","gutter_cm","header_cm","footer_cm"}`，厘米数值 |
| columns | `{"count": 1, "space_cm": 0.75}` |

## 5. numbering（可选）—— 标题多级编号

`mode` 字段控制整稿时标题编号的来源：

- `"reference"`（默认）：标题编号套用参考方案的编号体系（heading_scheme）；
- `"keep_target"`：**自定义编号样式模式**——只更改标题的级别与视觉样式
  （字体字号间距等），编号体系保留目标文档自己的定义，适合已排好编号、
  只想换版式的文档。

```jsonc
"numbering": {
  "mode": "reference",
  "heading_scheme": {
    "levels": [
      { "level": 1, "style": "heading_1", "text": "第%1章", "num_fmt": "chinese_counting", "suffix": "space" },
      { "level": 2, "style": "heading_2", "text": "%1.%2", "num_fmt": "decimal", "suffix": "space" },
      { "level": 3, "style": "heading_3", "text": "%1.%2.%3", "num_fmt": "decimal", "suffix": "space" }
    ]
  }
}
```

- `text` 中 `%1` 为本级编号，`%2`/`%3` 为上级编号占位符，**必须含 `%1`**；
- `num_fmt`：`decimal`(1,2,3) / `chinese_counting`(一,二,三) / `upper_roman` / `lower_roman` / `upper_letter`(A,B) / `lower_letter`(a,b)；
- `suffix`：编号与标题文字间的间隔，`space`/`tab`/`nothing`，默认 `space`。

## 5.4 citations（可选）—— 正文引用格式

```jsonc
"citations": { "superscript": true }
```

`superscript` 默认 true：整稿时把正文中**独立成 run** 的引用标记（`[1]`、`[2-4]`）
统一为 GB/T 7714 顺序编码制的上标格式。只动格式、不改文字；混合在长句里的
标记不处理。同时程序会检查引用完整性（悬空引用 / 未被引用的文献）并在
批注视图与整稿日志中提示。

## 5.5 tables（可选）—— 跨页表格规范与单元格文字样式

```jsonc
"tables": { "repeat_header": true, "cant_split_rows": true,
            "cell_style": { "font_cn": "宋体", "font_en": "Times New Roman",
                            "size_pt": 10.5, "alignment": "center" } }
```

- `repeat_header`：跨页表格的标题行在每一页重复（Word 原生 tblHeader）；
- `cant_split_rows`：表格行不跨页拆分（cantSplit），避免一行内容被拦腰截断；
- 程序会从参考文档的表格多数派自动提取这两个值；无表格时默认均为 true。
- 整稿时应用到目标文档的全部表格。
- `cell_style`（M2）：表格内部文字重排——把全部单元格的文字统一为该字体、
  字号与对齐（直接格式落地，不改文字内容、不动表格结构）；提取时自动学习
  参考文档表格单元格的多数派格式。可省略（省略时表格文字不动）。

## 6. header_footer（可选，M2 生效）

```jsonc
"header_footer": {
  "header": { "enabled": true, "content": "XX大学硕士学位论文", "align": "center",
              "font": {"font_cn": "宋体", "size_pt": 9},
              "different_first": false, "different_odd_even": false },
  "footer": { "enabled": true,
              "page_number": { "position": "center", "format": "第 {page} 页 共 {pages} 页",
                               "start_at": 1, "start_section": "body" } },
  "body_starts_new_section": true
}
```

`format` 中的 `{page}` `{pages}` 会替换为 Word 的 PAGE / NUMPAGES 域。

整稿行为（M2）：

- 多节文档按节处理：`body_starts_new_section=true` 时，页码重启写入**第一个
  正文级标题所在节**，该节之前各节（封面/摘要）页眉页脚清空并断开继承；
- `different_first`：只在页眉页脚区的第一节设首页不同（首页页眉页脚留空），
  多节文档不会让每一章的首页都丢页眉页脚；
- `different_odd_even`：仅打开 Word 的奇偶页开关，奇偶内容相同（同内容镜像
  写入偶数页），差异化内容由用户在 Word 里自行编辑；
- 从参考文档提取时：页眉取第一个有内容的节，页脚里的 PAGE/NUMPAGES 域还原
  为 `{page}`/`{pages}` 模板；参考文档没有页眉页脚时不写入方案；
- Word 打开整稿后的文档会自动刷新域（updateFields），页码即显示实际值。

## 7. toc / caption（可选，M2 已生效）

```jsonc
"toc":     { "levels": 3, "leader": "dot", "page_number_align": "right", "insert": true },
"caption": { "figure_pattern": "图{chap}-{seq}", "table_pattern": "表{chap}-{seq}",
             "position_figure": "below", "position_table": "above",
             "detect_patterns": ["^图\\s*\\d", "^表\\s*\\d", "^Fig", "^Table\\s*\\d"] }
```

`toc` 的整稿行为（M2）：`insert=true` 且文档内**尚无 TOC 域**时，在第一个
标题段之前插入 TOC 域（`\o "1-{levels}" \h \z \u`），生成的条目自动套用
方案的 toc_1~3 样式；文档已有 TOC 域则不重复插入。无论插入与否都会在
settings.xml 写 `updateFields`，Word 打开文档时自动刷新全部域。在 Word 中
右键目录处 →「更新域」即可生成/刷新目录条目。

`detect_patterns` 是正则表达式数组，用于归类时识别图注/表注段落，按需增删
（英文期刊可加 `"^Figure\\s*\\d"` 等）。

## 8. skeleton（可选，M3 生效）

新文档骨架，按顺序渲染占位段落：

```jsonc
"skeleton": [
  { "type": "cover", "text": "论文题目" },
  { "type": "abstract", "text": "摘要" },
  { "type": "keywords", "text": "关键词：；；" },
  { "type": "toc_field" },
  { "type": "heading_1", "text": "绪论" },
  { "type": "body_text", "text": "" },
  { "type": "reference", "text": "[1] " }
]
```

`type` 取内容类型或 `toc_field`（插入目录域）/ `blank`（空行）。

## 9. objects（可选，程序自动写入）

图片尺寸参考统计（信息性记录，应用时不会盲目修改目标文档的图片大小）：

```jsonc
"objects": {
  "figure": { "count": 59, "width_cm": { "median": 9.84, "min": 5.99, "max": 15.62 },
              "resize": false }
}
```

- `width_cm`：参考文档图片宽度的统计值（程序自动写入）；
- `resize`：**显式开关，默认 false**。为 true 时，整稿会把目标文档的内嵌图片
  宽度等比统一到 `median`（不超过版心宽，纵横比不变，浮动图不动）；
- 无论开关与否，宽度超过版心的图片都会被安全缩到版心宽以内（日志可见）。
- 注意：`resize` 开启会把小图放大到中位宽，放大可能发虚，请确认后再开启。

## 10. mapping（可选）

从 docx 成稿提取时，程序会自动写入"原样式名/特征 → 内容类型"的映射，
用于再次套用同名文档时跳过人工确认。手写方案无需提供。

---

## 完整示例

见 [`examples/学位论文示例.fprof.json`](../examples/学位论文示例.fprof.json)。
