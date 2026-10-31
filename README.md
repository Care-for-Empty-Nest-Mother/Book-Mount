# 书山 · Book Mount

<div align="center">

**把"中意文献的格式"变成一键可复用的排版方案**

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/platform-Windows%2010%2F11-blue)](#)
[![Python](https://img.shields.io/badge/python-3.10%2B-blue)](#)
[![Release](https://img.shields.io/badge/release-v0.2.0-green)](#)

</div>

---

写论文时最耗时间的不是写作，而是排版：期刊和学校的规定各不相同，人工逐段调整低效又易错。
**书山（Book Mount）** 换了一条路——从一篇排版规范的文献里**学习格式**，沉淀成可复用、
可分享的格式方案，再"按几下按键"套用到你自己的文档上。

> 书山有路勤为径 —— 让格式不再是障碍。

## ✨ 功能特性

- **格式精确提取**：不依赖文档样式表，解析样式继承链与直接格式的合成结果，
  读取每类段落的**实际生效格式**（字体、字号、行距、缩进、对齐、大纲级别）；
  `.doc` 老格式自动转换（借本机 Word，原文件只读）。
- **镜像 Word · 双栏格式工作台**：整稿前左手模板样式库、右手目标文档
  **所见即所得预览**（彩色类型标记保留，表格/图片按原文渲染）；
  顶部功能区与 Word"开始"选项卡同分布同外观（衬线 B/I、对齐线条图标、
  字体下拉按字体自身字形渲染），**快捷键与 Word 完全一致**
  （Ctrl+B/I/E/L/R/J/1/5/2/M/D/Q/Z/Y…）；点选段落即改格式，
  "更新到类型样式"一键把所选格式提升为类型样式；撤销/重做可回退；
  字体列表枚举系统全部字体并置顶论文常用（宋体/黑体/楷体/仿宋/Times 等）。
- **批注式归类界面**：像导师批注一样逐段显示将要套用的格式（彩色标签 + 底色），
  悬浮可查看该段提取到的完整格式明细；点击选中（Shift 连选 / Ctrl 加选），
  右键**修改格式**，树形导航按大纲级别分组、点击定位居中显示。
- **标题逐段编号控制**：每个标题带对号——勾选参与编号，取消则**只设定大纲级别**
  （样式、外观、编号字样全部保持原样），摘要、封面、参考文献等区域不再被误编号。
- **引用与文献一体化**：正文引用标记 `[1]`、`[2-4]` 与文献条目双向跳转定位；
  自动检查**悬空引用**与**未被引用的文献**；按 GB/T 7714 顺序编码制把引用标记
  统一为上标（不动文字）。
- **多种编号模式**：智能判断 / 保留目标文档编号（只改级别与样式）/ 套用参考方案
  编号 / 自定义编号样式（现场输入各级编号格式，实时预览）。
- **整稿安全机制**：任何写操作前自动备份到 `书山备份/`；图片、公式段防裁剪保护
  （固定值行距自动豁免）；超版心图片等比安全缩放；跨页表格规范
  （标题行重复 + 行不拆分）；表格与公式内容零改动。
- **页眉页脚与页码（M2）**：按方案给各节写入页眉文字与页码域
  （`第 {page} 页 共 {pages} 页`，Word 原生 PAGE/NUMPAGES 域），支持首页不同、
  奇偶页开关、**正文起编**（封面/摘要节页眉页脚清空、页码从正文第一章重新计数）；
  Word 打开时自动刷新域（updateFields）。
- **目录域（M2）**：文档没有目录时在第一个标题前自动插入 TOC 域
  （按方案目录级数），已有目录则不动；条目样式可由方案的 toc_1~3 定义。
- **表格内部文字重排（M2）**：按方案 `tables.cell_style` 统一全部表格单元格的
  字体、字号与对齐——不改文字、不动表格结构；提取时自动学习参考文档
  表格的多数派单元格格式。
- **格式方案文件**：提取结果沉淀为 `.fprof.json`——人可读、可手改、可分享，
  也支持让 AI 按规范生成后导入（内置校验器，错误定位到字段）。

## 📦 下载与安装

**方式一：绿色版（推荐）**

从 [Releases](../../releases) 下载 `BookMount-vX.X.X-windows-x64.zip`，解压后双击
`BookMount.exe` 即可，无需安装 Python。

**方式二：源码运行**

```bash
git clone https://github.com/Care-for-Empty-Nest-Mother/Book-Mount.git
cd Book-Mount
pip install -r requirements.txt
python main.py
```

## 🚀 快速上手

1. **提取格式**：点「从文献提取格式…」，选一篇排版规范的文献（期刊模板、
   导师改好的论文都行）→ 在批注视图里确认归类、按需改判 → 方案入库；
2. **整稿应用**：点「整稿应用到文档…」，选方案、选你的文档 → 归类确认 →
   选标题编号模式 → 完成。原文件自动备份在同级 `书山备份/` 目录；
3. **日常复用**：方案可反复使用，也可以分享给同门（直接发 `.fprof.json` 文件）。

命令行同样支持全部流程：

```bash
python main.py analyze 参考文献.docx -o 方案.fprof.json   # 提取
python main.py validate 方案.fprof.json                  # 校验
python main.py render 方案.fprof.json -o 方案.dotx       # 渲染 Word 模板
python main.py apply 方案.fprof.json 待整理.docx --auto   # 整稿
```

## 📖 文档

- [格式方案 JSON 规范](docs/格式方案JSON规范.md) —— schema 完整说明
- [给 AI 的生成提示词](docs/给AI的生成提示词.md) —— 把文字版格式要求变成方案文件

## 🗂 目录结构

```
├── main.py            # 入口：GUI / analyze / validate / render / apply
├── shushan/           # 核心包：分析器、方案、渲染、整稿、GUI
├── profiles/          # 格式方案库（你的方案保存在这里）
├── examples/          # 示例方案
└── docs/              # 规范文档
```

## 🧪 开发与构建

```bash
pip install -r requirements.txt

# 打包 exe（PyInstaller）
pip install pyinstaller
pyinstaller --noconfirm --clean --name BookMount --windowed ^
    --add-data "profiles/学位论文示例.fprof.json;profiles" main.py
```

## 🤝 适用于

学位论文排版 / 期刊投稿格式调整 / 团队格式统一 / 把导师改好的版式复用到自己的稿子。

> 已知边界：PDF 文献提取、WPS 适配、题注 SEQ 域（图表编号改写）与奇偶页
> 差异化页眉内容在计划中；奇偶页开关 v0.2 已支持但奇偶内容相同（镜像写入）。
> 文献条目文字本身的著录规范化不做自动改写。

## 📄 许可

[MIT](LICENSE) © 2026 Care-for-Empty-Nest-Mother
