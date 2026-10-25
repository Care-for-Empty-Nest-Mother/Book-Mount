# -*- coding: utf-8 -*-
"""书山 入口：无参数启动 GUI，子命令走命令行。

  python main.py                     图形面板
  python main.py analyze a.docx -o 方案.fprof.json [--name 名称]
  python main.py validate 方案.fprof.json
  python main.py render 方案.fprof.json -o 方案.dotx
  python main.py apply 方案.fprof.json b.docx [--auto] [--no-backup]
"""
import argparse
import sys


def cmd_analyze(args):
    from collections import Counter

    from shushan.analyzer.docx_analyzer import analyze_document
    from shushan.convert import ensure_docx
    from shushan.profile_builder import build_profile, auto_assignment
    from shushan.profile import save_profile

    analysis_path, converted = ensure_docx(args.input)
    if converted:
        print(f"[转换] {args.input} → {converted}（原文件未动）")
    analysis = analyze_document(analysis_path)
    analysis.source_path = str(args.input)
    assignment = auto_assignment(analysis)
    name = args.name or f"{args.stem_default}"
    profile = build_profile(analysis, assignment, name=name, source="docx")
    out = args.output or (args.input.rsplit(".", 1)[0] + ".fprof.json")
    save_profile(profile, out)

    para_counts = Counter(assignment[p.index] for p in analysis.nonempty())
    print(f"[完成] {out}")
    print(f"  方案名：{name}")
    print(f"  页面：{analysis.page.get('paper')}，边距 {analysis.page.get('margins')}")
    print("  归类统计：")
    for t, n in para_counts.most_common():
        print(f"    {t:<16} ×{n}")
    print("  提示：可打开 GUI 用批注视图复核与单段改判。")


def cmd_validate(args):
    import sys as _sys
    from shushan.profile import validate_or_report
    report = validate_or_report(args.file)
    print(report)
    if "[未通过]" in report:
        _sys.exit(1)


def cmd_render(args):
    from shushan.profile import load_profile
    from shushan.render import render_dotx
    profile = load_profile(args.file)
    out = args.output or args.file.rsplit(".", 1)[0] + ".dotx"
    render_dotx(profile, out)
    print(f"[完成] 已生成 Word 模板：{out}")
    print("  用法：Word → 文件 → 选项 → 加载项 → 管理:模板 → 转到 → 选用此模板，"
          "并勾选「自动更新文档样式」。")


def cmd_apply(args):
    from shushan.convert import ensure_docx
    from shushan.profile import load_profile
    from shushan.render import apply_profile_offline
    profile = load_profile(args.file)
    work_path, converted = ensure_docx(args.docx)
    if converted:
        print(f"[转换] {args.docx} → {converted}（整稿写入转换件，原 .doc 保持只读）")
    if not args.auto:
        print("[提示] 未加 --auto：交互式确认请使用 GUI（python main.py）。")
    report = apply_profile_offline(profile, work_path,
                                   make_backup=not args.no_backup)
    print(f"[完成] 整稿 {report['paragraphs']} 段 → {work_path}")
    for t, n in sorted(report["counts"].items(), key=lambda kv: -kv[1]):
        print(f"    {t:<16} ×{n}")
    if report.get("backup"):
        print(f"  备份：{report['backup']}")
    if report.get("manual_numbering_kept"):
        print("  检测到手工编号，已保留文字编号、未启用自动编号")
    for w in report["warnings"][:10]:
        print(f"  ⚠ {w}")


def main():
    parser = argparse.ArgumentParser(prog="书山", description="文献格式提取与复用工具")
    sub = parser.add_subparsers(dest="cmd")

    p = sub.add_parser("analyze", help="从 docx 成稿提取格式方案")
    p.add_argument("input", help="参考文献 docx 路径")
    p.add_argument("-o", "--output", help="输出方案文件路径")
    p.add_argument("--name", help="方案名称")
    p.add_argument("--stem-default", help=argparse.SUPPRESS)
    p.set_defaults(func=cmd_analyze, stem_default="我的格式方案")

    p = sub.add_parser("validate", help="校验格式方案 JSON")
    p.add_argument("file")
    p.set_defaults(func=cmd_validate)

    p = sub.add_parser("render", help="方案 JSON → .dotx 模板")
    p.add_argument("file")
    p.add_argument("-o", "--output")
    p.set_defaults(func=cmd_render)

    p = sub.add_parser("apply", help="离线整稿：方案应用到 docx")
    p.add_argument("file", help="方案文件")
    p.add_argument("docx", help="待整稿文档")
    p.add_argument("--auto", action="store_true", help="按自动建议归类，不询问")
    p.add_argument("--no-backup", action="store_true")
    p.set_defaults(func=cmd_apply)

    args = parser.parse_args()
    if not args.cmd:
        from shushan.gui.panel import run
        run()
        return
    try:
        args.func(args)
    except Exception as e:
        print(f"[错误] {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
