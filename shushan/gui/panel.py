# -*- coding: utf-8 -*-
"""书山主面板：格式库 + 提取/导入/渲染/整稿四个动作 + 日志。"""
import shutil
from pathlib import Path

from PySide6.QtGui import QIcon
from PySide6.QtWidgets import (QFileDialog, QHBoxLayout, QLabel, QListWidget,
                               QMainWindow, QMessageBox, QPlainTextEdit,
                               QPushButton, QVBoxLayout, QWidget)

from ..analyzer.docx_analyzer import analyze_document
from ..convert import ensure_docx
from ..profile import ProfileError, load_profile, save_profile
from ..profile_builder import build_profile
from ..render import apply_profile_offline
from ..render.dotx_renderer import render_dotx
from ..types import DEFAULT_CN_NAMES
from ..word import cleanup as word_cleanup
from .classify_dialog import ClassifyDialog

APP_TITLE = "书山 · 文献格式工坊"


class ShushanPanel(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        from ..paths import ICON_FILE
        if ICON_FILE.exists():
            self.setWindowIcon(QIcon(str(ICON_FILE)))
        self.resize(560, 640)
        self._build_ui()
        try:
            procs = word_cleanup.list_word_processes()
            bg = [p for p in procs if p[2]]
            self.word_status.setText(
                f"Word 后台：{len(procs)} 个进程，其中后台残留 {len(bg)} 个"
                + ("（可点右侧按钮清理）" if bg else ""))
        except Exception:
            self.word_status.setText("Word 后台：未检查")
        self.refresh_profiles()

    # ---------- UI ----------
    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        lay = QVBoxLayout(central)

        title = QLabel(APP_TITLE)
        title.setStyleSheet("font-size: 16px; font-weight: 600; color: #1a3c34;")
        lay.addWidget(title)

        word_row = QHBoxLayout()
        self.word_status = QLabel("Word 后台：未检查")
        word_row.addWidget(self.word_status, stretch=1)
        btn_cleanup = QPushButton("清理 Word 后台")
        btn_cleanup.setToolTip("结束后台残留的 Word 进程（自动化残留）；\n"
                               "你正在使用的、有窗口的 Word 不受影响")
        btn_cleanup.clicked.connect(self.on_cleanup_word)
        word_row.addWidget(btn_cleanup)
        lay.addLayout(word_row)

        lib_row = QHBoxLayout()
        lib_row.addWidget(QLabel("格式方案库（profiles/）："))
        lib_row.addStretch(1)
        btn_del = QPushButton("删除选中")
        btn_del.setToolTip("删除选中的格式方案文件（不可恢复，同名 .dotx 模板不受影响）")
        btn_del.clicked.connect(self.on_delete_profile)
        lib_row.addWidget(btn_del)
        lay.addLayout(lib_row)
        self.profile_list = QListWidget()
        lay.addWidget(self.profile_list, stretch=1)

        row1 = QHBoxLayout()
        btn_extract = QPushButton("① 从文献提取格式…")
        btn_extract.clicked.connect(self.on_extract)
        btn_import = QPushButton("② 导入 JSON 方案…")
        btn_import.clicked.connect(self.on_import)
        row1.addWidget(btn_extract)
        row1.addWidget(btn_import)
        lay.addLayout(row1)

        row2 = QHBoxLayout()
        btn_render = QPushButton("③ 渲染为 .dotx 模板…")
        btn_render.clicked.connect(self.on_render)
        btn_apply = QPushButton("④ 整稿应用到文档…")
        btn_apply.clicked.connect(self.on_apply)
        row2.addWidget(btn_render)
        row2.addWidget(btn_apply)
        lay.addLayout(row2)

        self.log = QPlainTextEdit()
        self.log.setReadOnly(True)
        self.log.setPlaceholderText("操作日志…")
        lay.addWidget(self.log, stretch=1)

    def log_line(self, text: str):
        self.log.appendPlainText(text)

    # ---------- 状态 ----------
    def on_cleanup_word(self):
        """清理自动化残留的 Word 后台进程（可见窗口的 Word 不受影响）。"""
        try:
            rep = word_cleanup.cleanup_word_background()
        except ImportError:
            QMessageBox.warning(self, APP_TITLE, "清理 Word 后台需要 pywin32 组件。")
            return
        except Exception as e:
            QMessageBox.critical(self, APP_TITLE, f"清理失败：{e}")
            return
        if rep["found"] == 0:
            self.word_status.setText("Word 后台：无残留进程 ✓")
            self.word_status.setStyleSheet("color: #1a6b45;")
            self.log_line("[Word后台] 未发现 Word 进程")
        else:
            msg = (f"发现 {rep['found']} 个 Word 进程："
                   f"结束后台 {len(rep['killed'])} 个")
            if rep["kept"]:
                msg += f"，保留可见窗口 {len(rep['kept'])} 个（你正在用的）"
            if rep["failed"]:
                msg += f"，{len(rep['failed'])} 个结束失败"
            self.word_status.setText("Word 后台：" + msg)
            self.word_status.setStyleSheet("color: #1a6b45;" if not rep["failed"] else "color: #8a6d1a;")
            self.log_line(f"[Word后台] {msg}"
                          + (f" PID={rep['killed']}" if rep["killed"] else ""))

    def refresh_profiles(self):
        from ..paths import PROFILES_DIR
        self.profile_list.clear()
        for p in sorted(PROFILES_DIR.glob("*.fprof.json")) + sorted(PROFILES_DIR.glob("*.fprof")):
            self.profile_list.addItem(p.name)

    def selected_profile_path(self) -> Path | None:
        item = self.profile_list.currentItem()
        if not item:
            QMessageBox.information(self, APP_TITLE, "请先在格式方案库中选择一个方案。")
            return None
        from ..paths import PROFILES_DIR
        return PROFILES_DIR / item.text()

    def _confirm(self, msg: str) -> bool:
        """删除等危险操作的确认弹窗（测试可替换）。"""
        box = QMessageBox(self)
        box.setWindowTitle(APP_TITLE)
        box.setIcon(QMessageBox.Warning)
        box.setText(msg)
        yes = box.addButton("删除", QMessageBox.DestructiveRole)
        box.addButton("取消", QMessageBox.RejectRole)
        box.exec()
        return box.clickedButton() is yes

    def on_delete_profile(self):
        ppath = self.selected_profile_path()
        if not ppath or not ppath.exists():
            if ppath is not None:
                QMessageBox.warning(self, APP_TITLE, f"方案文件不存在：{ppath.name}")
            return
        if not self._confirm(f"确定删除方案「{ppath.name}」吗？\n删除后不可恢复"
                             f"（渲染出的 .dotx 模板不受影响）。"):
            self.log_line(f"[删除] 已取消：{ppath.name}")
            return
        try:
            ppath.unlink()
        except Exception as e:
            QMessageBox.critical(self, APP_TITLE, f"删除失败：{e}")
            return
        self.refresh_profiles()
        self.log_line(f"[删除] {ppath.name} 已删除")

    # ---------- 动作 ----------
    def on_extract(self):
        src = self._pick_source_doc("提取格式")
        if not src:
            return
        try:
            analysis_path, converted = ensure_docx(src)
            if converted:
                self.log_line(f"[转换] {Path(src).name} → {Path(converted).name}（原文件未动）")
            analysis = analyze_document(analysis_path)
        except Exception as e:
            QMessageBox.critical(self, APP_TITLE, f"分析失败：{e}")
            return
        analysis.source_path = str(src)   # 方案备注记录原始文件
        dlg = ClassifyDialog(analysis, Path(src).name, self)
        if dlg.exec() != ClassifyDialog.Accepted:
            return
        default_name = Path(src).stem + " 的格式"
        name, ok = self._ask_text("方案名称", default_name)
        if not ok:
            return
        profile = build_profile(analysis, dlg.assignment, name=name, source="docx",
                                fmt_overrides=dlg.fmt_overrides)
        from ..paths import PROFILES_DIR
        out = PROFILES_DIR / f"{name}.fprof.json"
        try:
            save_profile(profile, out)
        except ProfileError as e:
            QMessageBox.critical(self, APP_TITLE, str(e))
            return
        self.refresh_profiles()
        self.log_line(f"[提取] {src} → {out.name}（{len(profile.get('styles', {}))} 类格式）")

    def on_import(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "选择格式方案 JSON", "", "书山格式方案 (*.fprof.json *.fprof *.json)")
        if not path:
            return
        try:
            data = load_profile(path)
        except ProfileError as e:
            QMessageBox.critical(self, APP_TITLE, str(e))
            return
        from ..paths import PROFILES_DIR
        dest = PROFILES_DIR / (Path(path).name if Path(path).suffix == ".fprof"
                               else Path(path).name)
        shutil.copy2(path, dest)
        self.refresh_profiles()
        self.log_line(f"[导入] {Path(path).name} —— {data.get('meta', {}).get('name', '')}")

    def on_render(self):
        ppath = self.selected_profile_path()
        if not ppath:
            return
        try:
            profile = load_profile(ppath)
        except ProfileError as e:
            QMessageBox.critical(self, APP_TITLE, str(e))
            return
        default_out = ppath.with_suffix(".dotx") if ppath.suffix == ".fprof" \
            else ppath.parent / (ppath.stem.replace(".fprof", "") + ".dotx")
        out, _ = QFileDialog.getSaveFileName(self, "保存 Word 模板",
                                             str(default_out), "Word 模板 (*.dotx)")
        if not out:
            return
        try:
            render_dotx(profile, out)
        except Exception as e:
            QMessageBox.critical(self, APP_TITLE, f"渲染失败：{e}")
            return
        self.log_line(f"[渲染] {ppath.name} → {out}\n        "
                      f"在 Word 中：文件→选项→加载项→管理:模板→选用 该模板并勾选"
                      f"「自动更新文档样式」即可换装。")

    def on_apply(self):
        ppath = self.selected_profile_path()
        if not ppath:
            return
        try:
            profile = load_profile(ppath)
        except ProfileError as e:
            QMessageBox.critical(self, APP_TITLE, str(e))
            return

        self._apply_to_file(ppath, profile)

    def _ask_numbering_mode(self, analysis, assignment):
        """整稿前询问标题编号模式；返回 numbering_override 或 None（取消）。"""
        from ..render.apply import detect_numbering_mode
        from .numbering_dialog import NumberingModeDialog
        detected = detect_numbering_mode(analysis, assignment)
        ndlg = NumberingModeDialog(detected, self)
        if ndlg.exec() != NumberingModeDialog.Accepted:
            self.log_line("[整稿] 已取消（编号模式未确认）")
            return None
        override = ndlg.result_override()
        self.log_line(f"[编号模式] {override[0]}"
                      + ("（自定义编号样式）" if override[0] == "custom" else ""))
        return override

    def _apply_to_file(self, ppath: Path, profile: dict):
        src, _ = QFileDialog.getOpenFileName(self, "选择待整稿的文档", "",
                                             "Word 文档 (*.docx *.doc)")
        if not src:
            return
        try:
            work_path, converted = ensure_docx(src)
            if converted:
                self.log_line(f"[转换] {Path(src).name} → {Path(converted).name}"
                              f"（整稿将写入转换件，原 .doc 保持只读）")
            analysis = analyze_document(work_path)
        except Exception as e:
            QMessageBox.critical(self, APP_TITLE, f"分析失败：{e}")
            return
        from .workbench import WorkbenchDialog
        dlg = WorkbenchDialog(analysis, profile, Path(src).name, work_path, self)
        if dlg.exec() != WorkbenchDialog.Accepted:
            return
        if dlg.style_edits:
            import copy as _copy
            profile = _copy.deepcopy(profile)
            profile.setdefault("styles", {}).update(dlg.style_edits)
            names = "、".join(dlg.style_edits)
            self.log_line(f"[样式] 已按所选内容更新类型样式：{names}（仅本次整稿生效）")
        numbering_override = self._ask_numbering_mode(analysis, dlg.assignment)
        if numbering_override is None:
            return
        try:
            report = apply_profile_offline(profile, work_path, para_override=dlg.assignment,
                                           numbering_override=numbering_override,
                                           numbered_map=dlg.numbered,
                                           fmt_overrides=dlg.fmt_overrides)
        except Exception as e:
            QMessageBox.critical(self, APP_TITLE, f"整稿失败：{e}")
            return
        msg = (f"已整稿 {report['paragraphs']} 段："
               + "、".join(f"{DEFAULT_CN_NAMES.get(k, k)}×{v}"
                           for k, v in sorted(report["counts"].items())))
        self.log_line(f"[整稿] {Path(work_path).name}\n        {msg}")
        if report.get("backup"):
            self.log_line(f"        备份：{report['backup']}")
        if report.get("manual_numbering_kept"):
            self.log_line("        检测到手工编号，已保留文字编号、未启用自动编号")
        if report.get("protected_objects"):
            self.log_line(f"        图形/公式段防裁剪保护：{report['protected_objects']} 段")
        if report.get("headings_outline_only"):
            self.log_line(f"        {report['headings_outline_only']} 个标题按对号只设定大纲级别（样式与外观保持原样）")
        if report.get("figures_resized"):
            self.log_line(f"        图片宽度调整：{report['figures_resized']} 张"
                          f"（方案 objects.figure.resize 开启）")
        if report.get("citations_superscripted"):
            self.log_line(f"        引用标记上标规范化：{report['citations_superscripted']} 处"
                          f"（citations.superscript）")
        if report.get("dangling_citations"):
            nums = "、".join(f"[{n}]" for n in report["dangling_citations"])
            self.log_line(f"        ⚠ 悬空引用 {nums}：正文引用了不存在的文献条目")
        if report.get("uncited_entries"):
            nums = "、".join(f"[{n}]" for n in report["uncited_entries"][:8])
            self.log_line(f"        ⚠ {len(report['uncited_entries'])} 条文献未被正文引用：{nums}")
        hf_rep = report.get("header_footer") or {}
        if hf_rep.get("applied"):
            self.log_line(f"        页眉页脚：已写入 {hf_rep.get('sections', 0)} 节"
                          f"（header_footer）")
        toc_rep = report.get("toc") or {}
        if toc_rep.get("inserted"):
            self.log_line("        目录域已插入（Word 中右键 → 更新域 生成条目）")
        elif toc_rep.get("reason"):
            self.log_line(f"        目录域：{toc_rep['reason']}")
        if report.get("table_cells"):
            self.log_line(f"        表格单元格重排：{report['table_cells']} 段"
                          f"（tables.cell_style）")
        for w in report.get("warnings", [])[:10]:
            self.log_line(f"        ⚠ {w}")
        QMessageBox.information(self, APP_TITLE, msg + "\n\n原文件已备份，可用 Word 打开查看效果。")

    # ---------- 小工具 ----------
    def _pick_source_doc(self, action: str):
        """提取格式的来源：磁盘上的文献文件（离线模式）。"""
        path, _ = QFileDialog.getOpenFileName(self, f"选择要{action}的参考文献",
                                              "", "Word 文档 (*.docx *.doc)")
        return path

    def _ask_text(self, title: str, default: str):
        from PySide6.QtWidgets import QInputDialog
        text, ok = QInputDialog.getText(self, title, "名称：", text=default)
        return text.strip(), ok and text.strip()


def run():
    import sys
    from PySide6.QtWidgets import QApplication
    app = QApplication(sys.argv)
    app.setApplicationName("书山")
    win = ShushanPanel()
    win.show()
    sys.exit(app.exec())
