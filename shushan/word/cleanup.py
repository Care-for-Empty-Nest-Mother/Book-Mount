# -*- coding: utf-8 -*-
"""Word 后台进程清理。

COM 自动化偶尔会在系统里留下看不见的 WINWORD.EXE 后台实例。本模块：
- 用 WMI 枚举所有 WINWORD.EXE 进程；
- 用窗口枚举标记"有可见窗口"的 Word（用户正在使用的）——这些绝不杀；
- 只结束无可见窗口的后台残留实例。
"""
import win32api
import win32com.client
import win32gui
import win32process
import pythoncom

PROCESS_TERMINATE = 0x0001


def _visible_window_pids() -> set:
    """当前有可见、有标题顶层窗口的进程 PID 集合。"""
    visible = set()

    def _enum(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if title.strip():
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                visible.add(pid)

    win32gui.EnumWindows(_enum, None)
    return visible


def list_word_processes() -> list:
    """返回 [(pid, command_line, is_background), ...]。"""
    pythoncom.CoInitialize()
    wmi = win32com.client.GetObject(r"winmgmts://./root/cimv2")
    procs = wmi.ExecQuery(
        "SELECT ProcessId, CommandLine FROM Win32_Process WHERE Name='WINWORD.EXE'")
    visible = _visible_window_pids()
    out = []
    for p in procs:
        pid = int(p.ProcessId)
        cmd = (p.CommandLine or "")
        out.append((pid, cmd, pid not in visible))
    return out


def cleanup_word_background() -> dict:
    """结束后台残留的 Word 实例（有可见窗口的一律不动）。

    返回 {"found": n, "killed": [pid], "kept": [pid], "failed": [(pid, 原因)]}
    """
    procs = list_word_processes()
    killed, kept, failed = [], [], []
    for pid, _cmd, is_background in procs:
        if not is_background:
            kept.append(pid)
            continue
        try:
            h = win32api.OpenProcess(PROCESS_TERMINATE, False, pid)
            win32api.TerminateProcess(h, 0)
            win32api.CloseHandle(h)
            killed.append(pid)
        except Exception as e:
            failed.append((pid, str(e)))
    return {"found": len(procs), "killed": killed, "kept": kept, "failed": failed}
