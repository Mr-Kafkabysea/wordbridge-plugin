"""First-call notification; reads one instance snapshot and never accesses Word."""

import json
from pathlib import Path
import sys
import time
import tkinter as tk
from tkinter import ttk


def presentation(data, now):
    if data.get("service") != "running" or not 0 <= now - data.get("heartbeat", 0) <= 15:
        return "连接状态待确认", "服务已停止或状态过期，请重新检查。", False
    state = data.get("first_check")
    if state == "connected":
        if now - data.get("first_checked_at", 0) > 15:
            return "连接状态待确认", "首次检查结果已过期，请重新检查。", False
        return "WordBridge 已就绪", "✓ MCP 通信正常\n✓ 已连接 Microsoft Word", True
    if state == "not_available":
        return "MCP 已连接，Word 暂不可访问", "请确认桌面 Word 已打开且处于可访问的用户会话。\n随后可让 Agent 再检查 Word 状态。", False
    if state in {"com_error", "unknown"}:
        return "MCP 已连接，Word 检查失败", "请检查 Word 和会话权限，具体操作结果以工具返回为准。", False
    return "正在检查 Word…", "✓ 已收到 MCP 工具调用\n正在只读检查 Word，不读取正文。", False


def create_popup(path):
    root = tk.Tk()
    root.title("WordBridge")
    root.geometry("520x260")
    root.resizable(False, False)
    frame = ttk.Frame(root, padding=24)
    frame.pack(fill="both", expand=True)
    title = ttk.Label(frame, text="正在连接 WordBridge…", font=("Microsoft YaHei UI", 16, "bold"))
    title.pack(anchor="w", pady=(0, 14))
    detail = ttk.Label(frame, text="等待连接检查结果", wraplength=470, font=("Microsoft YaHei UI", 10))
    detail.pack(anchor="w")
    footer = ttk.Label(frame, text="连接成功不代表当前文档可编辑；后续仍会核对操作条件。", wraplength=470)
    footer.pack(anchor="w", pady=(16, 8))
    ttk.Button(frame, text="关闭", command=root.destroy).pack(anchor="e")
    started = time.monotonic()
    closing = False

    def refresh():
        nonlocal closing
        try:
            data = json.loads(Path(path).read_text(encoding="utf-8"))
            heading, message, success = presentation(data, time.time())
        except (OSError, ValueError, TypeError):
            heading, message, success = "等待连接状态…", "暂未能读取状态记录。", False
        if not success and time.monotonic() - started > 20 and heading.startswith(("正在", "等待")):
            heading, message = "连接检查尚未完成", "请查看工具调用结果；本窗口不会重试操作。"
        title.config(text=heading)
        detail.config(text=message)
        if success and not closing:
            closing = True
            footer.config(text="首次连接检查成功，此窗口将在 3 秒后关闭。")
            root.after(3000, root.destroy)
        if not closing:
            root.after(300, refresh)

    refresh()
    return root


if __name__ == "__main__":
    create_popup(Path(sys.argv[1])).mainloop()
