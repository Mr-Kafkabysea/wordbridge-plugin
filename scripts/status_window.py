"""Show local service telemetry; opening this window does not connect to Word."""

import tkinter as tk
from tkinter import ttk

if __package__:
    from .connection_status import snapshots, STATUS_DIR
else:
    from connection_status import snapshots, STATUS_DIR


def create_window():
    root = tk.Tk()
    root.title("WordBridge Plugin · 连接状态")
    root.geometry("1120x360")
    root.minsize(900, 320)
    frame = ttk.Frame(root, padding=20)
    frame.pack(fill="both", expand=True)
    ttk.Label(frame, text="WordBridge Plugin 连接状态", font=("Microsoft YaHei UI", 18, "bold")).pack(anchor="w")
    ttk.Label(frame, text="每个服务实例单独显示 · 每 2 秒刷新 · 此窗口不连接或修改 Word").pack(anchor="w", pady=(6, 16))
    columns = ("进程", "服务", "MCP 客户端", "Word", "Word 检查时间", "最近收到的工具", "调用时间")
    tree = ttk.Treeview(frame, columns=columns, show="headings", height=5)
    for col, width in zip(columns, (65, 170, 160, 170, 120, 190, 120)):
        tree.heading(col, text=col)
        tree.column(col, width=width, minwidth=60)
    tree.pack(fill="both", expand=True)
    scrollbar = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
    scrollbar.pack(fill="x")
    tree.configure(xscrollcommand=scrollbar.set)
    note = ttk.Label(frame, text="")
    note.pack(anchor="w", pady=(12, 4))
    ttk.Label(frame, text="握手 ≠ Agent 已调用工具；Word 显示首次自动检查或最近 word_status 的结果，不是持续连接。\n未出现实例可能是旧版服务尚未重启，或状态文件不可写。模型名称不作推断。").pack(anchor="w")

    def refresh():
        for item in tree.get_children():
            tree.delete(item)
        rows = snapshots()
        for identifier, values in rows:
            tree.insert("", "end", iid=identifier, values=values)
        note.config(text=f"本地观测实例：{len(rows)}　|　{STATUS_DIR}")
        root.after(2000, refresh)

    refresh()
    return root


if __name__ == "__main__":
    create_window().mainloop()
