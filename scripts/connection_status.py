"""Local, metadata-only telemetry. Never writes to MCP stdout or touches Word."""

from contextlib import asynccontextmanager
import json
import os
import subprocess
import sys
from pathlib import Path
from threading import Event, Lock, Thread
import time
from uuid import uuid4

STATUS_DIR = Path(__file__).resolve().parents[1] / "local" / "runtime" / "connections"
STALE_SECONDS = 15


class ConnectionStatus:
    def __init__(self, directory=STATUS_DIR, *, popup=False):
        self.path = Path(directory) / (uuid4().hex + ".json")
        self.lock = Lock()
        self.first_lock = Lock()
        self.first_started = False
        self.popup = popup
        self.state = {"pid": os.getpid(), "service": "starting", "mcp": "waiting",
                      "first_check": "pending", "first_checked_at": None,
                      "word": "unchecked", "word_checked_at": None,
                      "last_tool": None, "last_tool_at": None, "heartbeat": 0}

    def launch_popup(self):
        executable = Path(sys.executable).with_name("pythonw.exe")
        if not executable.is_file():
            executable = Path(sys.executable)
        subprocess.Popen(
            [str(executable), str(Path(__file__).with_name("connection_popup.py")), str(self.path)],
            stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )

    def first_contact(self, probe):
        """Called under the server COM lock, after a valid tool reaches its body.

        Returns the initial probe result once; never authorizes or retries a write.
        """
        with self.first_lock:
            if self.first_started:
                return None
            self.first_started = True
            self.update(first_check="checking", mcp="request_received")
            if self.popup:
                try:
                    self.launch_popup()
                except OSError:
                    self.update(popup_error="launch_failed")
            try:
                result = probe()
            except Exception:
                # Diagnostic failure must not replace the requested tool result.
                result = {"status": "com_error", "connected": False}
            outcome = result.get("status")
            if outcome not in {"connected", "not_available", "com_error"}:
                outcome = "unknown"
            self.update(first_check=outcome, first_checked_at=time.time())
            self.word_result(result)
            return result

    def update(self, **fields):
        # Telemetry failure must never break a tool or leak arguments/content.
        with self.lock:
            self.state.update(fields)
            self.state["heartbeat"] = time.time()
            temporary = self.path.with_suffix(".tmp")
            try:
                self.path.parent.mkdir(parents=True, exist_ok=True)
                temporary.write_text(json.dumps(self.state), encoding="utf-8")
                os.replace(temporary, self.path)
            except OSError:
                pass

    @asynccontextmanager
    async def lifespan(self, server):
        stop = Event()
        self.update(service="running")

        def pulse():
            while not stop.wait(2):
                self.update()

        worker = Thread(target=pulse, daemon=True)
        worker.start()
        try:
            yield {}
        finally:
            stop.set()
            worker.join(timeout=3)
            self.update(service="stopped", mcp="disconnected")

    async def __call__(self, ctx, call_next):
        if ctx.method == "tools/call":
            name = (ctx.params or {}).get("name")
            # Only our known tool names; never persist arbitrary client input.
            if name in {"word_status", "get_active_document", "get_selection",
                        "preview_append_text", "append_text"}:
                self.update(last_tool=name, last_tool_at=time.time())
        result = await call_next(ctx)
        if ctx.method == "initialize":
            self.update(mcp="initializing")
        elif ctx.method == "notifications/initialized":
            self.update(mcp="initialized")
        elif ctx.method in {"tools/list", "tools/call"}:
            # Some protocol versions do not perform the legacy handshake.
            if self.state["mcp"] != "initialized":
                self.update(mcp="request_received")
        return result

    def word_result(self, result):
        self.update(word=result.get("status", "unknown"), word_checked_at=time.time())


def display_state(data, now=None):
    now = time.time() if now is None else now
    fresh = 0 <= now - data.get("heartbeat", 0) <= STALE_SECONDS
    running = fresh and data.get("service") == "running"
    service = "运行中" if running else ("已停止" if data.get("service") == "stopped" else "心跳过期 / 状态未知")
    mcp = {"waiting": "等待客户端", "initializing": "等待握手完成",
           "initialized": "握手已完成", "request_received": "已收到协议请求",
           "disconnected": "已断开"}.get(data.get("mcp"), "未知")
    if not running:
        mcp = "未确认在线"
    word = {"unchecked": "未检查", "connected": "可访问（上次检查）",
            "not_available": "不可访问（上次检查）", "com_error": "COM 错误（上次检查）"}.get(data.get("word"), "未知")
    def stamp(value):
        return time.strftime("%m-%d %H:%M:%S", time.localtime(value)) if value else "—"
    return (data.get("pid", "—"), service, mcp, word,
            stamp(data.get("word_checked_at")), data.get("last_tool") or "未收到工具调用",
            stamp(data.get("last_tool_at")))


def snapshots(directory=STATUS_DIR):
    rows = []
    for path in Path(directory).glob("*.json"):
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                rows.append((path.stem, display_state(data)))
        except (OSError, ValueError, TypeError, OverflowError):
            continue
    return rows
