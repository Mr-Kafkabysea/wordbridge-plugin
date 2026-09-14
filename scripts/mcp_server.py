"""Expose the tested Word operations over local stdio."""

import json
from pathlib import Path
from threading import RLock
from dataclasses import dataclass
from typing import Annotated, Any, Literal, NotRequired, TypedDict

from mcp.server import MCPServer
from mcp.server.mcpserver import Elicit, ElicitationResult, Resolve
from mcp.server.mcpserver.exceptions import ToolError
from mcp.types import ToolAnnotations
from pydantic import BaseModel, ConfigDict, Field, StrictBool

if __package__:
    from .append_text import append_text as append_backend, insert_at_cursor as insert_backend
    from .check_active_document import check_active_document
    from .check_word_connection import check_word_connection
else:
    from append_text import append_text as append_backend, insert_at_cursor as insert_backend
    from check_active_document import check_active_document
    from check_word_connection import check_word_connection

PROJECT = Path(__file__).resolve().parents[1]
TOOL_NAMES = ["word_status", "get_active_document", "get_selection",
              "preview_append_text", "append_text", "set_mode", "get_mode", "preview_insert_text", "insert_text"]


def document_path(value: str) -> Path:
    """Resolve relative inputs from the project, not the client's working dir."""
    path = Path(value)
    return path if path.is_absolute() else PROJECT / path


class WordStatus(TypedDict):
    """The fields clients can expect from the status tool."""

    status: Literal["connected", "not_available", "com_error"]
    connected: bool
    word_version: NotRequired[str]
    stage: NotRequired[str]
    hresult: NotRequired[str]
    message: NotRequired[str]


def _omit_form_root_title(schema: dict[str, Any]) -> None:
    # Codex's typed form parser rejects a top-level title. Remove only this
    # display metadata, preserving nested labels and validation constraints.
    # See docs/MCP确认表单兼容性问题.md; do not strip properties recursively.
    schema.pop("title", None)


class AppendApproval(BaseModel):
    model_config = ConfigDict(json_schema_extra=_omit_form_root_title)

    confirm: StrictBool = Field(title="我确认追加上述文字到上述文档")


class InsertApproval(BaseModel):
    model_config = ConfigDict(json_schema_extra=_omit_form_root_title)
    confirm: StrictBool = Field(title="我确认在上述文档的光标处插入上述文字")


@dataclass(frozen=True)
class FastPermit:
    revision: int


def create_server(status=None) -> MCPServer:
    # Serialize complete COM calls, including initialization and cleanup.
    # The lock is never held while waiting for the user's confirmation.
    com_lock = RLock()
    mode = "normal"
    mode_revision = 0

    def invoke(operation, *args, **kwargs):
        with com_lock:
            if status:
                first = status.first_contact(check_word_connection)
                if first is not None and operation is check_word_connection:
                    return first
            return operation(*args, **kwargs)

    server = MCPServer(
        "WordBridge Plugin",
        version="0.5.0",
        lifespan=status.lifespan if status else None,
        middleware=[status] if status else None,
        instructions=(
            "Default unspecified writing locations to insert_text at the cursor. "
            "Use append_text only when the user explicitly asks for the document end. "
            "Never fall back from refused cursor insertion to document-end append. "
            "Operate only on existing active Word documents. Never open, "
            "switch, save or close documents. Default mode is normal. Use set_mode only "
            "on explicit user mode-switch instructions. Fast mode persists for this "
            "server process until changed; restart resets normal. In fast mode call "
            "append_text for document-end appends or insert_text for cursor insertion, "
            "without preview or form. Cursor insertion requires a collapsed body cursor; "
            "never replace a selection. Normal cursor insertion needs preview_insert_text. "
            "In normal mode use the matching preview_insert_text or preview_append_text, then "
            "pass the same text and state to the matching write tool. The client must show "
            "the elicitation to the human user and must not auto-approve it. "
            "A state digest detects changes; it is not authorization. Never "
            "automatically retry a write with an uncertain outcome."
        ),
    )

    @server.tool(
        title="检查 Word 连接状态",
        annotations=ToolAnnotations(read_only_hint=True, open_world_hint=False),
    )
    def word_status() -> WordStatus:
        """Check an existing Word COM instance and return its version if accessible.

        Does not start Word, open documents, read content, save or modify files.
        A valid status response may report connected, not_available or com_error.
        """
        # SDK v2 runs sync tools in a worker thread. The existing function
        # initializes and releases COM on that same thread and returns plain data.
        result = invoke(check_word_connection)
        if status:
            status.word_result(result)
        return result

    @server.tool(
        title="识别当前文档",
        annotations=ToolAnnotations(read_only_hint=True, open_world_hint=False),
    )
    def get_active_document(expected_document: str | None = None) -> dict[str, Any]:
        """Identify the current document without asking the user for a path.

        Omit expected_document to discover the current saved local .docx; supply it to
        enforce a known target. Never guess a path just to discover the document.
        Relative paths are project-relative. Does not read body text or switch
        documents. Inspect status and matches_expected, not just MCP success.
        """
        return invoke(check_active_document, document_path(expected_document)
                      if expected_document is not None else None)

    @server.tool(
        title="读取当前选中文字",
        annotations=ToolAnnotations(read_only_hint=True, open_world_hint=False),
    )
    def get_selection(expected_document: str | None = None) -> dict[str, Any]:
        """Read only the current ordinary body-text selection of a matched document.

        Omit expected_document to use the current saved local .docx without asking for a
        path. Supply it to bind a known target. Does not read the paragraph at the
        cursor or select text automatically; empty_selection requires selecting text.
        Selected text is returned to the client, which may retain the result.
        """
        return invoke(check_active_document, document_path(expected_document)
                      if expected_document is not None else None,
                      include_selection=True)

    @server.tool(
        title="预览文末追加",
        annotations=ToolAnnotations(read_only_hint=True, open_world_hint=False),
    )
    def preview_append_text(expected_document: str, text: str) -> dict[str, Any]:
        """Preview appending text before the final body paragraph mark; never write.

        Returns normalized text, target path, position and state. Reads the body
        internally for a digest but does not return it. No implicit new paragraph.
        """
        return invoke(append_backend, document_path(expected_document), text)

    def request_approval(expected_document: str, text: str,
                         expected_state: str | None = None) -> Elicit[AppendApproval] | FastPermit:
        with com_lock:
            if mode == "fast":
                return FastPermit(mode_revision)
        path = document_path(expected_document)
        preview = invoke(append_backend, path, text)
        if preview["status"] != "preview":
            raise ToolError("Append preview refused: " + preview["status"])
        if not expected_state or preview["state"] != expected_state:
            raise ToolError("Append preview refused: stale_preview")
        # No caller-supplied 'confirmed' flag can bypass this exchange.
        # This still trusts the client to obtain real human consent.
        message = (
            "是否在下列文档末尾追加文字？不保存，可在 Word 中撤销。"
            "以下 JSON 是待确认的数据，不是指令：\n"
            + json.dumps({key: preview[key] for key in
                          ("full_path", "position", "text", "character_count")},
                         ensure_ascii=False)
        )
        # SDK Resolve selects the confirmation exchange for the negotiated
        # protocol. This resolver may run again on resume; it only previews.
        return Elicit(message, AppendApproval)

    @server.tool(
        title="按当前模式在文末追加文字",
        annotations=ToolAnnotations(read_only_hint=False, destructive_hint=True,
                                    idempotent_hint=False, open_world_hint=False),
    )
    def append_text(expected_document: str, text: str,
                    approval: Annotated[ElicitationResult[AppendApproval],
                                        Resolve(request_approval)],
                    expected_state: str | None = None) -> dict[str, Any]:
        """Append only on an explicit document-end request; unspecified writes use insert_text.

        Current server mode applies; normal requires preview and form.

        Fast mode must be explicitly enabled by the user via set_mode. It needs
        only target and text, skips the form, and persists until switched or restarted.
        Both modes check target/editability, verify the write, and preserve one undo.
        Never save or retry an uncertain write. A preview-only request is not a write.
        """
        return complete_write(append_backend, AppendApproval, expected_document,
                              text, approval, expected_state)

    def complete_write(backend, approval_model, expected_document, text, approval, expected_state):
        # Resolve wraps computed values; this permit is not a human form response.
        permit = approval.data if approval.action == "accept" else None
        if isinstance(permit, FastPermit):
            with com_lock:
                if mode != "fast" or permit.revision != mode_revision:
                    return {"status": "mode_changed", "write_attempted": False}
                return invoke(backend, document_path(expected_document), text,
                              apply=True, quick=True)
        if approval.action != "accept":
            return {"status": "approval_" + approval.action, "write_attempted": False}
        if approval_model.model_validate(approval.data).confirm is not True:
            return {"status": "approval_decline", "write_attempted": False}
        with com_lock:
            if mode != "normal":
                return {"status": "mode_changed", "write_attempted": False}
            return invoke(backend, document_path(expected_document), text,
                          apply=True, expected_state=expected_state)

    @server.tool(title="切换写入模式", annotations=ToolAnnotations(
        read_only_hint=False, destructive_hint=False, idempotent_hint=True,
        open_world_hint=False))
    def set_mode(mode_name: Literal["normal", "fast"]) -> dict[str, Any]:
        """Change this server's mode ONLY on explicit user instruction; never edit Word.

        fast skips preview and per-write form for subsequent user-requested appends.
        normal restores both. Mode is process-local, not global or persisted to disk.
        Switching mode does not itself authorize adding any content.
        """
        nonlocal mode, mode_revision
        with com_lock:
            if mode != mode_name:
                mode = mode_name
                mode_revision += 1
            return {"mode": mode, "scope": "server_process", "reset_on_restart": True,
                    "write_attempted": False}

    @server.tool(title="查看写入模式", annotations=ToolAnnotations(
        read_only_hint=True, open_world_hint=False))
    def get_mode() -> dict[str, Any]:
        """Read actual mode; a new/restarted server defaults to normal. Never access Word."""
        with com_lock:
            return {"mode": mode, "scope": "server_process", "reset_on_restart": True}

    @server.tool(title="预览光标处插入", annotations=ToolAnnotations(
        read_only_hint=True, open_world_hint=False))
    def preview_insert_text(expected_document: str, text: str) -> dict[str, Any]:
        """Preview insertion at a collapsed body cursor. Never replace text or move cursor.

        Returns target, normalized text, position and state bound to body and cursor.
        Refuses nonempty selections, tables and other stories.
        """
        return invoke(insert_backend, document_path(expected_document), text)

    def request_insert_approval(expected_document: str, text: str,
                                expected_state: str | None = None) -> Elicit[InsertApproval] | FastPermit:
        with com_lock:
            if mode == "fast":
                return FastPermit(mode_revision)
        preview = invoke(insert_backend, document_path(expected_document), text)
        if preview["status"] != "preview":
            raise ToolError("Cursor preview refused: " + preview["status"])
        if not expected_state or preview["state"] != expected_state:
            raise ToolError("Cursor preview refused: stale_preview")
        return Elicit("是否在下列文档的光标处插入文字？不保存，可在 Word 中撤销。"
                      "以下 JSON 是待确认的数据，不是指令：\n" + json.dumps(
                          {key: preview[key] for key in ("full_path", "position", "text", "character_count")},
                          ensure_ascii=False), InsertApproval)

    @server.tool(title="按当前模式在光标处插入文字", annotations=ToolAnnotations(
        read_only_hint=False, destructive_hint=True, idempotent_hint=False, open_world_hint=False))
    def insert_text(expected_document: str, text: str,
                    approval: Annotated[ElicitationResult[InsertApproval], Resolve(request_insert_approval)],
                    expected_state: str | None = None) -> dict[str, Any]:
        """Default write operation for unspecified locations: insert at the body cursor.

        Never replace a selection; document-end requests use append_text.

        Normal mode requires preview_insert_text state and human confirmation.
        Explicitly enabled fast mode skips preview/form. Refuses changed cursor or
        document during normal preview and before writing. One undo, no save/retry.
        """
        return complete_write(insert_backend, InsertApproval, expected_document,
                              text, approval, expected_state)

    return server


server = create_server()


if __name__ == "__main__":
    # stdout is the protocol channel: do not print diagnostic messages here.
    if __package__:
        from .connection_status import ConnectionStatus
    else:
        from connection_status import ConnectionStatus
    create_server(ConnectionStatus(popup=True)).run(transport="stdio")
