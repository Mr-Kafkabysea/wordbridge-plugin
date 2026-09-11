"""Expose the tested Word operations over local stdio."""

import json
from pathlib import Path
from threading import Lock
from typing import Annotated, Any, Literal, NotRequired, TypedDict

from mcp.server import MCPServer
from mcp.server.mcpserver import Elicit, ElicitationResult, Resolve
from mcp.server.mcpserver.exceptions import ToolError
from mcp.types import ToolAnnotations
from pydantic import BaseModel, ConfigDict, Field, StrictBool

if __package__:
    from .append_text import append_text as append_backend
    from .check_active_document import check_active_document
    from .check_word_connection import check_word_connection
else:
    from append_text import append_text as append_backend
    from check_active_document import check_active_document
    from check_word_connection import check_word_connection

PROJECT = Path(__file__).resolve().parents[1]
TOOL_NAMES = ["word_status", "get_active_document", "get_selection",
              "preview_append_text", "append_text"]


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


def create_server(status=None) -> MCPServer:
    # Serialize complete COM calls, including initialization and cleanup.
    # The lock is never held while waiting for the user's confirmation.
    com_lock = Lock()

    def invoke(operation, *args, **kwargs):
        with com_lock:
            if status:
                first = status.first_contact(check_word_connection)
                if first is not None and operation is check_word_connection:
                    return first
            return operation(*args, **kwargs)

    server = MCPServer(
        "WordBridge Plugin",
        version="0.3.0",
        lifespan=status.lifespan if status else None,
        middleware=[status] if status else None,
        instructions=(
            "Operate only on existing active Word documents. Never open, "
            "switch, save or close documents. Preview append_text first, then "
            "pass the same text and state to append_text. The client must show "
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
                         expected_state: str) -> Elicit[AppendApproval]:
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
        title="确认后在文末追加文字",
        annotations=ToolAnnotations(read_only_hint=False, destructive_hint=True,
                                    idempotent_hint=False, open_world_hint=False),
    )
    def append_text(expected_document: str, text: str, expected_state: str,
                    approval: Annotated[ElicitationResult[AppendApproval],
                                        Resolve(request_approval)]) -> dict[str, Any]:
        """Append the previously previewed text only after human form confirmation.

        Requires a client that presents elicitation to the user. Unsupported or
        invalid confirmation yields an MCP error without writing. Decline/cancel
        refuses writing. Rechecks state after confirmation. Never saves; one undo
        step. Inspect Word before retrying any failed or uncertain write.
        """
        if approval.action != "accept":
            return {"status": "approval_" + approval.action, "write_attempted": False}
        if AppendApproval.model_validate(approval.data).confirm is not True:
            return {"status": "approval_decline", "write_attempted": False}
        return invoke(append_backend, document_path(expected_document), text,
                      apply=True, expected_state=expected_state)

    return server


server = create_server()


if __name__ == "__main__":
    # stdout is the protocol channel: do not print diagnostic messages here.
    if __package__:
        from .connection_status import ConnectionStatus
    else:
        from connection_status import ConnectionStatus
    create_server(ConnectionStatus(popup=True)).run(transport="stdio")
