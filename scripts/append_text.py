"""Preview or explicitly apply a plain-text append to a Word document."""

import argparse
import hashlib
import json
from pathlib import Path

import pythoncom
import pywintypes
import win32com.client

if __package__:
    from .check_active_document import document_file_path
else:
    from check_active_document import document_file_path


MAX_TEXT_LENGTH = 10000
MAX_DOCUMENT_POSITIONS = 100000
UNDO_NAME = "WordBridge append text"


class Refused(Exception):
    """A named precondition failure with no document contents attached."""


def _normalize_text(text: str) -> str:
    if not isinstance(text, str):
        raise Refused("invalid_text")
    text = text.replace("\r\n", "\r").replace("\n", "\r")
    if not text or len(text) > MAX_TEXT_LENGTH:
        raise Refused("invalid_text")
    if any(ord(char) < 32 and char not in "\r\t" for char in text):
        raise Refused("unsupported_control_character")
    return text


def _snapshot(document, expected: Path, text: str) -> dict:
    """Read bounded document content for change detection, returning only a digest."""
    if not str(document.Path):
        raise Refused("different_document")
    full_name = str(document.FullName)
    if not Path(full_name).is_absolute() or Path(full_name).resolve() != expected:
        raise Refused("different_document")
    if bool(document.ReadOnly):
        raise Refused("read_only")
    if int(document.ProtectionType) != -1:  # wdNoProtection
        raise Refused("protected_document")
    # Rich structures/revision tracking require their own acceptance tests.
    if bool(document.TrackRevisions) or int(document.Revisions.Count):
        raise Refused("revisions_not_supported")
    if int(document.ContentControls.Count):
        raise Refused("content_controls_not_supported")
    content = None
    try:
        content = document.Content
        end = int(content.End)
        if not 1 <= end <= MAX_DOCUMENT_POSITIONS:
            raise Refused("document_size_not_supported")
        body = str(content.Text)
        if not body.endswith("\r"):
            raise Refused("unsupported_document_end")
        state = json.dumps(
            [str(expected), end, body, text], ensure_ascii=True, separators=(",", ":")
        ).encode("ascii")
        return {
            "end": end,
            "state": hashlib.sha256(state).hexdigest(),
        }
    finally:
        content = None


def _cursor_position(word, expected: Path, document_end: int) -> int:
    selection = selected = owner = None
    try:
        selection = word.Selection
        selected = selection.Range.Duplicate
        owner = selected.Document
        if not str(owner.Path) or Path(str(owner.FullName)).resolve() != expected:
            raise Refused("selection_document_mismatch")
        if int(selected.StoryType) != 1:
            raise Refused("unsupported_cursor_story")
        start, end = int(selected.Start), int(selected.End)
        if start != end:
            raise Refused("selection_not_collapsed")
        if int(selection.Type) != 1 or not 0 <= start < document_end:
            raise Refused("invalid_cursor")
        if bool(selected.Information(12)):
            raise Refused("cursor_in_table")
        return start
    finally:
        owner = selected = selection = None


def _cursor_state(snapshot, position):
    return hashlib.sha256(json.dumps(
        ["insert_at_cursor", snapshot["state"], position],
        separators=(",", ":")).encode("ascii")).hexdigest()


def insert_at_cursor(expected_document: Path, text: str, *, apply=False,
                     expected_state=None, quick=False) -> dict:
    """Insert only at a collapsed body cursor; never replace selected text."""
    return append_text(expected_document, text, apply=apply,
                       expected_state=expected_state, quick=quick, at_cursor=True)


def append_text(expected_document: Path, text: str, *, apply=False, expected_state=None, quick=False, at_cursor=False) -> dict:
    """Default to preview. An explicit apply must match the preview state.

    The state digest detects changes; it is NOT a user-authorization token.
    The invoking client/operator must obtain approval before requesting apply.
    quick=True is internal to an explicitly enabled MCP fast mode, not a CLI option.
    It skips a prior preview, retaining same-call snapshots and all write checks.
    """
    try:
        text = _normalize_text(text)
        expected = document_file_path(expected_document)
        if apply and not quick and not expected_state:
            raise Refused("preview_required")
    except Refused as error:
        return {"status": str(error), "write_attempted": False}
    except (OSError, ValueError, RuntimeError):
        return {"status": "invalid_document_file", "write_attempted": False}

    initialized = False
    started_record = False
    write_attempted = False
    word = document = insertion = undo = active_check = None
    stage = "initialize"
    result = None
    try:
        pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
        initialized = True
        stage = "connect"
        word = win32com.client.GetActiveObject("Word.Application")
        if int(word.Documents.Count) == 0:
            raise Refused("no_document")
        # Conservative first version: do not edit while a protected-view
        # window is present, even if Word exposes another ActiveDocument.
        if int(word.ProtectedViewWindows.Count):
            raise Refused("protected_view_present")
        document = word.ActiveDocument
        stage = "snapshot"
        snapshot = _snapshot(document, expected, text)
        position = (_cursor_position(word, expected, snapshot["end"]) if at_cursor
                    else snapshot["end"] - 1)
        preview_state = _cursor_state(snapshot, position) if at_cursor else snapshot["state"]
        undo_name = "WordBridge insert at cursor" if at_cursor else UNDO_NAME
        insertion = document.Range(position, position)
        if bool(insertion.Information(12)):  # wdWithInTable
            raise Refused("document_ends_in_table")
        undo = word.UndoRecord
        if int(undo.CustomRecordLevel) != 0:
            raise Refused("undo_record_busy")
        result = {
            "status": "preview",
            "full_path": str(expected),
            "position": position,
            "text": text,
            "character_count": len(text),
            "state": preview_state,
            "write_attempted": False,
        }
        if apply:
            if not quick and expected_state != preview_state:
                raise Refused("stale_preview")
            # Recheck the active target and captured document just before writing.
            active_check = word.ActiveDocument
            current = _snapshot(active_check, expected, text)
            active_check = None
            if current != snapshot or _snapshot(document, expected, text) != snapshot:
                raise Refused("stale_preview")
            if int(insertion.Start) != position or int(insertion.End) != position:
                raise Refused("stale_preview")
            if at_cursor and _cursor_position(word, expected, snapshot["end"]) != position:
                raise Refused("stale_preview")
            stage = "start_undo_record"
            undo.StartCustomRecord(undo_name)
            started_record = True
            stage = "insert"
            write_attempted = True
            insertion.InsertAfter(text)
            stage = "end_undo_record"
            undo.EndCustomRecord()
            started_record = False
            stage = "verify_inserted_text"
            inserted_end = int(insertion.End)
            if str(insertion.Text) != text:
                result = {
                    "status": "verification_failed", "write_attempted": True,
                    "message": "Inspect Word before retrying; the append may already exist.",
                }
            else:
                cursor_advanced = False
                if at_cursor:
                    # Do not reclaim a cursor the user moved during the operation.
                    stage = "advance_cursor"
                    try:
                        live_position = _cursor_position(word, expected, snapshot["end"] + inserted_end - position)
                    except Refused:
                        live_position = None
                    if live_position in (position, inserted_end):
                        word.Selection.SetRange(inserted_end, inserted_end)
                        cursor_advanced = True
                result = {
                    "status": "inserted" if at_cursor else "appended", "full_path": str(expected),
                    "start": position, "end": inserted_end,
                    "character_count": len(text), "undo_name": undo_name,
                    "write_attempted": True,
                }
                if at_cursor:
                    result["cursor_advanced"] = cursor_advanced
    except Refused as error:
        result = {"status": str(error), "write_attempted": write_attempted}
    except pywintypes.com_error as error:
        result = {
            "status": "write_outcome_unknown" if write_attempted else "com_error",
            "stage": stage, "hresult": f"0x{error.hresult & 0xFFFFFFFF:08X}",
            "write_attempted": write_attempted,
        }
    except (OSError, ValueError, RuntimeError):
        result = {
            "status": "write_outcome_unknown" if write_attempted else "path_error",
            "stage": stage, "write_attempted": write_attempted,
        }
    finally:
        if started_record:
            try:
                undo.EndCustomRecord()
            except pywintypes.com_error:
                if result is not None:
                    result["undo_cleanup_failed"] = True
        active_check = insertion = undo = document = word = None
        if initialized:
            pythoncom.CoUninitialize()
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--expected-document", required=True, type=Path)
    parser.add_argument("--text", required=True)
    parser.add_argument("--apply", action="store_true", help="Write only after user approval")
    parser.add_argument("--expected-state", help="State digest from the approved preview")
    args = parser.parse_args()
    result = append_text(
        args.expected_document, args.text, apply=args.apply, expected_state=args.expected_state
    )
    print(json.dumps(result, ensure_ascii=True, indent=2))
    return 0 if result["status"] in {"preview", "appended"} else 1


if __name__ == "__main__":
    raise SystemExit(main())
