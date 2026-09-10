"""Read-only document identity checks with optional bounded selection reading."""

import argparse
import json
from pathlib import Path

import pythoncom
import pywintypes
import win32com.client


TEST_DOCUMENTS = Path(__file__).resolve().parents[1] / "test-documents"
MK_E_UNAVAILABLE = 0x800401E3
MAX_SELECTION_POSITIONS = 10000


def _read_selection(word, expected: Path) -> dict:
    """Read only a bounded main-text selection belonging to the expected file."""
    selection = selected_range = selected_document = None
    try:
        selection = word.Selection
        selected_range = selection.Range.Duplicate
        selected_document = selected_range.Document
        # Recheck the captured range's own document: the user might switch
        # documents between ActiveDocument and Selection calls.
        folder = str(selected_document.Path)
        full_name = str(selected_document.FullName) if folder else None
        if not (
            full_name
            and Path(full_name).is_absolute()
            and Path(full_name).resolve() == expected
        ):
            return {"status": "selection_document_mismatch"}

        start = int(selected_range.Start)
        end = int(selected_range.End)
        story_type = int(selected_range.StoryType)
        positions = {"start": start, "end": end, "story_type": story_type}
        if start < 0 or end < start:
            return {"status": "invalid_selection"}
        if start == end:
            return {"status": "empty_selection", **positions}
        # First step: ordinary text in the main document body only.
        # wdSelectionNormal = 2, wdMainTextStory = 1.
        if int(selection.Type) != 2 or story_type != 1:
            return {"status": "unsupported_selection", **positions}
        if end - start > MAX_SELECTION_POSITIONS:
            return {"status": "selection_too_large", **positions}

        text = str(selected_range.Text)
        return {
            "status": "selected" if text else "empty_selection",
            **positions,
            "text": text,
            "character_count": len(text),
        }
    finally:
        selected_document = None
        selected_range = None
        selection = None


def check_active_document(expected_document: Path, *, include_selection=False) -> dict:
    """Compare metadata with a test file, optionally reading its selected text.

    Matching the path is not permission to write, and ReadOnly=False does not
    rule out editing protection. Future operations must recheck their target.
    """
    try:
        expected = expected_document.resolve(strict=True)
        if (
            not expected.is_relative_to(TEST_DOCUMENTS.resolve())
            or not expected.is_file()
            or expected.suffix.lower() != ".docx"
        ):
            raise ValueError("Expected a test .docx file")
    except (OSError, ValueError, RuntimeError):
        return {"status": "invalid_test_file", "matches_expected": False}

    initialized = False
    word = document = None
    stage = "initialize"
    try:
        pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
        initialized = True
        stage = "connect"
        word = win32com.client.GetActiveObject("Word.Application")
        stage = "active_document"
        if word.Documents.Count == 0:
            return {"status": "no_document", "matches_expected": False}
        document = word.ActiveDocument
        stage = "metadata"
        name = str(document.Name)
        folder = str(document.Path)
        full_name = str(document.FullName) if folder else None
        read_only = bool(document.ReadOnly)
        # Reject unsaved documents and web URLs rather than treating them as
        # relative local paths. Path comparison follows Windows case rules.
        matches = bool(
            full_name
            and Path(full_name).is_absolute()
            and Path(full_name).resolve() == expected
        )
        result = {
            "status": "matched" if matches else "different_document",
            "name": name,
            "full_path": full_name,
            "read_only": read_only,
            "expected_path": str(expected),
            "matches_expected": matches,
        }
        if matches and include_selection:
            stage = "selection"
            selection_result = _read_selection(word, expected)
            result["status"] = selection_result.pop("status")
            result["selection"] = selection_result
        return result
    except pywintypes.com_error as error:
        code = error.hresult & 0xFFFFFFFF
        return {
            "status": (
                "not_available"
                if stage == "connect" and code == MK_E_UNAVAILABLE
                else "com_error"
            ),
            "matches_expected": False,
            "stage": stage,
            "hresult": f"0x{code:08X}",
        }
    except (OSError, ValueError, RuntimeError):
        return {"status": "path_error", "matches_expected": False, "stage": stage}
    finally:
        document = None
        word = None
        if initialized:
            pythoncom.CoUninitialize()


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--expected-document", required=True, type=Path)
    args = parser.parse_args()
    result = check_active_document(args.expected_document)
    print(json.dumps(result, ensure_ascii=True, indent=2))
    return 0 if result["matches_expected"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
