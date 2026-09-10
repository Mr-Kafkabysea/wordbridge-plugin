"""Simulate append, refusal and uncertain COM outcomes; no Word is opened."""

from contextlib import ExitStack
from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

import pywintypes

from scripts import append_text as probe


class FakeRange:
    def __init__(self, document, start, end):
        self.document = document
        self.Start = start
        self.End = end
        self.Information = Mock(return_value=False)
        self.InsertAfter = Mock(side_effect=self.insert)

    @property
    def Text(self):
        return self.document.body[self.Start:self.End]

    def insert(self, text):
        self.document.body = (
            self.document.body[:self.End] + text + self.document.body[self.End:]
        )
        self.End += len(text)


class FakeDocument:
    def __init__(self, path):
        self.Path = str(path.parent)
        self.FullName = str(path)
        self.ReadOnly = False
        self.ProtectionType = -1
        self.TrackRevisions = False
        self.Revisions = SimpleNamespace(Count=0)
        self.ContentControls = SimpleNamespace(Count=0)
        self.body = "Original\r"
        self.ranges = []

    @property
    def Content(self):
        return SimpleNamespace(End=len(self.body), Text=self.body)

    def Range(self, start, end):
        result = FakeRange(self, start, end)
        self.ranges.append(result)
        return result


class AppendTests(unittest.TestCase):
    def setUp(self):
        stack = self.enterContext(ExitStack())
        self.root = Path(stack.enter_context(TemporaryDirectory()))
        self.path = self.root / "append.docx"
        self.path.touch()
        stack.enter_context(patch.object(probe, "TEST_DOCUMENTS", self.root))
        self.com = stack.enter_context(patch.object(probe, "pythoncom"))
        self.client = stack.enter_context(patch.object(probe.win32com, "client"))
        self.document = FakeDocument(self.path)
        self.undo = SimpleNamespace(
            CustomRecordLevel=0, StartCustomRecord=Mock(), EndCustomRecord=Mock()
        )
        # No Selection/Save/Close methods: none should be needed by append.
        self.word = SimpleNamespace(
            Documents=SimpleNamespace(Count=1), ProtectedViewWindows=SimpleNamespace(Count=0),
            ActiveDocument=self.document, UndoRecord=self.undo,
        )
        self.client.GetActiveObject.return_value = self.word
        self.text = " Added"

    def preview(self):
        return probe.append_text(self.path, self.text)

    def apply(self, state):
        return probe.append_text(self.path, self.text, apply=True, expected_state=state)

    def assert_no_write(self):
        self.assertEqual(self.document.body, "Original\r")
        self.undo.StartCustomRecord.assert_not_called()
        self.assertTrue(all(not item.InsertAfter.called for item in self.document.ranges))

    def test_default_is_preview_and_preserves_document(self):
        result = self.preview()
        self.assertEqual(result["status"], "preview")
        self.assertEqual(result["position"], 8)
        self.assertEqual(result["text"], self.text)
        self.assert_no_write()

    def test_apply_preserves_original_and_groups_one_undo_record(self):
        result = self.apply(self.preview()["state"])
        self.assertEqual(result["status"], "appended")
        self.assertEqual(self.document.body, "Original Added\r")
        self.undo.StartCustomRecord.assert_called_once_with(probe.UNDO_NAME)
        self.undo.EndCustomRecord.assert_called_once_with()
        self.assertEqual(self.document.ranges[-1].InsertAfter.call_count, 1)

    def test_apply_without_preview_state_refused_before_com(self):
        self.assertEqual(self.apply(None)["status"], "preview_required")
        self.com.CoInitializeEx.assert_not_called()
        self.assert_no_write()

    def test_same_length_content_change_invalidates_preview(self):
        state = self.preview()["state"]
        self.document.body = "Modified\r"
        self.assertEqual(self.apply(state)["status"], "stale_preview")
        self.assertEqual(self.document.body, "Modified\r")
        self.undo.StartCustomRecord.assert_not_called()

    def test_changing_append_text_invalidates_preview(self):
        state = self.preview()["state"]
        self.text = " Different"
        self.assertEqual(self.apply(state)["status"], "stale_preview")
        self.assert_no_write()

    def test_newlines_are_normalized_for_word(self):
        self.text = "\r\n新增第一段\n新增第二段"
        preview = self.preview()
        self.assertEqual(preview["text"], "\r新增第一段\r新增第二段")
        self.assertEqual(self.apply(preview["state"])["status"], "appended")
        self.assertEqual(self.document.body, "Original\r新增第一段\r新增第二段\r")

    def test_read_only_document_refused(self):
        self.document.ReadOnly = True
        self.assertEqual(self.preview()["status"], "read_only")
        self.assert_no_write()

    def test_protected_document_refused(self):
        self.document.ProtectionType = 3
        self.assertEqual(self.preview()["status"], "protected_document")
        self.assert_no_write()

    def test_revisions_and_content_controls_refused(self):
        for name, value, status in (
            ("TrackRevisions", True, "revisions_not_supported"),
            ("Revisions", SimpleNamespace(Count=1), "revisions_not_supported"),
            ("ContentControls", SimpleNamespace(Count=1), "content_controls_not_supported"),
        ):
            with self.subTest(name=name), patch.object(self.document, name, value):
                self.assertEqual(self.preview()["status"], status)
                self.assert_no_write()

    def test_wrong_target_refused_without_opening_another_document(self):
        self.document.FullName = str(self.root / "another.docx")
        self.assertEqual(self.preview()["status"], "different_document")
        self.client.Dispatch.assert_not_called()
        self.assert_no_write()

    def test_no_document_refused(self):
        self.word.Documents.Count = 0
        self.assertEqual(self.preview()["status"], "no_document")
        self.assert_no_write()

    def test_protected_view_refused(self):
        self.word.ProtectedViewWindows.Count = 1
        self.assertEqual(self.preview()["status"], "protected_view_present")
        self.assert_no_write()

    def test_nested_undo_record_refused(self):
        self.undo.CustomRecordLevel = 1
        self.assertEqual(self.preview()["status"], "undo_record_busy")
        self.assert_no_write()

    def test_invalid_text_and_outside_directory_refused(self):
        for text in ("", "bad\x00text", "x" * (probe.MAX_TEXT_LENGTH + 1)):
            with self.subTest(text_length=len(text)):
                self.assertFalse(probe.append_text(self.path, text)["write_attempted"])
        with patch.object(probe, "TEST_DOCUMENTS", self.root / "allowed"):
            self.assertEqual(self.preview()["status"], "invalid_test_file")
        self.com.CoInitializeEx.assert_not_called()

    def test_document_ending_in_table_refused(self):
        original_range = self.document.Range

        def table_range(start, end):
            value = original_range(start, end)
            value.Information.return_value = True
            return value

        self.document.Range = table_range
        self.assertEqual(self.preview()["status"], "document_ends_in_table")
        self.assert_no_write()

    def test_insert_error_is_uncertain_and_not_automatically_retried(self):
        state = self.preview()["state"]
        original_range = self.document.Range

        def failing_range(start, end):
            value = original_range(start, end)
            value.InsertAfter.side_effect = pywintypes.com_error(
                -2147417848, "private diagnostic", None, None
            )
            return value

        self.document.Range = failing_range
        result = self.apply(state)
        self.assertEqual(result["status"], "write_outcome_unknown")
        self.assertTrue(result["write_attempted"])
        self.assertNotIn("private diagnostic", str(result))
        self.assertEqual(self.document.ranges[-1].InsertAfter.call_count, 1)
        self.undo.EndCustomRecord.assert_called_once_with()

    def test_undo_end_error_reports_possible_write_and_cleanup_failure(self):
        state = self.preview()["state"]
        self.undo.EndCustomRecord.side_effect = pywintypes.com_error(
            -2147417848, "private diagnostic", None, None
        )
        result = self.apply(state)
        self.assertEqual(result["status"], "write_outcome_unknown")
        self.assertTrue(result["undo_cleanup_failed"])
        self.assertEqual(self.document.body, "Original Added\r")

    def test_inserted_text_verification_failure_is_not_reported_as_success(self):
        state = self.preview()["state"]
        original_range = self.document.Range

        def no_op_range(start, end):
            value = original_range(start, end)
            value.InsertAfter.side_effect = None
            return value

        self.document.Range = no_op_range
        self.assertEqual(self.apply(state)["status"], "verification_failed")

    def test_last_moment_change_refused(self):
        state = self.preview()["state"]
        original_snapshot = probe._snapshot
        calls = 0

        def changing_snapshot(document, expected, text):
            nonlocal calls
            calls += 1
            if calls == 2:
                document.body = "Changed just now\r"
            return original_snapshot(document, expected, text)

        with patch.object(probe, "_snapshot", side_effect=changing_snapshot):
            self.assertEqual(self.apply(state)["status"], "stale_preview")
        self.undo.StartCustomRecord.assert_not_called()


if __name__ == "__main__":
    unittest.main()
