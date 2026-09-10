"""Check selection boundaries without running Word or reading real documents."""

from contextlib import ExitStack
from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace
import unittest
from unittest.mock import patch

import pywintypes

from scripts import check_active_document as probe


class SelectionRange:
    def __init__(self, document):
        self.Document = document
        self.Start = 0
        self.End = 6
        self.StoryType = 1
        self.Duplicate = self
        self.text_reads = 0
        self.content = "测试😀。\r"

    @property
    def Text(self):
        self.text_reads += 1
        return self.content


class SelectionTests(unittest.TestCase):
    def setUp(self):
        stack = self.enterContext(ExitStack())
        self.root = Path(stack.enter_context(TemporaryDirectory()))
        self.expected = self.root / "selection.docx"
        self.expected.touch()
        stack.enter_context(patch.object(probe, "TEST_DOCUMENTS", self.root))
        self.com = stack.enter_context(patch.object(probe, "pythoncom"))
        self.client = stack.enter_context(patch.object(probe.win32com, "client"))
        self.document = SimpleNamespace(
            Name=self.expected.name, Path=str(self.root),
            FullName=str(self.expected), ReadOnly=False,
        )
        self.range = SelectionRange(self.document)
        self.word = SimpleNamespace(
            Documents=SimpleNamespace(Count=1), ActiveDocument=self.document,
            Selection=SimpleNamespace(Range=self.range, Type=2),
        )
        self.client.GetActiveObject.return_value = self.word

    def read(self):
        return probe.check_active_document(self.expected, include_selection=True)

    def test_preserves_chinese_emoji_and_paragraph_mark(self):
        result = self.read()
        self.assertEqual(result["status"], "selected")
        self.assertEqual(result["selection"]["text"], self.range.content)
        self.assertEqual(result["selection"]["character_count"], 5)
        self.assertEqual(result["selection"]["end"], 6)
        self.assertEqual(self.range.text_reads, 1)
        self.com.CoUninitialize.assert_called_once_with()

    def test_empty_cursor_does_not_read_text(self):
        self.range.End = self.range.Start
        self.assertEqual(self.read()["status"], "empty_selection")
        self.assertEqual(self.range.text_reads, 0)

    def test_different_active_document_does_not_read_selection(self):
        self.word.ActiveDocument = SimpleNamespace(
            Name="other.docx", Path=str(self.root),
            FullName=str(self.root / "other.docx"), ReadOnly=False,
        )
        del self.word.Selection
        self.assertEqual(self.read()["status"], "different_document")
        self.assertEqual(self.range.text_reads, 0)

    def test_selection_switched_to_other_document_does_not_leak_text(self):
        self.range.Document = SimpleNamespace(
            Path=str(self.root), FullName=str(self.root / "other.docx")
        )
        self.assertEqual(self.read()["status"], "selection_document_mismatch")
        self.assertEqual(self.range.text_reads, 0)

    def test_header_selection_is_not_read(self):
        self.range.StoryType = 7
        self.assertEqual(self.read()["status"], "unsupported_selection")
        self.assertEqual(self.range.text_reads, 0)

    def test_shape_selection_is_not_read(self):
        self.word.Selection.Type = 8
        self.assertEqual(self.read()["status"], "unsupported_selection")
        self.assertEqual(self.range.text_reads, 0)

    def test_large_selection_is_rejected_before_text_read(self):
        self.range.End = probe.MAX_SELECTION_POSITIONS + 1
        self.assertEqual(self.read()["status"], "selection_too_large")
        self.assertEqual(self.range.text_reads, 0)

    def test_metadata_only_mode_never_reads_selection(self):
        del self.word.Selection
        self.assertEqual(probe.check_active_document(self.expected)["status"], "matched")
        self.assertEqual(self.range.text_reads, 0)

    def test_com_error_during_selection_read_is_reported_without_text(self):
        class DisconnectedWord:
            Documents = SimpleNamespace(Count=1)
            ActiveDocument = self.document

            @property
            def Selection(self):
                raise pywintypes.com_error(-2147417848, "private data", None, None)

        self.client.GetActiveObject.return_value = DisconnectedWord()
        result = self.read()
        self.assertEqual(result["status"], "com_error")
        self.assertEqual(result["stage"], "selection")
        self.assertNotIn("private data", str(result))
        self.com.CoUninitialize.assert_called_once_with()


if __name__ == "__main__":
    unittest.main()
