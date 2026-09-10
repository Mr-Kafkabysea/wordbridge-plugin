"""Metadata-only tests; these do not attach to Word."""

from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace
import unittest
from unittest.mock import patch

import pywintypes

from scripts import check_active_document as probe


class ActiveDocumentTests(unittest.TestCase):
    def setUp(self):
        temporary = TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.root = Path(temporary.name)
        self.expected = self.root / "test document.docx"
        # Only a path placeholder: this probe never parses document contents.
        self.expected.touch()
        patcher = patch.object(probe, "pythoncom")
        self.com = patcher.start()
        self.addCleanup(patcher.stop)
        patcher = patch.object(probe.win32com, "client")
        self.client = patcher.start()
        self.addCleanup(patcher.stop)
        self.document = SimpleNamespace(
            Name=self.expected.name,
            Path=str(self.root),
            FullName=str(self.expected),
            ReadOnly=False,
        )
        self.word = SimpleNamespace(
            Documents=SimpleNamespace(Count=1), ActiveDocument=self.document
        )
        self.client.GetActiveObject.return_value = self.word

    def test_exact_target_matches_without_content_access(self):
        result = probe.check_active_document(self.expected)
        self.assertEqual(result["status"], "matched")
        self.assertEqual(result["name"], self.expected.name)
        self.assertFalse(result["read_only"])
        self.client.GetActiveObject.assert_called_once_with("Word.Application")
        self.client.Dispatch.assert_not_called()
        self.com.CoUninitialize.assert_called_once_with()

    def test_same_name_in_other_folder_does_not_match(self):
        self.document.FullName = str(self.root / "other" / self.expected.name)
        self.document.Path = str(self.root / "other")
        self.assertFalse(probe.check_active_document(self.expected)["matches_expected"])

    def test_discovers_active_test_document_without_path_or_body_access(self):
        result = probe.check_active_document()
        self.assertEqual(result["status"], "matched")
        self.assertEqual(result["full_path"], str(self.expected))
        self.assertEqual(result["expected_path"], str(self.expected))
        self.com.CoUninitialize.assert_called_once_with()

    def test_discovery_rejects_invalid_targets_without_selection_access(self):
        cases = [("", str(self.expected)),
                 ("https://example.invalid", "https://example.invalid/a.docx"),
                 (str(self.root), str(self.root / "missing.docx"))]
        other = self.root / "wrong.txt"
        other.touch()
        cases.append((str(self.root), str(other)))
        for folder, full_name in cases:
            with self.subTest(full_name=full_name):
                self.document.Path = folder
                self.document.FullName = full_name
                self.assertEqual(probe.check_active_document(include_selection=True)["status"],
                                 "invalid_document_file")

    def test_discovery_outside_project_is_allowed(self):
        result = probe.check_active_document()
        self.assertEqual(result["status"], "matched")

    def test_discovery_without_open_document(self):
        self.client.GetActiveObject.return_value = SimpleNamespace(Documents=SimpleNamespace(Count=0))
        self.assertEqual(probe.check_active_document()["status"], "no_document")

    def test_unsaved_document_has_no_full_path(self):
        self.document.Path = ""
        result = probe.check_active_document(self.expected)
        self.assertFalse(result["matches_expected"])
        self.assertIsNone(result["full_path"])

    def test_read_only_target_still_identified_but_flagged(self):
        self.document.ReadOnly = True
        result = probe.check_active_document(self.expected)
        self.assertTrue(result["matches_expected"])
        self.assertTrue(result["read_only"])

    def test_no_document_does_not_access_active_document(self):
        self.client.GetActiveObject.return_value = SimpleNamespace(
            Documents=SimpleNamespace(Count=0)
        )
        result = probe.check_active_document(self.expected)
        self.assertEqual(result["status"], "no_document")
        self.com.CoUninitialize.assert_called_once_with()

    def test_missing_file_rejected_before_com(self):
        result = probe.check_active_document(self.root / "missing.docx")
        self.assertEqual(result["status"], "invalid_document_file")
        self.com.CoInitializeEx.assert_not_called()
        self.client.GetActiveObject.assert_not_called()

    def test_com_failure_does_not_expose_error_details(self):
        self.client.GetActiveObject.side_effect = pywintypes.com_error(
            -2147221021, "private diagnostic", None, None
        )
        result = probe.check_active_document(self.expected)
        self.assertEqual(result["status"], "not_available")
        self.assertNotIn("private diagnostic", str(result))
        self.com.CoUninitialize.assert_called_once_with()

    def test_web_document_does_not_match_local_test_file(self):
        self.document.Path = "https://example.invalid"
        self.document.FullName = "https://example.invalid/test document.docx"
        self.assertFalse(probe.check_active_document(self.expected)["matches_expected"])


if __name__ == "__main__":
    unittest.main()
