"""Simulated failure-path checks, separate from real Word verification."""

import unittest
from types import SimpleNamespace
from unittest.mock import patch

import pywintypes

from scripts import check_word_connection as probe


def com_error(code):
    signed = code if code < 0x80000000 else code - 0x100000000
    return pywintypes.com_error(signed, "sensitive diagnostic", None, None)


class WordConnectionTests(unittest.TestCase):
    def setUp(self):
        self.com = patch.object(probe, "pythoncom").start()
        self.client = patch.object(probe.win32com, "client").start()
        self.addCleanup(patch.stopall)

    def test_connects_to_existing_instance_and_reads_only_version(self):
        # This object deliberately has no document access or mutation methods.
        self.client.GetActiveObject.return_value = SimpleNamespace(Version="16.0")
        result = probe.check_word_connection()
        self.assertEqual(result["status"], "connected")
        self.assertEqual(result["word_version"], "16.0")
        self.client.GetActiveObject.assert_called_once_with("Word.Application")
        self.client.Dispatch.assert_not_called()
        self.client.DispatchEx.assert_not_called()
        self.com.CoUninitialize.assert_called_once_with()

    def test_missing_active_object_is_not_reported_as_proven_not_running(self):
        self.client.GetActiveObject.side_effect = com_error(0x800401E3)
        result = probe.check_word_connection()
        self.assertEqual(result["status"], "not_available")
        self.assertFalse(result["connected"])
        self.assertEqual(result["hresult"], "0x800401E3")
        self.client.Dispatch.assert_not_called()
        self.com.CoUninitialize.assert_called_once_with()

    def test_access_denied_is_distinct_and_does_not_expose_raw_error(self):
        self.client.GetActiveObject.side_effect = com_error(0x80070005)
        result = probe.check_word_connection()
        self.assertEqual(result["status"], "com_error")
        self.assertEqual(result["stage"], "connect")
        self.assertNotIn("sensitive diagnostic", str(result))
        self.com.CoUninitialize.assert_called_once_with()

    def test_initialization_failure_does_not_uninitialize(self):
        self.com.CoInitializeEx.side_effect = com_error(0x80010106)
        result = probe.check_word_connection()
        self.assertEqual(result["stage"], "initialize")
        self.client.GetActiveObject.assert_not_called()
        self.com.CoUninitialize.assert_not_called()

    def test_word_closing_during_version_read_is_not_success(self):
        class ClosingWord:
            @property
            def Version(self):
                raise com_error(0x80010108)

        self.client.GetActiveObject.return_value = ClosingWord()
        result = probe.check_word_connection()
        self.assertFalse(result["connected"])
        self.assertEqual(result["stage"], "read_version")
        self.com.CoUninitialize.assert_called_once_with()


if __name__ == "__main__":
    unittest.main()
