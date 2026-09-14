"""Quick path write and refusal checks using simulated COM objects."""
from tests import test_append_text as base
from unittest.mock import patch

class QuickWriteTests(base.AppendTests):
    # Existing write-error, verification and late-change tests run against quick writes.
    def apply(self, state):
        return base.probe.append_text(self.path,self.text,apply=True,quick=True)

    def test_apply_without_preview_state_refused_before_com(self):
        self.assertEqual(self.apply(None)['status'],'appended')
        self.assertEqual(self.document.body,'Original Added\r')
        self.undo.StartCustomRecord.assert_called_once()
        self.undo.EndCustomRecord.assert_called_once()

    def test_same_length_content_change_invalidates_preview(self):
        self.document.body='Modified\r'
        self.assertEqual(self.apply(None)['status'],'appended')
        self.assertEqual(self.document.body,'Modified Added\r')

    def test_changing_append_text_invalidates_preview(self):
        self.text=' Different'
        self.assertEqual(self.apply(None)['status'],'appended')
        self.assertEqual(self.document.body,'Original Different\r')

    def test_quick_guards_refuse(self):
        for attr,value in [('ReadOnly',True),('ProtectionType',3),('TrackRevisions',True),
                           ('FullName',str(self.root/'other.docx'))]:
            with self.subTest(attr=attr), patch.object(self.document,attr,value):
                self.assertFalse(self.apply(None)['write_attempted'])
                self.assert_no_write()
        self.word.Documents.Count=0
        self.assertEqual(self.apply(None)['status'],'no_document')
        self.assert_no_write()