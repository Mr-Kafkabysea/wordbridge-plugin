"""Cursor insertion safety with simulated Word; no real document access."""
import unittest
from types import SimpleNamespace
from unittest.mock import Mock, patch
import pywintypes
from tests import test_append_text as base
from scripts import append_text as core

class Cursor:
    def __init__(self, document, position=3):
        self.document=document; self.start=position; self.end=position
        self.Type=1; self.story=1; self.table=False
        self.SetRange=Mock(side_effect=self.move)
    def move(self,start,end):
        self.start=start; self.end=end
    @property
    def Range(self):
        r=SimpleNamespace(Document=self.document,Start=self.start,End=self.end,
                          StoryType=self.story,Information=lambda _:self.table)
        return SimpleNamespace(Duplicate=r)

class CursorTests(unittest.TestCase):
    assert_no_write=base.AppendTests.assert_no_write
    def setUp(self):
        base.AppendTests.setUp(self)
        self.cursor=Cursor(self.document)
        self.word.Selection=self.cursor
    def insert(self,**kwargs):
        return core.insert_at_cursor(self.path,self.text,**kwargs)
    def test_preview_preserves_cursor_and_binds_position_and_operation(self):
        first=self.insert()
        self.assertEqual(first['position'],3)
        self.assert_no_write();self.cursor.SetRange.assert_not_called()
        self.cursor.move(4,4)
        self.assertNotEqual(self.insert()['state'],first['state'])
        self.cursor.move(8,8)
        self.assertNotEqual(self.insert()['state'],core.append_text(self.path,self.text)['state'])
    def test_normal_and_fast_middle_insert_and_continuation(self):
        for fast in (False,True):
            with self.subTest(fast=fast):
                self.document.body='Original\r';self.cursor.move(3,3)
                self.undo.StartCustomRecord.reset_mock();self.undo.EndCustomRecord.reset_mock()
                state=None if fast else self.insert()['state']
                r=self.insert(apply=True,quick=fast,expected_state=state)
                self.assertEqual(r['status'],'inserted')
                self.assertEqual(self.document.body,'Ori Addedginal\r')
                self.assertEqual(self.cursor.start,9)
                self.undo.StartCustomRecord.assert_called_once_with('WordBridge insert at cursor')
                self.undo.EndCustomRecord.assert_called_once()
                self.assertEqual(self.insert(apply=True,quick=True)['status'],'inserted')
                self.assertEqual(self.document.body,'Ori Added Addedginal\r')
    def test_start_end_and_newlines(self):
        for position in (0,8):
            self.document.body='Original\r';self.cursor.move(position,position);self.text='A\nB'
            r=self.insert(apply=True,quick=True)
            self.assertEqual(r['status'],'inserted')
            self.assertEqual(self.document.body,'Original\r'[:position]+'A\rB'+'Original\r'[position:])
    def test_cursor_move_invalidates_normal_preview(self):
        state=self.insert()['state'];self.cursor.move(4,4)
        self.assertEqual(self.insert(apply=True,expected_state=state)['status'],'stale_preview')
        self.assert_no_write()
    def test_noncollapsed_and_unsupported_cursor_never_write(self):
        cases=[('end',4,'selection_not_collapsed'),('story',2,'unsupported_cursor_story'),
               ('table',True,'cursor_in_table'),('Type',2,'invalid_cursor')]
        for attr,value,status in cases:
            with self.subTest(attr=attr),patch.object(self.cursor,attr,value):
                self.assertEqual(self.insert(apply=True,quick=True)['status'],status)
                self.assert_no_write()
        other=base.FakeDocument(self.root/'other.docx')
        with patch.object(self.cursor,'document',other):
            self.assertEqual(self.insert(apply=True,quick=True)['status'],'selection_document_mismatch')
            self.assert_no_write()
    def test_edit_guards_and_missing_document(self):
        for attr,value in [('ReadOnly',True),('ProtectionType',3),('TrackRevisions',True)]:
            with patch.object(self.document,attr,value):
                self.assertFalse(self.insert(apply=True,quick=True)['write_attempted'])
                self.assert_no_write()
        self.word.Documents.Count=0
        self.assertEqual(self.insert(apply=True,quick=True)['status'],'no_document')
    def test_last_moment_cursor_change_refused(self):
        original=core._snapshot; calls=0
        def snapshot(*args):
            nonlocal calls
            calls+=1
            if calls==2:self.cursor.move(5,5)
            return original(*args)
        with patch.object(core,'_snapshot',side_effect=snapshot):
            self.assertEqual(self.insert(apply=True,quick=True)['status'],'stale_preview')
        self.assert_no_write()
    def test_uncertain_insert_not_retried(self):
        original=self.document.Range
        def failing(*args):
            r=original(*args);r.InsertAfter.side_effect=pywintypes.com_error(-2147417848,'error',None,None);return r
        self.document.Range=failing
        r=self.insert(apply=True,quick=True)
        self.assertEqual(r['status'],'write_outcome_unknown')
        self.assertEqual(self.document.ranges[-1].InsertAfter.call_count,1)
        self.undo.EndCustomRecord.assert_called_once()
    def test_failed_verification_does_not_advance_cursor(self):
        original=self.document.Range
        def noop(*args):
            r=original(*args);r.InsertAfter.side_effect=None;return r
        self.document.Range=noop
        self.assertEqual(self.insert(apply=True,quick=True)['status'],'verification_failed')
        self.cursor.SetRange.assert_not_called()