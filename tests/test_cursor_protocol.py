"""Cursor tools over real in-memory MCP protocol with a simulated write backend."""
import unittest
from unittest.mock import patch
from mcp import Client
from mcp.types import ElicitResult
from scripts import mcp_server as server

class CursorProtocolTests(unittest.IsolatedAsyncioTestCase):
    args={'expected_document':'test-documents/cursor.docx','text':'hello'}
    preview=dict(status='preview',state='cursor-state',position=3,text='hello',character_count=5,full_path='cursor.docx')
    async def test_normal_confirmation_cancel_and_stale_cursor(self):
        for action in ('accept','cancel','decline'):
            forms=[]
            async def respond(ctx,params):
                forms.append(params)
                self.assertNotIn('title',params.requested_schema)
                self.assertEqual(params.requested_schema['required'],['confirm'])
                self.assertEqual(params.requested_schema['properties']['confirm']['title'],'我确认在上述文档的光标处插入上述文字')
                return ElicitResult(action=action,content={'confirm':True} if action=='accept' else None)
            def backend(*args,**kwargs):
                return {'status':'inserted','write_attempted':True} if kwargs.get('apply') else self.preview
            with patch.object(server,'insert_backend',side_effect=backend) as write:
                async with Client(server.create_server(),elicitation_callback=respond) as client:
                    p=await client.call_tool('preview_insert_text',self.args)
                    self.assertEqual(p.structured_content['state'],'cursor-state')
                    r=await client.call_tool('insert_text',dict(self.args,expected_state='cursor-state'))
                    self.assertFalse(r.is_error)
                    self.assertEqual(r.structured_content['write_attempted'],action=='accept')
                    self.assertEqual(len(forms),1)
                    if action!='accept':self.assertTrue(all(not c.kwargs.get('apply') for c in write.call_args_list))
            with patch.object(server,'insert_backend',return_value=dict(self.preview,state='moved')) as write:
                async with Client(server.create_server(),elicitation_callback=respond) as client:
                    r=await client.call_tool('insert_text',dict(self.args,expected_state='cursor-state'))
                    self.assertTrue(r.is_error)
                self.assertTrue(all(not c.kwargs.get('apply') for c in write.call_args_list))
    async def test_fast_uses_same_mode_without_form_and_normal_restores_guard(self):
        for status in ('inserted','selection_not_collapsed','write_outcome_unknown'):
            with patch.object(server,'insert_backend',return_value={'status':status}) as write, patch.object(server,'append_backend') as append:
                async with Client(server.create_server()) as client:
                    await client.call_tool('set_mode',{'mode_name':'fast'})
                    r=await client.call_tool('insert_text',self.args)
                    self.assertFalse(r.is_error);self.assertEqual(r.structured_content['status'],status)
                    write.assert_called_once_with(server.PROJECT/'test-documents/cursor.docx','hello',apply=True,quick=True)
                    await client.call_tool('set_mode',{'mode_name':'normal'})
                    write.return_value=self.preview
                    self.assertTrue((await client.call_tool('insert_text',self.args)).is_error)
                append.assert_not_called()