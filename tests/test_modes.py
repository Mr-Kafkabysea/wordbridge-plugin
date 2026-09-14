"""Persistent mode protocol tests; never access Word."""
import unittest
from unittest.mock import patch
from mcp import Client
from mcp.types import ElicitResult
from scripts import mcp_server as server

class ModeTests(unittest.IsolatedAsyncioTestCase):
    async def test_persistence_restore_and_single_call(self):
        forms = []
        args = {'expected_document': 'test-documents/a.docx', 'text': 'hello'}
        preview = dict(status='preview', state='digest', full_path='a.docx', position=0, text='hello', character_count=5)
        def backend(*args, **kwargs):
            return {'status':'appended','write_attempted':True} if kwargs.get('apply') else preview
        async def respond(ctx, params):
            forms.append(params)
            return ElicitResult(action='accept', content={'confirm':True})
        with patch.object(server, 'append_backend', side_effect=backend) as write:
            async with Client(server.create_server(), elicitation_callback=respond) as client:
                self.assertEqual((await client.call_tool('get_mode', {})).structured_content['mode'], 'normal')
                await client.call_tool('set_mode', {'mode_name':'fast'})
                write.assert_not_called()
                for _ in range(2):
                    r=await client.call_tool('append_text', args)
                    self.assertFalse(r.is_error)
                    self.assertEqual(r.structured_content['status'],'appended')
                self.assertEqual(write.call_count,2)
                self.assertTrue(all(c.kwargs == {'apply':True,'quick':True} for c in write.call_args_list))
                self.assertEqual(forms,[])
                r=await client.call_tool('preview_append_text',args)
                self.assertEqual(r.structured_content['status'],'preview')
                self.assertEqual(write.call_args.kwargs,{})
                await client.call_tool('set_mode', {'mode_name':'normal'})
                self.assertTrue((await client.call_tool('append_text',args)).is_error)
                r=await client.call_tool('append_text',dict(args,expected_state='digest'))
                self.assertFalse(r.is_error)
                self.assertEqual(len(forms),1)

    async def test_isolation_restart_invalid_mode_and_unknown_outcome(self):
        outcome={'status':'write_outcome_unknown','write_attempted':True}
        with patch.object(server,'append_backend',return_value=outcome) as write, patch.object(server,'check_word_connection') as word:
            async with Client(server.create_server()) as first:
                await first.call_tool('set_mode',{'mode_name':'fast'})
                self.assertTrue((await first.call_tool('set_mode',{'mode_name':'oops'})).is_error)
                async with Client(server.create_server()) as second:
                    self.assertEqual((await second.call_tool('get_mode',{})).structured_content['mode'],'normal')
                r=await first.call_tool('append_text',{'expected_document':'a.docx','text':'hello'})
                self.assertEqual(r.structured_content,outcome)
                self.assertEqual(write.call_count,1)
                self.assertEqual((await first.call_tool('get_mode',{})).structured_content['mode'],'fast')
            async with Client(server.create_server()) as restarted:
                self.assertEqual((await restarted.call_tool('get_mode',{})).structured_content['mode'],'normal')
            word.assert_not_called()