"""Verify telemetry against real in-memory MCP, without accessing Word."""

import json
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

from mcp import Client
from scripts.connection_status import ConnectionStatus, display_state, snapshots
from scripts.mcp_server import create_server
from scripts.connection_popup import presentation


class StatusTests(unittest.IsolatedAsyncioTestCase):
    def setUp(self):
        self.probe = self.enterContext(patch('scripts.mcp_server.check_word_connection',
            return_value={'status': 'connected', 'connected': True, 'word_version': '16.0'}))

    async def test_lifecycle_handshake_and_word_status(self):
        with TemporaryDirectory() as directory:
            status = ConnectionStatus(directory)
            with patch('scripts.mcp_server.check_word_connection', return_value={
                    'status': 'connected', 'connected': True, 'word_version': '16.0',
                    'message': 'must not be persisted'}):
                async with Client(create_server(status)) as client:
                    await client.list_tools()
                    self.assertEqual(status.state['service'], 'running')
                    self.assertIn(status.state['mcp'], ('initialized', 'request_received'))
                    self.assertEqual(status.state['word'], 'unchecked')
                    self.assertIsNone(status.state['last_tool'])
                    await client.call_tool('word_status', {})
                    self.assertEqual(status.state['word'], 'connected')
                    self.assertEqual(status.state['last_tool'], 'word_status')
                    self.assertNotIn('must not be persisted', status.path.read_text())
                self.assertEqual(status.state['service'], 'stopped')
                self.assertEqual(status.state['mcp'], 'disconnected')

    async def test_arguments_and_body_are_not_recorded(self):
        with TemporaryDirectory() as directory:
            status = ConnectionStatus(directory)
            with patch('scripts.mcp_server.check_active_document', return_value={
                    'status': 'selected', 'selection': {'text': 'PRIVATE_BODY'}}):
                async with Client(create_server(status)) as client:
                    await client.call_tool('get_selection', {'expected_document': 'PRIVATE_PATH.docx'})
                raw = status.path.read_text()
                self.assertNotIn('PRIVATE', raw)
                self.assertEqual(json.loads(raw)['word'], 'connected')

    async def test_popup_only_on_first_valid_call_and_probe_once(self):
        with TemporaryDirectory() as directory:
            status = ConnectionStatus(directory, popup=True)
            with patch.object(status, 'launch_popup') as popup, patch(
                    'scripts.mcp_server.check_active_document', return_value={'status': 'matched'}):
                async with Client(create_server(status)) as client:
                    await client.list_tools()
                    popup.assert_not_called()
                    self.probe.assert_not_called()
                    await client.call_tool('get_active_document', {})
                    await client.call_tool('get_selection', {})
                    popup.assert_called_once_with()
                    self.probe.assert_called_once_with()
                    self.assertEqual(status.state['first_check'], 'connected')

    async def test_popup_failure_does_not_break_tool(self):
        with TemporaryDirectory() as directory:
            status = ConnectionStatus(directory, popup=True)
            with patch.object(status, 'launch_popup', side_effect=OSError('private detail')):
                async with Client(create_server(status)) as client:
                    result = await client.call_tool('word_status', {})
                    self.assertFalse(result.is_error)
                    self.probe.assert_called_once_with()
                    self.assertEqual(status.state['popup_error'], 'launch_failed')
                    self.assertNotIn('private detail', status.path.read_text())

    async def test_unavailable_word_is_not_success_and_later_check_can_recover(self):
        self.probe.side_effect = [{'status': 'not_available', 'connected': False},
                                  {'status': 'connected', 'connected': True}]
        with TemporaryDirectory() as directory:
            status = ConnectionStatus(directory, popup=True)
            with patch.object(status, 'launch_popup') as popup:
                async with Client(create_server(status)) as client:
                    first = await client.call_tool('word_status', {})
                    self.assertFalse(first.structured_content['connected'])
                    self.assertEqual(status.state['first_check'], 'not_available')
                    await client.call_tool('word_status', {})
                    self.assertEqual(status.state['word'], 'connected')
                    popup.assert_called_once_with()

    def test_popup_presentation_success_failure_and_stale(self):
        data = {'service': 'running', 'heartbeat': 100, 'first_checked_at': 100}
        for state in ('checking', 'not_available', 'com_error', 'unknown'):
            self.assertFalse(presentation(dict(data, first_check=state), 101)[2])
        self.assertTrue(presentation(dict(data, first_check='connected'), 101)[2])
        self.assertFalse(presentation(dict(data, first_check='connected'), 120)[2])
        self.assertFalse(presentation(dict(data, first_check='connected', service='stopped'), 101)[2])

    async def test_write_failure_does_not_break_mcp(self):
        with TemporaryDirectory() as directory:
            file = Path(directory) / 'not-a-directory'
            file.touch()
            async with Client(create_server(ConnectionStatus(file))) as client:
                result = await client.list_tools()
                self.assertEqual(len(result.tools), 5)

    def test_stale_heartbeat_is_not_shown_online(self):
        data = {'service': 'running', 'mcp': 'initialized', 'heartbeat': 100,
                'word': 'connected', 'word_checked_at': 100}
        self.assertEqual(display_state(data, now=101)[1], '运行中')
        self.assertEqual(display_state(data, now=116)[2], '未确认在线')
        self.assertIn('上次检查', display_state(data, now=116)[3])

    def test_instances_are_separate_and_bad_json_is_ignored(self):
        with TemporaryDirectory() as directory:
            first, second = ConnectionStatus(directory), ConnectionStatus(directory)
            first.update(service='running')
            second.update(service='stopped')
            self.assertNotEqual(first.path, second.path)
            (Path(directory) / 'bad.json').write_text('{')
            self.assertEqual(len(snapshots(directory)), 2)


if __name__ == '__main__':
    unittest.main()
