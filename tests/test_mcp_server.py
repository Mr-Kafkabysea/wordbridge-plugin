"""Exercise real MCP discovery/calls in memory with a simulated Word backend."""

import unittest
from unittest.mock import patch

from mcp import Client
from mcp.server.elicitation import render_elicitation_schema
from mcp.types import ElicitResult
from mcp.shared.exceptions import MCPError
from pydantic import ValidationError

from scripts import mcp_server


class MCPServerTests(unittest.IsolatedAsyncioTestCase):
    def assert_codex_form_schema(self, schema):
        # Regression: current Codex rejects extra root fields such as title.
        # Check the SDK-rendered/wire schema, not just model field definitions.
        self.assertLessEqual(set(schema), {"$schema", "type", "properties", "required"})
        self.assertEqual(schema["type"], "object")
        self.assertEqual(schema["required"], ["confirm"])
        self.assertEqual(schema["properties"], {"confirm": {
            "type": "boolean", "title": "我确认追加上述文字到上述文档"}})

    def test_approval_schema_matches_codex_root_fields(self):
        self.assert_codex_form_schema(render_elicitation_schema(mcp_server.AppendApproval))

    def test_approval_schema_fix_preserves_strict_required_boolean(self):
        for value in (True, False):
            self.assertIs(mcp_server.AppendApproval(confirm=value).confirm, value)
        for data in ({}, {"confirm": "true"}, {"confirm": 1}, {"confirm": None}):
            with self.subTest(data=data), self.assertRaises(ValidationError):
                mcp_server.AppendApproval.model_validate(data)

    async def test_discovery_schemas_and_annotations_without_touching_word(self):
        with (patch.object(mcp_server, "check_word_connection") as backend,
              patch.object(mcp_server, "check_active_document") as document,
              patch.object(mcp_server, "append_backend") as append):
            async with Client(mcp_server.create_server()) as client:
                result = await client.list_tools()
                self.assertEqual([tool.name for tool in result.tools], mcp_server.TOOL_NAMES)
                tool = result.tools[0]
                self.assertTrue(tool.annotations.read_only_hint)
                self.assertFalse(tool.annotations.open_world_hint)
                self.assertEqual(tool.input_schema.get("properties", {}), {})
                self.assertIsNotNone(tool.output_schema)
                required = [set(), set(), set(),
                            {"expected_document", "text"},
                            {"expected_document", "text", "expected_state"}]
                for index, tool in enumerate(result.tools):
                    self.assertEqual(set(tool.input_schema.get("required", [])), required[index])
                    self.assertNotIn("ctx", tool.input_schema.get("properties", {}))
                    self.assertNotIn("approval", tool.input_schema.get("properties", {}))
                    self.assertIsNotNone(tool.output_schema)
                    self.assertEqual(tool.annotations.read_only_hint, index != 4)
                self.assertFalse(result.tools[-1].annotations.idempotent_hint)
                self.assertTrue(result.tools[-1].annotations.destructive_hint)
            backend.assert_not_called()
            document.assert_not_called()
            append.assert_not_called()

    async def test_status_call_returns_structured_backend_data(self):
        expected = {"status": "connected", "connected": True, "word_version": "16.0"}
        with patch.object(mcp_server, "check_word_connection", return_value=expected) as backend:
            async with Client(mcp_server.create_server()) as client:
                result = await client.call_tool("word_status", {})
                self.assertFalse(result.is_error)
                self.assertEqual(result.structured_content, expected)
            backend.assert_called_once_with()

    async def test_word_unavailable_is_valid_status_not_a_broken_protocol(self):
        expected = {"status": "not_available", "connected": False, "hresult": "0x800401E3"}
        with patch.object(mcp_server, "check_word_connection", return_value=expected):
            async with Client(mcp_server.create_server()) as client:
                result = await client.call_tool("word_status", {})
                self.assertFalse(result.is_error)
                self.assertEqual(result.structured_content, expected)

    async def test_missing_write_arguments_do_not_reach_backend(self):
        with patch.object(mcp_server, "append_backend") as backend:
            async with Client(mcp_server.create_server()) as client:
                result = await client.call_tool("append_text", {"text": "must not be written"})
                self.assertTrue(result.is_error)
            backend.assert_not_called()

    async def test_document_and_selection_forward_parameters(self):
        for tool, options in [("get_active_document", {}),
                              ("get_selection", {"include_selection": True})]:
            with self.subTest(tool=tool):
                expected = {"status": "matched", "matches_expected": True}
                with patch.object(mcp_server, "check_active_document", return_value=expected) as backend:
                    async with Client(mcp_server.create_server()) as client:
                        result = await client.call_tool(tool, {"expected_document": "test-documents/a.docx"})
                        self.assertFalse(result.is_error)
                        self.assertEqual(result.structured_content, expected)
                    backend.assert_called_once_with(mcp_server.PROJECT / "test-documents/a.docx", **options)

    async def test_preview_cannot_write(self):
        with patch.object(mcp_server, "append_backend", return_value=self.preview()) as backend:
            async with Client(mcp_server.create_server()) as client:
                result = await client.call_tool("preview_append_text", self.arguments(state=False))
                self.assertFalse(result.is_error)
                self.assertEqual(result.structured_content, self.preview())
            backend.assert_called_once_with(mcp_server.PROJECT / "test-documents/a.docx", "测试\n文字")

    async def test_read_tools_accept_no_path_over_mcp(self):
        for tool, options in [("get_active_document", {}),
                              ("get_selection", {"include_selection": True})]:
            with self.subTest(tool=tool):
                with patch.object(mcp_server, "check_active_document",
                                  return_value={"status": "matched"}) as backend:
                    async with Client(mcp_server.create_server()) as client:
                        result = await client.call_tool(tool, {})
                        self.assertFalse(result.is_error)
                    backend.assert_called_once_with(None, **options)

    @staticmethod
    def arguments(state=True):
        args = {"expected_document": "test-documents/a.docx", "text": "测试\n文字"}
        if state:
            args["expected_state"] = "digest"
        return args

    @staticmethod
    def preview():
        return {"status": "preview", "write_attempted": False,
                "full_path": str(mcp_server.PROJECT / "test-documents/a.docx"),
                "text": "测试\r文字", "character_count": 5, "position": 15, "state": "digest"}

    async def test_no_confirmation_support_refuses_write(self):
        with patch.object(mcp_server, "append_backend", return_value=self.preview()) as backend:
            async with Client(mcp_server.create_server()) as client:
                with self.assertRaises(MCPError):
                    await client.call_tool("append_text", self.arguments())
            self.assertEqual(backend.call_count, 1)
            self.assertEqual(backend.call_args.kwargs, {})

    async def test_decline_cancel_or_unchecked_form_never_write(self):
        for action, content in [("decline", None), ("cancel", None),
                                ("accept", {"confirm": False}), ("accept", {}),
                                ("accept", {"confirm": "true"})]:
            with self.subTest(action=action, content=content):
                async def respond(ctx, params):
                    return ElicitResult(action=action, content=content)
                with patch.object(mcp_server, "append_backend", return_value=self.preview()) as backend:
                    async with Client(mcp_server.create_server(), elicitation_callback=respond) as client:
                        result = await client.call_tool("append_text", self.arguments())
                        if action == "accept" and content in ({}, {"confirm": "true"}):
                            self.assertTrue(result.is_error)
                        else:
                            self.assertFalse(result.is_error)
                            self.assertFalse(result.structured_content["write_attempted"])
                            self.assertIn(result.structured_content["status"], {"approval_decline", "approval_cancel"})
                    self.assertTrue(all(not call.kwargs.get("apply") for call in backend.call_args_list))

    async def test_stale_or_refused_preview_never_requests_confirmation(self):
        for preview in [dict(self.preview(), state="changed"),
                        {"status": "different_document", "write_attempted": False}]:
            async def unexpected(ctx, params):
                self.fail("Must refuse before asking for approval")
            with patch.object(mcp_server, "append_backend", return_value=preview) as backend:
                async with Client(mcp_server.create_server(), elicitation_callback=unexpected) as client:
                    result = await client.call_tool("append_text", self.arguments())
                    self.assertTrue(result.is_error)
                self.assertEqual(backend.call_count, 1)

    async def test_confirmed_write_forwards_state_once_and_preserves_outcomes(self):
        for outcome in [{"status": "appended", "write_attempted": True},
                        {"status": "stale_preview", "write_attempted": False},
                        {"status": "write_outcome_unknown", "write_attempted": True}]:
            confirmations = []
            async def respond(ctx, params):
                confirmations.append(params)
                self.assert_codex_form_schema(params.requested_schema)
                self.assertIn("test-documents", params.message)
                self.assertIn("测试", params.message)
                return ElicitResult(action="accept", content={"confirm": True})
            with patch.object(mcp_server, "append_backend", side_effect=[self.preview(), self.preview(), outcome]) as backend:
                async with Client(mcp_server.create_server(), elicitation_callback=respond) as client:
                    result = await client.call_tool("append_text", self.arguments())
                    self.assertFalse(result.is_error)
                    self.assertEqual(result.structured_content, outcome)
                self.assertEqual(len(confirmations), 1)
                self.assertEqual(backend.call_count, 3)
                backend.assert_called_with(mcp_server.PROJECT / "test-documents/a.docx",
                                           "测试\n文字", apply=True, expected_state="digest")

    async def test_unimplemented_tools_are_not_exposed(self):
        async with Client(mcp_server.create_server()) as client:
            for name in ["replace_selection", "undo", "list_documents"]:
                result = await client.call_tool(name, {})
                self.assertTrue(result.is_error)

    async def test_document_change_during_confirmation_blocks_write(self):
        async def respond(ctx, params):
            return ElicitResult(action="accept", content={"confirm": True})
        with patch.object(mcp_server, "append_backend",
                          side_effect=[self.preview(), dict(self.preview(), state="changed")]) as backend:
            async with Client(mcp_server.create_server(), elicitation_callback=respond) as client:
                result = await client.call_tool("append_text", self.arguments())
                self.assertTrue(result.is_error)
            self.assertEqual(backend.call_count, 2)
            self.assertTrue(all(not call.kwargs.get("apply") for call in backend.call_args_list))

    async def test_caller_cannot_supply_approval_to_bypass_user(self):
        with patch.object(mcp_server, "append_backend", return_value=self.preview()) as backend:
            async with Client(mcp_server.create_server()) as client:
                args = dict(self.arguments(), approval={"action": "accept", "data": {"confirm": True}})
                try:
                    result = await client.call_tool("append_text", args)
                except MCPError:
                    pass  # Capability refusal is also a safe outcome.
                else:
                    self.assertTrue(result.is_error)
            self.assertTrue(all(not call.kwargs.get("apply") for call in backend.call_args_list))


if __name__ == "__main__":
    unittest.main()
