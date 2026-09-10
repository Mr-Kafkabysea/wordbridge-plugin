"""Discover tools and perform only read/preview calls; never append text."""

import argparse
import asyncio
import json
from pathlib import Path
import sys

from mcp import Client, StdioServerParameters


async def check_mcp(expected_document: str | None = None,
                    preview_text: str | None = None) -> dict:
    project = Path(__file__).resolve().parents[1]
    parameters = StdioServerParameters(
        command=sys.executable,
        args=[str(project / "scripts" / "mcp_server.py")],
        cwd=str(project),
    )
    # Closing this context closes the test server too; Word is left alone.
    async with Client(parameters) as client:
        tools = await client.list_tools()
        names = [tool.name for tool in tools.tools]
        if names != ["word_status", "get_active_document", "get_selection",
                     "preview_append_text", "append_text"]:
            raise RuntimeError("Unexpected MCP tool list")
        result = await client.call_tool("word_status", {})
        if result.is_error or not isinstance(result.structured_content, dict):
            raise RuntimeError("word_status did not return a structured result")
        status = result.structured_content
        if status.get("status") not in {"connected", "not_available", "com_error"}:
            raise RuntimeError("Unexpected Word status")
        report = {
            "mcp_check": "passed",
            "server_name": client.server_info.name if client.server_info else None,
            "protocol_version": client.protocol_version,
            "tools": names,
            "word_status": status,
        }
        if expected_document is not None:
            args = {"expected_document": expected_document}
            for name in ("get_active_document", "get_selection", "preview_append_text"):
                if name == "preview_append_text":
                    if preview_text is None:
                        continue
                    args = dict(args, text=preview_text)
                result = await client.call_tool(name, args)
                if result.is_error or not isinstance(result.structured_content, dict):
                    raise RuntimeError("Invalid structured MCP result")
                data = dict(result.structured_content)
                if "selection" in data:
                    data["selection"] = {k: v for k, v in data["selection"].items() if k != "text"}
                report[name] = data
        return report


async def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--expected-document")
    parser.add_argument("--preview-text")
    args = parser.parse_args()
    if args.preview_text is not None and args.expected_document is None:
        parser.error("--preview-text requires --expected-document")
    # A protocol success is separate from whether Word happens to be open.
    try:
        async with asyncio.timeout(30):
            result = await check_mcp(args.expected_document, args.preview_text)
    except Exception as error:
        print(json.dumps({"mcp_check": "failed", "error_type": type(error).__name__}))
        return 1
    print(json.dumps(result, ensure_ascii=True, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(asyncio.run(main()))
