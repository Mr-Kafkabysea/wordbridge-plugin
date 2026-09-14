"""Installation probe: initialize and list tools only; never call a Word tool."""
import asyncio
import json
from pathlib import Path
import sys
from mcp import Client, StdioServerParameters


async def check():
    root = Path(__file__).resolve().parents[1]
    async with Client(StdioServerParameters(
        command=sys.executable, args=[str(root / 'scripts/mcp_server.py')],
        cwd=str(root),
    )) as client:
        result = await client.list_tools()
        version = json.loads((root / '.codex-plugin/plugin.json').read_text(encoding='utf-8'))['version']
        if client.server_info is None or client.server_info.version != version:
            raise RuntimeError('Running MCP version does not match the plugin manifest')
        expected = {'word_status', 'get_active_document', 'get_selection',
                    'preview_append_text', 'append_text', 'set_mode', 'get_mode', 'preview_insert_text', 'insert_text'}
        if {tool.name for tool in result.tools} != expected:
            raise RuntimeError('Unexpected MCP tools')
        print(f'MCP {version} handshake and nine tool schemas: OK. Word was not accessed.')


if __name__ == '__main__':
    asyncio.run(check())
