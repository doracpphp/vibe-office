"""MCP サーバーを実際に起動して stdio 経由で呼び出すテスト"""
import asyncio
import json
import os
import sys

import openpyxl
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

import excel_tools
import text_tools
import word_tools

SERVER = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "mcp_server.py")


def call_server(workdir, calls):
    """サーバーを起動してツール一覧と各ツールの呼び出し結果を返す"""
    async def session_run():
        params = StdioServerParameters(
            command=sys.executable, args=[SERVER], env={"VIBE_OFFICE_WORKDIR": str(workdir)})
        async with stdio_client(params) as (read, write), ClientSession(read, write) as session:
            await session.initialize()
            tools = (await session.list_tools()).tools
            results = [await session.call_tool(name, args) for name, args in calls]
            return tools, results

    return asyncio.run(asyncio.wait_for(session_run(), timeout=60))


def test_mcp_server(tmp_path):
    tools, (written, missing, outside) = call_server(tmp_path, [
        ("write_cell", {"file_path": "a.xlsx", "cell_address": "A1", "value": "mcp"}),
        ("read_cell", {"file_path": "missing.xlsx", "cell_address": "A1"}),
        ("write_cell", {"file_path": "../escaped.xlsx", "cell_address": "A1", "value": 1}),
    ])

    names = [t.name for t in tools]
    assert len(names) == len(set(names))
    assert set(names) == {t["name"] for t in excel_tools.TOOLS + word_tools.TOOLS + text_tools.TOOLS}

    # VIBE_OFFICE_WORKDIR のディレクトリに書き込まれる
    assert written.isError is False
    assert json.loads(written.content[0].text)["success"] is True
    assert openpyxl.load_workbook(tmp_path / "a.xlsx").active["A1"].value == "mcp"

    # 失敗はクライアントにエラーとして返る
    assert missing.isError is True
    assert missing.content[0].text.startswith("エラー:")
    assert outside.isError is True
    assert not (tmp_path.parent / "escaped.xlsx").exists()
