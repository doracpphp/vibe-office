"""
vibe-office MCP サーバー

Claude Code のプラグインとして Excel / Word / Markdown ファイルを操作するツールを提供する。

起動方法（手動テスト用）:
  uv run python mcp_server.py

Claude Code への登録方法:
  claude mcp add vibe-office -- uv run --directory /path/to/vibe-office python mcp_server.py

操作対象のディレクトリ:
  既定ではこのプロジェクトのディレクトリ内のファイルだけを操作できる。
  別のディレクトリを対象にするには環境変数 VIBE_OFFICE_WORKDIR を指定する:
  claude mcp add vibe-office -e VIBE_OFFICE_WORKDIR=/path/to/documents -- uv run --directory /path/to/vibe-office python mcp_server.py
"""
import asyncio
import json
import os
import sys

# 作業ディレクトリを固定する（どこから起動されても同じディレクトリを使う）。
# ツールはこのディレクトリ外のファイルにはアクセスしない
PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
os.chdir(os.path.expanduser(os.environ.get("VIBE_OFFICE_WORKDIR") or PROJECT_DIR))
sys.path.insert(0, PROJECT_DIR)

from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import CallToolResult, Tool, TextContent

import excel_tools
import word_tools
import text_tools
from agent import execute_tool

# 全ツールをまとめる
_ALL_TOOLS = excel_tools.TOOLS + word_tools.TOOLS + text_tools.TOOLS

server = Server("vibe-office")


@server.list_tools()
async def list_tools() -> list[Tool]:
    return [
        Tool(
            name=t["name"],
            description=t["description"],
            inputSchema=t["input_schema"],
        )
        for t in _ALL_TOOLS
    ]


@server.call_tool()
async def call_tool(name: str, arguments: dict) -> CallToolResult:
    result_json = execute_tool(name, arguments or {})
    result = json.loads(result_json)

    # 失敗はクライアントがエラーとして扱えるよう isError を立てて返す
    if result.get("success") is False:
        return CallToolResult(
            content=[TextContent(type="text", text=f"エラー: {result.get('error', '不明なエラー')}")],
            isError=True,
        )
    return CallToolResult(content=[TextContent(type="text", text=result_json)])


async def main():
    async with stdio_server() as (read_stream, write_stream):
        await server.run(
            read_stream,
            write_stream,
            server.create_initialization_options(),
        )


if __name__ == "__main__":
    asyncio.run(main())
