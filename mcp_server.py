"""
vibe-office MCP サーバー

Claude Code のプラグインとして Excel / Word / Markdown ファイルを操作するツールを提供する。

起動方法（手動テスト用）:
  uv run python mcp_server.py

Claude Code への登録方法:
  ~/.claude/claude_code_config.json に以下を追加:
  {
    "mcpServers": {
      "vibe-office": {
        "command": "uv",
        "args": ["run", "python", "mcp_server.py"],
        "cwd": "/Users/admin/excel-agent"
      }
    }
  }
"""
import asyncio
import json
import os
import sys

# カレントディレクトリをプロジェクトルートに固定
# （MCP サーバーはどこから起動されても同じ作業ディレクトリを使う）
PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
os.chdir(PROJECT_DIR)
sys.path.insert(0, PROJECT_DIR)

from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import Tool, TextContent

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
async def call_tool(name: str, arguments: dict) -> list[TextContent]:
    try:
        result_json = execute_tool(name, arguments)
        result = json.loads(result_json)

        if result.get("success") is False:
            text = f"エラー: {result.get('error', '不明なエラー')}"
        else:
            # 結果を人が読みやすい形式に整形
            text = result_json
    except Exception as e:
        text = f"ツール実行エラー: {e}"

    return [TextContent(type="text", text=text)]


async def main():
    async with stdio_server() as (read_stream, write_stream):
        await server.run(
            read_stream,
            write_stream,
            server.create_initialization_options(),
        )


if __name__ == "__main__":
    asyncio.run(main())
