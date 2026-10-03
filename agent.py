"""
Excel / Word AIエージェント - Anthropic / OpenRouter / Ollama / Gemini 対応
"""
import os
import re
import json
from abc import ABC, abstractmethod
from typing import Callable
import excel_tools
import word_tools
import text_tools

# Excel + Word + Text ツールを統合
TOOLS = excel_tools.TOOLS + word_tools.TOOLS + text_tools.TOOLS

def execute_tool(name: str, tool_input: dict) -> str:
    if name in excel_tools.TOOL_FUNCTIONS:
        return excel_tools.execute_tool(name, tool_input)
    if name in word_tools.TOOL_FUNCTIONS:
        return word_tools.execute_tool(name, tool_input)
    if name in text_tools.TOOL_FUNCTIONS:
        return text_tools.execute_tool(name, tool_input)
    return json.dumps({"success": False, "error": f"不明なツール: {name}"}, ensure_ascii=False)

SYSTEM_PROMPT = """You are a specialized AI agent for manipulating Excel and Word files.
Follow the user's instructions and use the appropriate tools to read, edit, and format files.
Always respond to the user in Japanese.

## General guidelines
- Understand the user's intent and call the necessary tools in sequence to complete the task
- Write operations are saved automatically (no need to call save_excel / save_word separately)
- If an error occurs, explain the cause and try an alternative approach when possible
- When the task is complete, briefly report what was done in Japanese

## Excel operations
- Read and write using cell addresses (A1, B2, etc.)
- Formulas (e.g. =SUM(A1:A10)) can also be set
- Adding/deleting sheets and applying formatting are supported

## Handling large Excel files (important)
- Always call get_sheet_info first to check the total number of rows before reading
- If a sheet has more than 100 rows, do NOT read it all at once
- Read in chunks using min_row / max_row in read_sheet (e.g. rows 1-50, then 51-100)
- Chunk size guideline: 50-100 rows per call depending on column count
- For wide sheets (many columns), reduce chunk size accordingly
- When processing each chunk, complete the required operations before moving to the next chunk

## Word operations
- Paragraphs are managed by index (0-based). Use read_document first to confirm indices
- Use insert_paragraph (at a specific position) or append_paragraph (at the end) to insert text
- Use replace_text for editing and proofreading
- Use add_heading for headings and insert_image for images

## Text / Markdown operations
- .txt / .md files can be read as-is with read_text_file
- Use parse_markdown to structure content into headings, tables, lists, and paragraphs
- Parsed content can be reflected in Word headings/paragraphs or Excel cells
- For Markdown tables: parse_markdown → add_table (Word) or write_range (Excel)

## File path handling
- If only a filename is given, it is treated as a file in the current directory
- To write to a non-existent file, use create_if_missing=true to create it
"""

# Anthropic形式のツール定義をOpenAI形式に変換するヘルパー
def _to_openai(tools: list) -> list:
    return [
        {"type": "function", "function": {
            "name": t["name"],
            "description": t["description"],
            "parameters": t["input_schema"],
        }}
        for t in tools
    ]

OPENAI_TOOLS = _to_openai(TOOLS)

# ── ツール選択ロジック ─────────────────────────────────────────────────────

# 英単語は前後が英字でないときだけ一致させる（"password" や "keyword" を Word と誤判定しない）。
# 日本語は助詞が直後に続くため \b ではなく英字の前後判定を使う。
_EXCEL_RE = re.compile(r"\.xls[xm]?(?![a-z])|(?<![a-z])(?:excel|cells?|sheets?|spreadsheets?)(?![a-z])"
                       r"|エクセル|スプレッドシート|シート|セル")
_WORD_RE  = re.compile(r"\.docx?(?![a-z])|(?<![a-z])(?:word|documents?|paragraphs?)(?![a-z])"
                       r"|ワード|文書|段落|ドキュメント")
_TEXT_RE  = re.compile(r"\.(?:txt|md|csv)(?![a-z])|(?<![a-z])(?:markdown|readme|csv|text file)(?![a-z])"
                       r"|テキスト|マークダウン")


def _select_tools(history: list) -> tuple[list, list]:
    """
    ユーザーの発言からファイル種別を判定し、適切なツールセットを返す。
    テキスト/Markdownが含まれる場合は text_tools も追加する。
    Returns: (anthropic_tools, openai_tools)
    """
    # システムプロンプトやツール結果には "Excel" "Word" などが必ず含まれるため、
    # ユーザーが入力したテキストだけを判定対象にする
    text = " ".join(
        msg["content"].lower() for msg in history
        if msg.get("role") == "user" and isinstance(msg.get("content"), str)
    )

    has_excel = bool(_EXCEL_RE.search(text))
    has_word  = bool(_WORD_RE.search(text))
    has_text  = bool(_TEXT_RE.search(text))

    # テキスト/Markdownは単体で使うことはなく、必ずExcel/Wordと組み合わせる
    if has_excel and not has_word:
        tools = excel_tools.TOOLS
    elif has_word and not has_excel:
        tools = word_tools.TOOLS
    elif has_excel and has_word:
        tools = excel_tools.TOOLS + word_tools.TOOLS
    else:
        return TOOLS, OPENAI_TOOLS  # 判定不能 → 全ツール

    if has_text:
        tools = tools + text_tools.TOOLS
    return tools, _to_openai(tools)

_GRAY = "\033[90m"
_RESET = "\033[0m"
_CLEAR_LINE = "\033[2K\r"


_LOG_ARGS_MAX = 200


def _log_tool(name: str, input_dict: dict):
    # write_range などの大きな引数で画面が埋まらないよう省略する
    args = json.dumps(input_dict, ensure_ascii=False)
    if len(args) > _LOG_ARGS_MAX:
        args = args[:_LOG_ARGS_MAX] + "…"
    # スピナーが表示中の場合でも行を上書きして整合を保つ
    print(f"{_CLEAR_LINE}{_GRAY}  [tool] {name}({args}){_RESET}")


# Anthropic の1応答あたりの最大出力トークン数（大きな write_range でも途切れにくいように）
_ANTHROPIC_MAX_TOKENS = 16000
# OpenAI 互換プロバイダーは出力上限がモデルごとに異なるため控えめにする
_OPENAI_MAX_TOKENS = 4096
# 1回の発言で許可するツール呼び出しの往復回数（モデルがループした場合の歯止め）
_MAX_TOOL_ROUNDS = 50

_TRUNCATED_NOTE = "\n[出力トークンの上限に達したため応答が途中で終了しました]"
_TOOL_LIMIT_MESSAGE = f"[ツール呼び出しが {_MAX_TOOL_ROUNDS} 回に達したため中断しました]"


# ── ベースクラス ─────────────────────────────────────────────────────────────

class _BaseAgent(ABC):
    # スピナーのラベルを更新するコールバック（main.py から注入）
    on_tool_start: Callable[[str], None] | None = None

    def __init__(self):
        self._history: list[dict] = self._initial_history()

    def chat(self, user_message: str) -> str:
        checkpoint = len(self._history)
        self._history.append({"role": "user", "content": user_message})
        try:
            return self._run_loop()
        except BaseException:
            # API エラーや Ctrl+C で中断すると、tool_use に対応する結果が欠けた履歴が残り
            # 以降のリクエストがすべて失敗する。この発言の分を丸ごと巻き戻す
            del self._history[checkpoint:]
            raise

    def reset(self):
        self._history = self._initial_history()

    def _initial_history(self) -> list[dict]:
        return []

    def _call_tool(self, name: str, args: dict) -> str:
        if self.on_tool_start:
            self.on_tool_start(name)
        _log_tool(name, args)
        return execute_tool(name, args)

    @abstractmethod
    def _run_loop(self) -> str: ...


# ── Anthropic バックエンド ────────────────────────────────────────────────────

def _with_cache_breakpoint(messages: list[dict]) -> list[dict]:
    """最後のメッセージにキャッシュ指定を付けたコピーを返す。

    ツールループでは毎回会話全体を再送するため、前回までの部分をプロンプトキャッシュから読ませる。
    """
    last = messages[-1]
    content = last["content"]
    blocks = [{"type": "text", "text": content}] if isinstance(content, str) else list(content)
    blocks[-1] = {**blocks[-1], "cache_control": {"type": "ephemeral"}}
    return messages[:-1] + [{**last, "content": blocks}]


class _AnthropicAgent(_BaseAgent):
    def __init__(self, model: str):
        import anthropic
        self._client = anthropic.Anthropic()
        self._model = model
        super().__init__()

    def _run_loop(self) -> str:
        active_tools, _ = _select_tools(self._history)
        # ツール定義とシステムプロンプトは毎回同じなのでキャッシュする
        system = [{"type": "text", "text": SYSTEM_PROMPT, "cache_control": {"type": "ephemeral"}}]
        for _ in range(_MAX_TOOL_ROUNDS):
            resp = self._client.messages.create(
                model=self._model,
                max_tokens=_ANTHROPIC_MAX_TOKENS,
                system=system,
                tools=active_tools,
                messages=_with_cache_breakpoint(self._history),
            )
            text = "\n".join(b.text for b in resp.content if b.type == "text")

            if resp.stop_reason != "tool_use":
                # max_tokens で途切れた tool_use を履歴に残すと、対応する tool_result が無いため
                # 次のリクエストが失敗する。テキスト部分だけを残す
                if text:
                    self._history.append({"role": "assistant", "content": text})
                if resp.stop_reason == "max_tokens":
                    text += _TRUNCATED_NOTE
                elif resp.stop_reason not in ("end_turn", "stop_sequence"):
                    text += f"\n[終了理由: {resp.stop_reason}]"
                return text

            self._history.append({"role": "assistant", "content": resp.content})
            tool_results = [
                {
                    "type": "tool_result",
                    "tool_use_id": block.id,
                    "content": self._call_tool(block.name, block.input),
                }
                for block in resp.content if block.type == "tool_use"
            ]
            self._history.append({"role": "user", "content": tool_results})

        return _TOOL_LIMIT_MESSAGE


# ── OpenAI互換バックエンド（OpenRouter / Ollama 共通）────────────────────────

class _OpenAICompatAgent(_BaseAgent):
    def __init__(self, model: str, base_url: str, api_key: str):
        from openai import OpenAI
        self._client = OpenAI(base_url=base_url, api_key=api_key)
        self._model = model
        super().__init__()

    def _initial_history(self) -> list[dict]:
        return [{"role": "system", "content": SYSTEM_PROMPT}]

    def _run_loop(self) -> str:
        _, active_tools = _select_tools(self._history)
        for _ in range(_MAX_TOOL_ROUNDS):
            resp = self._client.chat.completions.create(
                model=self._model,
                max_tokens=_OPENAI_MAX_TOKENS,
                tools=active_tools,
                messages=self._history,
            )
            choice = resp.choices[0]
            msg = choice.message

            # アシスタントメッセージを履歴に追加
            self._history.append(msg.model_dump(exclude_none=True))

            # Gemini / Ollama などは tool_calls があっても finish_reason が "stop" になることがある。
            # ここで終了すると tool_calls に対応する結果が欠けて次のリクエストが失敗するため、
            # tool_calls の有無だけで判定する
            if not msg.tool_calls:
                text = msg.content or ""
                if choice.finish_reason == "length":
                    text += _TRUNCATED_NOTE
                return text

            # ツール実行
            for tc in msg.tool_calls:
                try:
                    args = json.loads(tc.function.arguments or "{}")
                except json.JSONDecodeError:
                    args = None
                if isinstance(args, dict):
                    result = self._call_tool(tc.function.name, args)
                else:
                    # 引数なしで実行せず、モデルに再試行させる
                    result = json.dumps(
                        {"success": False, "error": "ツール引数を JSON オブジェクトとして解釈できません"},
                        ensure_ascii=False,
                    )
                self._history.append({
                    "role": "tool",
                    "tool_call_id": tc.id,
                    "content": result,
                })

        return _TOOL_LIMIT_MESSAGE


# ── ファクトリ関数 ────────────────────────────────────────────────────────────

# プロバイダーごとのデフォルトモデル
_DEFAULT_MODELS = {
    "anthropic":  "claude-sonnet-5-5",
    "openrouter": "anthropic/claude-sonnet-5.5",
    "ollama":     "qwen2.5:7b",
    "gemini":     "gemini-3.8-flash",
}

_OPENROUTER_BASE_URL = "https://openrouter.ai/api/v1"
_OLLAMA_BASE_URL     = "http://localhost:11434/v1"
_GEMINI_BASE_URL     = "https://generativelanguage.googleapis.com/v1beta/openai/"


def create_agent(
    provider: str = "anthropic",
    model: str | None = None,
    base_url: str | None = None,
    api_key: str | None = None,
) -> _BaseAgent:
    """
    プロバイダーに応じたエージェントを生成する。

    provider: "anthropic" | "openrouter" | "ollama" | "gemini"
    model:    省略時はプロバイダーのデフォルトモデルを使用
    base_url: OpenAI互換エンドポイントのURL（ollama のカスタムポートなど）
    api_key:  APIキー（省略時は環境変数から取得）
    """
    provider = provider.lower()
    resolved_model = model or _DEFAULT_MODELS.get(provider, "")

    if base_url is not None:
        if not (base_url.startswith("http://") or base_url.startswith("https://")):
            raise ValueError(f"base_url は http:// または https:// で始まる必要があります: {base_url}")

    if provider == "anthropic":
        return _AnthropicAgent(model=resolved_model)

    if provider == "openrouter":
        key = api_key or os.environ.get("OPENROUTER_API_KEY", "")
        if not key:
            raise ValueError(
                "OPENROUTER_API_KEY が設定されていません。\n"
                "  export OPENROUTER_API_KEY='sk-or-...'\n"
                "または .env ファイルに記載してください。"
            )
        url = base_url or _OPENROUTER_BASE_URL
        return _OpenAICompatAgent(model=resolved_model, base_url=url, api_key=key)

    if provider == "ollama":
        key = api_key or "ollama"  # Ollama は認証不要なのでダミーキーでOK
        url = base_url or _OLLAMA_BASE_URL
        return _OpenAICompatAgent(model=resolved_model, base_url=url, api_key=key)

    if provider == "gemini":
        key = api_key or os.environ.get("GEMINI_API_KEY", "")
        if not key:
            raise ValueError(
                "GEMINI_API_KEY が設定されていません。\n"
                "  export GEMINI_API_KEY='AIza...'\n"
                "または .env ファイルに記載してください。"
            )
        url = base_url or _GEMINI_BASE_URL
        return _OpenAICompatAgent(model=resolved_model, base_url=url, api_key=key)

    raise ValueError(
        f"不明なプロバイダー: '{provider}'\n"
        "使用可能: anthropic / openrouter / ollama / gemini"
    )


# 後方互換のためのエイリアス
class ExcelAgent:
    """後方互換ラッパー（anthropic プロバイダー固定）"""
    def __init__(self, model: str = _DEFAULT_MODELS["anthropic"]):
        self._agent = create_agent("anthropic", model=model)

    def chat(self, msg: str) -> str:
        return self._agent.chat(msg)

    def reset(self):
        self._agent.reset()
