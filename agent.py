"""
Excel / Word AIエージェント - Anthropic / OpenRouter / Ollama 対応
"""
import os
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

SYSTEM_PROMPT = """あなたはExcelとWordファイルを操作する専門のAIエージェントです。
ユーザーの日本語の指示に従い、適切なツールを使ってファイルの読み取り・編集・書式設定などを行います。

## 基本方針
- ユーザーの指示を理解し、必要なツールを順番に呼び出してタスクを完了させてください
- 書き込み操作は自動的に保存されます（save_excel / save_word を別途呼ぶ必要はありません）
- エラーが発生した場合は原因を説明し、可能なら別のアプローチを試みてください
- 操作が完了したら、何をしたかを簡潔に日本語で報告してください

## Excelの操作
- セルアドレス（A1, B2など）で読み書きできます
- 数式（=SUM(A1:A10)など）も設定できます
- シートの追加・削除、書式設定も可能です

## Wordの操作
- 段落はインデックス（0始まり）で管理します。まず read_document でインデックスを確認してください
- テキストの挿入は insert_paragraph（特定位置）または append_paragraph（末尾）を使います
- 推敲・修正には replace_text が便利です
- 見出し追加は add_heading、画像挿入は insert_image を使います

## テキスト / Markdownの操作
- .txt / .md ファイルは read_text_file でそのまま読み込めます
- Markdownは parse_markdown で見出し・テーブル・リスト・段落に構造化できます
- 読み込んだ内容を Word の見出し/段落や Excel のセルに反映できます
- Markdownのテーブルは parse_markdown → add_table(Word) または write_range(Excel) で転記します

## ファイルパスの扱い
- ファイル名だけ指定した場合、カレントディレクトリのファイルとして扱います
- 存在しないファイルに書き込む場合は create_if_missing=true で新規作成します
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

_EXCEL_KW = {'.xlsx', '.xls', '.xlsm', 'excel', 'エクセル', 'スプレッドシート',
             'シート', 'セル', 'cell', 'sheet'}
_WORD_KW  = {'.docx', '.doc', 'word', 'ワード', '文書', '段落', 'paragraph',
             'document', 'ドキュメント'}
_TEXT_KW  = {'.txt', '.md', 'markdown', 'テキスト', 'text file', 'readme'}

def _select_tools(history: list) -> tuple[list, list]:
    """
    会話履歴全体からファイル種別を判定し、適切なツールセットを返す。
    テキスト/Markdownが含まれる場合は常に text_tools も追加する。
    Returns: (anthropic_tools, openai_tools)
    """
    text = ""
    for msg in history:
        content = msg.get("content", "")
        if isinstance(content, str):
            text += content.lower()
        elif isinstance(content, list):
            for block in content:
                if isinstance(block, dict) and block.get("type") == "text":
                    text += block.get("text", "").lower()

    has_excel = any(kw in text for kw in _EXCEL_KW)
    has_word  = any(kw in text for kw in _WORD_KW)
    has_text  = any(kw in text for kw in _TEXT_KW)

    # テキスト/Markdownは単体で使うことはなく、必ずExcel/Wordと組み合わせる
    if has_excel and not has_word:
        base = excel_tools.TOOLS
    elif has_word and not has_excel:
        base = word_tools.TOOLS
    elif has_excel and has_word:
        base = excel_tools.TOOLS + word_tools.TOOLS
    else:
        base = TOOLS  # 判定不能 → 全ツール

    # テキスト/Markdownキーワードがあれば text_tools を追加
    if has_text and text_tools.TOOLS not in [base]:
        combined = base + [t for t in text_tools.TOOLS if t not in base]
        return combined, _to_openai(combined)

    return base, _to_openai(base)

_GRAY = "\033[90m"
_RESET = "\033[0m"
_CLEAR_LINE = "\033[2K\r"


def _log_tool(name: str, input_dict: dict):
    # スピナーが表示中の場合でも行を上書きして整合を保つ
    print(f"{_CLEAR_LINE}{_GRAY}  [tool] {name}({json.dumps(input_dict, ensure_ascii=False)}){_RESET}")


# ── ベースクラス ─────────────────────────────────────────────────────────────

class _BaseAgent(ABC):
    # スピナーのラベルを更新するコールバック（main.py から注入）
    on_tool_start: Callable[[str], None] | None = None

    def chat(self, user_message: str) -> str:
        self._append_user(user_message)
        return self._run_loop()

    def reset(self):
        self._history_clear()

    @abstractmethod
    def _append_user(self, text: str): ...

    @abstractmethod
    def _history_clear(self): ...

    @abstractmethod
    def _run_loop(self) -> str: ...


# ── Anthropic バックエンド ────────────────────────────────────────────────────

class _AnthropicAgent(_BaseAgent):
    def __init__(self, model: str):
        import anthropic
        self._client = anthropic.Anthropic()
        self._model = model
        self._history: list[dict] = []

    def _history_clear(self):
        self._history = []

    def _append_user(self, text: str):
        self._history.append({"role": "user", "content": text})

    def _run_loop(self) -> str:
        active_tools, _ = _select_tools(self._history)
        while True:
            resp = self._client.messages.create(
                model=self._model,
                max_tokens=4096,
                system=SYSTEM_PROMPT,
                tools=active_tools,
                messages=self._history,
            )
            self._history.append({"role": "assistant", "content": resp.content})

            if resp.stop_reason == "end_turn":
                return "\n".join(
                    b.text for b in resp.content if hasattr(b, "text")
                )

            if resp.stop_reason != "tool_use":
                return f"[予期しない終了理由: {resp.stop_reason}]"

            tool_results = []
            for block in resp.content:
                if block.type != "tool_use":
                    continue
                if self.on_tool_start:
                    self.on_tool_start(block.name)
                _log_tool(block.name, block.input)
                result = execute_tool(block.name, block.input)
                tool_results.append({
                    "type": "tool_result",
                    "tool_use_id": block.id,
                    "content": result,
                })

            self._history.append({"role": "user", "content": tool_results})


# ── OpenAI互換バックエンド（OpenRouter / Ollama 共通）────────────────────────

class _OpenAICompatAgent(_BaseAgent):
    def __init__(self, model: str, base_url: str, api_key: str):
        from openai import OpenAI
        self._client = OpenAI(base_url=base_url, api_key=api_key)
        self._model = model
        self._history: list[dict] = [{"role": "system", "content": SYSTEM_PROMPT}]

    def _history_clear(self):
        self._history = [{"role": "system", "content": SYSTEM_PROMPT}]

    def _append_user(self, text: str):
        self._history.append({"role": "user", "content": text})

    def _run_loop(self) -> str:
        _, active_tools = _select_tools(self._history)
        while True:
            resp = self._client.chat.completions.create(
                model=self._model,
                max_tokens=4096,
                tools=active_tools,
                messages=self._history,
            )
            msg = resp.choices[0].message
            finish = resp.choices[0].finish_reason

            # アシスタントメッセージを履歴に追加
            self._history.append(msg.model_dump(exclude_unset=False))

            if finish == "stop" or not msg.tool_calls:
                return msg.content or ""

            # ツール実行
            for tc in msg.tool_calls:
                name = tc.function.name
                try:
                    args = json.loads(tc.function.arguments)
                except json.JSONDecodeError:
                    args = {}
                if self.on_tool_start:
                    self.on_tool_start(name)
                _log_tool(name, args)
                result = execute_tool(name, args)
                self._history.append({
                    "role": "tool",
                    "tool_call_id": tc.id,
                    "content": result,
                })


# ── ファクトリ関数 ────────────────────────────────────────────────────────────

# プロバイダーごとのデフォルトモデル
_DEFAULT_MODELS = {
    "anthropic":  "claude-sonnet-4-6",
    "openrouter": "anthropic/claude-3.5-sonnet",
    "ollama":     "qwen2.5:7b",
    "gemini":     "gemini-2.0-flash",
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
    def __init__(self, model: str = "claude-sonnet-4-6"):
        self._agent = create_agent("anthropic", model=model)

    def chat(self, msg: str) -> str:
        return self._agent.chat(msg)

    def reset(self):
        self._agent.reset()
