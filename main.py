#!/usr/bin/env python3
"""
vibe-office - チャットインターフェース

使い方:
  python3 main.py                              # Anthropic (デフォルト)
  python3 main.py --provider openrouter        # OpenRouter
  python3 main.py --provider ollama            # Ollama (ローカル)
  python3 main.py --provider ollama --model llama3.2
  python3 main.py --provider openrouter --model google/gemini-3.5-flash
  python3 main.py --provider ollama --base-url http://localhost:11434/v1
  python3 main.py --provider gemini
  python3 main.py --provider gemini --model gemini-3.5-flash-lite
"""
import sys
import os
import argparse
import threading
import itertools
import time
import unicodedata

try:
    import readline  # noqa: F401 - 矢印キー履歴ナビゲーション有効化
except ImportError:
    pass

# ANSI カラーコード
RESET  = "\033[0m"
BOLD   = "\033[1m"
CYAN   = "\033[36m"
GREEN  = "\033[32m"
YELLOW = "\033[33m"
RED    = "\033[31m"
GRAY   = "\033[90m"
BLUE   = "\033[34m"

# カーソル操作
HIDE_CURSOR = "\033[?25l"
SHOW_CURSOR = "\033[?25h"
CLEAR_LINE  = "\033[2K\r"


def enable_ansi_colors() -> None:
    """Windows の従来のコンソール（conhost）でも ANSI エスケープシーケンスを解釈させる。

    Windows Terminal では最初から有効なので影響しない。
    """
    if os.name != "nt":
        return
    import ctypes
    kernel32 = ctypes.windll.kernel32
    handle = kernel32.GetStdHandle(-11)  # STD_OUTPUT_HANDLE
    mode = ctypes.c_uint32()
    if kernel32.GetConsoleMode(handle, ctypes.byref(mode)):
        kernel32.SetConsoleMode(handle, mode.value | 0x0004)  # ENABLE_VIRTUAL_TERMINAL_PROCESSING


class Spinner:
    """AIが処理中であることを示すスピナー"""

    _FRAMES = ["⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏"]
    _PHASES = [
        "考えています",
        "処理中",
        "ツールを実行中",
    ]

    def __init__(self):
        self._stop_event = threading.Event()
        self._thread: threading.Thread | None = None
        self._current_label = self._PHASES[0]

    def set_label(self, label: str):
        self._current_label = label

    def start(self):
        self._stop_event.clear()
        self._current_label = self._PHASES[0]
        self._thread = threading.Thread(target=self._spin, daemon=True)
        print(HIDE_CURSOR, end="", flush=True)
        self._thread.start()

    def stop(self):
        self._stop_event.set()
        if self._thread:
            self._thread.join()
        # スピナー行を消してカーソルを戻す
        print(f"{CLEAR_LINE}{SHOW_CURSOR}", end="", flush=True)

    def _spin(self):
        for frame in itertools.cycle(self._FRAMES):
            if self._stop_event.is_set():
                break
            print(f"{CLEAR_LINE}{CYAN}{frame}{RESET} {GRAY}{self._current_label}...{RESET}",
                  end="", flush=True)
            time.sleep(0.08)


HELP_TEXT = """コマンド:
  /help   - このヘルプを表示
  /reset  - 会話履歴をリセット
  /cwd    - 作業ディレクトリを表示
  /cd     - 作業ディレクトリを変更
  /ls     - ファイル一覧を表示
  /quit   - 終了 (Ctrl+C でも可)"""


def load_dotenv(path: str) -> None:
    """.env を読み込んで環境変数に設定する（既に設定済みの変数は上書きしない）"""
    try:
        with open(path, encoding="utf-8") as f:
            lines = f.read().splitlines()
    except FileNotFoundError:
        return
    for line in lines:
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.removeprefix("export ").split("=", 1)
        key, value = key.strip(), value.strip()
        if len(value) >= 2 and value[0] == value[-1] and value[0] in "\"'":
            value = value[1:-1]
        os.environ.setdefault(key, value)


def _display_width(text: str) -> int:
    """全角文字を2桁として数えた表示幅"""
    return sum(2 if unicodedata.east_asian_width(c) in "WF" else 1 for c in text)


def print_banner(provider: str, model: str):
    provider_label = {
        "anthropic":  f"{CYAN}Anthropic{RESET}",
        "openrouter": f"{GREEN}OpenRouter{RESET}",
        "ollama":     f"{YELLOW}Ollama (local){RESET}",
        "gemini":     f"{BLUE}Google Gemini{RESET}",
    }.get(provider, provider)

    # 全角文字を含んでも枠線が揃うよう表示幅から余白を計算する
    lines = ["vibe-office", "自然言語でExcel・Wordを操作するAIエージェント"]
    inner = max(_display_width(l) for l in lines) + 4
    box = [f"╔{'═' * inner}╗"]
    for l in lines:
        pad = inner - _display_width(l)
        box.append(f"║{' ' * (pad // 2)}{l}{' ' * (pad - pad // 2)}║")
    box.append(f"╚{'═' * inner}╝")
    box_str = "\n".join(box)

    print(f"""
{CYAN}{BOLD}{box_str}{RESET}
  Provider : {provider_label}
  Model    : {BOLD}{model}{RESET}

{GRAY}{HELP_TEXT}{RESET}
""")


def handle_command(cmd: str):
    """組み込みコマンドを処理。
    True  → コマンド処理済み
    None  → /reset シグナル
    False → コマンドではない（agentへ渡す）
    """
    parts = cmd.strip().split(maxsplit=1)
    command = parts[0].lower()
    arg = parts[1] if len(parts) > 1 else ""

    if command == "/help":
        print(f"{GRAY}{HELP_TEXT}{RESET}")
        return True

    if command in ("/quit", "/exit"):
        print(f"{YELLOW}終了します。{RESET}")
        sys.exit(0)

    if command == "/reset":
        return None

    if command == "/cwd":
        print(f"{CYAN}作業ディレクトリ: {os.getcwd()}{RESET}")
        return True

    if command == "/cd":
        if not arg:
            print(f"{RED}使い方: /cd <ディレクトリ>{RESET}")
        else:
            target = os.path.expanduser(arg)
            if os.path.isdir(target):
                os.chdir(target)
                print(f"{GREEN}作業ディレクトリを変更しました: {os.getcwd()}{RESET}")
            else:
                print(f"{RED}ディレクトリが見つかりません: {target}{RESET}")
        return True

    if command == "/ls":
        target_dir = arg if arg else "."
        try:
            entries = sorted(os.listdir(target_dir))
            xlsx_files = [f for f in entries if f.endswith((".xlsx", ".xls", ".xlsm"))]
            docx_files = [f for f in entries if f.endswith((".docx", ".doc"))]
            other_files = [f for f in entries if f not in xlsx_files and f not in docx_files]

            print(f"\n{CYAN}[{os.path.abspath(target_dir)}]{RESET}")
            for f in xlsx_files:
                print(f"  {GREEN}{f}{RESET}  (Excel)")
            for f in docx_files:
                print(f"  {BLUE}{f}{RESET}  (Word)")
            for f in other_files:
                icon = "/" if os.path.isdir(os.path.join(target_dir, f)) else ""
                print(f"  {f}{icon}")
            print()
        except Exception as e:
            print(f"{RED}エラー: {e}{RESET}")
        return True

    # "/Users/me/data.xlsx を開いて" のような絶対パスで始まる入力はエージェントへ渡す
    if "/" in command[1:]:
        return False

    print(f"{RED}不明なコマンド: {command}（/help でコマンド一覧を表示）{RESET}")
    return True


def check_api_key(provider: str) -> bool:
    """必要なAPIキーが設定されているか確認する"""
    if provider == "anthropic":
        if not os.environ.get("ANTHROPIC_API_KEY"):
            print(f"{RED}エラー: ANTHROPIC_API_KEY が設定されていません{RESET}")
            print("  export ANTHROPIC_API_KEY='sk-ant-...'")
            return False

    elif provider == "openrouter":
        if not os.environ.get("OPENROUTER_API_KEY"):
            print(f"{RED}エラー: OPENROUTER_API_KEY が設定されていません{RESET}")
            print("  export OPENROUTER_API_KEY='sk-or-...'")
            return False

    elif provider == "gemini":
        if not os.environ.get("GEMINI_API_KEY"):
            print(f"{RED}エラー: GEMINI_API_KEY が設定されていません{RESET}")
            print("  export GEMINI_API_KEY='AIza...'")
            print(f"{GRAY}  または .env ファイルに GEMINI_API_KEY=AIza... と記載してください{RESET}")
            return False

    # ollama はキー不要
    return True


def parse_args():
    parser = argparse.ArgumentParser(
        description="vibe-office - 自然言語でExcel・Wordを操作するAIエージェント",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
例:
  python3 main.py
  python3 main.py --provider openrouter --model google/gemini-3.5-flash
  python3 main.py --provider ollama --model llama3.2
  python3 main.py --provider ollama --base-url http://192.168.1.10:11434/v1
        """
    )
    parser.add_argument(
        "--provider", "-p",
        choices=["anthropic", "openrouter", "ollama", "gemini"],
        default=os.environ.get("EXCEL_AGENT_PROVIDER", "anthropic"),
        help="AIプロバイダー (デフォルト: anthropic)",
    )
    parser.add_argument(
        "--model", "-m",
        default=os.environ.get("EXCEL_AGENT_MODEL"),
        help="使用するモデル名 (省略時はプロバイダーのデフォルト)",
    )
    parser.add_argument(
        "--base-url",
        default=os.environ.get("EXCEL_AGENT_BASE_URL"),
        help="OpenAI互換エンドポイントのベースURL (ollama/openrouter カスタム用)",
    )
    return parser.parse_args()


def main():
    enable_ansi_colors()
    # スクリプトと同じディレクトリの .env を読み込む（--provider などの既定値にも使うため最初に行う）
    load_dotenv(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env"))
    args = parse_args()

    if not check_api_key(args.provider):
        sys.exit(1)

    from agent import create_agent, _DEFAULT_MODELS

    resolved_model = args.model or _DEFAULT_MODELS.get(args.provider, "unknown")

    print_banner(args.provider, resolved_model)

    try:
        agent = create_agent(
            provider=args.provider,
            model=args.model,
            base_url=args.base_url,
        )
    except Exception as e:
        print(f"{RED}エージェントの初期化に失敗しました: {e}{RESET}")
        sys.exit(1)

    print(f"{GRAY}ヒント: 「sample.xlsxを開いてA1の値を見せて」のように話しかけてみてください{RESET}\n")

    spinner = Spinner()

    # ツール実行時にスピナーのラベルを切り替えるコールバックを注入
    def _on_tool_start(tool_name: str):
        spinner.set_label(f"{tool_name} を実行中")

    # create_agent は _BaseAgent を返す。ExcelAgentラッパー経由でも直接でも対応
    inner = getattr(agent, "_agent", agent)
    inner.on_tool_start = _on_tool_start

    while True:
        try:
            user_input = input(f"{BOLD}{BLUE}You>{RESET} ").strip()
        except (KeyboardInterrupt, EOFError):
            print(f"\n{YELLOW}終了します。{RESET}")
            break

        if not user_input:
            continue

        if user_input.startswith("/"):
            result = handle_command(user_input)
            if result is None:
                agent.reset()
                print(f"{GREEN}会話履歴をリセットしました。{RESET}")
            if result is not False:
                continue

        spinner.start()
        try:
            response = agent.chat(user_input)
            spinner.stop()
            print(f"{BOLD}Agent>{RESET}\n{response}\n")
        except KeyboardInterrupt:
            spinner.stop()
            print(f"{YELLOW}[中断]{RESET}\n")
        except Exception as e:
            spinner.stop()
            print(f"{RED}エラーが発生しました: {e}{RESET}\n")


if __name__ == "__main__":
    main()
