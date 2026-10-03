"""
excel_tools / word_tools / text_tools で共有するユーティリティ
"""
import os
import re
import json
from typing import Any, Callable

_HEX_COLOR_RE = re.compile(r'^[0-9A-Fa-f]{6}$')


def safe_path(file_path: str) -> str:
    """パストラバーサルを防ぐ。作業ディレクトリ外のパスは拒否する。

    作業ディレクトリは呼び出し時点の cwd。main.py の /cd で移動した場合も
    読み込み先と保存先が常に一致するよう、毎回 cwd から解決する。
    """
    base_dir = os.path.realpath(os.getcwd())
    abs_path = os.path.realpath(os.path.join(base_dir, file_path))
    if os.path.commonpath([base_dir, abs_path]) != base_dir:
        raise PermissionError(
            f"アクセス拒否: 作業ディレクトリ外のパスは操作できません: {abs_path}"
        )
    return abs_path


def validate_hex_color(value: str, name: str) -> str:
    """HEXカラー文字列を検証して正規化する（例: '#FF0000' → 'FF0000'）"""
    normalized = value.lstrip("#").upper()
    if not _HEX_COLOR_RE.match(normalized):
        raise ValueError(f"{name} は6桁の16進数カラーコードで指定してください（例: FF0000）")
    return normalized


def _mtime(abs_path: str) -> int | None:
    try:
        return os.stat(abs_path).st_mtime_ns
    except FileNotFoundError:
        return None


class FileCache:
    """開いたファイルのオブジェクトを保持するキャッシュ。

    ファイルの更新日時を記録しておき、ユーザーが Excel / Word で
    外部編集した場合は読み直す（古い内容での上書きを防ぐ）。
    """

    def __init__(self, load: Callable[[str], Any], create: Callable[[str], Any]):
        self._load = load
        self._create = create
        self._entries: dict[str, tuple[Any, int | None]] = {}

    def get(self, file_path: str, create_if_missing: bool = False) -> Any:
        abs_path = safe_path(file_path)
        mtime = _mtime(abs_path)
        cached = self._entries.get(abs_path)
        if cached and cached[1] == mtime:
            return cached[0]
        if mtime is not None:
            obj = self._load(abs_path)
        elif create_if_missing:
            obj = self._create(abs_path)
        else:
            raise FileNotFoundError(f"ファイルが見つかりません: {abs_path}")
        self._entries[abs_path] = (obj, mtime)
        return obj

    def save(self, file_path: str, obj: Any) -> str:
        abs_path = safe_path(file_path)
        obj.save(abs_path)
        self._entries[abs_path] = (obj, _mtime(abs_path))
        return abs_path

    def evict(self, file_path: str) -> None:
        try:
            self._entries.pop(safe_path(file_path), None)
        except PermissionError:
            pass


def run_tool(tool_functions: dict, tool_name: str, tool_input: dict,
             cache: FileCache | None = None) -> str:
    """ツールを実行してJSON文字列で結果を返す。

    引数の不足・過剰などで例外が出てもエラーとしてモデルに返し、
    エージェントループを止めない。失敗時は編集途中のキャッシュを破棄して
    次回ディスクから読み直す（中途半端な変更が後で保存されるのを防ぐ）。
    """
    func = tool_functions.get(tool_name)
    if not func:
        result = {"success": False, "error": f"不明なツール: {tool_name}"}
    else:
        try:
            result = func(**tool_input)
        except TypeError as e:
            result = {"success": False, "error": f"引数が不正です: {e}"}
        except Exception as e:
            result = {"success": False, "error": str(e)}
    if cache and not result.get("success", True) and isinstance(tool_input.get("file_path"), str):
        cache.evict(tool_input["file_path"])
    # 日付・時刻などJSON非対応の値は文字列化する
    return json.dumps(result, ensure_ascii=False, default=str)
