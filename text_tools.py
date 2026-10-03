"""
テキスト / Markdownファイル読み込みツール
.txt / .md を読んで Word・Excel に反映する際の前処理を担う
"""
import os
import re
from typing import Optional

from common import run_tool, safe_path


def _read_text(abs_path: str, encoding: Optional[str]) -> tuple[str, str]:
    """テキストを読み込んで (内容, 使用したエンコーディング) を返す。

    encoding 省略時は UTF-8（BOM 付きも可）→ Shift_JIS(cp932) の順に試す。
    """
    if encoding:
        with open(abs_path, encoding=encoding) as f:
            return f.read(), encoding
    with open(abs_path, "rb") as f:
        raw = f.read()
    for enc in ("utf-8-sig", "cp932"):
        try:
            return raw.decode(enc), enc
        except UnicodeDecodeError:
            continue
    raise ValueError("UTF-8 / Shift_JIS として読めません。encoding を指定してください")


# ── ツール関数 ──────────────────────────────────────────────────────────────


def read_text_file(file_path: str, encoding: Optional[str] = None) -> dict:
    """テキストファイル（.txt / .md など）の内容をそのまま読み取る"""
    try:
        abs_path = safe_path(file_path)
        if not os.path.exists(abs_path):
            return {"success": False, "error": f"ファイルが見つかりません: {abs_path}"}
        content, used_encoding = _read_text(abs_path, encoding)
        lines = content.splitlines()
        return {
            "success": True,
            "file_path": abs_path,
            "encoding": used_encoding,
            "content": content,
            "line_count": len(lines),
            "char_count": len(content),
        }
    except Exception as e:
        return {"success": False, "error": str(e)}


def parse_markdown(file_path: str, encoding: Optional[str] = None) -> dict:
    """
    Markdownファイルを解析して構造化データを返す。
    返却する blocks リストの各要素の type:
      - heading   : 見出し (level=1〜6, text)
      - paragraph : 本文段落 (text)
      - table     : テーブル (headers, rows)
      - list      : リスト (ordered, items)
      - code      : コードブロック (language, code)
      - hr        : 水平線
    """
    try:
        abs_path = safe_path(file_path)
        if not os.path.exists(abs_path):
            return {"success": False, "error": f"ファイルが見つかりません: {abs_path}"}
        raw, _ = _read_text(abs_path, encoding)

        blocks = _parse_blocks(raw)

        # テーブル一覧（Excel 向けに便利なので別出し）
        tables = [b for b in blocks if b["type"] == "table"]

        # 見出し一覧（Word 向けに便利なので別出し）
        headings = [
            {"level": b["level"], "text": b["text"]}
            for b in blocks if b["type"] == "heading"
        ]

        return {
            "success": True,
            "file_path": abs_path,
            "blocks": blocks,
            "headings": headings,
            "tables": tables,
            "block_count": len(blocks),
        }
    except Exception as e:
        return {"success": False, "error": str(e)}


# ── Markdownパーサー（外部ライブラリ不使用）─────────────────────────────────

def _parse_table(lines: list[str]) -> Optional[dict]:
    """Markdownテーブル行群を解析して dict を返す。失敗時は None"""
    if len(lines) < 2:
        return None
    # セパレータ行（ --- | --- 形式）チェック
    sep = lines[1].strip()
    if "-" not in sep or not re.match(r'^[\s|:\-]+$', sep):
        return None

    def split_row(line: str) -> list[str]:
        return [c.strip() for c in line.strip().strip("|").split("|")]

    headers = split_row(lines[0])
    rows = [split_row(line) for line in lines[2:] if line.strip()]
    return {"type": "table", "headers": headers, "rows": rows}


# バックスラッシュエスケープ（\*）とコードスパン（`code`）の中身は書式記号として扱わない
_LITERAL_RE = re.compile(r'\\([!-/:-@\[-`{-~])|(?<!`)(`+)(?!`)(.+?)(?<!`)\2(?!`)')
# [テキスト](URL) / ![代替テキスト](画像) はテキスト部分だけ残す
_LINK_RE = re.compile(r'!?\[([^\]]*)\]\((?:[^()]|\([^()]*\))*\)')
# * は単語の途中でも強調になるが、_ は単語の切れ目にあるときだけ（file_name_here 対策）。
# どちらも記号のすぐ内側が空白なら強調にしない（2 * 3 * 4 対策）
_EMPHASIS_RE = re.compile(
    r'(?P<star>\*{1,3})(?![\s*])(?P<star_text>.+?)(?<![\s*])(?P=star)(?!\*)'
    r'|(?<!\w)(?P<under>_{1,3})(?![\s_])(?P<under_text>.+?)(?<![\s_])(?P=under)(?!\w)'
)
# 退避したエスケープ・コードの中身は Unicode 私用領域の文字に置き換えておく
_PLACEHOLDER_BASE = 0xF0000


def _parse_emphasis(text: str, bold: bool, italic: bool) -> list[dict]:
    runs: list[dict] = []
    pos = 0
    for m in _EMPHASIS_RE.finditer(text):
        if m.start() > pos:
            runs.append({"text": text[pos:m.start()], "bold": bold, "italic": italic})
        if m.group("star"):
            delim, inner = m.group("star", "star_text")
        else:
            delim, inner = m.group("under", "under_text")
        # 内側の強調（**太字 *斜体* 太字**）も外側の書式を引き継いで解析する
        runs.extend(_parse_emphasis(inner, bold or len(delim) >= 2, italic or len(delim) != 2))
        pos = m.end()
    if pos < len(text):
        runs.append({"text": text[pos:], "bold": bold, "italic": italic})
    return runs


def parse_inline_formatting(text: str) -> list[dict]:
    """
    テキスト内の **bold** / *italic* / ***bold-italic*** を解析してランリストを返す。
    各ランは {"text": str, "bold": bool, "italic": bool} の辞書。
    Word の append_rich_paragraph に渡すことで書式付き段落を作れる。
    `code` は記号を外した書式なしのテキスト、[テキスト](URL) はテキスト部分だけになる。

    例:
      "Hello **world** and *foo*"
      → [{"text":"Hello ","bold":False,"italic":False},
         {"text":"world","bold":True,"italic":False},
         {"text":" and ","bold":False,"italic":False},
         {"text":"foo","bold":False,"italic":True}]
    """
    literals: list[str] = []

    def stash(m: re.Match) -> str:
        if m.group(1) is not None:
            literal = m.group(1)
        else:
            literal = m.group(3)
            # CommonMark と同様に両端の空白を1つずつ除く（`` `x` `` のような書き方のため）
            if len(literal) >= 2 and literal[0] == literal[-1] == " " and literal.strip():
                literal = literal[1:-1]
        literals.append(literal)
        return chr(_PLACEHOLDER_BASE + len(literals) - 1)

    masked = _LINK_RE.sub(r'\1', _LITERAL_RE.sub(stash, text))
    restore = {_PLACEHOLDER_BASE + i: literal for i, literal in enumerate(literals)}

    runs: list[dict] = []
    for run in _parse_emphasis(masked, False, False):
        run["text"] = run["text"].translate(restore)
        if not run["text"]:
            continue
        if runs and (runs[-1]["bold"], runs[-1]["italic"]) == (run["bold"], run["italic"]):
            runs[-1]["text"] += run["text"]
        else:
            runs.append(run)
    return runs or [{"text": "", "bold": False, "italic": False}]


_HEADING_RE = re.compile(r'^(#{1,6})\s+(.*)')
_LIST_RE = re.compile(r'^\s*(?:[-*+]|\d+\.)\s+')
_HR_RE = re.compile(r'^(\-{3,}|\*{3,}|_{3,})\s*$')


def _is_block_start(line: str) -> bool:
    """段落以外のブロック（見出し・コード・テーブル・リスト・水平線）の開始行か"""
    return bool(line.startswith("```") or _HR_RE.match(line) or _HEADING_RE.match(line)
                or "|" in line or _LIST_RE.match(line))


def _parse_blocks(src: str) -> list[dict]:
    blocks: list[dict] = []
    lines = src.splitlines()
    i = 0

    while i < len(lines):
        line = lines[i]

        # ── コードブロック ──
        if line.startswith("```"):
            lang = line[3:].strip()
            code_lines = []
            i += 1
            while i < len(lines) and not lines[i].startswith("```"):
                code_lines.append(lines[i])
                i += 1
            blocks.append({"type": "code", "language": lang, "code": "\n".join(code_lines)})
            i += 1
            continue

        # ── 水平線 ──
        if _HR_RE.match(line):
            blocks.append({"type": "hr"})
            i += 1
            continue

        # ── 見出し (ATX) ──
        m = _HEADING_RE.match(line)
        if m:
            blocks.append({
                "type": "heading",
                "level": len(m.group(1)),
                "text": m.group(2).strip(),
            })
            i += 1
            continue

        # ── テーブル ──
        if "|" in line:
            table_lines = []
            while i < len(lines) and "|" in lines[i]:
                table_lines.append(lines[i])
                i += 1
            parsed = _parse_table(table_lines)
            if parsed:
                blocks.append(parsed)
            else:
                # テーブルとして解釈できなければ段落扱い
                for tl in table_lines:
                    if tl.strip():
                        blocks.append({"type": "paragraph", "text": tl.strip(),
                                       "runs": parse_inline_formatting(tl.strip())})
            continue

        # ── リスト ──
        if _LIST_RE.match(line):
            list_lines = []
            ordered = bool(re.match(r'^\s*\d+\.', line))
            while i < len(lines) and _LIST_RE.match(lines[i]):
                item = _LIST_RE.sub('', lines[i], count=1).strip()
                list_lines.append(item)
                i += 1
            blocks.append({
                "type": "list",
                "ordered": ordered,
                "items": list_lines,
                "items_runs": [parse_inline_formatting(it) for it in list_lines],
            })
            continue

        # ── 空行スキップ ──
        if not line.strip():
            i += 1
            continue

        # ── 段落（複数行連続をまとめる）──
        # 先頭行は必ず消費する（ブロック開始と判定されない行で無限ループしないように）
        para_lines = [line.strip()]
        i += 1
        while i < len(lines) and lines[i].strip() and not _is_block_start(lines[i]):
            para_lines.append(lines[i].strip())
            i += 1
        joined = " ".join(para_lines)
        blocks.append({
            "type": "paragraph",
            "text": joined,
            "runs": parse_inline_formatting(joined),
        })

    return blocks


# ── ツール定義 ──────────────────────────────────────────────────────────────

TOOLS = [
    {
        "name": "parse_inline_formatting",
        "description": (
            "Parse inline Markdown formatting (**bold**, *italic*, ***bold-italic***) "
            "in a text string and return {'runs': [...]} with bold/italic flags per run. "
            "Inline `code` becomes plain text and [text](url) links keep only the text. "
            "Pass the result to append_rich_paragraph to write formatted text into Word."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "text": {"type": "string", "description": "Text containing Markdown inline markers"},
            },
            "required": ["text"],
        },
    },
    {
        "name": "read_text_file",
        "description": (
            "Read the raw content of a plain text file (.txt, .md, .csv, etc.). "
            "Use this to load file content before writing it into Word or Excel."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "Path to the text file"},
                "encoding": {"type": "string", "description": "File encoding, e.g. utf-8 or cp932 (default: auto-detect UTF-8 / Shift_JIS)"},
            },
            "required": ["file_path"],
        },
    },
    {
        "name": "parse_markdown",
        "description": (
            "Parse a Markdown file (.md) into structured blocks: headings, paragraphs, "
            "tables, lists, and code blocks. "
            "Use this to map Markdown content to Word headings/paragraphs or Excel tables."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "Path to the Markdown file (.md)"},
                "encoding": {"type": "string", "description": "File encoding, e.g. utf-8 or cp932 (default: auto-detect UTF-8 / Shift_JIS)"},
            },
            "required": ["file_path"],
        },
    },
]

TOOL_FUNCTIONS = {
    "parse_inline_formatting": lambda text: {"success": True, "runs": parse_inline_formatting(text)},
    "read_text_file": read_text_file,
    "parse_markdown": parse_markdown,
}


def execute_tool(tool_name: str, tool_input: dict) -> str:
    return run_tool(TOOL_FUNCTIONS, tool_name, tool_input)
