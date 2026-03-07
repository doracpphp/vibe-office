"""
テキスト / Markdownファイル読み込みツール
.txt / .md を読んで Word・Excel に反映する際の前処理を担う
"""
import os
import re
import json
from typing import Optional

_BASE_DIR = os.path.realpath(os.getcwd())


def _safe_path(file_path: str) -> str:
    abs_path = os.path.realpath(os.path.join(_BASE_DIR, file_path))
    if not abs_path.startswith(_BASE_DIR + os.sep) and abs_path != _BASE_DIR:
        raise PermissionError(
            f"アクセス拒否: 作業ディレクトリ外のパスは操作できません: {abs_path}"
        )
    return abs_path


def _reraise_if_fatal(e: Exception) -> None:
    if isinstance(e, (KeyboardInterrupt, SystemExit)):
        raise


# ── ツール関数 ──────────────────────────────────────────────────────────────


def read_text_file(file_path: str, encoding: str = "utf-8") -> dict:
    """テキストファイル（.txt / .md など）の内容をそのまま読み取る"""
    try:
        abs_path = _safe_path(file_path)
        if not os.path.exists(abs_path):
            return {"success": False, "error": f"ファイルが見つかりません: {abs_path}"}
        with open(abs_path, encoding=encoding) as f:
            content = f.read()
        lines = content.splitlines()
        return {
            "success": True,
            "file_path": abs_path,
            "content": content,
            "line_count": len(lines),
            "char_count": len(content),
        }
    except Exception as e:
        _reraise_if_fatal(e)
        return {"success": False, "error": str(e)}


def parse_markdown(file_path: str, encoding: str = "utf-8") -> dict:
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
        abs_path = _safe_path(file_path)
        if not os.path.exists(abs_path):
            return {"success": False, "error": f"ファイルが見つかりません: {abs_path}"}
        with open(abs_path, encoding=encoding) as f:
            raw = f.read()

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
        _reraise_if_fatal(e)
        return {"success": False, "error": str(e)}


# ── Markdownパーサー（外部ライブラリ不使用）─────────────────────────────────

def _parse_table(lines: list[str]) -> Optional[dict]:
    """Markdownテーブル行群を解析して dict を返す。失敗時は None"""
    if len(lines) < 2:
        return None
    # セパレータ行（ --- | --- 形式）チェック
    sep = lines[1].strip()
    if not re.match(r'^[\s|:\-]+$', sep):
        return None

    def split_row(line: str) -> list[str]:
        return [c.strip() for c in line.strip().strip("|").split("|")]

    headers = split_row(lines[0])
    rows = [split_row(l) for l in lines[2:] if l.strip()]
    return {"type": "table", "headers": headers, "rows": rows}


def parse_inline_formatting(text: str) -> list[dict]:
    """
    テキスト内の **bold** / *italic* / ***bold-italic*** を解析してランリストを返す。
    各ランは {"text": str, "bold": bool, "italic": bool} の辞書。
    Word の append_rich_paragraph に渡すことで書式付き段落を作れる。

    例:
      "Hello **world** and *foo*"
      → [{"text":"Hello ","bold":False,"italic":False},
         {"text":"world","bold":True,"italic":False},
         {"text":" and ","bold":False,"italic":False},
         {"text":"foo","bold":False,"italic":True}]
    """
    runs: list[dict] = []
    # ***bold-italic*** > **bold** > *italic* の順にマッチ
    pattern = re.compile(r'(\*\*\*(.+?)\*\*\*|\*\*(.+?)\*\*|\*(.+?)\*|__(.+?)__|_(.+?)_)')
    pos = 0
    for m in pattern.finditer(text):
        if m.start() > pos:
            runs.append({"text": text[pos:m.start()], "bold": False, "italic": False})
        raw = m.group(0)
        inner = m.group(2) or m.group(3) or m.group(4) or m.group(5) or m.group(6) or ""
        bold   = raw.startswith("***") or raw.startswith("**") or raw.startswith("__")
        italic = raw.startswith("***") or (raw.startswith("*") and not raw.startswith("**")) \
                 or (raw.startswith("_") and not raw.startswith("__"))
        runs.append({"text": inner, "bold": bold, "italic": italic})
        pos = m.end()
    if pos < len(text):
        runs.append({"text": text[pos:], "bold": False, "italic": False})
    return runs or [{"text": text, "bold": False, "italic": False}]


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
        if re.match(r'^(\-{3,}|\*{3,}|_{3,})\s*$', line):
            blocks.append({"type": "hr"})
            i += 1
            continue

        # ── 見出し (ATX) ──
        m = re.match(r'^(#{1,6})\s+(.*)', line)
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
                        blocks.append({"type": "paragraph", "text": tl.strip()})
            continue

        # ── リスト ──
        if re.match(r'^(\s*[-*+]\s|\s*\d+\.\s)', line):
            list_lines = []
            ordered = bool(re.match(r'^\s*\d+\.', line))
            while i < len(lines) and re.match(r'^(\s*[-*+]\s|\s*\d+\.\s)', lines[i]):
                item = re.sub(r'^\s*[-*+]\s+|\s*\d+\.\s+', '', lines[i]).strip()
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
        para_lines = []
        while i < len(lines) and lines[i].strip() and not lines[i].startswith("#") \
                and not lines[i].startswith("```") and "|" not in lines[i] \
                and not re.match(r'^(\s*[-*+]\s|\s*\d+\.\s)', lines[i]) \
                and not re.match(r'^(\-{3,}|\*{3,}|_{3,})\s*$', lines[i]):
            para_lines.append(lines[i].strip())
            i += 1
        if para_lines:
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
            "in a text string and return a list of runs with bold/italic flags. "
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
                "encoding": {"type": "string", "description": "File encoding (default: utf-8)"},
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
                "encoding": {"type": "string", "description": "File encoding (default: utf-8)"},
            },
            "required": ["file_path"],
        },
    },
]

TOOL_FUNCTIONS = {
    "parse_inline_formatting": lambda text: parse_inline_formatting(text),
    "read_text_file": read_text_file,
    "parse_markdown": parse_markdown,
}


def execute_tool(tool_name: str, tool_input: dict) -> str:
    func = TOOL_FUNCTIONS.get(tool_name)
    if not func:
        return json.dumps({"success": False, "error": f"不明なツール: {tool_name}"}, ensure_ascii=False)
    result = func(**tool_input)
    return json.dumps(result, ensure_ascii=False, indent=2)
