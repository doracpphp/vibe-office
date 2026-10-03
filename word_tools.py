"""
Word操作ツール群 - python-docxを使用してWordファイルを操作する
"""
import os
import re
from typing import Optional
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from common import FileCache, run_tool, safe_path, save_file, validate_hex_color


# 現在開いているドキュメントのキャッシュ
_cache = FileCache(load=Document, create=lambda _: Document())

# "Heading 1" / "Heading" などの見出しスタイル名
_HEADING_STYLE_RE = re.compile(r'^Heading\s*(\d*)$')

# w:tcPr 内で w:shd より後ろに置く必要がある要素（OOXML スキーマの順序）
_TCPR_AFTER_SHD = (
    "w:noWrap", "w:tcMar", "w:textDirection", "w:tcFitText", "w:vAlign",
    "w:hideMark", "w:headers", "w:cellIns", "w:cellDel", "w:cellMerge", "w:tcPrChange",
)


def _get_doc(file_path: str, create_if_missing: bool = False):
    return _cache.get(file_path, create_if_missing=create_if_missing)


def _save(file_path: str):
    return _cache.save(file_path, _get_doc(file_path))


def _iter_block_paragraphs(container, seen_cells: set):
    """本文・セル・ヘッダーなどに含まれる段落を、入れ子のテーブルも含めて返す"""
    yield from container.paragraphs
    for table in container.tables:
        for row in table.rows:
            for cell in row.cells:
                # 結合セルは row.cells に同じセルが複数回現れるので1回だけ処理する
                if cell._tc in seen_cells:
                    continue
                seen_cells.add(cell._tc)
                yield from _iter_block_paragraphs(cell, seen_cells)


def _iter_all_paragraphs(doc):
    """本文・テーブル・ヘッダー/フッターの全段落を返す"""
    seen_cells: set = set()
    yield from _iter_block_paragraphs(doc, seen_cells)
    for section in doc.sections:
        for part in (section.header, section.footer,
                     section.first_page_header, section.first_page_footer,
                     section.even_page_header, section.even_page_footer):
            # 前セクションにリンクされたものは同一内容（参照すると空の定義が作られてしまう）
            if not part.is_linked_to_previous:
                yield from _iter_block_paragraphs(part, seen_cells)


def _replace_in_paragraph(para, old: str, new: str, limit: Optional[int]) -> int:
    """段落内の old を new に置換して置換数を返す。

    Word は同じ見た目の文字列でも複数のラン（書式単位）に分割して保存することが多いため、
    ランをまたぐ一致も置換する。置換後の文字列は一致開始位置のランの書式を引き継ぐ。
    """
    runs = para.runs
    count = 0
    search_from = 0
    while limit is None or count < limit:
        texts = [r.text for r in runs]
        start = "".join(texts).find(old, search_from)
        if start < 0:
            break
        end = start + len(old)
        run_start = 0
        for run, text in zip(runs, texts, strict=True):
            run_end = run_start + len(text)
            if text and run_start < end and run_end > start:
                after = text[max(end - run_start, 0):]
                if run_start <= start:
                    run.text = text[:start - run_start] + new + after
                else:
                    run.text = after
            run_start = run_end
        # new に old が含まれていても無限ループしないよう、置換後の位置から探す
        search_from = start + len(new)
        count += 1
    return count


def _para_summary(para, index: int) -> dict:
    """段落の要約情報を返す"""
    return {
        "index": index,
        "style": para.style.name,
        "text": para.text,
        "bold": any(r.bold for r in para.runs if r.bold),
        "italic": any(r.italic for r in para.runs if r.italic),
    }


# ── ツール関数 ──────────────────────────────────────────────────────────────


def open_word(file_path: str, create_if_missing: bool = False) -> dict:
    """Wordファイルを開く（create_if_missing=True なら新規作成して保存する）"""
    try:
        abs_path = safe_path(file_path)
        is_new = not os.path.exists(abs_path)
        doc = _get_doc(file_path, create_if_missing=create_if_missing)
        if is_new:
            _save(file_path)
        para_count = len(doc.paragraphs)
        table_count = len(doc.tables)
        action = "新規作成しました" if is_new else "開きました"
        return {
            "success": True,
            "file_path": abs_path,
            "paragraph_count": para_count,
            "table_count": table_count,
            "message": f"ドキュメントを{action}。段落数: {para_count}, テーブル数: {table_count}"
        }
    except Exception as e:
        return {"success": False, "error": str(e)}


def read_document(file_path: str, include_tables: bool = True) -> dict:
    """ドキュメント全体の内容を読み取る（段落インデックス付き）"""
    try:
        doc = _get_doc(file_path)
        paragraphs = [_para_summary(p, i) for i, p in enumerate(doc.paragraphs)]

        tables = []
        if include_tables:
            for t_idx, table in enumerate(doc.tables):
                rows = []
                for row in table.rows:
                    rows.append([cell.text for cell in row.cells])
                tables.append({"table_index": t_idx, "rows": rows})

        return {
            "success": True,
            "paragraph_count": len(paragraphs),
            "paragraphs": paragraphs,
            "tables": tables,
        }
    except Exception as e:
        return {"success": False, "error": str(e)}


def read_paragraph(file_path: str, index: int) -> dict:
    """特定インデックスの段落を読み取る"""
    try:
        doc = _get_doc(file_path)
        if index < 0 or index >= len(doc.paragraphs):
            return {"success": False, "error": f"インデックス {index} は範囲外です（0〜{len(doc.paragraphs)-1}）"}
        para = doc.paragraphs[index]
        runs = [{"text": r.text, "bold": r.bold, "italic": r.italic,
                  "font_size": r.font.size.pt if r.font.size else None,
                  "font_color": str(r.font.color.rgb) if r.font.color and r.font.color.type and r.font.color.rgb else None}
                for r in para.runs]
        return {
            "success": True,
            "index": index,
            "style": para.style.name,
            "text": para.text,
            "runs": runs,
        }
    except Exception as e:
        return {"success": False, "error": str(e)}


def append_paragraph(file_path: str, text: str,
                     style: Optional[str] = None,
                     bold: bool = False,
                     italic: bool = False,
                     font_size: Optional[int] = None) -> dict:
    """ドキュメントの末尾に段落を追加する"""
    try:
        doc = _get_doc(file_path, create_if_missing=True)
        para = doc.add_paragraph(style=style)
        run = para.add_run(text)
        if bold:
            run.bold = True
        if italic:
            run.italic = True
        if font_size:
            run.font.size = Pt(font_size)
        _save(file_path)
        new_index = len(doc.paragraphs) - 1
        return {
            "success": True,
            "message": f"段落を末尾（インデックス {new_index}）に追加しました",
            "index": new_index,
        }
    except Exception as e:
        return {"success": False, "error": str(e)}


def append_rich_paragraph(file_path: str,
                           runs: list[dict],
                           style: Optional[str] = None,
                           font_size: Optional[int] = None) -> dict:
    """
    書式付きランのリストから段落を末尾に追加する。
    runs の各要素: {"text": str, "bold": bool, "italic": bool, "color": str}
    color は FF0000 のような6桁HEXまたは #FF0000 形式。
    parse_inline_formatting / parse_markdown の runs フィールドをそのまま渡せる。
    """
    try:
        doc = _get_doc(file_path, create_if_missing=True)
        para = doc.add_paragraph(style=style)
        for r in runs:
            run = para.add_run(r.get("text", ""))
            if r.get("bold"):
                run.bold = True
            if r.get("italic"):
                run.italic = True
            run_size = r.get("font_size") or font_size
            if run_size:
                run.font.size = Pt(run_size)
            if r.get("color"):
                hex_color = validate_hex_color(r["color"], "color")
                run.font.color.rgb = RGBColor(
                    int(hex_color[0:2], 16),
                    int(hex_color[2:4], 16),
                    int(hex_color[4:6], 16),
                )
        _save(file_path)
        new_index = len(doc.paragraphs) - 1
        return {
            "success": True,
            "message": f"書式付き段落をインデックス {new_index} に追加しました",
            "index": new_index,
        }
    except Exception as e:
        return {"success": False, "error": str(e)}


def insert_paragraph(file_path: str, index: int, text: str,
                     position: str = "before",
                     style: Optional[str] = None,
                     bold: bool = False,
                     italic: bool = False,
                     font_size: Optional[int] = None) -> dict:
    """指定インデックスの段落の前後にテキストを挿入する
    position: 'before' または 'after'
    """
    try:
        doc = _get_doc(file_path, create_if_missing=True)
        paras = doc.paragraphs

        if index < 0 or index >= len(paras):
            return {"success": False, "error": f"インデックス {index} は範囲外です（0〜{len(paras)-1}）"}

        ref_para = paras[index]

        # 新しい段落要素を作成
        new_para_elem = OxmlElement("w:p")
        if position == "after":
            ref_para._p.addnext(new_para_elem)
        else:
            ref_para._p.addprevious(new_para_elem)

        # 挿入された段落を探す（XMLから直接特定）
        inserted_index = index if position == "before" else index + 1
        target_para = doc.paragraphs[inserted_index]

        # スタイル設定
        if style:
            target_para.style = doc.styles[style]

        # テキストとフォーマットを設定
        run = target_para.add_run(text)
        if bold:
            run.bold = True
        if italic:
            run.italic = True
        if font_size:
            run.font.size = Pt(font_size)

        _save(file_path)
        return {
            "success": True,
            "message": f"インデックス {index} の{('前' if position == 'before' else '後')}に段落を挿入しました（新インデックス: {inserted_index}）",
            "inserted_index": inserted_index,
        }
    except Exception as e:
        return {"success": False, "error": str(e)}


def replace_text(file_path: str, old_text: str, new_text: str,
                 all_occurrences: bool = True) -> dict:
    """ドキュメント内のテキストを検索・置換する（本文・テーブル・ヘッダー/フッター）"""
    try:
        if not old_text:
            return {"success": False, "error": "old_text が空です"}
        doc = _get_doc(file_path)
        limit = None if all_occurrences else 1
        count = 0

        for para in _iter_all_paragraphs(doc):
            count += _replace_in_paragraph(
                para, old_text, new_text, None if limit is None else limit - count)
            if limit is not None and count >= limit:
                break

        if count:
            _save(file_path)
        return {
            "success": True,
            "replaced_count": count,
            "message": f"「{old_text}」→「{new_text}」に {count} 箇所置換しました"
        }
    except Exception as e:
        return {"success": False, "error": str(e)}


def delete_paragraph(file_path: str, index: int) -> dict:
    """指定インデックスの段落を削除する"""
    try:
        doc = _get_doc(file_path)
        if index < 0 or index >= len(doc.paragraphs):
            return {"success": False, "error": f"インデックス {index} は範囲外です（0〜{len(doc.paragraphs)-1}）"}

        para = doc.paragraphs[index]
        text_preview = para.text[:30]
        para._p.getparent().remove(para._p)

        _save(file_path)
        return {
            "success": True,
            "message": f"インデックス {index}「{text_preview}」を削除しました",
            "remaining_paragraphs": len(doc.paragraphs),
        }
    except Exception as e:
        return {"success": False, "error": str(e)}


def format_paragraph(file_path: str, index: int,
                     bold: Optional[bool] = None,
                     italic: Optional[bool] = None,
                     font_size: Optional[int] = None,
                     font_color: Optional[str] = None,
                     alignment: Optional[str] = None,
                     style: Optional[str] = None) -> dict:
    """段落全体の書式を変更する（alignment: left/center/right/justify）"""
    try:
        doc = _get_doc(file_path)
        if index < 0 or index >= len(doc.paragraphs):
            return {"success": False, "error": f"インデックス {index} は範囲外です"}

        para = doc.paragraphs[index]

        if style:
            para.style = doc.styles[style]

        align_map = {
            "left": WD_ALIGN_PARAGRAPH.LEFT,
            "center": WD_ALIGN_PARAGRAPH.CENTER,
            "right": WD_ALIGN_PARAGRAPH.RIGHT,
            "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
        }
        if alignment and alignment.lower() in align_map:
            para.alignment = align_map[alignment.lower()]

        for run in para.runs:
            if bold is not None:
                run.bold = bold
            if italic is not None:
                run.italic = italic
            if font_size is not None:
                run.font.size = Pt(font_size)
            if font_color is not None:
                hex_color = validate_hex_color(font_color, "font_color")
                r = int(hex_color[0:2], 16)
                g = int(hex_color[2:4], 16)
                b = int(hex_color[4:6], 16)
                run.font.color.rgb = RGBColor(r, g, b)

        _save(file_path)
        return {"success": True, "message": f"インデックス {index} の書式を更新しました"}
    except Exception as e:
        return {"success": False, "error": str(e)}


def insert_image(file_path: str, image_path: str,
                 paragraph_index: Optional[int] = None,
                 width_cm: Optional[float] = None) -> dict:
    """画像をドキュメントに挿入する（paragraph_index省略時は末尾）"""
    try:
        image_path = safe_path(image_path)
        if not os.path.exists(image_path):
            return {"success": False, "error": f"画像ファイルが見つかりません: {image_path}"}

        doc = _get_doc(file_path, create_if_missing=True)
        width = Cm(width_cm) if width_cm else None

        if paragraph_index is None:
            # 末尾に追加
            para = doc.add_paragraph()
            run = para.add_run()
            run.add_picture(image_path, width=width)
            inserted_index = len(doc.paragraphs) - 1
        else:
            if paragraph_index < 0 or paragraph_index >= len(doc.paragraphs):
                return {"success": False, "error": f"インデックス {paragraph_index} は範囲外です"}
            # 指定段落の後に画像段落を挿入
            ref_para = doc.paragraphs[paragraph_index]
            new_para_elem = OxmlElement("w:p")
            ref_para._p.addnext(new_para_elem)
            inserted_index = paragraph_index + 1
            target_para = doc.paragraphs[inserted_index]
            run = target_para.add_run()
            run.add_picture(image_path, width=width)

        _save(file_path)
        return {
            "success": True,
            "message": f"画像「{os.path.basename(image_path)}」をインデックス {inserted_index} に挿入しました",
            "inserted_index": inserted_index,
        }
    except Exception as e:
        return {"success": False, "error": str(e)}


def add_table(file_path: str, data: list[list],
              paragraph_index: Optional[int] = None,
              has_header: bool = True) -> dict:
    """テーブルを追加する（data: 2次元配列、先頭行をヘッダーとして太字に）"""
    try:
        if not isinstance(data, list) or not data:
            return {"success": False, "error": "data には空でない2次元配列を指定してください"}
        data = [r if isinstance(r, list) else [r] for r in data]

        doc = _get_doc(file_path, create_if_missing=True)
        if paragraph_index is not None and not 0 <= paragraph_index < len(doc.paragraphs):
            return {"success": False, "error": f"インデックス {paragraph_index} は範囲外です"}

        rows = len(data)
        cols = max(max(len(row) for row in data), 1)

        # テーブルを文書末尾に追加
        table = doc.add_table(rows=rows, cols=cols)
        try:
            table.style = "Table Grid"
        except KeyError:
            # Word で作成した文書には "Table Grid" スタイルが定義されていないことがある
            pass

        for r_idx, row_data in enumerate(data):
            for c_idx, cell_val in enumerate(row_data):
                cell = table.rows[r_idx].cells[c_idx]
                cell.text = str(cell_val) if cell_val is not None else ""
                if has_header and r_idx == 0:
                    for run in cell.paragraphs[0].runs:
                        run.bold = True

        # 指定段落の後ろに移動
        if paragraph_index is not None:
            ref_para = doc.paragraphs[paragraph_index]
            tbl_elem = table._tbl
            tbl_elem.getparent().remove(tbl_elem)
            ref_para._p.addnext(tbl_elem)

        _save(file_path)
        return {
            "success": True,
            "message": f"{rows}行×{cols}列のテーブルを追加しました",
        }
    except Exception as e:
        return {"success": False, "error": str(e)}


def add_heading(file_path: str, text: str, level: int = 1) -> dict:
    """見出しを末尾に追加する（level: 1〜9）"""
    try:
        doc = _get_doc(file_path, create_if_missing=True)
        doc.add_heading(text, level=level)
        _save(file_path)
        idx = len(doc.paragraphs) - 1
        return {
            "success": True,
            "message": f"見出し(H{level})「{text}」をインデックス {idx} に追加しました",
            "index": idx,
        }
    except Exception as e:
        return {"success": False, "error": str(e)}


def save_word(file_path: str, save_as: Optional[str] = None) -> dict:
    """Wordファイルを保存する（save_as を指定すると別名保存）"""
    try:
        doc = _get_doc(file_path)
        if save_as:
            # 別名保存先も作業ディレクトリ内に限定する。
            # 同じオブジェクトを2つのパスで共有しないよう、次回はディスクから読み直す
            target = safe_path(save_as)
            save_file(doc, target)
            _cache.evict(save_as)
        else:
            target = _save(file_path)
        return {"success": True, "message": f"'{target}' に保存しました"}
    except Exception as e:
        return {"success": False, "error": str(e)}


def get_document_info(file_path: str) -> dict:
    """ドキュメントの構造情報（見出し・段落数・テーブル数）を返す"""
    try:
        doc = _get_doc(file_path)
        headings = []
        for i, p in enumerate(doc.paragraphs):
            # "Heading 1" 形式以外（カスタムの "Heading Custom" など）でも落ちないようにする
            m = _HEADING_STYLE_RE.match(p.style.name if p.style is not None else "")
            if m:
                headings.append({"index": i, "level": int(m.group(1) or 0), "text": p.text})
        return {
            "success": True,
            "paragraph_count": len(doc.paragraphs),
            "table_count": len(doc.tables),
            "headings": headings,
        }
    except Exception as e:
        return {"success": False, "error": str(e)}


def set_page_layout(file_path: str,
                    orientation: Optional[str] = None,
                    top_cm: Optional[float] = None,
                    bottom_cm: Optional[float] = None,
                    left_cm: Optional[float] = None,
                    right_cm: Optional[float] = None) -> dict:
    """ページの向き・余白を設定する（orientation: portrait / landscape）"""
    try:
        from docx.enum.section import WD_ORIENT
        doc = _get_doc(file_path, create_if_missing=True)
        orient = None
        if orientation is not None:
            if orientation.lower() in ("landscape", "横"):
                orient = WD_ORIENT.LANDSCAPE
            elif orientation.lower() in ("portrait", "縦"):
                orient = WD_ORIENT.PORTRAIT
            else:
                return {"success": False, "error": "orientation は portrait / landscape のいずれかを指定してください"}

        # セクション区切りがある文書でも全ページに適用する
        for section in doc.sections:
            if orient is not None and section.orientation != orient:
                section.orientation = orient
                section.page_width, section.page_height = section.page_height, section.page_width

            if top_cm    is not None: section.top_margin    = Cm(top_cm)
            if bottom_cm is not None: section.bottom_margin = Cm(bottom_cm)
            if left_cm   is not None: section.left_margin   = Cm(left_cm)
            if right_cm  is not None: section.right_margin  = Cm(right_cm)

        _save(file_path)
        return {"success": True, "message": "ページレイアウトを設定しました"}
    except Exception as e:
        return {"success": False, "error": str(e)}


def add_page_break(file_path: str,
                   paragraph_index: Optional[int] = None) -> dict:
    """改ページを挿入する（paragraph_index 省略時は末尾）"""
    try:
        from docx.enum.text import WD_BREAK
        doc = _get_doc(file_path, create_if_missing=True)

        if paragraph_index is None:
            para = doc.add_paragraph()
            para.add_run().add_break(WD_BREAK.PAGE)
            inserted_index = len(doc.paragraphs) - 1
        else:
            if paragraph_index < 0 or paragraph_index >= len(doc.paragraphs):
                return {"success": False, "error": f"インデックス {paragraph_index} は範囲外です"}
            ref_para = doc.paragraphs[paragraph_index]
            new_para_elem = OxmlElement("w:p")
            ref_para._p.addnext(new_para_elem)
            inserted_index = paragraph_index + 1
            doc.paragraphs[inserted_index].add_run().add_break(WD_BREAK.PAGE)

        _save(file_path)
        return {"success": True, "message": f"改ページをインデックス {inserted_index} に挿入しました",
                "inserted_index": inserted_index}
    except Exception as e:
        return {"success": False, "error": str(e)}


def read_table(file_path: str, table_index: int) -> dict:
    """既存テーブルの内容を詳細に読み取る"""
    try:
        doc = _get_doc(file_path)
        if table_index < 0 or table_index >= len(doc.tables):
            return {"success": False,
                    "error": f"テーブルインデックス {table_index} は範囲外です（0〜{len(doc.tables)-1}）"}

        table = doc.tables[table_index]
        rows_data = [
            [{"row": r_idx, "col": c_idx, "text": cell.text}
             for c_idx, cell in enumerate(row.cells)]
            for r_idx, row in enumerate(table.rows)
        ]
        return {
            "success": True,
            "table_index": table_index,
            "row_count": len(table.rows),
            "col_count": len(table.columns),
            "rows": rows_data,
        }
    except Exception as e:
        return {"success": False, "error": str(e)}


def format_table(file_path: str, table_index: int,
                 row: int, col: int,
                 bold: Optional[bool] = None,
                 italic: Optional[bool] = None,
                 font_size: Optional[int] = None,
                 font_color: Optional[str] = None,
                 bg_color: Optional[str] = None,
                 alignment: Optional[str] = None) -> dict:
    """テーブルの特定セルに書式を適用する"""
    try:
        doc = _get_doc(file_path)
        if table_index < 0 or table_index >= len(doc.tables):
            return {"success": False, "error": f"テーブルインデックス {table_index} は範囲外です"}

        table = doc.tables[table_index]
        if row < 0 or row >= len(table.rows):
            return {"success": False, "error": f"行インデックス {row} は範囲外です（0〜{len(table.rows)-1}）"}
        if col < 0 or col >= len(table.columns):
            return {"success": False, "error": f"列インデックス {col} は範囲外です（0〜{len(table.columns)-1}）"}

        cell = table.rows[row].cells[col]
        align_map = {
            "left": WD_ALIGN_PARAGRAPH.LEFT, "center": WD_ALIGN_PARAGRAPH.CENTER,
            "right": WD_ALIGN_PARAGRAPH.RIGHT, "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
        }

        for para in cell.paragraphs:
            if alignment and alignment.lower() in align_map:
                para.alignment = align_map[alignment.lower()]
            for run in para.runs:
                if bold      is not None: run.bold   = bold
                if italic    is not None: run.italic = italic
                if font_size is not None: run.font.size = Pt(font_size)
                if font_color is not None:
                    hc = validate_hex_color(font_color, "font_color")
                    run.font.color.rgb = RGBColor(int(hc[0:2], 16), int(hc[2:4], 16), int(hc[4:6], 16))

        if bg_color is not None:
            hc = validate_hex_color(bg_color, "bg_color")
            tcPr = cell._tc.get_or_add_tcPr()
            # 既存の網掛けを置き換える（w:shd が重複すると Word が修復ダイアログを出す）
            for old_shd in tcPr.findall(qn("w:shd")):
                tcPr.remove(old_shd)
            shd = OxmlElement("w:shd")
            shd.set(qn("w:val"), "clear")
            shd.set(qn("w:color"), "auto")
            shd.set(qn("w:fill"), hc)
            # スキーマ上の順序を守って挿入する
            tcPr.insert_element_before(shd, *_TCPR_AFTER_SHD)

        _save(file_path)
        return {"success": True, "message": f"テーブル {table_index} の [{row}][{col}] に書式を適用しました"}
    except Exception as e:
        return {"success": False, "error": str(e)}


def add_header_footer(file_path: str,
                      header_text: Optional[str] = None,
                      footer_text: Optional[str] = None) -> dict:
    """ヘッダー・フッターを設定する"""
    try:
        doc = _get_doc(file_path, create_if_missing=True)
        section = doc.sections[0]

        if header_text is not None:
            header = section.header
            para = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
            para.text = header_text

        if footer_text is not None:
            footer = section.footer
            para = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
            para.text = footer_text

        _save(file_path)
        parts = []
        if header_text is not None: parts.append(f"ヘッダー「{header_text}」")
        if footer_text is not None: parts.append(f"フッター「{footer_text}」")
        return {"success": True, "message": f"{' / '.join(parts)} を設定しました"}
    except Exception as e:
        return {"success": False, "error": str(e)}


# ── ツール定義（Claude API用）──────────────────────────────────────────────

TOOLS = [
    {
        "name": "open_word",
        "description": "Open a Word document. If the file does not exist, set create_if_missing=true to create it.",
        "input_schema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "Path to the Word file (.docx)"},
                "create_if_missing": {"type": "boolean", "description": "Create the file if it does not exist"}
            },
            "required": ["file_path"]
        }
    },
    {
        "name": "read_document",
        "description": "Read the full content of a Word document with paragraph indices.",
        "input_schema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "Path to the Word file"},
                "include_tables": {"type": "boolean", "description": "Include table contents (default: true)"}
            },
            "required": ["file_path"]
        }
    },
    {
        "name": "read_paragraph",
        "description": "Read the text and formatting of a specific paragraph by index.",
        "input_schema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "Path to the Word file"},
                "index": {"type": "integer", "description": "Paragraph index (0-based)"}
            },
            "required": ["file_path", "index"]
        }
    },
    {
        "name": "append_paragraph",
        "description": "Append a new paragraph at the end of the document.",
        "input_schema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "Path to the Word file"},
                "text": {"type": "string", "description": "Text content to append"},
                "style": {"type": "string", "description": "Paragraph style, e.g. 'Normal' or 'Heading 1'"},
                "bold": {"type": "boolean", "description": "Set bold"},
                "italic": {"type": "boolean", "description": "Set italic"},
                "font_size": {"type": "integer", "description": "Font size in pt"}
            },
            "required": ["file_path", "text"]
        }
    },
    {
        "name": "insert_paragraph",
        "description": "Insert a paragraph before or after a specified paragraph index.",
        "input_schema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "Path to the Word file"},
                "index": {"type": "integer", "description": "Reference paragraph index"},
                "text": {"type": "string", "description": "Text to insert"},
                "position": {"type": "string", "enum": ["before", "after"], "description": "Insert before or after the reference paragraph (default: before)"},
                "style": {"type": "string", "description": "Paragraph style"},
                "bold": {"type": "boolean", "description": "Set bold"},
                "italic": {"type": "boolean", "description": "Set italic"},
                "font_size": {"type": "integer", "description": "Font size in pt"}
            },
            "required": ["file_path", "index", "text"]
        }
    },
    {
        "name": "replace_text",
        "description": (
            "Find and replace text in the document body, tables, and headers/footers. "
            "Matches that span multiple formatting runs are also replaced. "
            "Useful for editing and proofreading."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "Path to the Word file"},
                "old_text": {"type": "string", "description": "Text to search for"},
                "new_text": {"type": "string", "description": "Replacement text"},
                "all_occurrences": {"type": "boolean", "description": "Replace all occurrences (default: true)"}
            },
            "required": ["file_path", "old_text", "new_text"]
        }
    },
    {
        "name": "delete_paragraph",
        "description": "Delete the paragraph at the specified index.",
        "input_schema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "Path to the Word file"},
                "index": {"type": "integer", "description": "Index of the paragraph to delete"}
            },
            "required": ["file_path", "index"]
        }
    },
    {
        "name": "format_paragraph",
        "description": "Apply formatting to a paragraph: bold, italic, font size, color, alignment, or style.",
        "input_schema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "Path to the Word file"},
                "index": {"type": "integer", "description": "Target paragraph index"},
                "bold": {"type": "boolean", "description": "Set bold"},
                "italic": {"type": "boolean", "description": "Set italic"},
                "font_size": {"type": "integer", "description": "Font size in pt"},
                "font_color": {"type": "string", "description": "Font color as hex string, e.g. FF0000 for red"},
                "alignment": {"type": "string", "enum": ["left", "center", "right", "justify"], "description": "Text alignment"},
                "style": {"type": "string", "description": "Paragraph style name, e.g. 'Heading 1' or 'Normal'"}
            },
            "required": ["file_path", "index"]
        }
    },
    {
        "name": "insert_image",
        "description": "Insert an image file into the document.",
        "input_schema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "Path to the Word file"},
                "image_path": {"type": "string", "description": "Path to the image file (PNG, JPG, etc.)"},
                "paragraph_index": {"type": "integer", "description": "Insert after this paragraph index (default: append at end)"},
                "width_cm": {"type": "number", "description": "Image width in centimeters (default: original size)"}
            },
            "required": ["file_path", "image_path"]
        }
    },
    {
        "name": "add_table",
        "description": "Add a table to the document from a 2D data array.",
        "input_schema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "Path to the Word file"},
                "data": {
                    "type": "array",
                    "items": {"type": "array"},
                    "description": "2D array of table data (list of rows)"
                },
                "paragraph_index": {"type": "integer", "description": "Insert after this paragraph index (default: append at end)"},
                "has_header": {"type": "boolean", "description": "Bold the first row as a header (default: true)"}
            },
            "required": ["file_path", "data"]
        }
    },
    {
        "name": "add_heading",
        "description": "Append a heading to the end of the document.",
        "input_schema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "Path to the Word file"},
                "text": {"type": "string", "description": "Heading text"},
                "level": {"type": "integer", "description": "Heading level 1-9 (default: 1)"}
            },
            "required": ["file_path", "text"]
        }
    },
    {
        "name": "save_word",
        "description": "Save the Word document. Optionally save to a new path.",
        "input_schema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "Path to the Word file"},
                "save_as": {"type": "string", "description": "New path to save a copy (omit to overwrite original)"}
            },
            "required": ["file_path"]
        }
    },
    {
        "name": "get_document_info",
        "description": "Get structural info about a document: heading list, paragraph count, and table count.",
        "input_schema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "Path to the Word file"}
            },
            "required": ["file_path"]
        }
    },
    {
        "name": "set_page_layout",
        "description": "Set page orientation (portrait/landscape) and margins.",
        "input_schema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "Path to the Word file"},
                "orientation": {"type": "string", "enum": ["portrait", "landscape"], "description": "Page orientation"},
                "top_cm":    {"type": "number", "description": "Top margin in cm"},
                "bottom_cm": {"type": "number", "description": "Bottom margin in cm"},
                "left_cm":   {"type": "number", "description": "Left margin in cm"},
                "right_cm":  {"type": "number", "description": "Right margin in cm"}
            },
            "required": ["file_path"]
        }
    },
    {
        "name": "add_page_break",
        "description": "Insert a page break after a paragraph index, or at the end if omitted.",
        "input_schema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "Path to the Word file"},
                "paragraph_index": {"type": "integer", "description": "Insert after this paragraph index (default: append at end)"}
            },
            "required": ["file_path"]
        }
    },
    {
        "name": "read_table",
        "description": "Read the contents of an existing table by index.",
        "input_schema": {
            "type": "object",
            "properties": {
                "file_path":    {"type": "string",  "description": "Path to the Word file"},
                "table_index":  {"type": "integer", "description": "Table index (0-based)"}
            },
            "required": ["file_path", "table_index"]
        }
    },
    {
        "name": "format_table",
        "description": "Apply formatting to a specific cell in an existing table.",
        "input_schema": {
            "type": "object",
            "properties": {
                "file_path":    {"type": "string",  "description": "Path to the Word file"},
                "table_index":  {"type": "integer", "description": "Table index (0-based)"},
                "row":          {"type": "integer", "description": "Row index (0-based)"},
                "col":          {"type": "integer", "description": "Column index (0-based)"},
                "bold":         {"type": "boolean"},
                "italic":       {"type": "boolean"},
                "font_size":    {"type": "integer", "description": "Font size in pt"},
                "font_color":   {"type": "string",  "description": "Font color as hex, e.g. FF0000"},
                "bg_color":     {"type": "string",  "description": "Background color as hex, e.g. FFFF00"},
                "alignment":    {"type": "string",  "enum": ["left", "center", "right", "justify"]}
            },
            "required": ["file_path", "table_index", "row", "col"]
        }
    },
    {
        "name": "add_header_footer",
        "description": "Set the header and/or footer text for the document.",
        "input_schema": {
            "type": "object",
            "properties": {
                "file_path":    {"type": "string", "description": "Path to the Word file"},
                "header_text":  {"type": "string", "description": "Header text (omit to leave unchanged)"},
                "footer_text":  {"type": "string", "description": "Footer text (omit to leave unchanged)"}
            },
            "required": ["file_path"]
        }
    },
    {
        "name": "append_rich_paragraph",
        "description": (
            "Append a paragraph built from a list of runs with individual bold/italic formatting. "
            "Use the 'runs' output from parse_markdown or parse_inline_formatting to preserve "
            "Markdown bold (**text**) and italic (*text*) as real Word formatting."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "Path to the Word file"},
                "runs": {
                    "type": "array",
                    "description": "List of run objects: [{\"text\": str, \"bold\": bool, \"italic\": bool, \"color\": \"FF0000\"}]",
                    "items": {
                        "type": "object",
                        "properties": {
                            "text":   {"type": "string"},
                            "bold":   {"type": "boolean"},
                            "italic": {"type": "boolean"},
                            "color":     {"type": "string",  "description": "Font color as 6-digit hex, e.g. FF0000"},
                            "font_size": {"type": "integer", "description": "Font size in pt for this run only"},
                        },
                        "required": ["text"],
                    },
                },
                "style":     {"type": "string",  "description": "Paragraph style, e.g. 'Normal'"},
                "font_size": {"type": "integer", "description": "Font size in pt applied to all runs"},
            },
            "required": ["file_path", "runs"]
        }
    },
]

TOOL_FUNCTIONS = {
    "open_word": open_word,
    "read_document": read_document,
    "read_paragraph": read_paragraph,
    "append_paragraph": append_paragraph,
    "append_rich_paragraph": append_rich_paragraph,
    "insert_paragraph": insert_paragraph,
    "replace_text": replace_text,
    "delete_paragraph": delete_paragraph,
    "format_paragraph": format_paragraph,
    "insert_image": insert_image,
    "add_table": add_table,
    "add_heading": add_heading,
    "save_word": save_word,
    "get_document_info": get_document_info,
    "set_page_layout": set_page_layout,
    "add_page_break": add_page_break,
    "read_table": read_table,
    "format_table": format_table,
    "add_header_footer": add_header_footer,
}


def execute_tool(tool_name: str, tool_input: dict) -> str:
    return run_tool(TOOL_FUNCTIONS, tool_name, tool_input, cache=_cache)
