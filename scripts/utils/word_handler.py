"""Word テンプレート処理の共通ユーティリティ"""

import sys
from pathlib import Path

from docx import Document
from docx.shared import Pt


def load_template(template_path: str) -> Document:
    """テンプレートWordファイルを読み込む"""
    path = Path(template_path)
    if not path.exists():
        print(f"エラー: テンプレートファイルが見つかりません: {template_path}", file=sys.stderr)
        sys.exit(1)
    return Document(template_path)


def replace_placeholder(doc: Document, placeholder: str, value: str):
    """文書内のプレースホルダーを値で置換"""
    for paragraph in doc.paragraphs:
        if placeholder in paragraph.text:
            for run in paragraph.runs:
                if placeholder in run.text:
                    run.text = run.text.replace(placeholder, value)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if placeholder in cell.text:
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            if placeholder in run.text:
                                run.text = run.text.replace(placeholder, value)


def replace_placeholders(doc: Document, mappings: dict[str, str]):
    """複数のプレースホルダーを一括置換"""
    for placeholder, value in mappings.items():
        if value is not None and value != "":
            replace_placeholder(doc, placeholder, value)


def fill_table_cell(doc: Document, table_index: int, row: int, col: int, value: str):
    """指定テーブルの指定セルに値を書き込む"""
    if table_index >= len(doc.tables):
        print(f"エラー: テーブルインデックス {table_index} が範囲外です", file=sys.stderr)
        return
    table = doc.tables[table_index]
    if row >= len(table.rows) or col >= len(table.rows[0].cells):
        print(f"エラー: セル({row}, {col})が範囲外です", file=sys.stderr)
        return
    table.rows[row].cells[col].text = value


def save_output(doc: Document, output_path: str):
    """ファイルを保存"""
    path = Path(output_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(output_path)
    print(f"保存完了: {output_path}")


def get_document_structure(doc: Document) -> dict:
    """文書の構造情報を取得"""
    structure = {
        "paragraphs": len(doc.paragraphs),
        "tables": len(doc.tables),
        "headings": [],
        "placeholders": [],
        "table_sizes": [],
    }

    for para in doc.paragraphs:
        if para.style.name.startswith("Heading"):
            structure["headings"].append({
                "level": para.style.name,
                "text": para.text,
            })
        if "{{" in para.text and "}}" in para.text:
            structure["placeholders"].append(para.text)

    for i, table in enumerate(doc.tables):
        structure["table_sizes"].append({
            "index": i,
            "rows": len(table.rows),
            "cols": len(table.rows[0].cells) if table.rows else 0,
        })

    return structure


def list_placeholders(doc: Document) -> list[str]:
    """文書内の全プレースホルダー（{{...}}形式）を取得"""
    import re
    placeholders = set()
    pattern = re.compile(r"\{\{(.+?)\}\}")

    for para in doc.paragraphs:
        for match in pattern.finditer(para.text):
            placeholders.add(match.group(0))

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for match in pattern.finditer(cell.text):
                    placeholders.add(match.group(0))

    return sorted(placeholders)
