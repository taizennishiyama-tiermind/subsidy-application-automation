"""Excel テンプレート処理の共通ユーティリティ"""

import sys
from pathlib import Path
from copy import copy

from openpyxl import load_workbook
from openpyxl.cell.cell import Cell


def load_template(template_path: str):
    """テンプレートExcelファイルを読み込む"""
    path = Path(template_path)
    if not path.exists():
        print(f"エラー: テンプレートファイルが見つかりません: {template_path}", file=sys.stderr)
        sys.exit(1)
    return load_workbook(template_path)


def fill_cell(ws, cell_ref: str, value: str, preserve_format: bool = True):
    """指定セルに値を書き込む（書式を保持）"""
    cell = ws[cell_ref]
    cell.value = value


def fill_cells(ws, mappings: dict[str, str]):
    """複数セルに一括書き込み"""
    for cell_ref, value in mappings.items():
        if value is not None and value != "":
            fill_cell(ws, cell_ref, value)


def save_output(wb, output_path: str):
    """ファイルを保存"""
    path = Path(output_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    print(f"保存完了: {output_path}")


def read_cell(ws, cell_ref: str) -> str:
    """セルの値を読み取る"""
    value = ws[cell_ref].value
    return str(value) if value is not None else ""


def read_cells(ws, cell_refs: list[str]) -> dict[str, str]:
    """複数セルの値を一括読み取り"""
    return {ref: read_cell(ws, ref) for ref in cell_refs}


def list_sheets(wb) -> list[str]:
    """シート名一覧を取得"""
    return wb.sheetnames


def get_template_info(template_path: str) -> dict:
    """テンプレートの構造情報を取得（シート名、使用範囲等）"""
    wb = load_template(template_path)
    info = {}
    for name in wb.sheetnames:
        ws = wb[name]
        info[name] = {
            "min_row": ws.min_row,
            "max_row": ws.max_row,
            "min_col": ws.min_column,
            "max_col": ws.max_column,
            "dimensions": ws.dimensions,
        }
    return info
