"""新事業進出補助金 テンプレート書き込みスクリプト

使い方:
  python grants/新事業進出補助金/scripts/fill_template.py \
    fill --template grants/新事業進出補助金/templates/xxx.xlsx \
         --data workspace/{案件名}/application_data.json \
         --output workspace/{案件名}/申請書_新事業進出補助金.xlsx

  python grants/新事業進出補助金/scripts/fill_template.py \
    analyze --template grants/新事業進出補助金/templates/xxx.xlsx
"""

import argparse
import json
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(PROJECT_ROOT / "scripts"))

from utils.excel_handler import load_template, fill_cells, save_output, get_template_info
from utils.word_handler import (
    load_template as load_word_template,
    replace_placeholders,
    save_output as save_word_output,
    get_document_structure,
    list_placeholders,
)

# ===== セルマッピング定義 =====
# ※ 実際のテンプレートをアップロード後に更新すること

EXCEL_CELL_MAPPING = {
    # "事業計画": {
    #     "B3": "company_name",
    #     "B5": "current_business",
    #     "B10": "new_business_necessity",
    #     "B15": "new_business_detail",
    #     "B20": "synergy",
    # },
    # "収益計画": {
    #     "C3": "sales_existing_year0",
    #     "C4": "sales_new_year0",
    #     "C5": "sales_total_year0",
    #     "C8": "added_value_year0",
    # },
}

WORD_PLACEHOLDER_MAPPING = {
    # "{{会社名}}": "company_name",
    # "{{現在の事業}}": "current_business",
    # "{{新事業の必要性}}": "new_business_necessity",
    # "{{新事業の内容}}": "new_business_detail",
    # "{{シナジー}}": "synergy",
}


def load_data(data_path: str) -> dict:
    path = Path(data_path)
    if not path.exists():
        print(f"エラー: データファイルが見つかりません: {data_path}", file=sys.stderr)
        sys.exit(1)
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def fill_excel(template_path: str, data: dict, output_path: str):
    wb = load_template(template_path)
    for sheet_name, cell_map in EXCEL_CELL_MAPPING.items():
        if sheet_name not in wb.sheetnames:
            print(f"警告: シート '{sheet_name}' が見つかりません。スキップします。")
            continue
        ws = wb[sheet_name]
        resolved = {ref: str(data.get(key, "")) for ref, key in cell_map.items() if data.get(key)}
        fill_cells(ws, resolved)
    save_output(wb, output_path)


def fill_word(template_path: str, data: dict, output_path: str):
    doc = load_word_template(template_path)
    resolved = {ph: str(data.get(key, "")) for ph, key in WORD_PLACEHOLDER_MAPPING.items() if data.get(key)}
    replace_placeholders(doc, resolved)
    save_word_output(doc, output_path)


def analyze_template(template_path: str):
    path = Path(template_path)
    if path.suffix in (".xlsx", ".xls"):
        info = get_template_info(template_path)
        print("=== Excelテンプレート構造 ===")
        for sheet_name, sheet_info in info.items():
            print(f"\nシート: {sheet_name}")
            print(f"  範囲: {sheet_info['dimensions']}")
    elif path.suffix == ".docx":
        doc = load_word_template(template_path)
        structure = get_document_structure(doc)
        placeholders = list_placeholders(doc)
        print("=== Wordテンプレート構造 ===")
        print(f"段落数: {structure['paragraphs']}, テーブル数: {structure['tables']}")
        if placeholders:
            print(f"プレースホルダー: {placeholders}")


def main():
    parser = argparse.ArgumentParser(description="新事業進出補助金テンプレート処理")
    subparsers = parser.add_subparsers(dest="command")

    fill_p = subparsers.add_parser("fill")
    fill_p.add_argument("--template", required=True)
    fill_p.add_argument("--data", required=True)
    fill_p.add_argument("--output", required=True)

    analyze_p = subparsers.add_parser("analyze")
    analyze_p.add_argument("--template", required=True)

    args = parser.parse_args()

    if args.command == "fill":
        data = load_data(args.data)
        path = Path(args.template)
        if path.suffix in (".xlsx", ".xls"):
            fill_excel(args.template, data, args.output)
        elif path.suffix == ".docx":
            fill_word(args.template, data, args.output)
    elif args.command == "analyze":
        analyze_template(args.template)
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
