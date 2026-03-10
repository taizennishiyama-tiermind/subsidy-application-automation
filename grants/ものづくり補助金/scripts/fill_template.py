"""ものづくり補助金 テンプレート書き込みスクリプト

テンプレートExcel/Wordファイルにヒアリング結果・分析データを書き込む。

使い方:
  python grants/ものづくり補助金/scripts/fill_template.py \
    --template grants/ものづくり補助金/templates/事業計画書_テンプレート.xlsx \
    --data workspace/{案件名}/company_data.md \
    --output workspace/{案件名}/申請書_ものづくり補助金.xlsx

※ セルマッピングはテンプレートの構造に応じてカスタマイズすること
"""

import argparse
import json
import sys
from pathlib import Path

# プロジェクトルートをパスに追加
PROJECT_ROOT = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(PROJECT_ROOT / "scripts"))

from utils.excel_handler import load_template, fill_cells, save_output, get_template_info
from utils.word_handler import (
    load_template as load_word_template,
    replace_placeholders,
    save_output as save_word_output,
)


# ===== セルマッピング定義 =====
# テンプレートのセル位置 → データキーの対応表
# ※ 実際のテンプレートをアップロード後に更新すること

EXCEL_CELL_MAPPING = {
    # シート名: { セル参照: データキー }
    "事業計画": {
        # 例: "B3": "company_name",
        # 例: "B4": "representative",
        # 例: "B5": "established_year",
        # 例: "B6": "employees",
        # 例: "B7": "capital",
        # 例: "B10": "business_overview",
        # 例: "B15": "challenge_1",
        # 例: "B20": "solution_overview",
    },
    "数値計画": {
        # 例: "C3": "sales_year0",
        # 例: "D3": "sales_year1",
        # 例: "C5": "operating_profit_year0",
        # 例: "C7": "personnel_cost_year0",
        # 例: "C9": "added_value_year0",
    },
}

WORD_PLACEHOLDER_MAPPING = {
    # プレースホルダー: データキー
    # 例: "{{会社名}}": "company_name",
    # 例: "{{代表者名}}": "representative",
    # 例: "{{設立年}}": "established_year",
    # 例: "{{事業概要}}": "business_overview",
    # 例: "{{課題1}}": "challenge_1",
    # 例: "{{解決策}}": "solution_overview",
}


def load_data(data_path: str) -> dict:
    """案件データを読み込む（JSON形式）"""
    path = Path(data_path)
    if not path.exists():
        print(f"エラー: データファイルが見つかりません: {data_path}", file=sys.stderr)
        sys.exit(1)

    if path.suffix == ".json":
        with open(path, encoding="utf-8") as f:
            return json.load(f)

    print(f"エラー: JSONファイルを指定してください: {data_path}", file=sys.stderr)
    sys.exit(1)


def fill_excel(template_path: str, data: dict, output_path: str):
    """Excelテンプレートにデータを書き込む"""
    wb = load_template(template_path)

    for sheet_name, cell_map in EXCEL_CELL_MAPPING.items():
        if sheet_name not in wb.sheetnames:
            print(f"警告: シート '{sheet_name}' が見つかりません。スキップします。")
            continue

        ws = wb[sheet_name]
        resolved = {}
        for cell_ref, data_key in cell_map.items():
            value = data.get(data_key, "")
            if value:
                resolved[cell_ref] = str(value)

        fill_cells(ws, resolved)

    save_output(wb, output_path)


def fill_word(template_path: str, data: dict, output_path: str):
    """Wordテンプレートにデータを書き込む"""
    doc = load_word_template(template_path)

    resolved = {}
    for placeholder, data_key in WORD_PLACEHOLDER_MAPPING.items():
        value = data.get(data_key, "")
        if value:
            resolved[placeholder] = str(value)

    replace_placeholders(doc, resolved)
    save_word_output(doc, output_path)


def analyze_template(template_path: str):
    """テンプレートの構造を解析して表示"""
    path = Path(template_path)

    if path.suffix in (".xlsx", ".xls"):
        info = get_template_info(template_path)
        print("=== Excelテンプレート構造 ===")
        for sheet_name, sheet_info in info.items():
            print(f"\nシート: {sheet_name}")
            print(f"  範囲: {sheet_info['dimensions']}")
            print(f"  行: {sheet_info['min_row']}〜{sheet_info['max_row']}")
            print(f"  列: {sheet_info['min_col']}〜{sheet_info['max_col']}")

    elif path.suffix in (".docx",):
        from utils.word_handler import load_template as load_doc, get_document_structure, list_placeholders
        doc = load_doc(template_path)
        structure = get_document_structure(doc)
        placeholders = list_placeholders(doc)

        print("=== Wordテンプレート構造 ===")
        print(f"段落数: {structure['paragraphs']}")
        print(f"テーブル数: {structure['tables']}")
        if structure["headings"]:
            print("\n見出し:")
            for h in structure["headings"]:
                print(f"  [{h['level']}] {h['text']}")
        if placeholders:
            print(f"\nプレースホルダー ({len(placeholders)}個):")
            for p in placeholders:
                print(f"  {p}")
        if structure["table_sizes"]:
            print("\nテーブル:")
            for t in structure["table_sizes"]:
                print(f"  テーブル{t['index']}: {t['rows']}行 x {t['cols']}列")


def main():
    parser = argparse.ArgumentParser(description="ものづくり補助金テンプレート処理")
    subparsers = parser.add_subparsers(dest="command")

    # fill コマンド
    fill_parser = subparsers.add_parser("fill", help="テンプレートにデータを書き込む")
    fill_parser.add_argument("--template", required=True, help="テンプレートファイルパス")
    fill_parser.add_argument("--data", required=True, help="データファイルパス (JSON)")
    fill_parser.add_argument("--output", required=True, help="出力ファイルパス")

    # analyze コマンド
    analyze_parser = subparsers.add_parser("analyze", help="テンプレートの構造を解析")
    analyze_parser.add_argument("--template", required=True, help="テンプレートファイルパス")

    args = parser.parse_args()

    if args.command == "fill":
        data = load_data(args.data)
        path = Path(args.template)
        if path.suffix in (".xlsx", ".xls"):
            fill_excel(args.template, data, args.output)
        elif path.suffix == ".docx":
            fill_word(args.template, data, args.output)
        else:
            print(f"エラー: 未対応のファイル形式: {path.suffix}", file=sys.stderr)
            sys.exit(1)

    elif args.command == "analyze":
        analyze_template(args.template)

    else:
        parser.print_help()


if __name__ == "__main__":
    main()
