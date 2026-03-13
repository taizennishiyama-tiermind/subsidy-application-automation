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

from utils.excel_handler import fill_xlsx_safe, get_template_info, load_template
from utils.template_mapping import (
    build_sheet_cell_map,
    summarize_resolution_report,
    validate_resolution_report,
)
from utils.word_handler import (
    load_template as load_word_template,
    replace_placeholders,
    save_output as save_word_output,
)


# ===== セルマッピング定義 =====
# テンプレートのセル位置 → データキーの対応表
# ※ 実際のテンプレートをアップロード後に更新すること

EXCEL_FIELD_MAPPING = {
    "確認書": {
        "input_date": {
            "data_key": "input_date",
            "required": True,
            "targets": ["I8"],
            "anchors": ["入力日", "提出日"],
        },
        "address": {
            "data_key": "address",
            "required": True,
            "targets": ["I9"],
            "anchors": ["住所", "所在地"],
        },
        "company_name": {
            "data_key": "company_name",
            "required": True,
            "targets": ["I10"],
            "anchors": ["事業者名", "事業者名称", "会社名"],
        },
        "representative_title": {
            "data_key": "representative_title",
            "required": True,
            "targets": ["I11"],
            "anchors": ["役職", "代表者役職"],
        },
        "representative_name": {
            "data_key": "representative_name",
            "required": True,
            "targets": ["I12"],
            "anchors": ["氏名", "代表者氏名"],
        },
        # 確認書シート下部: 対象月・従業員数・最低賃金未満人数（TEMPLATE_SKILL.md 2026-03-13追加）
        "target_month_1_kakuninsho": {
            "data_key": "target_month_1",
            "required": True,
            "targets": ["I29"],
            "anchors": ["対象月①", "対象月"],
        },
        "target_month_2_kakuninsho": {
            "data_key": "target_month_2",
            "required": True,
            "targets": ["K29"],
            "anchors": ["対象月②"],
        },
        "target_month_3_kakuninsho": {
            "data_key": "target_month_3",
            "required": True,
            "targets": ["M29"],
            "anchors": ["対象月③"],
        },
        "total_employees_month_1_kakuninsho": {
            "data_key": "total_employees_month_1",
            "required": True,
            "targets": ["I30"],
            "anchors": ["全従業員数", "①全従業員"],
        },
        # K30/M30/K31/M31 は数式セル（対象月①〜③から自動集計）のため optional
        "total_employees_month_2_kakuninsho": {
            "data_key": "total_employees_month_2",
            "required": False,
            "targets": ["K30"],
        },
        "total_employees_month_3_kakuninsho": {
            "data_key": "total_employees_month_3",
            "required": False,
            "targets": ["M30"],
        },
        "under_min_wage_count_month_1": {
            "data_key": "under_min_wage_count_month_1",
            "required": False,
            "targets": ["I31"],
            "anchors": ["最低賃金未満", "②改定後"],
        },
        "under_min_wage_count_month_2": {
            "data_key": "under_min_wage_count_month_2",
            "required": False,
            "targets": ["K31"],
        },
        "under_min_wage_count_month_3": {
            "data_key": "under_min_wage_count_month_3",
            "required": False,
            "targets": ["M31"],
        },
    },
    "対象月①": {
        "company_name_m1": {"data_key": "company_name", "required": True, "targets": ["A4"], "anchors": ["事業者名"]},
        "total_employees_month_1": {"data_key": "total_employees_month_1", "required": True, "targets": ["D4"]},
        "target_month_1": {"data_key": "target_month_1", "required": True, "targets": ["F4"]},
        # K4 は qualifying_count（数式セルの可能性あり）: optional
        "qualifying_count_month_1": {"data_key": "qualifying_count_month_1", "required": False, "targets": ["K4"]},
        "month_1_emp_1_id": {"data_key": "month_1_emp_1_id", "required": True, "targets": ["B10"]},
        "month_1_emp_1_name": {"data_key": "month_1_emp_1_name", "required": True, "targets": ["C10"]},
        "month_1_emp_1_prefecture": {"data_key": "month_1_emp_1_prefecture", "required": True, "targets": ["D10"]},
        # E10 最低賃金額は数式または固定値セルの可能性あり: optional
        "month_1_emp_1_min_wage": {"data_key": "month_1_emp_1_min_wage", "required": False, "targets": ["E10"]},
        "month_1_emp_1_wage_type": {"data_key": "month_1_emp_1_wage_type", "required": True, "targets": ["G10"]},
        "month_1_emp_1_amount": {"data_key": "month_1_emp_1_amount", "required": True, "targets": ["H10"]},
        "month_1_emp_1_hours": {"data_key": "month_1_emp_1_hours", "required": True, "targets": ["I10"]},
        "month_1_emp_1_hourly_wage": {"data_key": "month_1_emp_1_hourly_wage", "targets": ["J10"]},
        "month_1_emp_2_id": {"data_key": "month_1_emp_2_id", "targets": ["B11"]},
        "month_1_emp_2_name": {"data_key": "month_1_emp_2_name", "targets": ["C11"]},
        "month_1_emp_2_prefecture": {"data_key": "month_1_emp_2_prefecture", "targets": ["D11"]},
        "month_1_emp_2_min_wage": {"data_key": "month_1_emp_2_min_wage", "targets": ["E11"]},
        "month_1_emp_2_wage_type": {"data_key": "month_1_emp_2_wage_type", "targets": ["G11"]},
        "month_1_emp_2_amount": {"data_key": "month_1_emp_2_amount", "targets": ["H11"]},
        "month_1_emp_2_hours": {"data_key": "month_1_emp_2_hours", "targets": ["I11"]},
        "month_1_emp_2_hourly_wage": {"data_key": "month_1_emp_2_hourly_wage", "targets": ["J11"]},
    },
    "対象月②": {
        "company_name_m2": {"data_key": "company_name", "required": True, "targets": ["A4"], "anchors": ["事業者名"]},
        "total_employees_month_2": {"data_key": "total_employees_month_2", "required": True, "targets": ["D4"]},
        "target_month_2": {"data_key": "target_month_2", "required": True, "targets": ["F4"]},
        "qualifying_count_month_2": {"data_key": "qualifying_count_month_2", "required": False, "targets": ["K4"]},
    },
    "対象月③": {
        "company_name_m3": {"data_key": "company_name", "required": True, "targets": ["A4"], "anchors": ["事業者名"]},
        "total_employees_month_3": {"data_key": "total_employees_month_3", "required": True, "targets": ["D4"]},
        "target_month_3": {"data_key": "target_month_3", "required": True, "targets": ["F4"]},
        "qualifying_count_month_3": {"data_key": "qualifying_count_month_3", "required": False, "targets": ["K4"]},
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


def flatten_data(data: dict) -> dict:
    """ネストしたJSONからスクリプトが参照しやすい平坦な辞書を作る"""
    flat = {}

    def _walk(value):
        if isinstance(value, dict):
            for k, v in value.items():
                if k not in flat and not isinstance(v, (dict, list)):
                    flat[k] = v
                _walk(v)
        elif isinstance(value, list):
            for item in value:
                _walk(item)

    _walk(data)
    return flat


def load_data(data_path: str) -> dict:
    """案件データを読み込む（JSON形式）"""
    path = Path(data_path)
    if not path.exists():
        print(f"エラー: データファイルが見つかりません: {data_path}", file=sys.stderr)
        sys.exit(1)

    if path.suffix == ".json":
        with open(path, encoding="utf-8") as f:
            raw = json.load(f)
            if isinstance(raw, dict):
                flat = flatten_data(raw)
                flat["_raw"] = raw
                return flat
            return raw

    print(f"エラー: JSONファイルを指定してください: {data_path}", file=sys.stderr)
    sys.exit(1)


def fill_excel(template_path: str, data: dict, output_path: str):
    """Excelテンプレートにデータを書き込む"""
    wb = load_template(template_path)
    sheet_cell_map, report = build_sheet_cell_map(wb, EXCEL_FIELD_MAPPING, data)
    validate_resolution_report(report)
    print(f"マッピング解決: {summarize_resolution_report(report)}")
    fill_xlsx_safe(template_path, sheet_cell_map, output_path)


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
