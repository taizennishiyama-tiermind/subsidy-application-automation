"""IT導入補助金 / デジタル化・AI導入補助金 テンプレート書き込みスクリプト

使い方:
  python grants/IT導入補助金/scripts/fill_template.py \
    fill --template grants/IT導入補助金/templates/xxx.xlsx \
         --data workspace/{案件名}/application_data.json \
         --output workspace/{案件名}/申請書_IT導入補助金.xlsx

  python grants/IT導入補助金/scripts/fill_template.py \
    analyze --template grants/IT導入補助金/templates/xxx.xlsx
"""

import argparse
import json
import sys
from pathlib import Path

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
    get_document_structure,
    list_placeholders,
)

# ===== セルマッピング定義 =====
# ※ 実際のテンプレートをアップロード後に更新すること

EXCEL_CELL_MAPPING = {
    # --- カテゴリー６ 価格説明資料 (it2026_kakakusetsumei_cate6.xlsx) ---
    #
    # 【注意】業務内容シートの Z列・BN列 はITツール登録コード（Tコード）専用の
    # 大規模結合セルのため openpyxl では書き込み不可。
    # 実際の作業内容・備考は「価格の内訳」シートの BJ列 に記載する。
    #
    # ヘッダー情報（業務内容シートに入力 → 価格の内訳シートは数式で自動反映）
    "業務内容": {
        "K10": "it_vendor_name",        # IT導入支援事業者名
        "K12": "category6_name",        # カテゴリー６の名称
        "K14": "target_software",       # 役務提供の対象となるソフトウェア
    },
    # 価格の内訳シート（時間単価・時間・人数・備考を入力、金額は数式で自動計算）
    # 行番号: 導入設定=23〜41, CSV=49〜67, カスタマイズ=75〜93,
    #         研修=101〜119, マニュアル=127〜145, RPA/AI=153〜171
    "価格の内訳": {
        # 導入設定費用・テーブル設定費用等 (行23〜)
        "AA23": "setup_rate_1",
        "AJ23": "setup_hours_1",
        "AS23": "setup_headcount_1",
        "BJ23": "setup_work_1",         # 実施作業内容
        "AA25": "setup_rate_2",
        "AJ25": "setup_hours_2",
        "AS25": "setup_headcount_2",
        "BJ25": "setup_work_2",
        # 研修資料作成・研修実施費用 (行101〜)
        "AA101": "training_rate_1",
        "AJ101": "training_hours_1",
        "AS101": "training_headcount_1",
        "BJ101": "training_work_1",
        # 運用マニュアル作成費用 (行127〜)
        "AA127": "manual_rate_1",
        "AJ127": "manual_hours_1",
        "AS127": "manual_headcount_1",
        "BJ127": "manual_work_1",
    },
}


# --- カテゴリー７ 価格説明資料 (it2026_kakakusetsumei_cate7.xlsx) ---
# IT導入支援事業者（ベンダー）が事務局にITツールを登録する際の価格説明資料。
# カテゴリー7 = 保守費用・問合せ窓口費用のみ含む役務提供サービス。
# 書き込み禁止: 価格の内訳シートのK10/K12/K14（数式で自動反映）、
#               BB列金額セル（自動計算）、AH14（標準販売価格）、BW〜CB列（T-コード検証）
EXCEL_CELL_MAPPING_CATE7 = {
    "業務内容": {
        "K10": "it_vendor_name",            # IT導入支援事業者名
        "K12": "category7_name",            # カテゴリー７の名称
        "K14": "target_software",           # 役務提供の対象となるソフトウェア
        "B21": "maintenance_description",   # 保守費用の業務内容説明（結合セル B21:BU30）
        "B33": "support_description",       # 問合せ窓口費用の業務内容説明（結合セル B33:BU42）
    },
    "価格の内訳": {
        # 保守費用テーブル（行23〜42、奇数行のみ。各行は2行結合）時間単価上限10,000円
        "B23":  "maintenance_work_1",  "AA23": "maintenance_rate_1",  "AJ23": "maintenance_hours_1",  "AS23": "maintenance_headcount_1",  "BJ23": "maintenance_note_1",
        "B25":  "maintenance_work_2",  "AA25": "maintenance_rate_2",  "AJ25": "maintenance_hours_2",  "AS25": "maintenance_headcount_2",  "BJ25": "maintenance_note_2",
        "B27":  "maintenance_work_3",  "AA27": "maintenance_rate_3",  "AJ27": "maintenance_hours_3",  "AS27": "maintenance_headcount_3",  "BJ27": "maintenance_note_3",
        "B29":  "maintenance_work_4",  "AA29": "maintenance_rate_4",  "AJ29": "maintenance_hours_4",  "AS29": "maintenance_headcount_4",  "BJ29": "maintenance_note_4",
        "B31":  "maintenance_work_5",  "AA31": "maintenance_rate_5",  "AJ31": "maintenance_hours_5",  "AS31": "maintenance_headcount_5",  "BJ31": "maintenance_note_5",
        "B33":  "maintenance_work_6",  "AA33": "maintenance_rate_6",  "AJ33": "maintenance_hours_6",  "AS33": "maintenance_headcount_6",  "BJ33": "maintenance_note_6",
        "B35":  "maintenance_work_7",  "AA35": "maintenance_rate_7",  "AJ35": "maintenance_hours_7",  "AS35": "maintenance_headcount_7",  "BJ35": "maintenance_note_7",
        "B37":  "maintenance_work_8",  "AA37": "maintenance_rate_8",  "AJ37": "maintenance_hours_8",  "AS37": "maintenance_headcount_8",  "BJ37": "maintenance_note_8",
        "B39":  "maintenance_work_9",  "AA39": "maintenance_rate_9",  "AJ39": "maintenance_hours_9",  "AS39": "maintenance_headcount_9",  "BJ39": "maintenance_note_9",
        "B41":  "maintenance_work_10", "AA41": "maintenance_rate_10", "AJ41": "maintenance_hours_10", "AS41": "maintenance_headcount_10", "BJ41": "maintenance_note_10",
        # 問合せ窓口費用テーブル（行49〜68、奇数行のみ）
        "B49":  "support_work_1",  "AA49": "support_rate_1",  "AJ49": "support_hours_1",  "AS49": "support_headcount_1",  "BJ49": "support_note_1",
        "B51":  "support_work_2",  "AA51": "support_rate_2",  "AJ51": "support_hours_2",  "AS51": "support_headcount_2",  "BJ51": "support_note_2",
        "B53":  "support_work_3",  "AA53": "support_rate_3",  "AJ53": "support_hours_3",  "AS53": "support_headcount_3",  "BJ53": "support_note_3",
        "B55":  "support_work_4",  "AA55": "support_rate_4",  "AJ55": "support_hours_4",  "AS55": "support_headcount_4",  "BJ55": "support_note_4",
        "B57":  "support_work_5",  "AA57": "support_rate_5",  "AJ57": "support_hours_5",  "AS57": "support_headcount_5",  "BJ57": "support_note_5",
        "B59":  "support_work_6",  "AA59": "support_rate_6",  "AJ59": "support_hours_6",  "AS59": "support_headcount_6",  "BJ59": "support_note_6",
        "B61":  "support_work_7",  "AA61": "support_rate_7",  "AJ61": "support_hours_7",  "AS61": "support_headcount_7",  "BJ61": "support_note_7",
        "B63":  "support_work_8",  "AA63": "support_rate_8",  "AJ63": "support_hours_8",  "AS63": "support_headcount_8",  "BJ63": "support_note_8",
        "B65":  "support_work_9",  "AA65": "support_rate_9",  "AJ65": "support_hours_9",  "AS65": "support_headcount_9",  "BJ65": "support_note_9",
        "B67":  "support_work_10", "AA67": "support_rate_10", "AJ67": "support_hours_10", "AS67": "support_headcount_10", "BJ67": "support_note_10",
    },
}

WORD_PLACEHOLDER_MAPPING = {
    # "{{会社名}}": "company_name",
    # "{{経営課題}}": "business_challenge",
    # "{{ITツール名}}": "it_tool_name",
    # "{{導入効果}}": "expected_effect",
}


def to_field_profile(cell_mapping: dict[str, dict[str, str]]) -> dict[str, dict[str, dict]]:
    profile: dict[str, dict[str, dict]] = {}
    for sheet_name, sheet_map in cell_mapping.items():
        profile[sheet_name] = {}
        for cell_ref, data_key in sheet_map.items():
            field_name = f"{data_key}_{cell_ref.lower()}"
            profile[sheet_name][field_name] = {
                "data_key": data_key,
                "targets": [cell_ref],
                "required": False,
            }
    return profile


def load_data(data_path: str) -> dict:
    path = Path(data_path)
    if not path.exists():
        print(f"エラー: データファイルが見つかりません: {data_path}", file=sys.stderr)
        sys.exit(1)
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def fill_excel(template_path: str, data: dict, output_path: str):
    # テンプレートファイル名でマッピング辞書を自動選択
    tpl_name = Path(template_path).name.lower()
    if "cate7" in tpl_name:
        mapping = EXCEL_CELL_MAPPING_CATE7
    else:
        mapping = EXCEL_CELL_MAPPING

    wb = load_template(template_path)
    field_profile = to_field_profile(mapping)
    sheet_cell_map, report = build_sheet_cell_map(wb, field_profile, data)
    validate_resolution_report(report)
    print(f"マッピング解決: {summarize_resolution_report(report)}")
    fill_xlsx_safe(template_path, sheet_cell_map, output_path)


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
    parser = argparse.ArgumentParser(description="IT導入補助金テンプレート処理")
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
