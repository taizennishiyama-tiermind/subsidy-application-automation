#!/usr/bin/env python3
"""
Hearing Data Extractor

ヒアリング内容(自然言語テキスト、PDF、音声書き起こし等)から
補助金申請に必要な構造化データを抽出します。

Claude APIを使用して高精度な情報抽出を実現。
"""

import anthropic
import json
import sys
from pathlib import Path


class HearingDataExtractor:
    def __init__(self, schema_path=None):
        """
        Args:
            schema_path: 抽出スキーマJSONのパス(オプション)
        """
        self.client = anthropic.Anthropic()
        
        # デフォルトスキーマ
        self.schema = {
            "company_info": {
                "name": "",
                "name_kana": "",
                "representative": "",
                "representative_kana": "",
                "address": "",
                "postal_code": "",
                "tel": "",
                "fax": "",
                "email": "",
                "employee_count": 0,
                "capital": 0,
                "established_date": "",
                "business_type": "",
                "industry": ""
            },
            "project_info": {
                "title": "",
                "purpose": "",
                "background": "",
                "expected_outcome": "",
                "period_start": "",
                "period_end": "",
                "total_budget": 0,
                "subsidy_amount": 0
            },
            "expense_details": [
                {
                    "category": "",
                    "item_name": "",
                    "specification": "",
                    "quantity": 0,
                    "unit": "",
                    "unit_price": 0,
                    "amount": 0,
                    "vendor": "",
                    "note": ""
                }
            ],
            "implementation_plan": [
                {
                    "phase": "",
                    "period": "",
                    "content": "",
                    "deliverables": ""
                }
            ]
        }
        
        # カスタムスキーマがあれば上書き
        if schema_path:
            with open(schema_path, 'r', encoding='utf-8') as f:
                self.schema = json.load(f)
    
    def extract_from_text(self, hearing_text):
        """
        自然言語テキストから構造化データを抽出
        
        Args:
            hearing_text: ヒアリング内容のテキスト
        
        Returns:
            抽出されたデータ(dict)
        """
        print("🔍 情報抽出開始...")
        
        prompt = self._build_extraction_prompt(hearing_text)
        
        message = self.client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=8000,
            temperature=0,  # 再現性を高めるため0に設定
            messages=[{
                "role": "user",
                "content": prompt
            }]
        )
        
        # レスポンスからJSON部分を抽出
        response_text = message.content[0].text
        extracted_data = self._parse_json_response(response_text)
        
        print("✅ 抽出完了")
        return extracted_data
    
    def _build_extraction_prompt(self, hearing_text):
        """情報抽出用のプロンプトを構築"""
        return f"""あなたは補助金申請のエキスパートです。
以下のヒアリング内容から、補助金申請に必要な情報を正確に抽出し、
指定されたJSON形式で出力してください。

<extraction_schema>
{json.dumps(self.schema, indent=2, ensure_ascii=False)}
</extraction_schema>

<hearing_content>
{hearing_text}
</hearing_content>

抽出ルール:
1. 明示的に記載されている情報のみ抽出する
2. 推測や補完は行わない(不明な項目は空文字または0のまま)
3. 日付は必ず YYYY-MM-DD 形式に統一
4. 金額はカンマなしの数値(例: 1000000)
5. 複数の解釈が可能な場合、最も確実性の高い解釈を採用
6. expense_detailsやimplementation_planは、該当項目が複数あればすべて配列に含める

出力形式:
- JSON形式のみを出力(説明文、Markdownのコードブロック記号は不要)
- スキーマの構造を厳密に守る
- 日本語の値はそのまま使用

それでは抽出を開始してください。"""
    
    def _parse_json_response(self, response_text):
        """
        Claude APIのレスポンスからJSONを抽出・パース
        
        ```json ``` のようなマーカーを除去
        """
        # Markdownコードブロックを除去
        clean_text = response_text.strip()
        
        if clean_text.startswith('```json'):
            clean_text = clean_text[7:]  # ```json を削除
        if clean_text.startswith('```'):
            clean_text = clean_text[3:]  # ``` を削除
        if clean_text.endswith('```'):
            clean_text = clean_text[:-3]  # ``` を削除
        
        clean_text = clean_text.strip()
        
        try:
            return json.loads(clean_text)
        except json.JSONDecodeError as e:
            print(f"❌ JSON解析エラー: {e}")
            print(f"レスポンス:\n{response_text[:500]}")
            raise
    
    def validate(self, extracted_data):
        """
        抽出データの妥当性を検証
        
        Returns:
            (is_valid, issues)のタプル
        """
        issues = []
        
        # 必須項目チェック
        if not extracted_data.get("company_info", {}).get("name"):
            issues.append("❌ 会社名が抽出されていません")
        
        if not extracted_data.get("project_info", {}).get("title"):
            issues.append("❌ 事業名が抽出されていません")
        
        # 日付フォーマットチェック
        import re
        date_pattern = r'^\d{4}-\d{2}-\d{2}$'
        
        date_fields = [
            ("project_info.period_start", "事業開始日"),
            ("project_info.period_end", "事業終了日"),
            ("company_info.established_date", "設立日")
        ]
        
        for field_path, field_name in date_fields:
            keys = field_path.split(".")
            value = extracted_data
            for key in keys:
                value = value.get(key, "")
            
            if value and not re.match(date_pattern, value):
                issues.append(f"⚠️ {field_name}のフォーマットが不正: {value} (期待: YYYY-MM-DD)")
        
        # 金額の整合性チェック
        expense_details = extracted_data.get("expense_details", [])
        if expense_details:
            expense_total = sum(item.get("amount", 0) for item in expense_details)
            declared_total = extracted_data.get("project_info", {}).get("total_budget", 0)
            
            if expense_total > 0 and declared_total > 0:
                diff = abs(expense_total - declared_total)
                if diff > 1000:  # 1000円以上の差異
                    issues.append(
                        f"⚠️ 経費合計(¥{expense_total:,})と"
                        f"予算総額(¥{declared_total:,})が一致しません (差額: ¥{diff:,})"
                    )
        
        # 論理的な日付順序チェック
        period_start = extracted_data.get("project_info", {}).get("period_start")
        period_end = extracted_data.get("project_info", {}).get("period_end")
        
        if period_start and period_end:
            if period_start > period_end:
                issues.append(f"❌ 事業期間が不正: 開始日({period_start}) > 終了日({period_end})")
        
        is_valid = len([i for i in issues if i.startswith("❌")]) == 0
        
        return is_valid, issues
    
    def save_to_file(self, extracted_data, output_path):
        """抽出結果をJSONファイルに保存"""
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(extracted_data, f, indent=2, ensure_ascii=False)
        
        print(f"💾 保存完了: {output_path}")
    
    def print_summary(self, extracted_data):
        """抽出結果のサマリーを表示"""
        print("\n" + "=" * 60)
        print("📊 抽出結果サマリー")
        print("=" * 60)
        
        company_name = extracted_data.get("company_info", {}).get("name", "不明")
        project_title = extracted_data.get("project_info", {}).get("title", "不明")
        expense_count = len(extracted_data.get("expense_details", []))
        
        print(f"会社名: {company_name}")
        print(f"事業名: {project_title}")
        print(f"経費項目数: {expense_count}件")
        
        total_budget = extracted_data.get("project_info", {}).get("total_budget", 0)
        if total_budget > 0:
            print(f"総予算: ¥{total_budget:,}")


def main():
    if len(sys.argv) < 2:
        print("使い方: python extract_data.py <ヒアリング内容.txt> [出力先.json] [スキーマ.json]")
        print("\n例:")
        print("  python extract_data.py hearing.txt extracted_data.json")
        print("  python extract_data.py hearing.txt extracted_data.json custom_schema.json")
        sys.exit(1)
    
    hearing_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else "extracted_data.json"
    schema_file = sys.argv[3] if len(sys.argv) > 3 else None
    
    # ヒアリング内容を読み込み
    with open(hearing_file, 'r', encoding='utf-8') as f:
        hearing_text = f.read()
    
    # 抽出実行
    extractor = HearingDataExtractor(schema_file)
    extracted_data = extractor.extract_from_text(hearing_text)
    
    # サマリー表示
    extractor.print_summary(extracted_data)
    
    # 検証
    is_valid, issues = extractor.validate(extracted_data)
    
    if issues:
        print("\n" + "=" * 60)
        print("⚠️ 検証結果")
        print("=" * 60)
        for issue in issues:
            print(issue)
    
    if not is_valid:
        print("\n❌ 重大なエラーがあります。修正が必要です。")
    else:
        print("\n✅ 検証OK")
    
    # 保存
    extractor.save_to_file(extracted_data, output_file)


if __name__ == "__main__":
    main()
