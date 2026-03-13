

## 使い方



### 詳細な手順

#### フェーズ0: テンプレート分析

最初に、Excelテンプレートの構造を詳細に分析します:

```bash
python scripts/analyze_template.py \
  /mnt/user-data/uploads/template.xlsx \
  template_analysis.json
```

これにより以下の情報が取得できます:
- シート構成
- セル結合パターン(左上セルを自動検出)
- Named Range一覧
- 動的表の検出
- 数式セル(上書き禁止)

#### フェーズ1: 情報抽出

ヒアリング内容から構造化データを抽出:

```bash
python scripts/extract_data.py \
  /mnt/user-data/uploads/hearing_notes.txt \
  extracted_data.json
```

抽出される情報:
- 会社情報(名称、住所、代表者など)
- 事業情報(事業名、目的、期間など)
- 経費明細(項目、金額、仕様など)
- 実施計画

カスタムスキーマを使用する場合:

```bash
python scripts/extract_data.py \
  hearing_notes.txt \
  extracted_data.json \
  custom_schema.json
```

#### フェーズ2: マッピング定義作成

テンプレート分析結果と抽出データを元に、マッピング定義を作成します。

サンプルを参考にしてください:
```bash
cp examples/sample_mapping.json my_mapping.json
# my_mapping.json を編集
```

詳細は `references/mapping_schema.md` を参照。

#### フェーズ3: データ埋め込み

マッピング定義に従ってデータを埋め込みます:

```bash
python scripts/fill_template.py \
  /mnt/user-data/uploads/template.xlsx \
  my_mapping.json \
  extracted_data.json \
  /mnt/user-data/outputs/completed_application.xlsx
```

#### フェーズ4: 検証

埋め込みスクリプトが自動的に検証を実行しますが、手動でも確認してください:

1. 完成ファイルをExcelで開く
2. 主要セルが埋まっているか確認
3. 数式セルが壊れていないか確認
4. 日付・金額のフォーマットが正しいか確認



## マッピング定義の構造

マッピング定義は3種類の埋め込みタイプをサポート:

### 1. simple_cells (単純セル)

単一のセルに単一の値を埋め込む:

```json
{
  "cell": "B3",
  "data_path": "company_info.name",
  "description": "会社名",
  "format": "text"
}
```

### 2. named_ranges (Named Range)

Excelの名前付き範囲に埋め込む:

```json
{
  "name": "CompanyName",
  "data_path": "company_info.name",
  "description": "会社名(Named Range)"
}
```

### 3. dynamic_tables (動的表)

行数可変の表に埋め込む(経費明細など):

```json
{
  "table_id": "expense_table",
  "header_row": 15,
  "data_start_row": 16,
  "columns": [
    {"col": "B", "data_field": "category"},
    {"col": "C", "data_field": "item_name"},
    {"col": "F", "data_field": "amount", "format": "currency"}
  ],
  "data_path": "expense_details",
  "auto_sum": {
    "row_offset": 1,
    "col": "F",
    "formula_template": "=SUM(F{start}:F{end})"
  }
}
```

## 重要な注意点

### セル結合の扱い

⚠️ セル結合がある場合、必ず**左上セル**に書き込んでください。

```
B3:D3が結合されている場合
  → cellには "B3" を指定
```

`analyze_template.py`が自動的に左上セルを検出します。

### 数式セルの保護

🚫 数式が入っているセルは絶対に上書きしないでください。

マッピング定義に含めず、`auto_sum`を使って数式を自動生成してください。

### 日付・金額フォーマット

📅 日付や金額には必ず適切なフォーマットを指定:

```json
{"format": "date"}       // YYYY/MM/DD
{"format": "currency"}   // ¥#,##0
{"format": "number"}     // #,##0
```

## トラブルシューティング

### Q: セルに値が埋まらない

**原因と対策:**
1. `data_path`のタイポ → 抽出データのキー名と一致しているか確認
2. セル結合 → 左上セルを指定しているか確認
3. データ不足 → 抽出データに該当フィールドが存在するか確認

### Q: 数式が壊れた

**原因:** 数式セルをマッピング定義に含めてしまった

**対策:** 
- 数式セルはマッピングから除外
- `auto_sum`を使用して数式を自動生成

### Q: 日付が数値で表示される

**原因:** フォーマットが指定されていない

**対策:** `"format": "date"` を追加

### Q: 動的表の行数が足りない

**原因:** `data_end_row`の設定が小さい

**対策:** 
- `data_end_row`を大きめに設定
- またはテンプレートに十分な行を用意

## ベストプラクティス

### 1. マッピング定義の再利用

同じ補助金プログラムでは、テンプレートが統一されることが多いため、
一度作成したマッピング定義を保存・再利用してください。

```bash
# プログラムごとにマッピングを保存
mkdir -p mappings
cp my_mapping.json mappings/ものづくり補助金_2024.json
```

### 2. ヒアリングシートの標準化

情報抽出の精度を上げるため、ヒアリング時に構造化シートを使用:

```
【基本情報】
会社名:
代表者:
...

【事業内容】
事業名:
目的:
...
```

### 3. テンプレート分析の徹底

フェーズ0のテンプレート分析は絶対に省略しないでください。
これが埋め込み精度の要です。

### 4. 検証の実施

埋め込み後は必ず検証を実施:
- スクリプトの自動検証
- 目視での最終確認
- 数式の動作確認

### 5. エラーログの保持

本番運用では、エラーログを保存してトラブル対応に活用:

```python
import logging
logging.basicConfig(filename='subsidy_filler.log', level=logging.INFO)
```



## ライセンスと使用条件

このスキルは補助金申請支援業務の効率化を目的としています。
実際の申請前には必ず人間による最終確認を行ってください。

## サポート

詳細なドキュメント:
- `SKILL.md` - スキル全体の詳細説明
- `references/mapping_schema.md` - マッピング定義の詳細仕様

質問や問題がある場合は、Claudeに直接お尋ねください。
