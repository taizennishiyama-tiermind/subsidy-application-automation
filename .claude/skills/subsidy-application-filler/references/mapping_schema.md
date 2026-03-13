# マッピング定義リファレンス

このドキュメントでは、マッピング定義JSONの構造と各フィールドの説明を提供します。

## 概要

マッピング定義は、抽出データをExcelテンプレートのどのセルに埋め込むかを定義するJSONファイルです。

## 基本構造

```json
{
  "version": "1.0",
  "template_file": "template.xlsx",
  "description": "○○補助金申請書用マッピング",
  "mappings": {
    "Sheet1": {
      "simple_cells": [...],
      "named_ranges": [...],
      "dynamic_tables": [...],
      "field_resolution_rules": [...]
    }
  }
}
```

---

## フィールド詳細

### トップレベル

| フィールド | 型 | 必須 | 説明 |
|-----------|-----|------|------|
| `version` | string | Yes | マッピング定義のバージョン(現在は"1.0") |
| `template_file` | string | Yes | 対象テンプレートファイル名 |
| `description` | string | No | このマッピングの説明 |
| `mappings` | object | Yes | シート名をキーとしたマッピング定義 |

### field_resolution_rules (差し込み位置の再解決ルール)

テンプレート版が変わっても位置ずれしにくくするため、固定セルだけでなく
ラベルアンカーと周辺構造を使って毎回差し込み先を再解決する。

```json
{
  "field_name": "company_name",
  "data_key": "company_name",
  "required": true,
  "targets": ["I10"],
  "anchors": ["事業者名", "会社名"],
  "preferred_offsets": [[0, 1], [0, 2], [1, 0], [1, 1]]
}
```

| フィールド | 型 | 必須 | 説明 |
|-----------|-----|------|------|
| `field_name` | string | Yes | 項目の識別子 |
| `data_key` | string | Yes | `application_data.json` 側のキー |
| `required` | boolean | No | 必須項目か。`true` のとき未解決なら出力を失敗にする |
| `targets` | array | No | 既知の固定セル候補 |
| `anchors` | array | No | ラベル候補。毎回テンプレート内で検索する |
| `preferred_offsets` | array | No | ラベルセルからどの方向に入力欄があるかの優先順 |

運用ルール:
- `targets` は初期候補であり、毎回そのまま書く前に `anchors` と周辺構造で再確認する
- `anchors` に一致したセルの右隣、直下、結合セルの右側を優先して入力欄を推定する
- `required = true` の項目は、差し込み先が解決できないまま出力しない
- `記入例` `記入方法` `説明` `見本` などの参照専用シートはマッピング対象から除外する

### simple_cells (単純セル)

単一のセルに単一の値を埋め込む場合に使用。

```json
{
  "cell": "B3",
  "data_path": "company_info.name",
  "description": "会社名",
  "format": "text"
}
```

| フィールド | 型 | 必須 | 説明 |
|-----------|-----|------|------|
| `cell` | string | Yes | セル座標(例: "B3", "AA10") |
| `data_path` | string | Yes | 抽出データ内のパス(ドット区切り) |
| `description` | string | No | このマッピングの説明 |
| `format` | string | No | セルフォーマット("date", "number", "currency", "percentage", "text") |

**data_path の例:**
- `"company_info.name"` → extracted_data["company_info"]["name"]
- `"project_info.total_budget"` → extracted_data["project_info"]["total_budget"]

### named_ranges (Named Range)

Excelの名前付き範囲への埋め込み。

```json
{
  "name": "CompanyName",
  "data_path": "company_info.name",
  "description": "会社名(Named Range)",
  "format": "text"
}
```

| フィールド | 型 | 必須 | 説明 |
|-----------|-----|------|------|
| `name` | string | Yes | Named Range の名前 |
| `data_path` | string | Yes | 抽出データ内のパス |
| `description` | string | No | このマッピングの説明 |
| `format` | string | No | セルフォーマット |

### dynamic_tables (動的表)

行数が可変の表(経費明細、従業員リストなど)への埋め込み。

```json
{
  "table_id": "expense_table",
  "header_row": 15,
  "data_start_row": 16,
  "data_end_row": 100,
  "clear_existing": true,
  "columns": [
    {
      "col": "B",
      "data_field": "category",
      "format": "text"
    },
    {
      "col": "C",
      "data_field": "item_name",
      "format": "text"
    },
    {
      "col": "F",
      "data_field": "amount",
      "format": "currency"
    }
  ],
  "data_path": "expense_details",
  "auto_sum": {
    "row_offset": 1,
    "col": "F",
    "formula_template": "=SUM(F{start}:F{end})",
    "format": "currency"
  }
}
```

| フィールド | 型 | 必須 | 説明 |
|-----------|-----|------|------|
| `table_id` | string | Yes | テーブルの識別子 |
| `header_row` | integer | Yes | ヘッダー行の行番号 |
| `data_start_row` | integer | Yes | データ開始行の行番号 |
| `data_end_row` | integer | No | データ終了行(クリア範囲の指定) |
| `clear_existing` | boolean | No | 既存データをクリアするか(デフォルト: false) |
| `columns` | array | Yes | 列定義の配列 |
| `data_path` | string | Yes | 抽出データ内の配列パス |
| `auto_sum` | object | No | 合計行の自動生成設定 |

**columns の各要素:**

| フィールド | 型 | 必須 | 説明 |
|-----------|-----|------|------|
| `col` | string | Yes | 列文字(例: "B", "AA") |
| `data_field` | string | Yes | 配列要素内のフィールド名 |
| `format` | string | No | セルフォーマット |

**auto_sum の詳細:**

| フィールド | 型 | 必須 | 説明 |
|-----------|-----|------|------|
| `row_offset` | integer | Yes | データ最終行からのオフセット(通常1) |
| `col` | string | Yes | 合計を表示する列 |
| `formula_template` | string | Yes | 数式テンプレート({start}, {end}, {col}が置換される) |
| `format` | string | No | セルフォーマット |

---



```json
{
  "version": "1.0",
  "template_file": "ものづくり補助金申請書.xlsx",
  "description": "ものづくり補助金 第XX次公募用",
  "mappings": {
    "申請書": {
      "simple_cells": [
        {
          "cell": "B3",
          "data_path": "company_info.name",
          "description": "会社名",
          "format": "text"
        },
        {
          "cell": "B5",
          "data_path": "company_info.representative",
          "description": "代表者氏名",
          "format": "text"
        },
        {
          "cell": "B7",
          "data_path": "company_info.address",
          "description": "所在地",
          "format": "text"
        },
        {
          "cell": "D10",
          "data_path": "project_info.period_start",
          "description": "事業開始日",
          "format": "date"
        },
        {
          "cell": "F10",
          "data_path": "project_info.period_end",
          "description": "事業終了日",
          "format": "date"
        }
      ],
      "named_ranges": [
        {
          "name": "CompanyTel",
          "data_path": "company_info.tel",
          "description": "電話番号",
          "format": "text"
        },
        {
          "name": "ProjectTitle",
          "data_path": "project_info.title",
          "description": "事業名",
          "format": "text"
        }
      ],
      "dynamic_tables": [
        {
          "table_id": "expense_table",
          "header_row": 15,
          "data_start_row": 16,
          "data_end_row": 50,
          "clear_existing": true,
          "columns": [
            {
              "col": "B",
              "data_field": "category",
              "format": "text"
            },
            {
              "col": "C",
              "data_field": "item_name",
              "format": "text"
            },
            {
              "col": "D",
              "data_field": "quantity",
              "format": "number"
            },
            {
              "col": "E",
              "data_field": "unit_price",
              "format": "currency"
            },
            {
              "col": "F",
              "data_field": "amount",
              "format": "currency"
            }
          ],
          "data_path": "expense_details",
          "auto_sum": {
            "row_offset": 1,
            "col": "F",
            "formula_template": "=SUM(F{start}:F{end})",
            "format": "currency"
          }
        }
      ]
    },
    "事業計画": {
      "simple_cells": [
        {
          "cell": "B3",
          "data_path": "project_info.purpose",
          "description": "事業の目的",
          "format": "text"
        },
        {
          "cell": "B10",
          "data_path": "project_info.expected_outcome",
          "description": "期待される効果",
          "format": "text"
        }
      ],
      "dynamic_tables": [
        {
          "table_id": "implementation_plan",
          "header_row": 20,
          "data_start_row": 21,
          "columns": [
            {
              "col": "B",
              "data_field": "phase",
              "format": "text"
            },
            {
              "col": "C",
              "data_field": "period",
              "format": "text"
            },
            {
              "col": "D",
              "data_field": "content",
              "format": "text"
            }
          ],
          "data_path": "implementation_plan"
        }
      ]
    }
  }
}
```

---

## ベストプラクティス

### 1. セル結合への対応

テンプレートにセル結合がある場合、`cell`には常に**結合範囲の左上セル**を指定してください。
スクリプトが自動的に左上セルを検出しますが、明示的に指定する方が安全です。

```json
// セル B3:D3 が結合されている場合
{
  "cell": "B3",  // 左上セル
  "data_path": "company_info.name"
}
```

### 2. 数式セルの保護

数式が入っているセル(合計セルなど)は、マッピング定義に含めないでください。
誤って上書きすると計算が壊れます。

`auto_sum`を使用すると、数式を自動生成できます。

### 3. フォーマットの指定

日付や金額には必ず適切なフォーマットを指定してください:

```json
{
  "cell": "D10",
  "data_path": "project_info.period_start",
  "format": "date"  // YYYY/MM/DD で表示
}
```

### 4. data_path の検証

`data_path`は抽出データの構造と完全に一致している必要があります。
タイポがあるとデータが埋まりません。

### 5. 動的表の設計

`data_start_row`と`data_end_row`の間に十分な余裕を持たせてください。
データ数が想定を超えても問題ないようにします。

---

## トラブルシューティング

### Q: セルに値が埋まらない

**A:** 以下を確認:
1. `data_path`のタイポがないか
2. 抽出データに該当フィールドが存在するか
3. セル座標が正しいか
4. セル結合の場合、左上セルを指定しているか

### Q: 数式が壊れた

**A:** 数式セルをマッピング定義に含めていないか確認。
`auto_sum`を使用して数式を自動生成してください。

### Q: 日付が数値として表示される

**A:** `format: "date"`を指定してください。

### Q: 動的表の行数が足りない

**A:** `data_end_row`を大きめに設定してください。
または、テンプレート自体に十分な行を用意してください。
