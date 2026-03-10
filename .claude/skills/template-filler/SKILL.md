# テンプレート書き込みスキル

## トリガー

`/fill` または `/add-template` コマンド実行時に自動で参照される。
Excel/Word テンプレートへの自動書き込み機能の操作方法。

---

## 仕組み

各補助金フォルダに以下がセットになっている:

```
grants/{補助金名}/
├── templates/              ← 公式テンプレート（Excel/Word）を配置
│   └── README.md           ← 配置手順
└── scripts/
    └── fill_template.py    ← 書き込みスクリプト
```

スクリプトは共通ユーティリティ（`scripts/utils/`）を使用する:
- `excel_handler.py` - Excel操作（openpyxl）
- `word_handler.py` - Word操作（python-docx）

---

## テンプレート分析（/add-template で使用）

### Excelファイルの分析

```bash
.venv/bin/python "grants/{補助金名}/scripts/fill_template.py" analyze \
  --template "grants/{補助金名}/templates/{ファイル名}.xlsx"
```

出力される情報:
- シート名の一覧
- 各シートのセル範囲
- 主要セルの内容（ラベル・見出し）

### Wordファイルの分析

同じ `analyze` コマンドで実行:
- 見出し構成
- テーブル構造
- プレースホルダー（`{{...}}`形式）の一覧

---

## セルマッピングの設定

`fill_template.py` 内の2つの辞書を設定する:

### EXCEL_CELL_MAPPING

```python
EXCEL_CELL_MAPPING = {
    "シート名": {
        "company_name": "B3",      # 会社名
        "representative": "B4",     # 代表者名
        "sales_year0": "D10",       # 直近期売上
    }
}
```

### WORD_PLACEHOLDER_MAPPING

```python
WORD_PLACEHOLDER_MAPPING = {
    "company_name": "{{会社名}}",
    "business_overview": "{{事業概要}}",
}
```

### ルール
- データキーは英語のスネークケースにする
- 必須項目と任意項目を区別するコメントを付ける
- テンプレートの特殊なセル（結合セル等）には注意書きを付ける

---

## テンプレートへの書き込み（/fill で使用）

### 入力データ

`workspace/{案件名}/application_data.json`:

```json
{
  "company_name": "株式会社XXX",
  "representative": "山田太郎",
  "business_overview": "...",
  "sales_year0": "85000",
  "sales_year1": "90000"
}
```

### 実行コマンド

```bash
.venv/bin/python "grants/{補助金名}/scripts/fill_template.py" fill \
  --template "grants/{補助金名}/templates/{テンプレート}.xlsx" \
  --data "workspace/{案件名}/application_data.json" \
  --output "workspace/{案件名}/申請書_{補助金名}.xlsx"
```

### 出力
- 書き込み済みのExcel/Wordファイルが `workspace/{案件名}/` に出力される
- 複数テンプレートがある場合はそれぞれに対して実行する

---

## Python環境のセットアップ

初回のみ実行:

```bash
bash scripts/setup.sh
```

依存パッケージ:
- `openpyxl` >= 3.1.0（Excel操作）
- `python-docx` >= 1.1.0（Word操作）

実行時は必ず `.venv/bin/python` を使用する。

---

## トラブルシューティング

| 症状 | 対処 |
|------|------|
| スクリプトが実行できない | `bash scripts/setup.sh` でPython環境をセットアップ |
| テンプレートが見つからない | `grants/{補助金名}/templates/` にファイルを配置 |
| 書き込み位置がずれる | `/add-template` を再実行してマッピングを修正 |
| PDFが読めない | テキスト抽出可能なPDFか確認 |
