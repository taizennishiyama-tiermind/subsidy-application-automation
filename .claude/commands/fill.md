# テンプレートへの書き込み実行

申請書データをExcel/Wordテンプレートに自動で書き込みます。

## 引数

$ARGUMENTS = 案件名（省略時はユーザーに確認する）

## 前提条件

- `workspace/{案件名}/draft_application.md` が作成済みであること
- `workspace/{案件名}/application_data.json` が存在すること（なければ Step 2 で生成）
- 対象補助金のテンプレートが `grants/{補助金名}/templates/` に配置済みであること
- セルマッピングが `/add-template` で設定済みであること

## 手順

### Step 1: データの確認

1. `workspace/{案件名}/strategy.md` から対象補助金名を取得する
2. `grants/{補助金名}/templates/` にテンプレートファイルがあるか確認する
3. `workspace/{案件名}/application_data.json` を読み込む（存在する場合）

### Step 2: application_data.json の生成（存在しない場合）

`workspace/{案件名}/draft_application.md` と `workspace/{案件名}/company_data.md` から、
テンプレートの各フィールドに対応するデータをJSON形式で抽出・整形する。

```json
{
  "company_name": "株式会社XXX",
  "representative": "山田太郎",
  "business_overview": "...",
  "challenge_1": "...",
  "sales_year0": "85000",
  "sales_year1": "90000"
}
```

JSONのキーは `grants/{補助金名}/scripts/fill_template.py` のセルマッピングに対応させる。

### Step 3: スクリプト実行

```bash
.venv/bin/python "grants/{補助金名}/scripts/fill_template.py" fill \
  --template "grants/{補助金名}/templates/{テンプレートファイル}" \
  --data "workspace/{案件名}/application_data.json" \
  --output "workspace/{案件名}/申請書_{補助金名}.xlsx"
```

複数のテンプレートファイルがある場合は、それぞれに対して実行する。

### Step 4: 確認

1. 出力ファイルのパスをユーザーに伝える
2. 書き込まれた項目数を報告する
3. 内容の確認を依頼する

## 完了メッセージ

「テンプレートへの書き込みが完了しました。出力ファイル:
- `workspace/{案件名}/申請書_{補助金名}.xlsx`
内容を確認し、必要に応じて手動で微調整してください」

## エラー時

- テンプレートが見つからない → `grants/{補助金名}/templates/` への配置と `/add-template` の実施を案内
- セルマッピングが未設定 → `/add-template {補助金名}` の実施を案内
- Python環境がない → `bash scripts/setup.sh` の実行を案内
- データが不足 → 不足項目を表示し `/draft` での追記を案内
