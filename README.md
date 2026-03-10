# 補助金申請書作成支援 AI 環境

Claude Code を使って、補助金申請書の作成を AI が支援する環境です。

---

## できること

- AI がプロの視点でクライアントにヒアリング（逆質問）
- 決算書 PDF の読み取り・財務分析
- 最適な補助金のマッチング提案
- 採択率の高い申請書の自動生成
- Excel/Word テンプレートへの自動書き込み
- 審査基準に基づくセルフレビュー

## 対応補助金

1. ものづくり補助金
2. IT導入補助金 / デジタル化・AI導入補助金
3. 新事業進出補助金（旧: 事業再構築補助金）
4. 小規模事業者持続化補助金

`/add-grant` コマンドでいつでも追加可能。

---

## セットアップ

### 1. Claude Code をインストール

```bash
npm install -g @anthropic-ai/claude-code
```

### 2. Python 環境をセットアップ

```bash
cd /path/to/補助金申請自動化環境
bash scripts/setup.sh
```

### 3. 起動

```bash
cd /path/to/補助金申請自動化環境
claude
```

---

## 使い方

### 全体の流れ

```
/new-case → /hearing → /analyze → /match → /draft → /review → /fill
 案件開始    ヒアリング   財務分析   補助金選定  申請書作成   レビュー   テンプレート書込
```

### コマンド一覧

| コマンド | 説明 |
|---------|------|
| `/new-case` | 新規案件を開始してヒアリング開始 |
| `/hearing {案件名}` | ヒアリングを再開する |
| `/analyze {案件名}` | 決算書等の財務データを分析 |
| `/match {案件名}` | 最適な補助金を提案 |
| `/draft {案件名}` | 申請書のドラフトを作成 |
| `/review {案件名}` | 審査基準でセルフレビュー |
| `/fill {案件名}` | Excel/Word テンプレートに自動書き込み |
| `/status` | 全案件の進捗を表示 |
| `/add-grant {補助金名}` | 新しい補助金を追加 |
| `/add-template {補助金名}` | テンプレートを解析・登録 |

### クイックスタート

1. `claude` で起動
2. `/new-case` → 案件名を入力
3. AI の質問に答えていく（対面でクライアントと一緒に使うのがおすすめ）
4. ヒアリング完了後 `/analyze` → `/match` → `/draft` → `/review` と順に進める
5. `/fill` で Excel/Word テンプレートに自動書き込み

---

## フォルダ構成

```
補助金申請自動化環境/
│
├── ノウハウ/                      ← ★ 自社独自のノウハウをここに追加
│
├── grants/                       ← 補助金の情報はここ
│   └── {補助金名}/
│       ├── overview.md           ← 制度概要・要件・審査基準
│       ├── templates/            ← ★ 公式テンプレート（Excel/Word）をここに配置
│       └── scripts/
│           └── fill_template.py  ← テンプレート書き込みスクリプト
│
├── workspace/                    ← 案件データはここに自動生成される
│   └── {案件名}/
│       ├── hearing_notes.md      ← ヒアリング記録
│       ├── company_data.md       ← 会社・財務データ
│       ├── strategy.md           ← 補助金選定・戦略
│       ├── draft_application.md  ← 申請書ドラフト
│       └── application_data.json ← テンプレート書き込み用データ
│
├── scripts/                      ← セットアップ・共通ツール
│
├── .claude/                      ← AI の設定（通常触らない）
│   ├── commands/                 ← スラッシュコマンド定義
│   └── skills/                   ← AI の専門知識
│       ├── hearing-expert/       ← ヒアリング手法
│       ├── storytelling/         ← ストーリー構成パターン
│       ├── financial-analysis/   ← 財務分析ガイド
│       ├── review-criteria/      ← 審査基準・採択事例
│       └── template-filler/      ← テンプレート書き込み操作
│
├── CLAUDE.md                     ← AI への指示書（通常触らない）
└── README.md                     ← このファイル
```

---

## テンプレートの管理

### テンプレートの配置

1. 補助金の公式サイトからテンプレート（Excel/Word）をダウンロード
2. `grants/{補助金名}/templates/` フォルダに配置
3. `/add-template {補助金名}` を実行 → AI がテンプレート構造を解析

### テンプレートが更新された場合

1. 新しいテンプレートを `grants/{補助金名}/templates/` に上書き配置
2. `/add-template {補助金名}` を再実行

---

## よくある質問

**Q: テンプレートがなくても使える？**
A: はい。`/draft` まではテンプレートなしで動作します。Markdown の申請書ドラフトが生成されます。

**Q: 途中で中断できる？**
A: はい。Claude Code を終了しても `workspace/{案件名}/` にデータが保存されます。`/hearing {案件名}` で再開できます。

**Q: 新しい補助金に対応したい**
A: `/add-grant {補助金名}` で簡単に追加できます。

---

## トラブルシューティング

| 症状 | 対処 |
|------|------|
| `/fill` でエラー | `bash scripts/setup.sh` を実行 |
| 書き込み位置がずれる | `/add-template {補助金名}` を再実行 |
| PDF が読めない | テキスト抽出可能な PDF か確認 |
| Claude Code が起動しない | `node --version` と `claude --version` を確認 |

---

## 必要なもの

- Mac または Windows PC
- Node.js v18 以上
- Python 3.10 以上
- Claude Code サブスクリプション（月額 $20〜）
