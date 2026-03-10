#!/bin/bash
# 補助金申請自動化環境 セットアップスクリプト
# Usage: bash scripts/setup.sh

set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PROJECT_DIR="$(dirname "$SCRIPT_DIR")"

echo "=== 補助金申請自動化環境 セットアップ ==="

# Python仮想環境の作成
if [ ! -d "$PROJECT_DIR/.venv" ]; then
  echo "Python仮想環境を作成中..."
  python3 -m venv "$PROJECT_DIR/.venv"
fi

# 依存パッケージのインストール
echo "依存パッケージをインストール中..."
"$PROJECT_DIR/.venv/bin/pip" install -r "$SCRIPT_DIR/requirements.txt" --quiet

echo "=== セットアップ完了 ==="
echo "仮想環境: $PROJECT_DIR/.venv"
echo ""
echo "使い方:"
echo "  claude  # Claude Codeを起動して /new-case で案件開始"
