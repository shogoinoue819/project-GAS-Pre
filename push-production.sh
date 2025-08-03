#!/bin/bash

# 本番環境へのpush自動化スクリプト

echo "🚀 本番環境へのpushを開始します..."

# 確認プロンプト
read -p "⚠️  本番環境にpushしますか？ (y/N): " -n 1 -r
echo
if [[ ! $REPLY =~ ^[Yy]$ ]]; then
    echo "❌ pushをキャンセルしました"
    exit 1
fi

# 1. 本番環境に切り替え
echo "📝 本番環境に切り替え中..."
node switch-env.js production

if [ $? -ne 0 ]; then
    echo "❌ 環境切り替えに失敗しました"
    exit 1
fi

echo "✅ 環境切り替え完了"

# 2. 本番環境にpush
echo "📤 本番環境のGASにpush中..."
clasp --project .clasp.json push

if [ $? -ne 0 ]; then
    echo "❌ pushに失敗しました"
    exit 1
fi

echo "✅ 本番環境へのpushが完了しました！" 