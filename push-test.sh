#!/bin/bash

# テスト環境へのpush自動化スクリプト

echo "🚀 テスト環境へのpushを開始します..."

# 1. テスト環境に切り替え
echo "📝 テスト環境に切り替え中..."
node switch-env.js test

if [ $? -ne 0 ]; then
    echo "❌ 環境切り替えに失敗しました"
    exit 1
fi

echo "✅ 環境切り替え完了"

# 2. テスト環境にpush
echo "📤 テスト環境のGASにpush中..."
clasp --project .clasp-test.json push

if [ $? -ne 0 ]; then
    echo "❌ pushに失敗しました"
    exit 1
fi

echo "✅ テスト環境へのpushが完了しました！" 