# Google Apps Script プロジェクト

## 環境設定

このプロジェクトは本番環境とテスト環境を分離して管理しています。

### 本番環境

- 設定ファイル: `.clasp.json`
- 環境変数: なし（直接 GAS 内で設定）

### テスト環境

- 設定ファイル: `.clasp-test.json`
- 環境変数: `.env.test`

## 開発環境の切り替え

### テスト環境で開発する場合

1. **clasp 設定を切り替え**

   ```bash
   # テスト環境用の設定に切り替え
   cp .clasp-test.json .clasp.json
   ```

2. **環境変数を設定**

   - `.env.test`ファイルの ID を実際のテスト用 ID に変更
   - `SPREADSHEET_ID`: テスト用スプレッドシート ID
   - `FOLDER_ID`: テスト用フォルダ ID

3. **GAS にデプロイ**
   ```bash
   clasp push
   ```

### 本番環境に戻す場合

1. **本番設定を復元**
   ```bash
   # 本番環境用の設定に戻す
   git checkout .clasp.json
   ```

## ファイル構成

```
project-GAS-Pre/
├── .clasp.json          # 本番環境設定
├── .clasp-test.json     # テスト環境設定
├── .env.test            # テスト環境変数
├── .gitignore           # Git除外設定
├── appsscript.json      # GAS設定
├── const.js             # 定数定義
├── utils.js             # 共通ユーティリティ
├── createMenu.js        # メニュー作成
├── createDailySheets.js # 日次シート作成
├── createStaffSheets.js # スタッフシート作成
├── linkStaffList.js     # スタッフ情報反映
├── reflectWish.js       # 希望シフト反映
├── lessonManager.js     # 講義管理
├── teacherAssigner.js   # 講師割り当て
└── reflectLessons.js    # 全日程授業反映
```

## 注意事項

- `.clasp.json`と`.env*`ファイルは`.gitignore`に含まれているため、Git で管理されません
- 本番環境の ID は直接`.clasp.json`に記載してください
- テスト環境の ID は`.env.test`に記載してください
