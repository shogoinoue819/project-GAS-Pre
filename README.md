# Google Apps Script プロジェクト

授業シフト管理システムの Google Apps Script プロジェクトです。

## 機能

- **日次シート作成**: 指定期間の日次シートを自動生成
- **スタッフシート作成**: スタッフ別の希望シフトシートを自動生成
- **希望シフト反映**: スタッフの希望を日次シートに反映
- **講師割り当て**: 授業に講師を自動割り当て
- **PDF エクスポート**: 日次シートを PDF 化して Google Drive に保存

## 環境管理

このプロジェクトは本番環境とテスト環境を分離して管理しています。

### 環境設定ファイル

| 環境       | clasp 設定         | 説明                                     |
| ---------- | ------------------ | ---------------------------------------- |
| 本番環境   | `.clasp.json`      | 本番用 Google Apps Script プロジェクト   |
| テスト環境 | `.clasp-test.json` | テスト用 Google Apps Script プロジェクト |

### 環境別 ID 管理

環境別の ID（スプレッドシート ID、フォルダ ID など）は`switch-env.js`で管理し、`const-env.js`に自動反映されます。

## セットアップ

### 1. 必要なファイルの作成

このプロジェクトを clone した後、以下のファイルを手動で作成する必要があります：

#### **switch-env.js の作成**

```javascript
#!/usr/bin/env node

const fs = require("fs");
const path = require("path");

const args = process.argv.slice(2);
const environment = args[0] || "test";

const ENV_CONFIG = {
  production: {
    SPREADSHEET_ID: "自分の本番用スプレッドシートID",
    FOLDER_ID: "自分の本番用フォルダID",
    PDF_FOLDER_ID: "自分の本番用PDFフォルダID",
  },
  test: {
    SPREADSHEET_ID: "自分のテスト用スプレッドシートID",
    FOLDER_ID: "自分のテスト用フォルダID",
    PDF_FOLDER_ID: "自分のテスト用PDFフォルダID",
  },
};

// env-constants.jsのパス
const envConstantsPath = path.join(__dirname, "const-env.js");

// ファイル内容を読み込み
function readEnvConstants() {
  try {
    return fs.readFileSync(envConstantsPath, "utf8");
  } catch (error) {
    console.error(
      "エラー: const-env.js を読み込めませんでした:",
      error.message
    );
    process.exit(1);
  }
}

// 環境に応じてIDを置換
function replaceIds(content, config) {
  content = content.replace(
    /const SPREADSHEET_ID = "[^"]*";/,
    `const SPREADSHEET_ID = "${config.SPREADSHEET_ID}";`
  );
  content = content.replace(
    /const FOLDER_ID = "[^"]*";/,
    `const FOLDER_ID = "${config.FOLDER_ID}";`
  );
  content = content.replace(
    /const PDF_FOLDER_ID = "[^"]*";/,
    `const PDF_FOLDER_ID = "${config.PDF_FOLDER_ID}";`
  );
  return content;
}

// ファイルに書き込み
function writeEnvConstants(content) {
  try {
    fs.writeFileSync(envConstantsPath, content, "utf8");
    console.log("const-env.js を更新しました");
  } catch (error) {
    console.error(
      "エラー: const-env.js の書き込みに失敗しました:",
      error.message
    );
    process.exit(1);
  }
}

// メイン処理
function main() {
  if (!ENV_CONFIG[environment]) {
    console.error(`エラー: 不明な環境 "${environment}"`);
    console.log("使用可能な環境: production, test");
    process.exit(1);
  }

  const config = ENV_CONFIG[environment];
  console.log(`環境設定:`, config);

  let content = readEnvConstants();
  content = replaceIds(content, config);
  writeEnvConstants(content);

  console.log(`${environment}環境への切り替えが完了しました`);
}

main();
```

#### **clasp 設定ファイルの作成**

**.clasp.json（本番環境）**:

```json
{
  "scriptId": "自分の本番用scriptId",
  "rootDir": "./"
}
```

**.clasp-test.json（テスト環境）**:

```json
{
  "scriptId": "自分のテスト用scriptId",
  "rootDir": "./"
}
```

### 2. 環境 ID の設定

作成した`switch-env.js`ファイルの`ENV_CONFIG`に環境別の ID を設定してください：

```javascript
const ENV_CONFIG = {
  production: {
    SPREADSHEET_ID: "本番用スプレッドシートID",
    FOLDER_ID: "本番用フォルダID",
    PDF_FOLDER_ID: "本番用PDF保存フォルダID",
  },
  test: {
    SPREADSHEET_ID: "テスト用スプレッドシートID",
    FOLDER_ID: "テスト用フォルダID",
    PDF_FOLDER_ID: "テスト用PDF保存フォルダID",
  },
};
```

### 3. clasp の認証

```bash
# claspにログイン
clasp login
```

## 使用方法

### テスト環境での開発

```bash
# 自動化スクリプトを使用（推奨）
./push-test.sh

# または手動で実行
node switch-env.js test
clasp --project .clasp-test.json push
```

### 本番環境へのデプロイ

```bash
# 自動化スクリプトを使用（推奨）
./push-production.sh

# または手動で実行
node switch-env.js production
clasp --project .clasp.json push
```

## ファイル構成

```
project-GAS-Pre/
├── GAS機能ファイル（push対象）
│   ├── appsscript.json          # GASプロジェクト設定
│   ├── const.js                 # アプリケーション定数
│   ├── const-env.js             # 環境別ID定数
│   ├── pdfExporter.js           # PDF化機能
│   ├── utils.js                 # ユーティリティ関数
│   ├── createMenu.js            # メニュー作成
│   ├── createDailySheets.js     # 日次シート作成
│   ├── createStaffSheets.js     # スタッフシート作成
│   ├── lessonManager.js         # 授業管理
│   ├── linkStaffList.js         # スタッフリスト連携
│   ├── reflectLessons.js        # 授業反映
│   ├── reflectWish.js           # 希望反映
│   └── teacherAssigner.js       # 講師割り当て
│
├── 開発用ファイル（push非対象）
│   ├── switch-env.js            # 環境切り替えスクリプト
│   ├── .clasp.json              # 本番環境clasp設定
│   ├── .clasp-test.json         # テスト環境clasp設定
│   ├── .claspignore             # push除外設定
│   ├── .gitignore               # Git除外設定
│   └── README.md                # ドキュメント
│
└── Git管理
    └── .git/                    # Gitリポジトリ
```

## 主要機能

### PDF エクスポート機能

```javascript
// 環境確認
checkEnvironmentVariables();

// 全日次シートをPDF化
exportAllDailySheetsAsPDF();

// 特定の日付シートをPDF化
exportSpecificDailySheetAsPDF("2025-01-15");
```

### 定数管理

- **`const.js`**: アプリケーション固有の定数（環境非依存）
- **`const-env.js`**: 環境別の ID 定数（自動生成）

## セキュリティ

| ファイル           | GitHub 公開 | 理由                             |
| ------------------ | ----------- | -------------------------------- |
| `switch-env.js`    | ❌          | ID が含まれるため非公開          |
| `const-env.js`     | ❌          | ID が含まれるため非公開          |
| `const.js`         | ✅          | アプリケーション定数のみ         |
| `.clasp.json`      | ❌          | 本番環境の scriptId が含まれる   |
| `.clasp-test.json` | ❌          | テスト環境の scriptId が含まれる |

## 注意事項

- **環境切り替え**: push 前に必ず`node switch-env.js [環境名]`を実行してください
- **ID 管理**: 環境別の ID は`switch-env.js`で一元管理します
- **セキュリティ**: 機密情報を含むファイルは`.gitignore`で除外されています
- **重要**: `const-env.js`は自動生成されるファイルのため、手動で編集しないでください

## トラブルシューティング

### clasp push エラー

```bash
# 認証の確認
clasp login

# プロジェクト設定の確認
clasp --project .clasp-test.json info
```

### 環境変数エラー

```javascript
// GASエディタで環境確認を実行
checkEnvironmentVariables();
```
