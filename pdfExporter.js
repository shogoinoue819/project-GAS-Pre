/**
 * PDFエクスポート機能
 * 日次シートをPDF化してGoogle Driveに保存する機能を提供
 */

// 環境変数初期設定関数は削除（env-constants.jsで直接定数を使用）

/**
 * すべての日次シートをPDFにエクスポートしてGoogle Driveに保存する
 * スプレッドシート内のシートのうち、「yyyy-mm-dd」の形式の日付を名前に持つシートのみを対象とする
 */
function exportAllDailySheetsAsPDF() {
  try {
    Logger.log("PDFエクスポート処理を開始します...");

    // スプレッドシートを取得
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();

    // 日次シート（yyyy-mm-dd形式）を抽出
    const dailySheets = sheets.filter((sheet) => {
      const sheetName = sheet.getName();
      return isYYYYMMDD(sheetName);
    });

    if (dailySheets.length === 0) {
      Logger.log("対象となる日次シートが見つかりませんでした。");
      return;
    }

    Logger.log(`${dailySheets.length}件の日次シートが見つかりました。`);

    // 保存先フォルダIDを取得
    const folderId = PDF_FOLDER_ID;
    if (!folderId) {
      throw new Error("PDF_FOLDER_IDが設定されていません。");
    }

    // 保存先フォルダを取得
    const folder = DriveApp.getFolderById(folderId);
    if (!folder) {
      throw new Error(`指定されたフォルダID (${folderId}) が見つかりません。`);
    }

    // 現在の日時を取得（ファイル名用）
    const now = new Date();
    const timestamp = Utilities.formatDate(
      now,
      Session.getScriptTimeZone(),
      "yyyy-MM-dd_HH-mm"
    );

    let successCount = 0;
    let errorCount = 0;

    // 各日次シートをPDF化
    dailySheets.forEach((sheet) => {
      try {
        const sheetName = sheet.getName();
        const fileName = `授業シフト_${sheetName}_${timestamp}.pdf`;

        Logger.log(`シート "${sheetName}" をPDF化中...`);

        // PDFをエクスポート
        const pdfBlob = exportSheetAsPDF(sheet, fileName);

        // Google Driveに保存
        const file = folder.createFile(pdfBlob);

        Logger.log(`PDF保存完了: ${fileName} (ID: ${file.getId()})`);
        successCount++;
      } catch (error) {
        logError(`シート "${sheet.getName()}" のPDF化に失敗しました`, error);
        errorCount++;
      }
    });

    // 結果をログに出力
    Logger.log(`=== PDFエクスポート完了 ===`);
    Logger.log(`成功: ${successCount}件`);
    Logger.log(`失敗: ${errorCount}件`);
    Logger.log(`合計: ${dailySheets.length}件のPDFを出力しました`);
  } catch (error) {
    logError("PDFエクスポート処理でエラーが発生しました", error);
    throw error;
  }
}

/**
 * 指定されたシートをPDFにエクスポートする
 * @param {Sheet} sheet - エクスポート対象のシート
 * @param {string} fileName - ファイル名
 * @returns {Blob} PDFファイルのBlobオブジェクト
 */
function exportSheetAsPDF(sheet, fileName) {
  try {
    // シートのURLを取得
    const sheetUrl = sheet.getParent().getUrl();
    const sheetId = sheet.getParent().getId();
    const gid = sheet.getSheetId();

    // PDFエクスポート用のURLを構築
    const pdfUrl = `${sheetUrl}/export?format=pdf&gid=${gid}&portrait=false&size=A4&fzr=true&gridlines=false&printtitle=false&sheetnames=false&pagenum=false&horizontal_alignment=CENTER&vertical_alignment=TOP&top_margin=0.5&bottom_margin=0.5&left_margin=0.5&right_margin=0.5`;

    // OAuth2トークンを取得
    const token = ScriptApp.getOAuthToken();

    // HTTPリクエストのオプション
    const options = {
      method: "GET",
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/pdf",
      },
      muteHttpExceptions: true,
    };

    // PDFをダウンロード
    const response = UrlFetchApp.fetch(pdfUrl, options);

    if (response.getResponseCode() !== 200) {
      throw new Error(
        `PDFエクスポートに失敗しました。HTTPステータス: ${response.getResponseCode()}`
      );
    }

    // Blobを作成
    const blob = response.getBlob();
    blob.setName(fileName);

    return blob;
  } catch (error) {
    logError(
      `シート "${sheet.getName()}" のPDFエクスポートに失敗しました`,
      error
    );
    throw error;
  }
}

/**
 * 特定の日付のシートのみをPDF化する（テスト用）
 * @param {string} dateString - 日付文字列（yyyy-mm-dd形式）
 */
function exportSpecificDailySheetAsPDF(dateString) {
  try {
    if (!isYYYYMMDD(dateString)) {
      throw new Error("日付はyyyy-mm-dd形式で指定してください。");
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(dateString);

    if (!sheet) {
      throw new Error(`シート "${dateString}" が見つかりません。`);
    }

    Logger.log(`シート "${dateString}" をPDF化します...`);

    // 保存先フォルダIDを取得
    const folderId = PDF_FOLDER_ID;
    if (!folderId) {
      throw new Error("PDF_FOLDER_IDが設定されていません。");
    }

    // 保存先フォルダを取得
    const folder = DriveApp.getFolderById(folderId);
    if (!folder) {
      throw new Error(`指定されたフォルダID (${folderId}) が見つかりません。`);
    }

    // 現在の日時を取得
    const now = new Date();
    const timestamp = Utilities.formatDate(
      now,
      Session.getScriptTimeZone(),
      "yyyy-MM-dd_HH-mm"
    );
    const fileName = `授業シフト_${dateString}_${timestamp}.pdf`;

    // PDFをエクスポート
    const pdfBlob = exportSheetAsPDF(sheet, fileName);

    // Google Driveに保存
    const file = folder.createFile(pdfBlob);

    Logger.log(`PDF保存完了: ${fileName} (ID: ${file.getId()})`);
    Logger.log(`ファイルURL: ${file.getUrl()}`);
  } catch (error) {
    logError("特定シートのPDFエクスポートでエラーが発生しました", error);
    throw error;
  }
}

/**
 * 現在の環境変数設定を確認する（デバッグ用）
 */
function checkEnvironmentVariables() {
  try {
    Logger.log("=== 環境変数設定確認 ===");
    Logger.log(`SPREADSHEET_ID: ${SPREADSHEET_ID}`);
    Logger.log(`FOLDER_ID: ${FOLDER_ID}`);
    Logger.log(`PDF_FOLDER_ID: ${PDF_FOLDER_ID}`);

    if (PDF_FOLDER_ID) {
      try {
        const folder = DriveApp.getFolderById(PDF_FOLDER_ID);
        Logger.log(`PDFフォルダ名: ${folder.getName()}`);
        Logger.log(`PDFフォルダURL: ${folder.getUrl()}`);
      } catch (error) {
        Logger.log(`PDFフォルダアクセスエラー: ${error.message}`);
      }
    }
  } catch (error) {
    logError("環境変数確認でエラーが発生しました", error);
  }
}
