/**
 * PDFエクスポート機能
 * 日次シートをPDF化してGoogle Driveに保存する機能を提供
 */

/**
 * すべての日次シートをPDFにエクスポートしてGoogle Driveに保存する
 * スプレッドシート内のシートのうち、「M/d」の形式の日付を名前に持つシートのみを対象とする
 */
function exportAllDailySheetsAsPDF() {
  try {
    Logger.log("PDFエクスポート処理を開始します...");

    // スプレッドシートを取得
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 日次シート（M/d形式）を抽出
    const dailySheetNames = getDailySheetNames(ss);

    if (dailySheetNames.length === 0) {
      Logger.log("対象となる日次シートが見つかりませんでした。");
      return;
    }

    Logger.log(`${dailySheetNames.length}件の日次シートが見つかりました。`);

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
    dailySheetNames.forEach((sheetName) => {
      try {
        const sheet = getSheetSafely(ss, sheetName);
        if (!sheet) {
          Logger.log(`シート "${sheetName}" が見つかりません`);
          errorCount++;
          return;
        }

        const fileName = `授業シフト_${sheetName}_${timestamp}.pdf`;

        Logger.log(`シート "${sheetName}" をPDF化中...`);

        // PDFをエクスポート
        const pdfBlob = exportSheetAsPDF(sheet, fileName);

        // Google Driveに保存
        const file = folder.createFile(pdfBlob);

        Logger.log(`PDF保存完了: ${fileName} (ID: ${file.getId()})`);
        successCount++;
      } catch (error) {
        logError(`シート "${sheetName}" のPDF化に失敗しました`, error);
        errorCount++;
      }
    });

    // 結果をログに出力
    Logger.log(`=== PDFエクスポート完了 ===`);
    Logger.log(`成功: ${successCount}件`);
    Logger.log(`失敗: ${errorCount}件`);
    Logger.log(`合計: ${dailySheetNames.length}件のPDFを出力しました`);
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
    const pdfUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=pdf&gid=${gid}&portrait=true&size=A4`;

    // PDFをダウンロード
    const response = UrlFetchApp.fetch(pdfUrl, {
      headers: {
        Authorization: `Bearer ${ScriptApp.getOAuthToken()}`,
      },
    });

    if (response.getResponseCode() !== 200) {
      throw new Error(
        `PDFエクスポートに失敗しました。レスポンスコード: ${response.getResponseCode()}`
      );
    }

    // Blobオブジェクトを作成
    const blob = response.getBlob().setName(fileName);

    return blob;
  } catch (error) {
    logError(
      `シート "${sheet.getName()}" のPDFエクスポートでエラーが発生しました`,
      error
    );
    throw error;
  }
}

/**
 * 特定の日付の日次シートをPDFにエクスポート
 * @param {string} dateString - 日付文字列（M/d形式）
 */
function exportSpecificDailySheetAsPDF(dateString) {
  try {
    if (!isDailySheetName(dateString)) {
      throw new Error(
        `不正な日付形式です: ${dateString}。M/d形式で指定してください。`
      );
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getSheetSafely(ss, dateString);

    if (!sheet) {
      throw new Error(`日次シート "${dateString}" が見つかりません。`);
    }

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

    const fileName = `授業シフト_${dateString}_${timestamp}.pdf`;

    Logger.log(`シート "${dateString}" をPDF化中...`);

    // PDFをエクスポート
    const pdfBlob = exportSheetAsPDF(sheet, fileName);

    // Google Driveに保存
    const file = folder.createFile(pdfBlob);

    Logger.log(`PDF保存完了: ${fileName} (ID: ${file.getId()})`);

    // 成功メッセージを表示
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      "PDFエクスポート完了",
      `日次シート "${dateString}" のPDFエクスポートが完了しました。\nファイル名: ${fileName}`,
      ui.ButtonSet.OK
    );
  } catch (error) {
    logError(`特定日次シートのPDFエクスポートでエラーが発生しました`, error);

    // エラーメッセージを表示
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      "PDFエクスポートエラー",
      `PDFエクスポートに失敗しました。\nエラー: ${error.message}`,
      ui.ButtonSet.OK
    );

    throw error;
  }
}

/**
 * 環境変数の設定状況をチェック
 */
function checkEnvironmentVariables() {
  Logger.log("=== 環境変数チェック ===");
  Logger.log(`SPREADSHEET_ID: ${SPREADSHEET_ID ? "設定済み" : "未設定"}`);
  Logger.log(`FOLDER_ID: ${FOLDER_ID ? "設定済み" : "未設定"}`);
  Logger.log(`PDF_FOLDER_ID: ${PDF_FOLDER_ID ? "設定済み" : "未設定"}`);

  if (!SPREADSHEET_ID || !FOLDER_ID || !PDF_FOLDER_ID) {
    Logger.log("警告: 一部の環境変数が未設定です。");
  } else {
    Logger.log("全ての環境変数が正常に設定されています。");
  }
}
