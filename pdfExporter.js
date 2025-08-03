/**
 * PDFエクスポート機能
 * 日次シートをPDF化してGoogle Driveに保存する機能を提供
 */

/**
 * リトライ機能付きでシートをPDFにエクスポート
 * @param {Sheet} sheet - エクスポート対象のシート
 * @param {string} fileName - ファイル名
 * @param {number} maxRetries - 最大リトライ回数（デフォルト: 3）
 * @returns {Blob} PDFファイルのBlobオブジェクト
 */
function exportSheetAsPDFWithRetry(sheet, fileName, maxRetries = 3) {
  let lastError = null;

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      return exportSheetAsPDF(sheet, fileName);
    } catch (error) {
      lastError = error;

      // 429エラー（レート制限）の場合は待機してリトライ
      if (error.message.includes("429") && attempt < maxRetries) {
        const waitTime = attempt * 5000; // 5秒、10秒、15秒と増加
        Logger.log(
          `レート制限エラー (429) が発生しました。${waitTime}秒待機してリトライします... (試行 ${attempt}/${maxRetries})`
        );
        Utilities.sleep(waitTime);
        continue;
      }

      // その他のエラーまたは最後の試行の場合はエラーを投げる
      throw error;
    }
  }

  throw lastError;
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

    // PDFエクスポート用のURLを構築（A4サイズ、セル区切れ線なし、最大拡大）
    const pdfUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=pdf&gid=${gid}&portrait=true&size=A4&scale=1&top_margin=0.25&bottom_margin=0.25&left_margin=0.25&right_margin=0.25&gridlines=false&printnotes=false&pageorder=1&horizontal_alignment=CENTER&vertical_alignment=TOP&printtitle=false&sheetnames=false&fzr=false&fzc=false&attachment=false`;

    // PDFをダウンロード
    const response = UrlFetchApp.fetch(pdfUrl, {
      headers: {
        Authorization: `Bearer ${ScriptApp.getOAuthToken()}`,
      },
      muteHttpExceptions: true, // エラーレスポンスも取得
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
 * すべての日次シートを1つのPDFにまとめてエクスポートしてGoogle Driveに保存する
 * 1ページ1日程で、A4サイズに最適化されたPDFを作成
 * 注意: Google Apps Scriptの制限により、実際には個別のPDFファイルとして保存されます
 */
function exportAllDailySheetsAsCombinedPDF() {
  try {
    Logger.log("統合PDFエクスポート処理を開始します...");

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

    // 日程範囲を取得してフォルダ名を生成
    const dateRange = getDateRangeForFileName(dailySheetNames);
    const folderName = `授業シフト_${dateRange}`;

    // 専用フォルダを作成
    const subFolder = folder.createFolder(folderName);
    Logger.log(`専用フォルダを作成: ${folderName} (ID: ${subFolder.getId()})`);

    let successCount = 0;
    let errorCount = 0;

    // 各日次シートをPDF化して専用フォルダに保存
    dailySheetNames.forEach((sheetName, index) => {
      try {
        // レート制限対策：シート間でディレイを入れる
        if (index > 0) {
          Logger.log("レート制限対策のため2秒待機中...");
          Utilities.sleep(2000); // 2秒待機
        }

        const sheet = getSheetSafely(ss, sheetName);
        if (!sheet) {
          Logger.log(`シート "${sheetName}" が見つかりません`);
          errorCount++;
          return;
        }

        // 日付をyyyymmdd形式に変換
        const date = getDateFromDailySheetName(sheetName);
        const dateStr = date ? formatDateToYYYYMMDD(date) : sheetName;
        const fileName = `授業シフト_${dateStr}.pdf`;

        Logger.log(`シート "${sheetName}" をPDF化中...`);

        // PDFをエクスポート（リトライ機能付き）
        const pdfBlob = exportSheetAsPDFWithRetry(sheet, fileName);

        // 専用フォルダに保存
        const file = subFolder.createFile(pdfBlob);

        Logger.log(`PDF保存完了: ${fileName} (ID: ${file.getId()})`);
        successCount++;
      } catch (error) {
        logError(`シート "${sheetName}" のPDF化に失敗しました`, error);
        errorCount++;
      }
    });

    // 結果をログに出力
    Logger.log(`=== 統合PDFエクスポート完了 ===`);
    Logger.log(`成功: ${successCount}件`);
    Logger.log(`失敗: ${errorCount}件`);
    Logger.log(`合計: ${dailySheetNames.length}件のPDFを出力しました`);
    Logger.log(`保存先フォルダ: ${folderName} (ID: ${subFolder.getId()})`);

    // 成功メッセージを表示
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      "統合PDFエクスポート完了",
      `全${dailySheetNames.length}件の日次シートのPDFエクスポートが完了しました。\n保存先フォルダ: ${folderName}\n成功: ${successCount}件\n失敗: ${errorCount}件`,
      ui.ButtonSet.OK
    );
  } catch (error) {
    logError("統合PDFエクスポート処理でエラーが発生しました", error);
    throw error;
  }
}

/**
 * 日付をyyyymmdd形式に変換
 * @param {Date} date - 日付オブジェクト
 * @returns {string} yyyymmdd形式の文字列
 */
function formatDateToYYYYMMDD(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}${month}${day}`;
}

/**
 * 日程範囲からファイル名用の日付文字列を生成
 * @param {Array} dailySheetNames - 日次シート名の配列
 * @returns {string} ファイル名用の日付文字列（yyyymmdd形式、複数の場合は開始-終了）
 */
function getDateRangeForFileName(dailySheetNames) {
  if (dailySheetNames.length === 0) {
    return "";
  }

  // 日次シート名を日付オブジェクトに変換してソート
  const dates = dailySheetNames
    .map((name) => getDateFromDailySheetName(name))
    .filter((date) => date !== null)
    .sort((a, b) => a - b);

  if (dates.length === 0) {
    return "";
  }

  if (dates.length === 1) {
    // 1つの日程の場合
    return formatDateToYYYYMMDD(dates[0]);
  } else {
    // 複数の日程の場合、最初と最後を-で繋ぐ
    const startDate = formatDateToYYYYMMDD(dates[0]);
    const endDate = formatDateToYYYYMMDD(dates[dates.length - 1]);
    return `${startDate}-${endDate}`;
  }
}
