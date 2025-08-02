/**
 * メインシートの日程リストから日次シートを生成する
 */
function createDailySheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(MAIN);

  if (!mainSheet) {
    throw new Error(`メインシート「${MAIN}」が見つかりません。`);
  }

  const templateSheet = ss.getSheetByName(TEMPLATE_DAILY);

  if (!templateSheet) {
    throw new Error(
      `テンプレートシート「${TEMPLATE_DAILY}」が見つかりません。\nテンプレートシートを作成してから実行してください。`
    );
  }

  for (let row = MAIN_DATE_START_ROW; row <= MAIN_DATE_END_ROW; row++) {
    const dateValue = mainSheet.getRange(row, MAIN_DATE_COL).getValue();
    if (!dateValue) continue;

    // M/d 形式に整形（例：7/30）
    const sheetName = Utilities.formatDate(
      new Date(dateValue),
      Session.getScriptTimeZone(),
      "M/d"
    );

    // 既に同名シートがあれば削除（上書き）
    const existingSheet = ss.getSheetByName(sheetName);
    if (existingSheet) {
      ss.deleteSheet(existingSheet);
    }

    // テンプレートを複製して、名前を設定
    const newSheet = templateSheet.copyTo(ss);
    newSheet.setName(sheetName);

    // 日付セルに日付をセット（M/d形式で）
    newSheet.getRange(DAILY_DATE_ROW, DAILY_DATE_COL).setValue(sheetName);

    // 日程ごとに完了ログを出力
    Logger.log(`${sheetName} のシート生成が完了しました`);
  }

  Logger.log("日次シートの生成が完了しました！");
}
