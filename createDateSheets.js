function createDateSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName("Main");
  const templateSheet = ss.getSheetByName("Template_Date");

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

    // A1セルに日付をセット（M/d形式で）
    newSheet.getRange("A1").setValue(sheetName);

    // 日程ごとに完了ログを出力
    Logger.log(`${sheetName} のシート生成が完了しました`);
  }

  Logger.log("日程シートの生成が完了しました！");
}
