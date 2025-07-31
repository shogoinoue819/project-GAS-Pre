function createStaffSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName("Main");
  const templateSheet = ss.getSheetByName("Template_Staff");

  for (let row = MAIN_STAFF_START_ROW; row <= MAIN_STAFF_END_ROW; row++) {
    const name = mainSheet.getRange(row, MAIN_STAFF_NAME_COL).getValue();
    if (!name) continue;

    // 既存シートがあれば削除
    const existingSheet = ss.getSheetByName(name);
    if (existingSheet) {
      ss.deleteSheet(existingSheet);
    }

    // テンプレートを複製＆リネーム
    const newSheet = templateSheet.copyTo(ss);
    newSheet.setName(name);

    // A2セルに名前をセット
    newSheet.getRange("B1").setValue(name);
  }

  Logger.log("スタッフシートの生成が完了しました！");
}
