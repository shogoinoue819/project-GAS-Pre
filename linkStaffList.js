function linkStaffList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName("Main");
  const templateDateSheet = ss.getSheetByName("Template_Date");

  // 表示名と背景色をMainシートから取得
  const names = mainSheet
    .getRange(MAIN_STAFF_START_ROW, MAIN_STAFF_DISPLAY_COL, MAIN_STAFF_END_ROW - MAIN_STAFF_START_ROW + 1, 1)
    .getValues()
    .flat();

  const bgColors = mainSheet
    .getRange(MAIN_STAFF_START_ROW, MAIN_STAFF_DISPLAY_COL, MAIN_STAFF_END_ROW - MAIN_STAFF_START_ROW + 1, 1)
    .getBackgrounds()
    .flat();

  // 既存の値・背景色をクリア
  const clearRange = templateDateSheet.getRange(DATE_NAME_ROW, DATE_NAME_START_COL, 1, names.length);
  clearRange.clearContent();
  clearRange.setBackground(null);

  // 表示名と背景色を反映
  for (let i = 0; i < names.length; i++) {
    const cell = templateDateSheet.getRange(DATE_NAME_ROW, DATE_NAME_START_COL + i);
    cell.setValue(names[i]);
    cell.setBackground(bgColors[i]);
  }

  Logger.log("Template_Dateシートに表示名と背景色を反映しました！");
}
