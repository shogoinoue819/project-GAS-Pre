/**
 * メインシートのスタッフ表示名と背景色を日次シートテンプレートに反映する
 */
function linkStaffList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(MAIN);

  if (!mainSheet) {
    throw new Error(`メインシート「${MAIN}」が見つかりません。`);
  }

  const templateDailySheet = ss.getSheetByName(TEMPLATE_DAILY);

  if (!templateDailySheet) {
    throw new Error(
      `テンプレートシート「${TEMPLATE_DAILY}」が見つかりません。\nテンプレートシートを作成してから実行してください。`
    );
  }

  // 表示名と背景色をメインシートから取得
  const displayNames = mainSheet
    .getRange(
      MAIN_STAFF_START_ROW,
      MAIN_STAFF_DISPLAY_COL,
      MAIN_STAFF_END_ROW - MAIN_STAFF_START_ROW + 1,
      1
    )
    .getValues()
    .flat();

  const backgroundColors = mainSheet
    .getRange(
      MAIN_STAFF_START_ROW,
      MAIN_STAFF_DISPLAY_COL,
      MAIN_STAFF_END_ROW - MAIN_STAFF_START_ROW + 1,
      1
    )
    .getBackgrounds()
    .flat();

  // 既存の値・背景色をクリア
  const clearRange = templateDailySheet.getRange(
    DAILY_STAFF_ROW,
    DAILY_STAFF_START_COL,
    1,
    displayNames.length
  );
  clearRange.clearContent();
  clearRange.setBackground(null);

  // 表示名と背景色を反映
  for (let i = 0; i < displayNames.length; i++) {
    const cell = templateDailySheet.getRange(
      DAILY_STAFF_ROW,
      DAILY_STAFF_START_COL + i
    );
    cell.setValue(displayNames[i]);
    cell.setBackground(backgroundColors[i]);
  }

  Logger.log("Template_Dailyシートに表示名と背景色を反映しました！");
}
