/**
 * Google Apps Scriptのカスタムメニューを作成する
 */
function createCustomMenu() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu("管理メニュー")
    .addItem("日次シート作成", "createDailySheets")
    .addItem("スタッフシート作成", "createStaffSheets")
    .addItem("スタッフ情報反映", "linkStaffList")
    .addItem("希望シフト反映", "reflectWish")
    .addItem("教科担当更新", "updatePrioritySheetFromStaffSheets")
    .addItem("全日程授業反映", "reflectLessons")
    .addToUi();
}

/**
 * スプレッドシートが開かれた時に自動的にメニューを作成する
 */
function onOpen() {
  createCustomMenu();
}
