/**
 * メインシートのスタッフリストからスタッフシートを生成する
 */
function createStaffSheets() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mainSheet = ss.getSheetByName(MAIN);

    if (!mainSheet) {
      throw new Error(`メインシート「${MAIN}」が見つかりません。`);
    }

    const templateSheet = ss.getSheetByName(TEMPLATE_STAFF);

    if (!templateSheet) {
      throw new Error(
        `テンプレートシート「${TEMPLATE_STAFF}」が見つかりません。\nテンプレートシートを作成してから実行してください。`
      );
    }

    for (let row = MAIN_STAFF_START_ROW; row <= MAIN_STAFF_END_ROW; row++) {
      const staffName = mainSheet.getRange(row, MAIN_STAFF_NAME_COL).getValue();
      if (!staffName) continue;

      // 既存シートがあれば削除
      const existingSheet = ss.getSheetByName(staffName);
      if (existingSheet) {
        ss.deleteSheet(existingSheet);
      }

      // テンプレートを複製＆リネーム
      const newSheet = templateSheet.copyTo(ss);
      newSheet.setName(staffName);

      // スタッフ名セルに名前をセット
      newSheet.getRange(STAFF_NAME_ROW, STAFF_NAME_COL).setValue(staffName);
    }

    Logger.log("スタッフシートの生成が完了しました！");
  } catch (error) {
    Logger.log(`エラーが発生しました: ${error.message}`);
    throw error;
  }
}
