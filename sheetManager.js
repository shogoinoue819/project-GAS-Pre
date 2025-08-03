/**
 * シート管理機能
 * シートの作成、削除、操作に関する共通処理を集約
 */

/**
 * 日次シートを一括作成
 */
function createDailySheets() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 必須シートの存在チェック
    const validation = validateRequiredSheets(ss, [
      SHEET_NAMES.MAIN,
      SHEET_NAMES.TEMPLATE_DAILY,
    ]);

    if (!validation.success) {
      throw new Error(
        `必須シートが見つかりません: ${validation.missing.join(", ")}`
      );
    }

    const mainSheet = getSheetSafely(ss, SHEET_NAMES.MAIN);
    const templateSheet = getSheetSafely(ss, SHEET_NAMES.TEMPLATE_DAILY);

    // 日程リストから日次シートを生成
    for (
      let row = MAIN_SHEET.DATE.START_ROW;
      row <= MAIN_SHEET.DATE.END_ROW;
      row++
    ) {
      const dateValue = mainSheet.getRange(row, MAIN_SHEET.DATE.COL).getValue();
      if (!dateValue) continue;

      // 日次シート名を生成
      const sheetName = generateDailySheetName(dateValue);
      if (!sheetName) continue;

      // 既に同名シートがあれば削除（上書き）
      const existingSheet = getSheetSafely(ss, sheetName);
      if (existingSheet) {
        ss.deleteSheet(existingSheet);
      }

      // テンプレートを複製して、名前を設定
      const newSheet = templateSheet.copyTo(ss);
      newSheet.setName(sheetName);

      // 日付セルに日付をセット（M/d形式で）
      newSheet
        .getRange(DAILY_SHEET.DATE_ROW, DAILY_SHEET.DATE_COL)
        .setValue(sheetName);

      Logger.log(`${sheetName} のシート生成が完了しました`);
    }

    Logger.log("日次シートの生成が完了しました！");
  } catch (error) {
    logError("日次シート作成でエラーが発生しました", error);
    throw error;
  }
}

/**
 * スタッフシートを一括作成
 */
function createStaffSheets() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 必須シートの存在チェック
    const validation = validateRequiredSheets(ss, [
      SHEET_NAMES.MAIN,
      SHEET_NAMES.TEMPLATE_STAFF,
    ]);

    if (!validation.success) {
      throw new Error(
        `必須シートが見つかりません: ${validation.missing.join(", ")}`
      );
    }

    const mainSheet = getSheetSafely(ss, SHEET_NAMES.MAIN);
    const templateSheet = getSheetSafely(ss, SHEET_NAMES.TEMPLATE_STAFF);

    // スタッフリストからスタッフシートを生成
    for (
      let row = MAIN_SHEET.STAFF.START_ROW;
      row <= MAIN_SHEET.STAFF.END_ROW;
      row++
    ) {
      const staffName = mainSheet
        .getRange(row, MAIN_SHEET.STAFF.NAME_COL)
        .getValue();
      if (!staffName) continue;

      // 既存シートがあれば削除
      const existingSheet = getSheetSafely(ss, staffName);
      if (existingSheet) {
        ss.deleteSheet(existingSheet);
      }

      // テンプレートを複製＆リネーム
      const newSheet = templateSheet.copyTo(ss);
      newSheet.setName(staffName);

      // スタッフ名セルに名前をセット
      newSheet
        .getRange(STAFF_SHEET.NAME_ROW, STAFF_SHEET.NAME_COL)
        .setValue(staffName);
    }

    Logger.log("スタッフシートの生成が完了しました！");
  } catch (error) {
    logError("スタッフシート作成でエラーが発生しました", error);
    throw error;
  }
}

/**
 * メインシートのスタッフ表示名と背景色を日次シートテンプレートに反映
 */
function linkStaffList() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 必須シートの存在チェック
    const validation = validateRequiredSheets(ss, [
      SHEET_NAMES.MAIN,
      SHEET_NAMES.TEMPLATE_DAILY,
    ]);

    if (!validation.success) {
      throw new Error(
        `必須シートが見つかりません: ${validation.missing.join(", ")}`
      );
    }

    const mainSheet = getSheetSafely(ss, SHEET_NAMES.MAIN);
    const templateDailySheet = getSheetSafely(ss, SHEET_NAMES.TEMPLATE_DAILY);

    // 表示名と背景色をメインシートから取得
    const displayNames = mainSheet
      .getRange(
        MAIN_SHEET.STAFF.START_ROW,
        MAIN_SHEET.STAFF.DISPLAY_COL,
        MAIN_SHEET.STAFF.END_ROW - MAIN_SHEET.STAFF.START_ROW + 1,
        1
      )
      .getValues()
      .flat();

    const backgroundColors = mainSheet
      .getRange(
        MAIN_SHEET.STAFF.START_ROW,
        MAIN_SHEET.STAFF.DISPLAY_COL,
        MAIN_SHEET.STAFF.END_ROW - MAIN_SHEET.STAFF.START_ROW + 1,
        1
      )
      .getBackgrounds()
      .flat();

    // 既存の値・背景色をクリア
    const clearRange = templateDailySheet.getRange(
      DAILY_SHEET.STAFF_ROW,
      DAILY_SHEET.STAFF_START_COL,
      1,
      displayNames.length
    );
    clearRange.clearContent();
    clearRange.setBackground(null);

    // 表示名と背景色を反映
    for (let i = 0; i < displayNames.length; i++) {
      const cell = templateDailySheet.getRange(
        DAILY_SHEET.STAFF_ROW,
        DAILY_SHEET.STAFF_START_COL + i
      );
      cell.setValue(displayNames[i]);
      cell.setBackground(backgroundColors[i]);
    }

    Logger.log("Template_Dailyシートに表示名と背景色を反映しました！");
  } catch (error) {
    logError("スタッフ情報反映でエラーが発生しました", error);
    throw error;
  }
}

/**
 * スタッフ情報を取得
 * @param {Spreadsheet} ss - スプレッドシート
 * @returns {Array} スタッフ情報の配列 [{fullName, displayName, index}]
 */
function getStaffInfo(ss) {
  const mainSheet = getSheetSafely(ss, SHEET_NAMES.MAIN);
  if (!mainSheet) return [];

  const staffData = mainSheet
    .getRange(
      MAIN_SHEET.STAFF.START_ROW,
      MAIN_SHEET.STAFF.NAME_COL,
      MAIN_SHEET.STAFF.END_ROW - MAIN_SHEET.STAFF.START_ROW + 1,
      2
    )
    .getValues();

  return staffData
    .map((row, index) => ({
      fullName: row[0], // 氏名（フルネーム）
      displayName: row[1], // 表示名（苗字）
      index: index,
    }))
    .filter((staff) => staff.fullName && staff.displayName);
}
