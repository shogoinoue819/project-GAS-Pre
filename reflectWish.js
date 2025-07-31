/**
 * 各日付シートの各スタッフ列に対し、該当スタッフの個人シートから希望シフトを取得し、
 * 希望がWISH_TRUEなら日付シートの希望行にWISH_TRUEを、そうでなければfalseを記入する。
 */
function reflectStaffWishesToScheduleSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // 全シート名を取得
  const sheetNames = ss.getSheets().map(sheet => sheet.getName());

  // 日付シートのみを抽出（テンプレートやメイン、スタッフ個人シートを除外）
  const scheduleSheetNames = sheetNames.filter(name => isScheduleSheetName(name));

  scheduleSheetNames.forEach(scheduleSheetName => {
    const scheduleSheet = ss.getSheetByName(scheduleSheetName);
    if (!scheduleSheet) return;

    // ヘッダー行からスタッフ表示名を取得
    const staffNames = scheduleSheet
      .getRange(DATE_NAME_ROW, DATE_NAME_START_COL, 1, scheduleSheet.getLastColumn() - DATE_NAME_START_COL + 1)
      .getValues()[0];

    staffNames.forEach((staffName, i) => {
      if (!staffName) return; // 空欄はスキップ

      const staffSheet = ss.getSheetByName(staffName);
      if (!staffSheet) return; // 個人シートがなければスキップ

      // 日付シートのA1セルなどから日付を取得（A1に日付が入っている前提）
      const dateValue = scheduleSheet.getRange("A1").getValue();
      if (!dateValue) return;

      // 個人シート内で該当日付の行を特定
      const staffDates = staffSheet
        .getRange(STAFF_DATE_START_ROW, STAFF_DATE_COL, staffSheet.getLastRow() - STAFF_DATE_START_ROW + 1, 1)
        .getValues()
        .flat();

      // 日付が一致する行を探す
      const dateRowOffset = staffDates.findIndex(d => isSameDate(d, dateValue));
      if (dateRowOffset === -1) {
        // 一致する日付がなければfalseを書き込む
        scheduleSheet.getRange(DATE_WISH_ROW, DATE_NAME_START_COL + i).setValue(false);
        return;
      }

      // 希望値を取得
      const wishValue = staffSheet.getRange(STAFF_WISH_ROW + dateRowOffset, STAFF_WISH_COL).getValue();

      // 希望がWISH_TRUEならWISH_TRUE、そうでなければfalse
      const result = (wishValue === WISH_TRUE) ? WISH_TRUE : false;
      scheduleSheet.getRange(DATE_WISH_ROW, DATE_NAME_START_COL + i).setValue(result);
    });
  });
}

/**
 * スケジュールシート名かどうかを判定するヘルパー
 * 例: "7/30" のような日付形式のみtrue
 */
function isScheduleSheetName(name) {
  // 必要に応じて除外シート名を追加
  const exclude = [MAIN, DATE_STAFF, TEMPLATE_STAFF, "Template_Date"];
  if (exclude.includes(name)) return false;
  // "M/d"形式（例: 7/30）かどうか
  return /^\d{1,2}\/\d{1,2}$/.test(name);
}

/**
 * 日付の一致判定（時刻部分を無視して比較）
 */
function isSameDate(date1, date2) {
  if (!date1 || !date2) return false;
  const d1 = new Date(date1);
  const d2 = new Date(date2);
  return (
    d1.getFullYear() === d2.getFullYear() &&
    d1.getMonth() === d2.getMonth() &&
    d1.getDate() === d2.getDate()
  );
}