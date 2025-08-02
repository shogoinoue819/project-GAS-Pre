/**
 * 各日次シートの各スタッフ列に対し、該当スタッフの個人シートから希望シフトを取得し、
 * 希望がWISH_TRUEなら日次シートの希望行にWISH_TRUEを、そうでなければWISH_FALSEを記入する。
 */
function reflectWish() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(MAIN);
  if (!mainSheet) return;

  // メインシートからスタッフ情報を取得（氏名と表示名）
  const staffData = mainSheet
    .getRange(
      MAIN_STAFF_START_ROW,
      MAIN_STAFF_NAME_COL,
      MAIN_STAFF_END_ROW - MAIN_STAFF_START_ROW + 1,
      2
    )
    .getValues();

  // 氏名と表示名のマッピングを作成（空でないもののみ）
  const staffMapping = staffData
    .map((row, index) => ({
      fullName: row[0], // 氏名（フルネーム）
      displayName: row[1], // 表示名（苗字）
      index: index,
    }))
    .filter((staff) => staff.fullName && staff.displayName); // 空でないもののみ

  // 全シート名を取得
  const sheetNames = ss.getSheets().map((sheet) => sheet.getName());

  // 日次シートのみを抽出（テンプレートやメイン、スタッフ個人シートを除外）
  const dailySheetNames = sheetNames.filter((name) => isDailySheetName(name));

  dailySheetNames.forEach((dailySheetName) => {
    const dailySheet = ss.getSheetByName(dailySheetName);
    if (!dailySheet) return;

    // 日次シートから日付を取得
    const dateValue = dailySheet
      .getRange(DAILY_DATE_ROW, DAILY_DATE_COL)
      .getValue();
    if (!dateValue) return;

    // 各スタッフについて処理
    staffMapping.forEach((staff) => {
      const staffSheet = ss.getSheetByName(staff.fullName);
      if (!staffSheet) return; // 個人シートがなければスキップ

      // 個人シート内で該当日付の行を特定
      const staffDates = staffSheet
        .getRange(
          STAFF_DATE_START_ROW,
          STAFF_DATE_COL,
          staffSheet.getLastRow() - STAFF_DATE_START_ROW + 1,
          1
        )
        .getValues()
        .flat();

      // 日付が一致する行を探す
      const dateRowOffset = staffDates.findIndex((d) =>
        isSameDate(d, dateValue)
      );
      if (dateRowOffset === -1) {
        // 一致する日付がなければWISH_FALSEを書き込む
        dailySheet
          .getRange(DAILY_WISH_ROW, DAILY_STAFF_START_COL + staff.index)
          .setValue(WISH_FALSE);
        return;
      }

      // 希望値を取得
      const wishValue = staffSheet
        .getRange(STAFF_DATE_START_ROW + dateRowOffset, STAFF_WISH_COL)
        .getValue();

      // 希望がWISH_TRUEならWISH_TRUE、そうでなければWISH_FALSE
      const result = wishValue === WISH_TRUE ? WISH_TRUE : WISH_FALSE;
      dailySheet
        .getRange(DAILY_WISH_ROW, DAILY_STAFF_START_COL + staff.index)
        .setValue(result);
    });
  });
}

/**
 * 日次シート名かどうかを判定するヘルパー
 * 例: "7/30" のような日付形式のみtrue
 */
function isDailySheetName(name) {
  // 必要に応じて除外シート名を追加
  const exclude = [MAIN, TEMPLATE_DAILY, TEMPLATE_STAFF];
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
