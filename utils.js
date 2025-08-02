/**
 * 共通ユーティリティ関数
 * プロジェクト全体で使用される汎用的な関数を集約
 */

/**
 * 日付から曜日を取得する
 * @param {Date} date - 日付オブジェクト
 * @returns {string} 曜日（例：月曜日、火曜日...）
 */
function getDayOfWeek(date) {
  const weekdays = [
    WEEK_SUN, // 日曜日 (index 0)
    WEEK_MON, // 月曜日 (index 1)
    WEEK_TUE, // 火曜日 (index 2)
    WEEK_WED, // 水曜日 (index 3)
    WEEK_THU, // 木曜日 (index 4)
    WEEK_FRI, // 金曜日 (index 5)
    WEEK_SAT, // 土曜日 (index 6)
  ];
  return weekdays[date.getDay()];
}

/**
 * 日次シート名かどうかを判定するヘルパー
 * 例: "7/30" のような日付形式のみtrue
 * @param {string} name - シート名
 * @returns {boolean} 日次シート名の場合true
 */
function isDailySheetName(name) {
  // 除外シート名を定義
  const exclude = [MAIN, TEMPLATE_DAILY, TEMPLATE_STAFF, PRIORITY];

  // 除外シート名に含まれる場合はfalse
  if (exclude.includes(name)) return false;

  // スタッフ個人シートを除外（メインシートからスタッフ名を取得して判定）
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(MAIN);
  if (mainSheet) {
    const staffNames = mainSheet
      .getRange(
        MAIN_STAFF_START_ROW,
        MAIN_STAFF_NAME_COL,
        MAIN_STAFF_END_ROW - MAIN_STAFF_START_ROW + 1,
        1
      )
      .getValues()
      .flat()
      .filter((name) => name && name !== "");

    if (staffNames.includes(name)) return false;
  }

  // "M/d"形式（例: 7/30）かどうか
  return /^\d{1,2}\/\d{1,2}$/.test(name);
}

/**
 * 日付の一致判定（時刻部分を無視して比較）
 * @param {Date|string} date1 - 比較対象日付1
 * @param {Date|string} date2 - 比較対象日付2
 * @returns {boolean} 日付が一致する場合true
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

/**
 * 日付文字列からDateオブジェクトを作成
 * @param {string|Date} dateValue - 日付値
 * @returns {Date|null} Dateオブジェクト（変換できない場合はnull）
 */
function parseDateValue(dateValue) {
  if (typeof dateValue === "string") {
    // 文字列の場合（例："7/30"）
    const dateParts = dateValue.split("/");
    if (dateParts.length !== 2) return null;

    const month = parseInt(dateParts[0]);
    const day = parseInt(dateParts[1]);
    const currentYear = new Date().getFullYear();

    if (isNaN(month) || isNaN(day)) return null;

    return new Date(currentYear, month - 1, day);
  } else if (dateValue instanceof Date) {
    // Dateオブジェクトの場合
    return dateValue;
  }

  return null;
}

/**
 * 日付をM/d形式の文字列にフォーマット
 * @param {Date} date - 日付オブジェクト
 * @returns {string} M/d形式の文字列
 */
function formatDateToMD(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "M/d");
}

/**
 * エラーログを出力
 * @param {string} message - エラーメッセージ
 * @param {Error} error - エラーオブジェクト
 */
function logError(message, error) {
  Logger.log(`${message}: ${error.message}`);
  if (error.stack) {
    Logger.log(`Stack trace: ${error.stack}`);
  }
}
