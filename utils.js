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
    SHEET_NAMES.WEEKDAYS.SUN, // 日曜日 (index 0)
    SHEET_NAMES.WEEKDAYS.MON, // 月曜日 (index 1)
    SHEET_NAMES.WEEKDAYS.TUE, // 火曜日 (index 2)
    SHEET_NAMES.WEEKDAYS.WED, // 水曜日 (index 3)
    SHEET_NAMES.WEEKDAYS.THU, // 木曜日 (index 4)
    SHEET_NAMES.WEEKDAYS.FRI, // 金曜日 (index 5)
    SHEET_NAMES.WEEKDAYS.SAT, // 土曜日 (index 6)
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
  const exclude = [
    SHEET_NAMES.MAIN,
    SHEET_NAMES.TEMPLATE_DAILY,
    SHEET_NAMES.TEMPLATE_STAFF,
    SHEET_NAMES.PRIORITY,
  ];

  // 除外シート名に含まれる場合はfalse
  if (exclude.includes(name)) return false;

  // スタッフ個人シートを除外（メインシートからスタッフ名を取得して判定）
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(SHEET_NAMES.MAIN);
  if (mainSheet) {
    const staffNames = getStaffNamesFromMainSheet(mainSheet);
    if (staffNames.includes(name)) return false;
  }

  // "M/d"形式（例: 7/30）かどうか
  return /^\d{1,2}\/\d{1,2}$/.test(name);
}

/**
 * 日次シート名の一覧を取得
 * @param {Spreadsheet} ss - スプレッドシート
 * @returns {Array} 日次シート名の配列
 */
function getDailySheetNames(ss) {
  const sheetNames = ss.getSheets().map((sheet) => sheet.getName());
  return sheetNames.filter((name) => isDailySheetName(name));
}

/**
 * 日付から日次シート名を生成
 * @param {Date|string} dateValue - 日付値
 * @returns {string} 日次シート名（M/d形式）
 */
function generateDailySheetName(dateValue) {
  if (typeof dateValue === "string") {
    // 既にM/d形式の場合はそのまま返す
    if (isDailySheetName(dateValue)) {
      return dateValue;
    }
    // 他の形式の場合はDateオブジェクトに変換
    const date = parseDateValue(dateValue);
    return date ? formatDateToMD(date) : null;
  } else if (dateValue instanceof Date) {
    return formatDateToMD(dateValue);
  }
  return null;
}

/**
 * 日次シート名からDateオブジェクトを取得
 * @param {string} sheetName - 日次シート名（M/d形式）
 * @returns {Date|null} Dateオブジェクト（変換できない場合はnull）
 */
function getDateFromDailySheetName(sheetName) {
  if (!isDailySheetName(sheetName)) {
    return null;
  }
  return parseDateValue(sheetName);
}

/**
 * メインシートからスタッフ名を取得
 * @param {Sheet} mainSheet - メインシート
 * @returns {Array} スタッフ名の配列
 */
function getStaffNamesFromMainSheet(mainSheet) {
  return mainSheet
    .getRange(
      MAIN_SHEET.STAFF.START_ROW,
      MAIN_SHEET.STAFF.NAME_COL,
      MAIN_SHEET.STAFF.END_ROW - MAIN_SHEET.STAFF.START_ROW + 1,
      1
    )
    .getValues()
    .flat()
    .filter((name) => name && name !== "");
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

/**
 * 日付文字列がyyyy-mm-dd形式かどうかを判定する
 * @param {string} dateString - 判定対象の文字列
 * @returns {boolean} yyyy-mm-dd形式の場合true
 */
function isYYYYMMDD(dateString) {
  return /^\d{4}-\d{2}-\d{2}$/.test(dateString);
}

/**
 * スプレッドシートからシートを安全に取得
 * @param {Spreadsheet} ss - スプレッドシート
 * @param {string} sheetName - シート名
 * @returns {Sheet|null} シートオブジェクト（存在しない場合はnull）
 */
function getSheetSafely(ss, sheetName) {
  try {
    return ss.getSheetByName(sheetName);
  } catch (error) {
    logError(`シート「${sheetName}」の取得に失敗しました`, error);
    return null;
  }
}

/**
 * 必須シートの存在チェック
 * @param {Spreadsheet} ss - スプレッドシート
 * @param {Array} requiredSheets - 必須シート名の配列
 * @returns {Object} チェック結果 {success: boolean, missing: Array}
 */
function validateRequiredSheets(ss, requiredSheets) {
  const missing = [];

  requiredSheets.forEach((sheetName) => {
    if (!getSheetSafely(ss, sheetName)) {
      missing.push(sheetName);
    }
  });

  return {
    success: missing.length === 0,
    missing: missing,
  };
}

/**
 * 講義コードが有効かどうかを判定
 * @param {string} lessonCode - 講義コード
 * @returns {boolean} 有効な講義コードの場合true
 */
function isValidLessonCode(lessonCode) {
  return (
    typeof lessonCode === "string" &&
    /^[1-6][MJRS]$/.test(lessonCode) &&
    LESSON_CODES.hasOwnProperty(lessonCode)
  );
}

/**
 * 講義コードから講義情報を取得
 * @param {string} lessonCode - 講義コード
 * @returns {Object|null} 講義情報オブジェクト
 */
function getLessonInfo(lessonCode) {
  return isValidLessonCode(lessonCode) ? LESSON_CODES[lessonCode] : null;
}
