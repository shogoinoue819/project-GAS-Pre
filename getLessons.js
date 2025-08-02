/**
 * 講義取得機能
 * 曜日テンプレートシートから講義情報を取得する
 */

/**
 * 日付から曜日を取得する
 * @param {Date} date - 日付オブジェクト
 * @returns {string} 曜日（例：月曜日、火曜日...）
 */
function getDayOfWeek(date) {
  const weekdays = [
    WEEK_MON, // 月曜日 (index 0)
    WEEK_TUE, // 火曜日 (index 1)
    WEEK_WED, // 水曜日 (index 2)
    WEEK_THU, // 木曜日 (index 3)
    WEEK_FRI, // 金曜日 (index 4)
    WEEK_SAT, // 土曜日 (index 5)
    WEEK_SUN, // 日曜日 (index 6)
  ];
  return weekdays[date.getDay()];
}

/**
 * 現在の日付の講義情報を取得する
 * @returns {Array} 講義オブジェクトの配列
 */
function getCurrentDateLessons() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 現在の日付を取得
  const currentDate = new Date();
  const currentDateFormatted = Utilities.formatDate(
    currentDate,
    Session.getScriptTimeZone(),
    "M/d"
  );

  // 現在の日付シートを取得
  const currentDateSheet = ss.getSheetByName(currentDateFormatted);
  if (!currentDateSheet) {
    Logger.log(`現在の日付シート「${currentDateFormatted}」が見つかりません。`);
    return [];
  }

  // 曜日を判定
  const dayOfWeek = getDayOfWeek(currentDate);
  Logger.log(`現在の日付: ${currentDateFormatted}, 曜日: ${dayOfWeek}`);

  // 曜日テンプレートシートを取得
  const weekdayTemplateSheet = ss.getSheetByName(dayOfWeek);
  if (!weekdayTemplateSheet) {
    Logger.log(`曜日テンプレートシート「${dayOfWeek}」が見つかりません。`);
    return [];
  }

  // 講義情報を取得（lessonManager.jsの関数を使用）
  const lessons = extractLessonsFromTemplate(weekdayTemplateSheet);

  Logger.log(`取得した講義数: ${lessons.length}`);
  return lessons;
}

/**
 * 指定した日付の講義情報を取得する
 * @param {Date} targetDate - 対象日付
 * @returns {Array} 講義オブジェクトの配列
 */
function getTargetDateLessons(targetDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const targetDateFormatted = Utilities.formatDate(
    targetDate,
    Session.getScriptTimeZone(),
    "M/d"
  );

  // 指定した日付シートを取得
  const targetDateSheet = ss.getSheetByName(targetDateFormatted);
  if (!targetDateSheet) {
    Logger.log(
      `指定した日付シート「${targetDateFormatted}」が見つかりません。`
    );
    return [];
  }

  // 曜日を判定
  const dayOfWeek = getDayOfWeek(targetDate);
  Logger.log(`指定した日付: ${targetDateFormatted}, 曜日: ${dayOfWeek}`);

  // 曜日テンプレートシートを取得
  const weekdayTemplateSheet = ss.getSheetByName(dayOfWeek);
  if (!weekdayTemplateSheet) {
    Logger.log(`曜日テンプレートシート「${dayOfWeek}」が見つかりません。`);
    return [];
  }

  // 講義情報を取得（lessonManager.jsの関数を使用）
  const lessons = extractLessonsFromTemplate(weekdayTemplateSheet);

  Logger.log(`取得した講義数: ${lessons.length}`);
  return lessons;
}
