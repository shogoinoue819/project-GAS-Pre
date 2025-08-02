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

  // 講義情報を取得
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

  // 講義情報を取得
  const lessons = extractLessonsFromTemplate(weekdayTemplateSheet);

  Logger.log(`取得した講義数: ${lessons.length}`);
  return lessons;
}

/**
 * 曜日テンプレートシートから授業情報を抽出する
 * @param {Sheet} templateSheet - 曜日テンプレートシート
 * @returns {Array} 授業情報の配列 [{row, col, lessonName}, ...]
 */
function extractLessonsFromTemplate(templateSheet) {
  const lessons = [];

  // 定数で定義された範囲のみを走査
  // 行：WEEK_PERIOD1_ROW から WEEK_PERIOD3_ROW
  // 列：WEEK_YOUNG_COL から WEEK_SIXTH_COL
  for (let row = WEEK_PERIOD1_ROW; row <= WEEK_PERIOD3_ROW; row++) {
    for (let col = WEEK_YOUNG_COL; col <= WEEK_SIXTH_COL; col++) {
      const cell = templateSheet.getRange(row, col);
      const cellValue = cell.getValue();

      // セルが空白または「なし」の場合はスキップ
      if (!cellValue || cellValue === "" || cellValue === "なし") {
        continue;
      }

      // 講義コードとして認識できる文字列かチェック
      // 例：「1M」「2J」「3R」「4S」など
      if (isLessonCell(cellValue)) {
        // 講義情報を解析
        const lessonInfo = parseLessonCode(cellValue);

        // コマ数を行番号から判定（3行目=1コマ目、4行目=2コマ目、5行目=3コマ目）
        const periodNumber = row - WEEK_PERIOD1_ROW + 1;

        lessons.push({
          lessonCode: cellValue, // 講義コード（例："1M"）
          period: periodNumber, // コマ数（1, 2, 3）
          grade: lessonInfo.grade, // 学年（例："小1"）
          subject: lessonInfo.subject, // 教科（例："算数"）
          gradeNumber: lessonInfo.gradeNumber, // 学年番号（1-6）
          subjectCode: lessonInfo.subjectCode, // 教科コード（M/J/R/S）
          row: row, // 行番号
          col: col, // 列番号
          assignedTeacher: null, // 担当講師の表示名（後で設定）
        });
      }
    }
  }

  return lessons;
}

/**
 * セルの値が講義コードかどうかを判定する
 * @param {string} cellValue - セルの値
 * @returns {boolean} 講義コードの場合true
 */
function isLessonCell(cellValue) {
  if (typeof cellValue !== "string") {
    return false;
  }

  // 講義コードのパターンを定義（学年+教科の形式）
  // 例：「1M」「2J」「3R」「4S」など
  const lessonPattern = /^[1-6][MJRS]$/;

  return lessonPattern.test(cellValue);
}

/**
 * 講義コードから講義情報を取得する
 * @param {string} lessonCode - 講義コード（例：「1M」「2J」）
 * @returns {Object} 講義情報オブジェクト または null
 */
function parseLessonCode(lessonCode) {
  if (!isLessonCell(lessonCode)) {
    return null;
  }

  return LESSON_CODES[lessonCode] || null;
}
