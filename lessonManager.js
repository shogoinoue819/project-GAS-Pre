/**
 * 講義管理機能
 * 講義情報の取得、解析、管理を行う
 */

/**
 * 現在の日付の講義情報を取得する
 * @returns {Array} 講義オブジェクトの配列
 */
function getCurrentDateLessons() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 現在の日付を取得
  const currentDate = new Date();
  const currentDateFormatted = generateDailySheetName(currentDate);

  // 現在の日付シートを取得
  const currentDateSheet = getSheetSafely(ss, currentDateFormatted);
  if (!currentDateSheet) {
    Logger.log(`現在の日付シート「${currentDateFormatted}」が見つかりません。`);
    return [];
  }

  // 曜日を判定
  const dayOfWeek = getDayOfWeek(currentDate);
  Logger.log(`現在の日付: ${currentDateFormatted}, 曜日: ${dayOfWeek}`);

  // 曜日テンプレートシートを取得
  const weekdayTemplateSheet = getSheetSafely(ss, dayOfWeek);
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
  // 行：WEEK_SHEET.PERIOD_ROWS.FIRST から WEEK_SHEET.PERIOD_ROWS.THIRD
  // 列：WEEK_SHEET.GRADE_COLS.YOUNG から WEEK_SHEET.GRADE_COLS.SIXTH
  for (
    let row = WEEK_SHEET.PERIOD_ROWS.FIRST;
    row <= WEEK_SHEET.PERIOD_ROWS.THIRD;
    row++
  ) {
    for (
      let col = WEEK_SHEET.GRADE_COLS.YOUNG;
      col <= WEEK_SHEET.GRADE_COLS.SIXTH;
      col++
    ) {
      const cell = templateSheet.getRange(row, col);
      const cellValue = cell.getValue();

      // セルが空白または「なし」の場合はスキップ
      if (!cellValue || cellValue === "" || cellValue === "なし") {
        continue;
      }

      // 講義コードとして認識できる文字列かチェック
      if (isValidLessonCode(cellValue)) {
        // 講義情報を解析
        const lessonInfo = getLessonInfo(cellValue);

        // コマ数を行番号から判定
        const periodNumber = row - WEEK_SHEET.PERIOD_ROWS.FIRST + 1;

        // セルのスタイル情報を取得
        const cellStyle = extractCellStyle(cell);

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
          // スタイル情報
          style: cellStyle,
        });
      }
    }
  }

  return lessons;
}

/**
 * セルのスタイル情報を抽出
 * @param {Range} cell - セルオブジェクト
 * @returns {Object} スタイル情報オブジェクト
 */
function extractCellStyle(cell) {
  return {
    backgroundColor: cell.getBackground(),
    fontColor: cell.getFontColor(),
    borders: cell.getBorder(),
    fontFamily: cell.getFontFamily(),
    fontSize: cell.getFontSize(),
    fontBold: cell.getFontWeight() === "bold",
    horizontalAlignment: cell.getHorizontalAlignment(),
    verticalAlignment: cell.getVerticalAlignment(),
  };
}

/**
 * セルの値が講義コードかどうかを判定する
 * @param {string} cellValue - セルの値
 * @returns {boolean} 講義コードの場合true
 */
function isValidLessonCode(cellValue) {
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
function getLessonInfo(lessonCode) {
  if (!isValidLessonCode(lessonCode)) {
    return null;
  }

  return LESSON_CODES[lessonCode] || null;
}

/**
 * 指定された日付の講義情報を取得
 * @param {string} dateString - 日付文字列（M/d形式）
 * @returns {Array} 講義オブジェクトの配列
 */
function getLessonsForDate(dateString) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 日付からDateオブジェクトを作成
  const targetDate = getDateFromDailySheetName(dateString);
  if (!targetDate) {
    Logger.log(`日付形式が不正です: ${dateString}`);
    return [];
  }

  // 曜日を判定
  const dayOfWeek = getDayOfWeek(targetDate);
  Logger.log(`日付: ${dateString}, 曜日: ${dayOfWeek}`);

  // 曜日テンプレートシートを取得
  const weekdayTemplateSheet = getSheetSafely(ss, dayOfWeek);
  if (!weekdayTemplateSheet) {
    Logger.log(`曜日テンプレートシート「${dayOfWeek}」が見つかりません`);
    return [];
  }

  // 講義情報を取得
  return extractLessonsFromTemplate(weekdayTemplateSheet);
}
