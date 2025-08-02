/**
 * 現在の日付シートから授業情報を取得する（ステップ①）
 * 1. 現在の日付シートから曜日を判定
 * 2. 曜日テンプレートシートを参照
 * 3. 授業情報（行番号・列番号・授業名）を配列で返す
 */
function getCurrentDateLessons() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 特定の日付を指定
  const targetDate = new Date("2024-08-04");
  const targetDateFormatted = Utilities.formatDate(
    targetDate,
    Session.getScriptTimeZone(),
    "M/d"
  );

  // 指定した日付シートを取得
  const targetDateSheet = ss.getSheetByName(targetDateFormatted);
  if (!targetDateSheet) {
    throw new Error(
      `指定した日付シート「${targetDateFormatted}」が見つかりません。`
    );
  }

  // 曜日を判定
  const dayOfWeek = getDayOfWeek(targetDate);
  Logger.log(`指定した日付: ${targetDateFormatted}, 曜日: ${dayOfWeek}`);

  // 曜日テンプレートシートを取得
  const weekdayTemplateSheet = ss.getSheetByName(dayOfWeek);
  if (!weekdayTemplateSheet) {
    throw new Error(`曜日テンプレートシート「${dayOfWeek}」が見つかりません。`);
  }

  // 授業情報を取得
  const lessons = extractLessonsFromTemplate(weekdayTemplateSheet);

  // 結果をログ出力
  Logger.log(`取得した授業数: ${lessons.length}`);
  lessons.forEach((lesson, index) => {
    Logger.log(
      `授業${index + 1}: 行${lesson.row}, 列${lesson.col}, 「${
        lesson.lessonName
      }」`
    );
  });

  return lessons;
}

/**
 * 日付から曜日を取得する
 * @param {Date} date - 日付オブジェクト
 * @returns {string} 曜日（例：月曜日、火曜日...）
 */
function getDayOfWeek(date) {
  const weekdays = [
    WEEK_SUN, // 日曜日
    WEEK_MON, // 月曜日
    WEEK_TUE, // 火曜日
    WEEK_WED, // 水曜日
    WEEK_THU, // 木曜日
    WEEK_FRI, // 金曜日
    WEEK_SAT, // 土曜日
  ];
  return weekdays[date.getDay()];
}

/**
 * 曜日テンプレートシートから授業情報を抽出する
 * @param {Sheet} templateSheet - 曜日テンプレートシート
 * @returns {Array} 授業情報の配列 [{row, col, lessonName}, ...]
 */
function extractLessonsFromTemplate(templateSheet) {
  const lessons = [];

  // 定数で定義された範囲のみを走査
  // 行：WEEK_LESSON1_ROW から WEEK_LESSON3_ROW
  // 列：WEEK_YOUNG_COL から WEEK_SIXTH_COL
  for (let row = WEEK_LESSON1_ROW; row <= WEEK_LESSON3_ROW; row++) {
    for (let col = WEEK_YOUNG_COL; col <= WEEK_SIXTH_COL; col++) {
      const cell = templateSheet.getRange(row, col);
      const cellValue = cell.getValue();

      // セルが空白または「なし」の場合はスキップ
      if (!cellValue || cellValue === "" || cellValue === "なし") {
        continue;
      }

      // 授業名として認識できる文字列かチェック
      // 例：「１M 英語」「２M 数学」「３M 国語」など
      if (isLessonCell(cellValue)) {
        lessons.push({
          row: row,
          col: col,
          lessonName: cellValue,
        });
      }
    }
  }

  return lessons;
}

/**
 * 講義コードの定義（24通り）
 * 形式: 学年(1-6) + 教科(M/J/R/S)
 */
const LESSON_CODES = {
  // 小1
  "1M": { grade: "小1", subject: "算数", gradeNumber: 1, subjectCode: "M" },
  "1J": { grade: "小1", subject: "国語", gradeNumber: 1, subjectCode: "J" },
  "1R": { grade: "小1", subject: "理科", gradeNumber: 1, subjectCode: "R" },
  "1S": { grade: "小1", subject: "社会", gradeNumber: 1, subjectCode: "S" },

  // 小2
  "2M": { grade: "小2", subject: "算数", gradeNumber: 2, subjectCode: "M" },
  "2J": { grade: "小2", subject: "国語", gradeNumber: 2, subjectCode: "J" },
  "2R": { grade: "小2", subject: "理科", gradeNumber: 2, subjectCode: "R" },
  "2S": { grade: "小2", subject: "社会", gradeNumber: 2, subjectCode: "S" },

  // 小3
  "3M": { grade: "小3", subject: "算数", gradeNumber: 3, subjectCode: "M" },
  "3J": { grade: "小3", subject: "国語", gradeNumber: 3, subjectCode: "J" },
  "3R": { grade: "小3", subject: "理科", gradeNumber: 3, subjectCode: "R" },
  "3S": { grade: "小3", subject: "社会", gradeNumber: 3, subjectCode: "S" },

  // 小4
  "4M": { grade: "小4", subject: "算数", gradeNumber: 4, subjectCode: "M" },
  "4J": { grade: "小4", subject: "国語", gradeNumber: 4, subjectCode: "J" },
  "4R": { grade: "小4", subject: "理科", gradeNumber: 4, subjectCode: "R" },
  "4S": { grade: "小4", subject: "社会", gradeNumber: 4, subjectCode: "S" },

  // 小5
  "5M": { grade: "小5", subject: "算数", gradeNumber: 5, subjectCode: "M" },
  "5J": { grade: "小5", subject: "国語", gradeNumber: 5, subjectCode: "J" },
  "5R": { grade: "小5", subject: "理科", gradeNumber: 5, subjectCode: "R" },
  "5S": { grade: "小5", subject: "社会", gradeNumber: 5, subjectCode: "S" },

  // 小6
  "6M": { grade: "小6", subject: "算数", gradeNumber: 6, subjectCode: "M" },
  "6J": { grade: "小6", subject: "国語", gradeNumber: 6, subjectCode: "J" },
  "6R": { grade: "小6", subject: "理科", gradeNumber: 6, subjectCode: "R" },
  "6S": { grade: "小6", subject: "社会", gradeNumber: 6, subjectCode: "S" },
};

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
