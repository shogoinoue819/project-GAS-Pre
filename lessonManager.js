/**
 * 講義管理機能
 * 講義情報の取得、解析、管理を行う
 */

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
