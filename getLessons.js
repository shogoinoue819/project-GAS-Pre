/**
 * 統合関数：講義取得から講師割り当てまで一括実行
 * 1. 現在の日付シートから曜日を判定
 * 2. 曜日テンプレートシートを参照して講義情報を取得
 * 3. 優先順位と勤務希望を加味して講師を自動割り当て
 * 4. 結果を日次シートに反映
 */
function getLessons() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 特定の日付を指定
  const targetDate = new Date("2024-08-04");
  const targetDateFormatted = Utilities.formatDate(
    targetDate,
    Session.getScriptTimeZone(),
    "M/d"
  );

  Logger.log("=== 講義取得・講師割り当て処理開始 ===");

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

  // 講義情報を取得
  const lessons = extractLessonsFromTemplate(weekdayTemplateSheet);

  // 講義取得結果をログ出力
  Logger.log(`取得した講義数: ${lessons.length}`);
  lessons.forEach((lesson, index) => {
    Logger.log(
      `講義${index + 1}: ${lesson.period}コマ目, ${lesson.grade}${
        lesson.subject
      }, 講義コード「${lesson.lessonCode}」, 位置(${lesson.row}行, ${
        lesson.col
      }列)`
    );
  });

  // 優先順位シートを取得
  const prioritySheet = ss.getSheetByName(PRIORITY);
  if (!prioritySheet) {
    Logger.log(
      "優先順位シート「Priority」が見つかりません。講師割り当てをスキップします。"
    );
    return lessons;
  }

  Logger.log("=== 講師割り当て処理開始 ===");

  // 講師割り当てを実行
  assignTeachersToLessons(lessons, targetDateSheet, prioritySheet);

  // 結果サマリーを表示
  printAssignmentSummary(lessons);

  Logger.log("=== 処理完了 ===");

  return lessons;
}

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

/**
 * 講師割り当て機能
 * 各授業に対して、優先順位と勤務希望を加味し、自動で講師を割り当てて日次シートに反映する
 */

/**
 * メイン関数：講師割り当て処理を実行
 * @param {Array} lessons - 講義オブジェクトの配列
 * @param {Sheet} dateSheet - 日次シート
 * @param {Sheet} prioritySheet - 優先順位シート
 */
function assignTeachersToLessons(lessons, dateSheet, prioritySheet) {
  Logger.log("講師割り当て処理を開始します...");

  // 各講義に対して講師を割り当て
  lessons.forEach((lesson, index) => {
    Logger.log(
      `講義${index + 1} (${lesson.lessonCode}) の講師割り当てを開始...`
    );

    // 優先順位リストを取得
    const priorityList = getPriorityList(lesson.lessonCode, prioritySheet);
    if (!priorityList || priorityList.length === 0) {
      Logger.log(`講義${lesson.lessonCode}の優先順位リストが見つかりません`);
      return;
    }

    // 優先順位上位から講師の希望を確認
    for (const teacherName of priorityList) {
      if (isAvailable(teacherName, dateSheet)) {
        lesson.assignedTeacher = teacherName;
        Logger.log(
          `講義${lesson.lessonCode}に講師「${teacherName}」を割り当てました`
        );
        break;
      }
    }

    if (!lesson.assignedTeacher) {
      Logger.log(`講義${lesson.lessonCode}の講師割り当てができませんでした`);
    }
  });

  // 割り当て結果を日次シートに反映
  fillLessonCodesToSheet(lessons, dateSheet);

  Logger.log("講師割り当て処理が完了しました");
}

/**
 * 講義コードに対応する優先順位リストを取得
 * @param {string} lessonCode - 講義コード（例："1M"）
 * @param {Sheet} prioritySheet - 優先順位シート
 * @returns {Array} 講師名の配列（優先順位順）
 */
function getPriorityList(lessonCode, prioritySheet) {
  // 優先順位シートから該当する列を検索
  const lastCol = prioritySheet.getLastColumn();

  for (let col = 1; col <= lastCol; col++) {
    const lessonCodeCell = prioritySheet
      .getRange(PRIORITY_LESSON_ROW, col)
      .getValue();

    if (lessonCodeCell === lessonCode) {
      // 優先順位リストを取得（2行目から4行目）
      const priorityList = [];

      // 優先順位①
      const firstTeacher = prioritySheet
        .getRange(PRIORITY_FIRST_ROW, col)
        .getValue();
      if (firstTeacher && firstTeacher !== "") {
        priorityList.push(firstTeacher);
      }

      // 優先順位②
      const secondTeacher = prioritySheet
        .getRange(PRIORITY_SECOND_ROW, col)
        .getValue();
      if (secondTeacher && secondTeacher !== "") {
        priorityList.push(secondTeacher);
      }

      // 優先順位③
      const thirdTeacher = prioritySheet
        .getRange(PRIORITY_THIRD_ROW, col)
        .getValue();
      if (thirdTeacher && thirdTeacher !== "") {
        priorityList.push(thirdTeacher);
      }

      return priorityList;
    }
  }

  return null;
}

/**
 * 講師の勤務希望を確認
 * @param {string} teacherName - 講師名（表示名）
 * @param {Sheet} dateSheet - 日次シート
 * @returns {boolean} 勤務可能な場合true
 */
function isAvailable(teacherName, dateSheet) {
  // 日次シートで講師の列を検索
  const teacherCol = findTeacherColumn(teacherName, dateSheet);
  if (teacherCol === -1) {
    Logger.log(`講師「${teacherName}」の列が見つかりません`);
    return false;
  }

  // 希望行から勤務希望を取得
  const wishValue = dateSheet.getRange(DAILY_WISH_ROW, teacherCol).getValue();

  // 希望が「◯」の場合のみ勤務可能
  return wishValue === WISH_TRUE;
}

/**
 * 講師名に対応する列番号を検索
 * @param {string} teacherName - 講師名
 * @param {Sheet} dateSheet - 日次シート
 * @returns {number} 列番号（見つからない場合は-1）
 */
function findTeacherColumn(teacherName, dateSheet) {
  const lastCol = dateSheet.getLastColumn();

  for (let col = DAILY_STAFF_START_COL; col <= lastCol; col++) {
    const staffName = dateSheet.getRange(DAILY_STAFF_ROW, col).getValue();
    if (staffName === teacherName) {
      return col;
    }
  }

  return -1;
}

/**
 * 割り当て結果を日次シートに反映
 * @param {Array} lessons - 講義オブジェクトの配列
 * @param {Sheet} dateSheet - 日次シート
 */
function fillLessonCodesToSheet(lessons, dateSheet) {
  Logger.log("日次シートへの反映を開始します...");

  lessons.forEach((lesson) => {
    if (lesson.assignedTeacher) {
      // 講師の列を検索
      const teacherCol = findTeacherColumn(lesson.assignedTeacher, dateSheet);
      if (teacherCol !== -1) {
        // コマ位置の行を計算
        const lessonRow = WEEK_PERIOD1_ROW + lesson.period - 1;

        // 講義コードをセット
        dateSheet.getRange(lessonRow, teacherCol).setValue(lesson.lessonCode);

        Logger.log(
          `講義コード「${lesson.lessonCode}」を${lesson.assignedTeacher}の${lesson.period}コマ目に設定しました`
        );
      }
    }
  });

  Logger.log("日次シートへの反映が完了しました");
}

/**
 * 講師割り当て結果のサマリーを出力
 * @param {Array} lessons - 講義オブジェクトの配列
 */
function printAssignmentSummary(lessons) {
  Logger.log("=== 講師割り当て結果サマリー ===");

  const assignedCount = lessons.filter(
    (lesson) => lesson.assignedTeacher
  ).length;
  const totalCount = lessons.length;

  Logger.log(`総講義数: ${totalCount}`);
  Logger.log(`割り当て済み: ${assignedCount}`);
  Logger.log(`未割り当て: ${totalCount - assignedCount}`);

  lessons.forEach((lesson, index) => {
    const status = lesson.assignedTeacher
      ? `✅ ${lesson.assignedTeacher}`
      : "❌ 未割り当て";

    Logger.log(
      `${index + 1}. ${lesson.lessonCode} (${lesson.grade}${
        lesson.subject
      }) - ${status}`
    );
  });

  Logger.log("================================");
}
