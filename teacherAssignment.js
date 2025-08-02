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
