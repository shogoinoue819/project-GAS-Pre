/**
 * メイン実行関数
 * 講義取得から講師割り当てまで一括実行
 */

/**
 * 統合実行関数：講義取得から講師割り当てまで一括実行
 * @returns {Array} 講義オブジェクトの配列
 */
function reflectLessons() {
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
    showAssignmentResultUI(lessons, targetDateFormatted);
    return lessons;
  }

  Logger.log("=== 講師割り当て処理開始 ===");

  // 講師割り当てを実行
  assignTeachersToLessons(lessons, targetDateSheet, prioritySheet);

  // 結果サマリーを表示
  printAssignmentSummary(lessons);

  // UIアラートで結果を表示
  showAssignmentResultUI(lessons, targetDateFormatted);

  Logger.log("=== 処理完了 ===");

  return lessons;
}

/**
 * 現在の日付で講義取得から講師割り当てまで一括実行
 * @returns {Array} 講義オブジェクトの配列
 */
function getCurrentDateLessonsWithAssignment() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 現在の日付を取得
  const currentDate = new Date();
  const currentDateFormatted = Utilities.formatDate(
    currentDate,
    Session.getScriptTimeZone(),
    "M/d"
  );

  Logger.log("=== 現在の日付での講義取得・講師割り当て処理開始 ===");

  // 現在の日付シートを取得
  const currentDateSheet = ss.getSheetByName(currentDateFormatted);
  if (!currentDateSheet) {
    throw new Error(
      `現在の日付シート「${currentDateFormatted}」が見つかりません。`
    );
  }

  // 曜日を判定
  const dayOfWeek = getDayOfWeek(currentDate);
  Logger.log(`現在の日付: ${currentDateFormatted}, 曜日: ${dayOfWeek}`);

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
    showAssignmentResultUI(lessons, currentDateFormatted);
    return lessons;
  }

  Logger.log("=== 講師割り当て処理開始 ===");

  // 講師割り当てを実行
  assignTeachersToLessons(lessons, currentDateSheet, prioritySheet);

  // 結果サマリーを表示
  printAssignmentSummary(lessons);

  // UIアラートで結果を表示
  showAssignmentResultUI(lessons, currentDateFormatted);

  Logger.log("=== 処理完了 ===");

  return lessons;
}

/**
 * 割り当て結果をUIアラートで表示
 * @param {Array} lessons - 講義オブジェクトの配列
 * @param {string} dateFormatted - 日付文字列
 */
function showAssignmentResultUI(lessons, dateFormatted) {
  const ui = SpreadsheetApp.getUi();

  // 割り当て済みと未割り当ての講義を分類
  const assignedLessons = lessons.filter((lesson) => lesson.assignedTeacher);
  const unassignedLessons = lessons.filter((lesson) => !lesson.assignedTeacher);

  let message = `【${dateFormatted}の講師割り当て結果】\n\n`;

  if (unassignedLessons.length === 0) {
    // 全て割り当て完了
    message += `✅ 全ての講義の割り当てが完了しました！\n\n`;
    message += `総講義数: ${lessons.length}\n`;
    message += `割り当て済み: ${assignedLessons.length}\n`;
    message += `未割り当て: 0`;

    ui.alert("講師割り当て完了", message, ui.ButtonSet.OK);
  } else {
    // 未割り当てがある場合
    message += `⚠️ 一部の講義で割り当てができませんでした\n\n`;
    message += `総講義数: ${lessons.length}\n`;
    message += `割り当て済み: ${assignedLessons.length}\n`;
    message += `未割り当て: ${unassignedLessons.length}\n\n`;
    message += `【未割り当て講義一覧】\n`;

    unassignedLessons.forEach((lesson, index) => {
      message += `${index + 1}. ${dateFormatted} ${lesson.period}コマ目 ${
        lesson.lessonCode
      } (${lesson.grade}${lesson.subject})\n`;
    });

    message += `\n※未割り当ての講義は日次シートに反映されていません。`;

    ui.alert("講師割り当て結果", message, ui.ButtonSet.OK);
  }
}

/**
 * テスト用関数：講義情報のみ取得して確認
 */
function testGetLessons() {
  try {
    const lessons = getCurrentDateLessons();
    if (lessons && lessons.length > 0) {
      Logger.log("講義情報の取得テスト成功");
      lessons.forEach((lesson, index) => {
        Logger.log(
          `講義${index + 1}: ${lesson.lessonCode} (${lesson.grade}${
            lesson.subject
          })`
        );
      });
    } else {
      Logger.log("講義が見つかりませんでした");
    }
  } catch (error) {
    Logger.log(`エラーが発生しました: ${error.message}`);
  }
}
