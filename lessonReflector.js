/**
 * 授業反映機能
 * 全日程の講師割り当てを一括実行
 */

/**
 * 全日次シートで講師割り当てを一括実行
 */
function reflectLessons() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    Logger.log("=== 全日次シートでの講師割り当て処理開始 ===");

    // 必須シートの存在チェック
    const validation = validateRequiredSheets(ss, [SHEET_NAMES.PRIORITY]);
    if (!validation.success) {
      throw new Error(
        `必須シートが見つかりません: ${validation.missing.join(", ")}`
      );
    }

    // 優先順位シートを取得
    const prioritySheet = getSheetSafely(ss, SHEET_NAMES.PRIORITY);

    // 日次シート名を取得
    const dailySheetNames = getDailySheetNames(ss);

    Logger.log(`処理対象日次シート数: ${dailySheetNames.length}`);

    // 全体の結果を格納
    const allResults = {
      totalSheets: dailySheetNames.length,
      processedSheets: 0,
      totalLessons: 0,
      totalAssigned: 0,
      totalUnassigned: 0,
      unassignedDetails: [],
    };

    // 各日次シートを処理
    dailySheetNames.forEach((dailySheetName) => {
      Logger.log(`=== ${dailySheetName}の処理開始 ===`);

      const dailySheet = getSheetSafely(ss, dailySheetName);
      if (!dailySheet) {
        Logger.log(`日次シート「${dailySheetName}」が見つかりません`);
        return;
      }

      // 日次シートから日付を取得
      const dateValue = dailySheet
        .getRange(DAILY_SHEET.DATE_ROW, DAILY_SHEET.DATE_COL)
        .getValue();
      if (!dateValue) {
        Logger.log(`日次シート「${dailySheetName}」の日付が取得できません`);
        return;
      }

      // 日付からDateオブジェクトを作成
      const targetDate = getDateFromDailySheetName(dateValue);
      if (!targetDate) {
        Logger.log(
          `日次シート「${dailySheetName}」の日付形式が不正です: ${dateValue}`
        );
        return;
      }

      // 曜日を判定
      const dayOfWeek = getDayOfWeek(targetDate);
      const displayDate = dateValue; // 既にM/d形式なのでそのまま使用
      Logger.log(`日付: ${displayDate}, 曜日: ${dayOfWeek}`);

      // 曜日テンプレートシートを取得
      const weekdayTemplateSheet = getSheetSafely(ss, dayOfWeek);
      if (!weekdayTemplateSheet) {
        Logger.log(`曜日テンプレートシート「${dayOfWeek}」が見つかりません`);
        return;
      }

      // 講義情報を取得
      const lessons = extractLessonsFromTemplate(weekdayTemplateSheet);

      if (lessons.length === 0) {
        Logger.log(`日次シート「${dailySheetName}」に講義が見つかりません`);
        return;
      }

      // 講師割り当てを実行
      assignTeachersToLessons(lessons, dailySheet, prioritySheet);

      // 結果を集計
      const assignedLessons = lessons.filter(
        (lesson) => lesson.assignedTeacher
      );
      const unassignedLessons = lessons.filter(
        (lesson) => !lesson.assignedTeacher
      );

      allResults.totalLessons += lessons.length;
      allResults.totalAssigned += assignedLessons.length;
      allResults.totalUnassigned += unassignedLessons.length;
      allResults.processedSheets++;

      // 未割り当て講義の詳細を記録
      if (unassignedLessons.length > 0) {
        unassignedLessons.forEach((lesson) => {
          allResults.unassignedDetails.push({
            date: dailySheetName,
            lessonCode: lesson.lessonCode,
            grade: lesson.grade,
            subject: lesson.subject,
            period: lesson.period,
          });
        });
      }

      Logger.log(
        `${dailySheetName}: ${assignedLessons.length}/${lessons.length} 講義を割り当て完了`
      );
    });

    // 全体結果を表示
    showAllAssignmentResultUI(allResults);

    Logger.log("=== 全日次シートでの講師割り当て処理完了 ===");
  } catch (error) {
    logError("全日程授業反映でエラーが発生しました", error);
    throw error;
  }
}

/**
 * 全体の割り当て結果をUIに表示
 * @param {Object} allResults - 全体結果オブジェクト
 */
function showAllAssignmentResultUI(allResults) {
  const ui = SpreadsheetApp.getUi();

  let message = `全日程の講師割り当てが完了しました。\n\n`;
  message += `処理対象シート数: ${allResults.processedSheets}/${allResults.totalSheets}\n`;
  message += `総講義数: ${allResults.totalLessons}\n`;
  message += `割り当て成功: ${allResults.totalAssigned}\n`;
  message += `割り当て失敗: ${allResults.totalUnassigned}\n`;

  if (allResults.unassignedDetails.length > 0) {
    message += `\n割り当て失敗した講義:\n`;
    allResults.unassignedDetails.forEach((detail) => {
      message += `・${detail.date}: ${detail.lessonCode} (${detail.grade}${detail.subject}) - ${detail.period}コマ目\n`;
    });
  }

  ui.alert("講師割り当て完了", message, ui.ButtonSet.OK);
}

/**
 * テスト用：現在の日付の講義情報を取得
 */
function testGetLessons() {
  const lessons = getCurrentDateLessons();
  Logger.log(`取得した講義数: ${lessons.length}`);

  lessons.forEach((lesson, index) => {
    Logger.log(
      `${index + 1}. ${lesson.lessonCode} (${lesson.grade}${
        lesson.subject
      }) - ${lesson.period}コマ目`
    );
  });
}
