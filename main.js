/**
 * 講師割り当て機能
 * 全日程の講師割り当てを実行
 */

/**
 * 全日次シートで講師割り当てを一括実行
 */
function reflectLessons() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    Logger.log("=== 全日次シートでの講師割り当て処理開始 ===");

    // 優先順位シートを取得
    const prioritySheet = ss.getSheetByName(PRIORITY);
    if (!prioritySheet) {
      Logger.log(
        "優先順位シート「Priority」が見つかりません。処理を中止します。"
      );
      return;
    }

    // 全シート名を取得
    const sheetNames = ss.getSheets().map((sheet) => sheet.getName());

    // 日次シートのみを抽出
    const dailySheetNames = sheetNames.filter((name) => isDailySheetName(name));

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

      const dailySheet = ss.getSheetByName(dailySheetName);
      if (!dailySheet) {
        Logger.log(`日次シート「${dailySheetName}」が見つかりません`);
        return;
      }

      // 日付から曜日を判定
      const dateValue = dailySheet
        .getRange(DAILY_DATE_ROW, DAILY_DATE_COL)
        .getValue();
      if (!dateValue) {
        Logger.log(`日次シート「${dailySheetName}」の日付が取得できません`);
        return;
      }

      // 日付からDateオブジェクトを作成
      let targetDate;
      if (typeof dateValue === "string") {
        // 文字列の場合（例："7/30"）
        const dateParts = dateValue.split("/");
        const month = parseInt(dateParts[0]);
        const day = parseInt(dateParts[1]);
        const currentYear = new Date().getFullYear();
        targetDate = new Date(currentYear, month - 1, day);
      } else if (dateValue instanceof Date) {
        // Dateオブジェクトの場合
        targetDate = dateValue;
      } else {
        Logger.log(
          `日次シート「${dailySheetName}」の日付形式が不正です: ${dateValue}`
        );
        return;
      }

      // 曜日を判定
      const dayOfWeek = getDayOfWeek(targetDate);
      const displayDate =
        typeof dateValue === "string"
          ? dateValue
          : Utilities.formatDate(
              targetDate,
              Session.getScriptTimeZone(),
              "M/d"
            );
      Logger.log(`日付: ${displayDate}, 曜日: ${dayOfWeek}`);

      // 曜日テンプレートシートを取得
      const weekdayTemplateSheet = ss.getSheetByName(dayOfWeek);
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

      allResults.processedSheets++;
      allResults.totalLessons += lessons.length;
      allResults.totalAssigned += assignedLessons.length;
      allResults.totalUnassigned += unassignedLessons.length;

      // 未割り当て講義の詳細を記録
      const recordDate =
        typeof dateValue === "string"
          ? dateValue
          : Utilities.formatDate(
              targetDate,
              Session.getScriptTimeZone(),
              "M/d"
            );
      unassignedLessons.forEach((lesson) => {
        allResults.unassignedDetails.push({
          date: recordDate,
          period: lesson.period,
          lessonCode: lesson.lessonCode,
          grade: lesson.grade,
          subject: lesson.subject,
        });
      });

      Logger.log(`=== ${dailySheetName}の処理完了 ===`);
      Logger.log(
        `講義数: ${lessons.length}, 割り当て済み: ${assignedLessons.length}, 未割り当て: ${unassignedLessons.length}`
      );
    });

    // 全体の結果をUIアラートで表示
    showAllAssignmentResultUI(allResults);

    Logger.log("=== 全日次シートでの講師割り当て処理完了 ===");
  } catch (error) {
    Logger.log(`全日次シート処理でエラーが発生しました: ${error.message}`);
    throw error;
  }
}

/**
 * 全日次シートの割り当て結果をUIアラートで表示
 * @param {Object} allResults - 全体の結果オブジェクト
 */
function showAllAssignmentResultUI(allResults) {
  const ui = SpreadsheetApp.getUi();

  let message = `【全日次シートの講師割り当て結果】\n\n`;

  if (allResults.totalUnassigned === 0) {
    // 全て割り当て完了
    message += `✅ 全ての日次シートで講師割り当てが完了しました！\n\n`;
    message += `処理対象シート数: ${allResults.processedSheets}\n`;
    message += `総講義数: ${allResults.totalLessons}\n`;
    message += `割り当て済み: ${allResults.totalAssigned}\n`;
    message += `未割り当て: 0`;

    ui.alert("全講師割り当て完了", message, ui.ButtonSet.OK);
  } else {
    // 未割り当てがある場合
    message += `⚠️ 一部の講義で割り当てができませんでした\n\n`;
    message += `処理対象シート数: ${allResults.processedSheets}\n`;
    message += `総講義数: ${allResults.totalLessons}\n`;
    message += `割り当て済み: ${allResults.totalAssigned}\n`;
    message += `未割り当て: ${allResults.totalUnassigned}\n\n`;
    message += `【未割り当て講義一覧】\n`;

    allResults.unassignedDetails.forEach((detail, index) => {
      message += `${index + 1}. ${detail.date} ${detail.period}コマ目 ${
        detail.lessonCode
      } (${detail.grade}${detail.subject})\n`;
    });

    message += `\n※未割り当ての講義は日次シートに反映されていません。`;

    ui.alert("全講師割り当て結果", message, ui.ButtonSet.OK);
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
