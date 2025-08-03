/**
 * 希望シフト管理機能
 * スタッフの希望シフトの反映と管理を行う
 */

/**
 * 各日次シートの各スタッフ列に対し、該当スタッフの個人シートから希望シフトを取得し、
 * 希望がWISH_TRUEなら日次シートの希望行にWISH_TRUEを、そうでなければWISH_FALSEを記入する。
 */
function reflectWish() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 必須シートの存在チェック
    const validation = validateRequiredSheets(ss, [SHEET_NAMES.MAIN]);
    if (!validation.success) {
      throw new Error(
        `必須シートが見つかりません: ${validation.missing.join(", ")}`
      );
    }

    // スタッフ情報を取得
    const staffInfo = getStaffInfo(ss);
    if (staffInfo.length === 0) {
      Logger.log("スタッフ情報が取得できませんでした");
      return;
    }

    // 日次シート名を取得
    const dailySheetNames = getDailySheetNames(ss);

    // 各日次シートを処理
    dailySheetNames.forEach((dailySheetName) => {
      const dailySheet = getSheetSafely(ss, dailySheetName);
      if (!dailySheet) return;

      // 日次シートから日付を取得
      const dateValue = dailySheet
        .getRange(DAILY_SHEET.DATE_ROW, DAILY_SHEET.DATE_COL)
        .getValue();
      if (!dateValue) return;

      // 各スタッフについて処理
      staffInfo.forEach((staff) => {
        const staffSheet = getSheetSafely(ss, staff.fullName);
        if (!staffSheet) return; // 個人シートがなければスキップ

        // 個人シート内で該当日付の行を特定
        const staffDates = staffSheet
          .getRange(
            STAFF_SHEET.DATE.START_ROW,
            STAFF_SHEET.DATE.COL,
            staffSheet.getLastRow() - STAFF_SHEET.DATE.START_ROW + 1,
            1
          )
          .getValues()
          .flat();

        // 日付が一致する行を探す
        const dateRowOffset = staffDates.findIndex((d) =>
          isSameDate(d, dateValue)
        );
        if (dateRowOffset === -1) {
          // 一致する日付がなければWISH_FALSEを書き込む
          dailySheet
            .getRange(
              DAILY_SHEET.WISH_ROW,
              DAILY_SHEET.STAFF_START_COL + staff.index
            )
            .setValue(STRINGS.WISH.FALSE);
          return;
        }

        // 希望値を取得
        const wishValue = staffSheet
          .getRange(
            STAFF_SHEET.DATE.START_ROW + dateRowOffset,
            STAFF_SHEET.WISH_COL
          )
          .getValue();

        // 希望がWISH_TRUEならWISH_TRUE、そうでなければWISH_FALSE
        const result =
          wishValue === STRINGS.WISH.TRUE
            ? STRINGS.WISH.TRUE
            : STRINGS.WISH.FALSE;
        dailySheet
          .getRange(
            DAILY_SHEET.WISH_ROW,
            DAILY_SHEET.STAFF_START_COL + staff.index
          )
          .setValue(result);
      });
    });

    Logger.log("希望シフトの反映が完了しました");
  } catch (error) {
    logError("希望シフト反映でエラーが発生しました", error);
    throw error;
  }
}

/**
 * スタッフシートの授業可能欄から情報を取得し、Priorityシートに表示名を追加する
 */
function updatePrioritySheetFromStaffSheets() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 必須シートの存在チェック
    const validation = validateRequiredSheets(ss, [
      SHEET_NAMES.MAIN,
      SHEET_NAMES.PRIORITY,
    ]);

    if (!validation.success) {
      throw new Error(
        `必須シートが見つかりません: ${validation.missing.join(", ")}`
      );
    }

    const prioritySheet = getSheetSafely(ss, SHEET_NAMES.PRIORITY);
    const staffInfo = getStaffInfo(ss);

    // 各スタッフの個人シートを処理
    staffInfo.forEach((staff) => {
      const staffSheet = getSheetSafely(ss, staff.fullName);
      if (!staffSheet) return;

      // スタッフシートから授業可能情報を取得
      const availableLessons = getAvailableLessonsFromStaffSheet(staffSheet);

      // 各授業についてPriorityシートを更新
      availableLessons.forEach((lessonCode) => {
        updatePrioritySheetForLesson(
          prioritySheet,
          lessonCode,
          staff.displayName
        );
      });
    });

    Logger.log("Priorityシートの更新が完了しました");
  } catch (error) {
    logError("Priorityシート更新でエラーが発生しました", error);
    throw error;
  }
}

/**
 * スタッフシートから授業可能な講義コードを取得
 * @param {Sheet} staffSheet - スタッフシート
 * @returns {Array} 授業可能な講義コードの配列
 */
function getAvailableLessonsFromStaffSheet(staffSheet) {
  const availableLessons = [];

  // 各学年・教科の組み合わせをチェック
  Object.keys(LESSON_CODES).forEach((lessonCode) => {
    const lessonInfo = LESSON_CODES[lessonCode];
    const gradeRow =
      STAFF_SHEET.GRADE_ROWS[getGradeKey(lessonInfo.gradeNumber)];
    const subjectCol =
      STAFF_SHEET.SUBJECT_COLS[getSubjectKey(lessonInfo.subjectCode)];

    if (gradeRow && subjectCol) {
      const cellValue = staffSheet.getRange(gradeRow, subjectCol).getValue();
      if (cellValue === STRINGS.WISH.TRUE) {
        availableLessons.push(lessonCode);
      }
    }
  });

  return availableLessons;
}

/**
 * 学年番号から設定キーを取得
 * @param {number} gradeNumber - 学年番号
 * @returns {string} 設定キー
 */
function getGradeKey(gradeNumber) {
  const gradeMap = {
    0: "YOUNG",
    1: "FIRST",
    2: "SECOND",
    3: "THIRD",
    4: "FOURTH",
    5: "FIFTH",
    6: "SIXTH",
  };
  return gradeMap[gradeNumber];
}

/**
 * 教科コードから設定キーを取得
 * @param {string} subjectCode - 教科コード
 * @returns {string} 設定キー
 */
function getSubjectKey(subjectCode) {
  const subjectMap = {
    M: "MATH",
    J: "JAPANESE",
    R: "SCIENCE",
    S: "SOCIAL",
  };
  return subjectMap[subjectCode];
}

/**
 * Priorityシートの特定の講義列に表示名を追加
 * @param {Sheet} prioritySheet - 優先順位シート
 * @param {string} lessonCode - 講義コード
 * @param {string} displayName - 表示名
 */
function updatePrioritySheetForLesson(prioritySheet, lessonCode, displayName) {
  // 優先順位シートから該当する列を検索
  const lastCol = prioritySheet.getLastColumn();

  for (let col = 1; col <= lastCol; col++) {
    const lessonCodeCell = prioritySheet
      .getRange(PRIORITY_SHEET.LESSON_ROW, col)
      .getValue();

    if (lessonCodeCell === lessonCode) {
      // 優先順位リストを取得（2行目から最後まで）
      const priorityList = prioritySheet
        .getRange(
          PRIORITY_SHEET.PRIORITY_ROWS.FIRST,
          col,
          prioritySheet.getLastRow() - PRIORITY_SHEET.PRIORITY_ROWS.FIRST + 1,
          1
        )
        .getValues()
        .flat()
        .filter((name) => name && name !== "");

      // 既に表示名が含まれている場合はスキップ
      if (priorityList.includes(displayName)) {
        return;
      }

      // 空のセルを探して表示名を追加
      for (
        let row = PRIORITY_SHEET.PRIORITY_ROWS.FIRST;
        row <= prioritySheet.getLastRow();
        row++
      ) {
        const cellValue = prioritySheet.getRange(row, col).getValue();
        if (!cellValue || cellValue === "") {
          prioritySheet.getRange(row, col).setValue(displayName);
          Logger.log(
            `Priorityシート: ${lessonCode}列に「${displayName}」を追加しました`
          );
          break;
        }
      }
      break;
    }
  }
}
