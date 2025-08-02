/**
 * 各日次シートの各スタッフ列に対し、該当スタッフの個人シートから希望シフトを取得し、
 * 希望がWISH_TRUEなら日次シートの希望行にWISH_TRUEを、そうでなければWISH_FALSEを記入する。
 */
function reflectWish() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(MAIN);
  if (!mainSheet) return;

  // メインシートからスタッフ情報を取得（氏名と表示名）
  const staffData = mainSheet
    .getRange(
      MAIN_STAFF_START_ROW,
      MAIN_STAFF_NAME_COL,
      MAIN_STAFF_END_ROW - MAIN_STAFF_START_ROW + 1,
      2
    )
    .getValues();

  // 氏名と表示名のマッピングを作成（空でないもののみ）
  const staffMapping = staffData
    .map((row, index) => ({
      fullName: row[0], // 氏名（フルネーム）
      displayName: row[1], // 表示名（苗字）
      index: index,
    }))
    .filter((staff) => staff.fullName && staff.displayName); // 空でないもののみ

  // 全シート名を取得
  const sheetNames = ss.getSheets().map((sheet) => sheet.getName());

  // 日次シートのみを抽出（テンプレートやメイン、スタッフ個人シートを除外）
  const dailySheetNames = sheetNames.filter((name) => isDailySheetName(name));

  dailySheetNames.forEach((dailySheetName) => {
    const dailySheet = ss.getSheetByName(dailySheetName);
    if (!dailySheet) return;

    // 日次シートから日付を取得
    const dateValue = dailySheet
      .getRange(DAILY_DATE_ROW, DAILY_DATE_COL)
      .getValue();
    if (!dateValue) return;

    // 各スタッフについて処理
    staffMapping.forEach((staff) => {
      const staffSheet = ss.getSheetByName(staff.fullName);
      if (!staffSheet) return; // 個人シートがなければスキップ

      // 個人シート内で該当日付の行を特定
      const staffDates = staffSheet
        .getRange(
          STAFF_DATE_START_ROW,
          STAFF_DATE_COL,
          staffSheet.getLastRow() - STAFF_DATE_START_ROW + 1,
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
          .getRange(DAILY_WISH_ROW, DAILY_STAFF_START_COL + staff.index)
          .setValue(WISH_FALSE);
        return;
      }

      // 希望値を取得
      const wishValue = staffSheet
        .getRange(STAFF_DATE_START_ROW + dateRowOffset, STAFF_WISH_COL)
        .getValue();

      // 希望がWISH_TRUEならWISH_TRUE、そうでなければWISH_FALSE
      const result = wishValue === WISH_TRUE ? WISH_TRUE : WISH_FALSE;
      dailySheet
        .getRange(DAILY_WISH_ROW, DAILY_STAFF_START_COL + staff.index)
        .setValue(result);
    });
  });

  Logger.log("希望シフトの反映が完了しました");
}

/**
 * スタッフシートの授業可能欄から情報を取得し、Priorityシートに表示名を追加する
 */
function updatePrioritySheetFromStaffSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(MAIN);
  const prioritySheet = ss.getSheetByName(PRIORITY);

  if (!mainSheet || !prioritySheet) {
    Logger.log("メインシートまたはPriorityシートが見つかりません");
    return;
  }

  // メインシートからスタッフ情報を取得（氏名と表示名）
  const staffData = mainSheet
    .getRange(
      MAIN_STAFF_START_ROW,
      MAIN_STAFF_NAME_COL,
      MAIN_STAFF_END_ROW - MAIN_STAFF_START_ROW + 1,
      2
    )
    .getValues();

  // 氏名と表示名のマッピングを作成（空でないもののみ）
  const staffMapping = staffData
    .map((row) => ({
      fullName: row[0], // 氏名（フルネーム）
      displayName: row[1], // 表示名（苗字）
    }))
    .filter((staff) => staff.fullName && staff.displayName);

  Logger.log("Priorityシートの更新を開始します...");

  // 各スタッフについて処理
  staffMapping.forEach((staff) => {
    const staffSheet = ss.getSheetByName(staff.fullName);
    if (!staffSheet) {
      Logger.log(`スタッフシート「${staff.fullName}」が見つかりません`);
      return;
    }

    // スタッフシートから授業可能な学年・教科を取得
    const availableLessons = getAvailableLessonsFromStaffSheet(staffSheet);

    if (availableLessons.length > 0) {
      Logger.log(
        `スタッフ「${
          staff.displayName
        }」の授業可能科目: ${availableLessons.join(", ")}`
      );

      // 各授業コードについてPriorityシートを更新
      availableLessons.forEach((lessonCode) => {
        updatePrioritySheetForLesson(
          prioritySheet,
          lessonCode,
          staff.displayName
        );
      });
    }
  });

  Logger.log("Priorityシートの更新が完了しました");
}

/**
 * スタッフシートから授業可能な学年・教科を取得
 * @param {Sheet} staffSheet - スタッフシート
 * @returns {Array} 授業可能な講義コードの配列
 */
function getAvailableLessonsFromStaffSheet(staffSheet) {
  const availableLessons = [];

  // 学年別の授業可能チェック
  const gradeRows = [
    { row: STAFF_YOUNG_ROW, grade: "年長" },
    { row: STAFF_FIRST_ROW, grade: "小1" },
    { row: STAFF_SECOND_ROW, grade: "小2" },
    { row: STAFF_THIRD_ROW, grade: "小3" },
    { row: STAFF_FOURTH_ROW, grade: "小4" },
    { row: STAFF_FIFTH_ROW, grade: "小5" },
    { row: STAFF_SIXTH_ROW, grade: "小6" },
  ];

  // 教科別の授業可能チェック
  const subjectCols = [
    { col: STAFF_MAT_COL, subject: "算数", code: "M" },
    { col: STAFF_JAP_COL, subject: "国語", code: "J" },
    { col: STAFF_SCI_COL, subject: "理科", code: "R" },
    { col: STAFF_SOC_COL, subject: "社会", code: "S" },
  ];

  // 各学年・教科の組み合わせをチェック
  gradeRows.forEach((gradeInfo) => {
    subjectCols.forEach((subjectInfo) => {
      const cellValue = staffSheet
        .getRange(gradeInfo.row, subjectInfo.col)
        .getValue();

      if (cellValue === WISH_TRUE) {
        // 学年番号を取得（年長は0、小1は1、...、小6は6）
        const gradeNumber =
          gradeInfo.grade === "年長"
            ? 0
            : parseInt(gradeInfo.grade.replace("小", ""));

        // 年長の場合はスキップ（講義コードに含まれないため）
        if (gradeNumber > 0) {
          const lessonCode = `${gradeNumber}${subjectInfo.code}`;
          availableLessons.push(lessonCode);
        }
      }
    });
  });

  return availableLessons;
}

/**
 * Priorityシートの特定の講義コード列に表示名を追加
 * @param {Sheet} prioritySheet - Priorityシート
 * @param {string} lessonCode - 講義コード
 * @param {string} displayName - 表示名
 */
function updatePrioritySheetForLesson(prioritySheet, lessonCode, displayName) {
  const lastCol = prioritySheet.getLastColumn();

  // 該当する講義コードの列を検索
  for (let col = 1; col <= lastCol; col++) {
    const lessonCodeCell = prioritySheet
      .getRange(PRIORITY_LESSON_ROW, col)
      .getValue();

    if (lessonCodeCell === lessonCode) {
      // 優先順位①～③をチェック
      const firstTeacher = prioritySheet
        .getRange(PRIORITY_FIRST_ROW, col)
        .getValue();
      const secondTeacher = prioritySheet
        .getRange(PRIORITY_SECOND_ROW, col)
        .getValue();
      const thirdTeacher = prioritySheet
        .getRange(PRIORITY_THIRD_ROW, col)
        .getValue();

      // 既に優先順位①～③に含まれている場合はスキップ
      if (
        firstTeacher === displayName ||
        secondTeacher === displayName ||
        thirdTeacher === displayName
      ) {
        Logger.log(
          `講義${lessonCode}の優先順位①～③に「${displayName}」が既に含まれています`
        );
        return;
      }

      // 優先順位③の下から順番に空いているセルを探す
      let targetRow = PRIORITY_THIRD_ROW + 1;
      let added = false;

      while (targetRow <= prioritySheet.getLastRow()) {
        const cellValue = prioritySheet.getRange(targetRow, col).getValue();

        if (!cellValue || cellValue === "") {
          // 空いているセルに表示名を追加
          prioritySheet.getRange(targetRow, col).setValue(displayName);
          Logger.log(
            `講義${lessonCode}の${targetRow}行目に「${displayName}」を追加しました`
          );
          added = true;
          break;
        }

        targetRow++;
      }

      if (!added) {
        // 最後の行の下に新しい行を追加
        prioritySheet.getRange(targetRow, col).setValue(displayName);
        Logger.log(
          `講義${lessonCode}の${targetRow}行目に「${displayName}」を追加しました（新規行）`
        );
      }

      break; // 該当する列が見つかったらループを抜ける
    }
  }
}

/**
 * 日次シート名かどうかを判定するヘルパー
 * 例: "7/30" のような日付形式のみtrue
 */
function isDailySheetName(name) {
  // 除外シート名を定義
  const exclude = [MAIN, TEMPLATE_DAILY, TEMPLATE_STAFF, PRIORITY];

  // 除外シート名に含まれる場合はfalse
  if (exclude.includes(name)) return false;

  // スタッフ個人シートを除外（メインシートからスタッフ名を取得して判定）
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(MAIN);
  if (mainSheet) {
    const staffNames = mainSheet
      .getRange(
        MAIN_STAFF_START_ROW,
        MAIN_STAFF_NAME_COL,
        MAIN_STAFF_END_ROW - MAIN_STAFF_START_ROW + 1,
        1
      )
      .getValues()
      .flat()
      .filter((name) => name && name !== "");

    if (staffNames.includes(name)) return false;
  }

  // "M/d"形式（例: 7/30）かどうか
  return /^\d{1,2}\/\d{1,2}$/.test(name);
}

/**
 * 日付の一致判定（時刻部分を無視して比較）
 */
function isSameDate(date1, date2) {
  if (!date1 || !date2) return false;
  const d1 = new Date(date1);
  const d2 = new Date(date2);
  return (
    d1.getFullYear() === d2.getFullYear() &&
    d1.getMonth() === d2.getMonth() &&
    d1.getDate() === d2.getDate()
  );
}
