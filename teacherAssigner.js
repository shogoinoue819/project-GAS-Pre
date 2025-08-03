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
  try {
    Logger.log("講師割り当て処理を開始します...");

    // 割り当て済み講師の管理（コマ別）
    const assignedTeachers = {
      1: new Set(), // 1コマ目に割り当て済みの講師
      2: new Set(), // 2コマ目に割り当て済みの講師
      3: new Set(), // 3コマ目に割り当て済みの講師
    };

    // 授業の優先度を考慮してソート
    const sortedLessons = sortLessonsByPriority(lessons);

    Logger.log("授業の優先度順でソートしました:");
    sortedLessons.forEach((lesson, index) => {
      Logger.log(
        `${index + 1}. ${lesson.lessonCode} (${lesson.grade}${
          lesson.subject
        }) - 優先度: ${calculateLessonPriority(lesson)}`
      );
    });

    // 各講義に対して講師を割り当て
    sortedLessons.forEach((lesson, index) => {
      Logger.log(
        `講義${index + 1} (${lesson.lessonCode}) の講師割り当てを開始...`
      );

      // 優先順位リストを取得（Priorityシートの該当授業列の上から順番）
      const priorityList = getPriorityList(lesson.lessonCode, prioritySheet);
      if (!priorityList || priorityList.length === 0) {
        Logger.log(`講義${lesson.lessonCode}の優先順位リストが見つかりません`);
        return;
      }

      // 優先順位リストの上から順番に講師の希望を確認
      let assigned = false;
      for (const teacherName of priorityList) {
        if (
          isAvailable(teacherName, dateSheet, assignedTeachers, lesson.period)
        ) {
          lesson.assignedTeacher = teacherName;
          // 割り当て済み講師リストに追加
          assignedTeachers[lesson.period].add(teacherName);

          Logger.log(
            `講義${lesson.lessonCode}に講師「${teacherName}」を割り当てました（${lesson.period}コマ目）`
          );
          assigned = true;
          break;
        }
      }

      // Priorityシートの該当授業列の一番下まで当たっても不可能だった場合
      if (!assigned) {
        Logger.log(
          `講義${lesson.lessonCode}のPriorityシート記載講師での割り当てができませんでした（割り当て失敗）`
        );
      }
    });

    // 割り当て結果を日次シートに反映
    fillLessonCodesToSheet(sortedLessons, dateSheet);

    Logger.log("講師割り当て処理が完了しました");
  } catch (error) {
    logError("講師割り当て処理でエラーが発生しました", error);
    throw error;
  }
}

/**
 * 講義コードに対応する優先順位リストを取得（優先順位③の下も含む）
 * @param {string} lessonCode - 講義コード（例："1M"）
 * @param {Sheet} prioritySheet - 優先順位シート
 * @returns {Array} 講師名の配列（優先順位順）
 */
function getPriorityList(lessonCode, prioritySheet) {
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

      return priorityList;
    }
  }

  return [];
}

/**
 * 講師が指定されたコマで利用可能かどうかを判定
 * @param {string} teacherName - 講師名
 * @param {Sheet} dateSheet - 日次シート
 * @param {Object} assignedTeachers - 割り当て済み講師の管理オブジェクト
 * @param {number} period - コマ数
 * @returns {boolean} 利用可能な場合true
 */
function isAvailable(teacherName, dateSheet, assignedTeachers, period) {
  // 既にそのコマに割り当て済みの場合はfalse
  if (assignedTeachers[period].has(teacherName)) {
    return false;
  }

  // 講師の列を特定
  const teacherCol = findTeacherColumn(teacherName, dateSheet);
  if (teacherCol === -1) {
    return false;
  }

  // 希望行の値を確認
  const wishValue = dateSheet
    .getRange(DAILY_SHEET.WISH_ROW, teacherCol)
    .getValue();

  // 希望がWISH_TRUEの場合のみtrue
  return wishValue === STRINGS.WISH.TRUE;
}

/**
 * 講師名から日次シートの列番号を取得
 * @param {string} teacherName - 講師名
 * @param {Sheet} dateSheet - 日次シート
 * @returns {number} 列番号（見つからない場合は-1）
 */
function findTeacherColumn(teacherName, dateSheet) {
  const staffRow = DAILY_SHEET.STAFF_ROW;
  const lastCol = dateSheet.getLastColumn();

  for (let col = DAILY_SHEET.STAFF_START_COL; col <= lastCol; col++) {
    const cellValue = dateSheet.getRange(staffRow, col).getValue();
    if (cellValue === teacherName) {
      return col;
    }
  }

  return -1;
}

/**
 * 講義情報を日次シートに反映
 * @param {Array} lessons - 講義オブジェクトの配列
 * @param {Sheet} dateSheet - 日次シート
 */
function fillLessonCodesToSheet(lessons, dateSheet) {
  try {
    // 既存の授業エリアをクリア
    resetLessonArea(dateSheet);

    // 各講義の情報をシートに反映
    lessons.forEach((lesson) => {
      if (!lesson.assignedTeacher) return;

      // 講師の列を特定
      const teacherCol = findTeacherColumn(lesson.assignedTeacher, dateSheet);
      if (teacherCol === -1) return;

      // コマ数に応じた行を特定
      const lessonRow = DAILY_SHEET.LESSON_ROWS[getPeriodKey(lesson.period)];
      if (!lessonRow) return;

      // セルに講義コードを設定
      const cell = dateSheet.getRange(lessonRow, teacherCol);
      cell.setValue(lesson.lessonCode);

      // スタイルを適用
      if (lesson.style) {
        applyCellStyle(cell, lesson.style);
      }
    });

    Logger.log("講義情報のシート反映が完了しました");
  } catch (error) {
    logError("講義情報のシート反映でエラーが発生しました", error);
    throw error;
  }
}

/**
 * 授業エリアをリセット
 * @param {Sheet} dateSheet - 日次シート
 */
function resetLessonArea(dateSheet) {
  const staffRow = DAILY_SHEET.STAFF_ROW;
  const lastCol = dateSheet.getLastColumn();

  // 各コマの行をクリア
  Object.values(DAILY_SHEET.LESSON_ROWS).forEach((row) => {
    const clearRange = dateSheet.getRange(
      row,
      DAILY_SHEET.STAFF_START_COL,
      1,
      lastCol - DAILY_SHEET.STAFF_START_COL + 1
    );
    clearRange.clearContent();
    clearRange.setBackground(null);
  });
}

/**
 * コマ数から設定キーを取得
 * @param {number} period - コマ数
 * @returns {string} 設定キー
 */
function getPeriodKey(period) {
  const periodMap = {
    1: "FIRST",
    2: "SECOND",
    3: "THIRD",
  };
  return periodMap[period];
}

/**
 * セルにスタイルを適用
 * @param {Range} cell - セルオブジェクト
 * @param {Object} style - スタイル情報オブジェクト
 */
function applyCellStyle(cell, style) {
  try {
    if (style.backgroundColor) {
      cell.setBackground(style.backgroundColor);
    }
    if (style.fontColor) {
      cell.setFontColor(style.fontColor);
    }
    if (style.borders) {
      cell.setBorder(
        style.borders.top,
        style.borders.left,
        style.borders.bottom,
        style.borders.right,
        style.borders.vertical,
        style.borders.horizontal
      );
    }
    if (style.fontFamily) {
      cell.setFontFamily(style.fontFamily);
    }
    if (style.fontSize) {
      cell.setFontSize(style.fontSize);
    }
    if (style.fontBold !== undefined) {
      cell.setFontWeight(style.fontBold ? "bold" : "normal");
    }
    if (style.horizontalAlignment) {
      cell.setHorizontalAlignment(style.horizontalAlignment);
    }
    if (style.verticalAlignment) {
      cell.setVerticalAlignment(style.verticalAlignment);
    }
  } catch (error) {
    logError("セルスタイル適用でエラーが発生しました", error);
  }
}

/**
 * 複数のセルにスタイルを一括適用
 * @param {Array} cells - セルオブジェクトの配列
 * @param {Array} styles - スタイル情報オブジェクトの配列
 */
function applyCellStylesBatch(cells, styles) {
  cells.forEach((cell, index) => {
    if (styles[index]) {
      applyCellStyle(cell, styles[index]);
    }
  });
}

/**
 * 割り当て結果のサマリーを出力
 * @param {Array} lessons - 講義オブジェクトの配列
 */
function printAssignmentSummary(lessons) {
  const assignedLessons = lessons.filter((lesson) => lesson.assignedTeacher);
  const unassignedLessons = lessons.filter((lesson) => !lesson.assignedTeacher);

  Logger.log("=== 講師割り当て結果サマリー ===");
  Logger.log(`総講義数: ${lessons.length}`);
  Logger.log(`割り当て成功: ${assignedLessons.length}`);
  Logger.log(`割り当て失敗: ${unassignedLessons.length}`);

  if (unassignedLessons.length > 0) {
    Logger.log("割り当て失敗した講義:");
    unassignedLessons.forEach((lesson) => {
      Logger.log(`- ${lesson.lessonCode} (${lesson.grade}${lesson.subject})`);
    });
  }
}

/**
 * 講義の優先度を計算
 * @param {Object} lesson - 講義オブジェクト
 * @returns {number} 優先度（数値が大きいほど優先度が高い）
 */
function calculateLessonPriority(lesson) {
  let priority = 0;

  // 学年による優先度（高学年ほど優先度が高い）
  priority += lesson.gradeNumber * 10;

  // 教科による優先度（算数 > 国語 > 理科 > 社会）
  const subjectPriority = { M: 4, J: 3, R: 2, S: 1 };
  priority += subjectPriority[lesson.subjectCode] || 0;

  return priority;
}

/**
 * 講義を優先度順にソート
 * @param {Array} lessons - 講義オブジェクトの配列
 * @returns {Array} ソートされた講義オブジェクトの配列
 */
function sortLessonsByPriority(lessons) {
  return lessons.slice().sort((a, b) => {
    const priorityA = calculateLessonPriority(a);
    const priorityB = calculateLessonPriority(b);
    return priorityB - priorityA; // 降順（優先度が高い順）
  });
}
