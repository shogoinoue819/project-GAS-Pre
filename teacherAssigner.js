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
      .getRange(PRIORITY_LESSON_ROW, col)
      .getValue();

    if (lessonCodeCell === lessonCode) {
      // 優先順位リストを取得（2行目から最後まで）
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

      // 優先順位③の下に追加された講師たちも取得
      const lastRow = prioritySheet.getLastRow();
      for (let row = PRIORITY_THIRD_ROW + 1; row <= lastRow; row++) {
        const additionalTeacher = prioritySheet.getRange(row, col).getValue();
        if (additionalTeacher && additionalTeacher !== "") {
          priorityList.push(additionalTeacher);
        }
      }

      Logger.log(
        `講義${lessonCode}の優先順位リスト: ${priorityList.join(", ")}`
      );
      return priorityList;
    }
  }

  return null;
}

/**
 * 講師の勤務希望を確認（重複割り当てチェック付き）
 * @param {string} teacherName - 講師名（表示名）
 * @param {Sheet} dateSheet - 日次シート
 * @param {Object} assignedTeachers - 割り当て済み講師の管理オブジェクト
 * @param {number} period - コマ数
 * @returns {boolean} 勤務可能な場合true
 */
function isAvailable(teacherName, dateSheet, assignedTeachers, period) {
  // 日次シートで講師の列を検索
  const teacherCol = findTeacherColumn(teacherName, dateSheet);
  if (teacherCol === -1) {
    Logger.log(`講師「${teacherName}」の列が見つかりません`);
    return false;
  }

  // 希望行から勤務希望を取得
  const wishValue = dateSheet.getRange(DAILY_WISH_ROW, teacherCol).getValue();

  // 希望が「◯」でない場合は勤務不可
  if (wishValue !== WISH_TRUE) {
    Logger.log(`講師「${teacherName}」は勤務希望がありません（${wishValue}）`);
    return false;
  }

  // 同じコマに既に割り当て済みかチェック
  if (assignedTeachers[period].has(teacherName)) {
    Logger.log(`講師「${teacherName}」は${period}コマ目に既に割り当て済みです`);
    return false;
  }

  return true;
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

  // コマ部分を一旦リセット
  resetLessonArea(dateSheet);

  // 割り当て済みの講義をフィルタリング
  const assignedLessons = lessons.filter((lesson) => lesson.assignedTeacher);
  const unassignedLessons = lessons.filter((lesson) => !lesson.assignedTeacher);

  Logger.log(`割り当て済み講義数: ${assignedLessons.length}`);
  Logger.log(`未割り当て講義数: ${unassignedLessons.length}`);

  // 未割り当て講義の詳細をログ出力
  if (unassignedLessons.length > 0) {
    Logger.log("【未割り当て講義一覧】");
    unassignedLessons.forEach((lesson, index) => {
      Logger.log(
        `${index + 1}. ${lesson.lessonCode} (${lesson.grade}${
          lesson.subject
        }) - ${lesson.period}コマ目`
      );
    });
    Logger.log("※未割り当ての講義は日次シートに反映されません");
  }

  assignedLessons.forEach((lesson) => {
    // 講師の列を検索
    const teacherCol = findTeacherColumn(lesson.assignedTeacher, dateSheet);
    if (teacherCol !== -1) {
      // コマ位置の行を計算
      const lessonRow = WEEK_PERIOD1_ROW + lesson.period - 1;

      // 対象セルを取得
      const targetCell = dateSheet.getRange(lessonRow, teacherCol);

      // 講義コードをセット
      targetCell.setValue(lesson.lessonCode);

      // スタイル情報を適用
      if (lesson.style) {
        applyCellStyle(targetCell, lesson.style);
      }

      Logger.log(
        `講義コード「${lesson.lessonCode}」を${lesson.assignedTeacher}の${lesson.period}コマ目に設定しました（スタイル付き）`
      );
    } else {
      Logger.log(`講師「${lesson.assignedTeacher}」の列が見つかりません`);
    }
  });

  Logger.log("日次シートへの反映が完了しました");
}

/**
 * 日次シートのコマ部分をリセット
 * @param {Sheet} dateSheet - 日次シート
 */
function resetLessonArea(dateSheet) {
  Logger.log("コマ部分のリセットを開始します...");

  // スタッフの列数を取得
  const lastCol = dateSheet.getLastColumn();

  // コマ部分の範囲を定義（1行目と1列目は除外）
  // 行：DAILY_LESSON1_ROW から DAILY_LESSON3_ROW（3行目から5行目）
  // 列：DAILY_STAFF_START_COL から lastCol（2列目から最後まで）
  const startRow = DAILY_LESSON1_ROW;
  const endRow = DAILY_LESSON3_ROW;
  const startCol = DAILY_STAFF_START_COL;
  const endCol = lastCol;

  // コマ部分の範囲を取得
  const lessonRange = dateSheet.getRange(
    startRow,
    startCol,
    endRow - startRow + 1,
    endCol - startCol + 1
  );

  // 内容とスタイルをクリア
  lessonRange.clearContent();
  lessonRange.setBackground(null);
  lessonRange.setFontColor(null);
  lessonRange.setFontWeight("normal");
  lessonRange.setFontSize(10);
  lessonRange.setFontFamily("Arial");
  lessonRange.setHorizontalAlignment("center");
  lessonRange.setVerticalAlignment("middle");

  Logger.log(
    `コマ部分をリセットしました（${startRow}行目-${endRow}行目、${startCol}列目-${endCol}列目）`
  );
}

/**
 * セルにスタイル情報を適用
 * @param {Range} cell - 対象セル
 * @param {Object} style - スタイル情報オブジェクト
 */
function applyCellStyle(cell, style) {
  try {
    // 背景色を設定
    if (style.backgroundColor) {
      cell.setBackground(style.backgroundColor);
    }

    // 文字色を設定
    if (style.fontColor) {
      cell.setFontColor(style.fontColor);
    }

    // フォントファミリーを設定
    if (style.fontFamily) {
      cell.setFontFamily(style.fontFamily);
    }

    // フォントサイズを設定
    if (style.fontSize) {
      cell.setFontSize(style.fontSize);
    }

    // 太字設定
    if (style.fontBold !== undefined) {
      cell.setFontWeight(style.fontBold ? "bold" : "normal");
    }

    // 水平方向の配置を設定
    if (style.horizontalAlignment) {
      cell.setHorizontalAlignment(style.horizontalAlignment);
    }

    // 垂直方向の配置を設定
    if (style.verticalAlignment) {
      cell.setVerticalAlignment(style.verticalAlignment);
    }

    // ボーダーを設定
    if (style.borders) {
      cell.setBorder(
        style.borders.top,
        style.borders.right,
        style.borders.bottom,
        style.borders.left,
        style.borders.vertical,
        style.borders.horizontal
      );
    }
  } catch (error) {
    Logger.log(`スタイル適用エラー: ${error.message}`);
  }
}

/**
 * 複数セルに一括でスタイルを適用（将来的な拡張用）
 * @param {Array} cells - セルオブジェクトの配列
 * @param {Array} styles - スタイル情報オブジェクトの配列
 */
function applyCellStylesBatch(cells, styles) {
  if (cells.length !== styles.length) {
    Logger.log("セル数とスタイル数が一致しません");
    return;
  }

  cells.forEach((cell, index) => {
    applyCellStyle(cell, styles[index]);
  });
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

/**
 * 授業の優先度を計算
 * @param {Object} lesson - 講義オブジェクト
 * @returns {number} 優先度スコア（高いほど優先）
 */
function calculateLessonPriority(lesson) {
  let priority = 0;

  // 学年による優先度（高学年ほど優先）
  priority += lesson.gradeNumber * 10;

  // 教科による優先度
  const subjectPriority = {
    算数: 5, // 算数は最重要
    国語: 4, // 国語は重要
    理科: 3, // 理科は中程度
    社会: 2, // 社会は低め
  };
  priority += subjectPriority[lesson.subject] || 0;

  // コマによる優先度（早いコマほど優先）
  priority += (4 - lesson.period) * 2;

  return priority;
}

/**
 * 授業を優先度順にソート
 * @param {Array} lessons - 講義オブジェクトの配列
 * @returns {Array} 優先度順にソートされた講義配列
 */
function sortLessonsByPriority(lessons) {
  return lessons.slice().sort((a, b) => {
    const priorityA = calculateLessonPriority(a);
    const priorityB = calculateLessonPriority(b);
    return priorityB - priorityA; // 降順（優先度の高い順）
  });
}
