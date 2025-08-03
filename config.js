/**
 * 設定ファイル
 * アプリケーション固有の定数を管理
 *
 * 注意: 環境別設定（スプレッドシートID、フォルダIDなど）は
 * env-config.js で管理しています。
 */

// ===== シート名 =====
const SHEET_NAMES = {
  MAIN: "Main",
  TEMPLATE_DAILY: "Template_Daily",
  TEMPLATE_STAFF: "Template_Staff",
  PRIORITY: "Priority",
  WEEKDAYS: {
    MON: "月",
    TUE: "火",
    WED: "水",
    THU: "木",
    FRI: "金",
    SAT: "土",
    SUN: "日",
  },
};

// ===== メインシート設定 =====
const MAIN_SHEET = {
  // 日程リスト
  DATE: {
    COL: 1,
    START_ROW: 4,
    END_ROW: 10,
  },
  // スタッフリスト
  STAFF: {
    ID_COL: 3,
    NAME_COL: 4,
    DISPLAY_COL: 5,
    SUBMIT_COL: 6,
    REFLECT_COL: 7,
    START_ROW: 4,
    END_ROW: 13,
  },
};

// ===== 日次シート設定 =====
const DAILY_SHEET = {
  DATE_ROW: 1,
  DATE_COL: 1,
  STAFF_ROW: 1,
  STAFF_START_COL: 2,
  WISH_ROW: 2,
  LESSON_ROWS: {
    FIRST: 3,
    SECOND: 4,
    THIRD: 5,
  },
};

// ===== 曜日シート設定 =====
const WEEK_SHEET = {
  PERIOD_ROWS: {
    FIRST: 3,
    SECOND: 4,
    THIRD: 5,
  },
  GRADE_COLS: {
    YOUNG: 2,
    FIRST: 3,
    SECOND: 4,
    THIRD: 5,
    FOURTH: 6,
    FIFTH: 7,
    SIXTH: 8,
  },
};

// ===== スタッフシート設定 =====
const STAFF_SHEET = {
  NAME_ROW: 1,
  NAME_COL: 2,
  CHECK_ROW: 1,
  CHECK_COL: 4,
  DATE: {
    START_ROW: 5,
    END_ROW: 11,
    COL: 1,
  },
  WISH_COL: 2,
  GRADE_ROWS: {
    YOUNG: 5,
    FIRST: 6,
    SECOND: 7,
    THIRD: 8,
    FOURTH: 9,
    FIFTH: 10,
    SIXTH: 11,
  },
  SUBJECT_COLS: {
    MATH: 5,
    JAPANESE: 6,
    SCIENCE: 7,
    SOCIAL: 8,
    ALGO: 9,
    PRE: 10,
  },
};

// ===== 優先順位シート設定 =====
const PRIORITY_SHEET = {
  LESSON_ROW: 1,
  PRIORITY_ROWS: {
    FIRST: 2,
    SECOND: 3,
    THIRD: 4,
  },
};

// ===== 講義コード定義 =====
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

// ===== 文字列定数 =====
const STRINGS = {
  WISH: {
    TRUE: "◯",
    FALSE: "×",
  },
  SUBMIT: {
    TRUE: "✅提出済み",
    FALSE: "未提出",
  },
  REFLECT: {
    TRUE: "✅反映済み",
    FALSE: "未反映",
  },
};

// ===== 注意事項 =====
//
// このファイル（config.js）には環境に依存しない
// アプリケーション固有の定数のみを定義してください。
//
// 環境別設定（スプレッドシートID、フォルダIDなど）は
// env-config.js で管理しています。
