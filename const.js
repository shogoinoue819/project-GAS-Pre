// ===== シート名 =====

// メインシート名
const MAIN = "Main";
// 日次テンプレートシート名
const TEMPLATE_DAILY = "Template_Daily";
// スタッフテンプレートシート名
const TEMPLATE_STAFF = "Template_Staff";
// 曜日シート名
const WEEK_MON = "月";
const WEEK_TUE = "火";
const WEEK_WED = "水";
const WEEK_THU = "木";
const WEEK_FRI = "金";
const WEEK_SAT = "土";
const WEEK_SUN = "日";
// 優先順位シート名
const PRIORITY = "Priority";

// =====　メインシート 行列インデックス =====

// 日程リスト列
const MAIN_DATE_COL = 1;
// 日程リスト開始行
const MAIN_DATE_START_ROW = 4;
// 日程リスト終了行
const MAIN_DATE_END_ROW = 10;

// スタッフリストID列
const MAIN_STAFF_ID_COL = 3;
// スタッフリスト氏名列
const MAIN_STAFF_NAME_COL = 4;
// スタッフリスト表示名列
const MAIN_STAFF_DISPLAY_COL = 5;
// スタッフリスト提出列
const MAIN_STAFF_SUBMIT_COL = 6;
// スタッフリスト反映列
const MAIN_STAFF_REFLECT_COL = 7;
// スタッフリスト開始行
const MAIN_STAFF_START_ROW = 4;
// スタッフリスト終了行
const MAIN_STAFF_END_ROW = 13;

// =====　日次シート 行列インデックス =====

// 日付行
const DAILY_DATE_ROW = 1;
// 日付列
const DAILY_DATE_COL = 1;
// スタッフ表示名行
const DAILY_STAFF_ROW = 1;
// スタッフ表示名開始列
const DAILY_STAFF_START_COL = 2;
// 希望行
const DAILY_WISH_ROW = 2;
// １コマ目行
const DAILY_LESSON1_ROW = 3;
// ２コマ目行
const DAILY_LESSON2_ROW = 4;
// ３コマ目行
const DAILY_LESSON3_ROW = 5;

// ===== 曜日シート 行列インデックス =====

// 1コマ目行
const WEEK_PERIOD1_ROW = 3;
// 2コマ目行
const WEEK_PERIOD2_ROW = 4;
// 3コマ目行
const WEEK_PERIOD3_ROW = 5;

// 年長列
const WEEK_YOUNG_COL = 2;
// 小１列
const WEEK_FIRST_COL = 3;
// 小２列
const WEEK_SECOND_COL = 4;
// 小３列
const WEEK_THIRD_COL = 5;
// 小４列
const WEEK_FOURTH_COL = 6;
// 小５列
const WEEK_FIFTH_COL = 7;
// 小６列
const WEEK_SIXTH_COL = 8;

// ===== スタッフシート 行列インデックス =====

// スタッフ氏名行
const STAFF_NAME_ROW = 1;
// スタッフ氏名列
const STAFF_NAME_COL = 2;
// チェック行
const STAFF_CHECK_ROW = 1;
// チェック列
const STAFF_CHECK_COL = 4;

// 日程リスト開始行
const STAFF_DATE_START_ROW = 5;
// 日程リスト終了行
const STAFF_DATE_END_ROW = 11;
// 日程リスト列
const STAFF_DATE_COL = 1;
// 希望列
const STAFF_WISH_COL = 2;

// 年長行
const STAFF_YOUNG_ROW = 5;
// 小１行
const STAFF_FIRST_ROW = 6;
// 小２行
const STAFF_SECOND_ROW = 7;
// 小３行
const STAFF_THIRD_ROW = 8;
// 小４行
const STAFF_FOURTH_ROW = 9;
// 小５行
const STAFF_FIFTH_ROW = 10;
// 小６行
const STAFF_SIXTH_ROW = 11;

// 算数列
const STAFF_MAT_COL = 5;
// 国語列
const STAFF_JAP_COL = 6;
// 理科列
const STAFF_SCI_COL = 7;
// 社会列
const STAFF_SOC_COL = 8;
// アルゴ列
const STAFF_ALGO_COL = 9;
// プレ列
const STAFF_PRE_COL = 10;

// ===== 優先順位シート 行列インデックス =====

// 講義コード行
const PRIORITY_LESSON_ROW = 1;
// 優先順位①行
const PRIORITY_FIRST_ROW = 2;
// 優先順位②行
const PRIORITY_SECOND_ROW = 3;
// 優先順位③行
const PRIORITY_THIRD_ROW = 4;

// ===== 講義コード定義 =====

// 講義コードの定義（24通り）
// 形式: 学年(1-6) + 教科(M/J/R/S)
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

// 希望◯
const WISH_TRUE = "◯";
// 希望×
const WISH_FALSE = "×";

// 提出済み
const SUBMIT_TRUE = "✅提出済み";
// 未提出
const SUBMIT_FALSE = "未提出";

// 反映済み
const REFLECT_TRUE = "✅反映済み";
// 未反映
const REFLECT_FALSE = "未反映";
