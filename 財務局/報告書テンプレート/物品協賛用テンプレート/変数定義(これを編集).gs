// グローバル変数
const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
const reportSheet1 = spreadsheet.getSheetByName("①報告書1");
const reportSheet2 = spreadsheet.getSheetByName("②報告書2");
const clientSheet = spreadsheet.getSheetByName("③送付一覧");
const executionSheet = spreadsheet.getSheetByName("④実行");
const logSheet = spreadsheet.getSheetByName("実行ログ");

// セルの定義
const CELL_MAPPING = {
  issueNumber: "A",         // 発行番号
  companyName: "B",         // 企業名
  contactName: "C",         // 担当者名
  sponsorshipItem: "D",     // 協賛品
  distributionCount: "E",   // 配布数
  remainingCount: "F",      // 残数
  distributionPlace: "G",   // 配布場所
  distributionDetails: "H", // 配布の詳細
  otherDetails: "I",        // その他
  folderUrl: "A3",          // フォルダのURL
  issueDate: "A7",          // 発行日
  string: "A11",            // 中間文字
  folderUrlForPDF: "A21",   // PDF変換後のフォルダ
};

// テンプレートのセル
const TEMPLATE_CELLS = {
  issueNumber: "Z1",         // 発行番号
  issueDate: "Z2",           // 発行日
  companyName: "A4",         // 企業名
  contactName: "Y5",         // 担当者名
  sponsorshipItem: "G39",    // 協賛品名
  distributionCount: "G41",  // 配布数
  remainingCount: "V41",     // 残数
  distributionPlace: "G43",  // 配布場所
  distributionDetails: "G45",// 配布の詳細
  otherDetails: "G47"        // その他
};

// ファイル名のフォーマット
const FILE_NAME_FORMAT = "{issueNumber}_{string}_{companyName}";
