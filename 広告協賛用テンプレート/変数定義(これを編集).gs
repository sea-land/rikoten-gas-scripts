// グローバル変数
const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
const reportSheet1 = spreadsheet.getSheetByName("①報告書1");
const reportSheet2 = spreadsheet.getSheetByName("②報告書2");
const clientSheet = spreadsheet.getSheetByName("③送付一覧");
const executionSheet = spreadsheet.getSheetByName("④実行");
const logSheet = spreadsheet.getSheetByName("実行ログ");

// セルの定義
const CELL_MAPPING = {
  issueNumber: "A", // 発行番号
  companyName: "B", // 企業名
  contactName: "C", // 担当者名
  folderUrl: "A3", // フォルダのURL
  issueDate: "A7", // 発行日
};

// テンプレートのセル
const TEMPLATE_CELLS = {
  issueNumber: "Z1", // 発行番号
  issueDate: "Z2", // 発行日
  companyName: "A4", // 企業名
  contactName: "Y5", // 担当者名
};

// ファイル名のフォーマット
const FILE_NAME_FORMAT = "{issueNumber}_第70回理工展報告書_{companyName}";
