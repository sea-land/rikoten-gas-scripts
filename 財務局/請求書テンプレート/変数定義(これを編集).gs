// グローバル変数
const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
const reportSheet1 = spreadsheet.getSheetByName("①請求書");
const reportSheet2 = spreadsheet.getSheetByName("②広告料金表");
const clientSheet = spreadsheet.getSheetByName("③送付一覧");
const executionSheet = spreadsheet.getSheetByName("④実行");
const logSheet = spreadsheet.getSheetByName("実行ログ");

// セルの定義
const CELL_MAPPING = {
  issueNumber: "A", // 発行番号
  companyName: "B", // 企業名
  contactName: "C", // 担当者名
  ad1: "D", // ad1
  ad2: "G", // ad2
  ad3: "J", // ad3
  ad4: "M", // ad4
  ad5: "P", // ad5
  ad1count: "E", // ad1
  ad2count: "H", // ad2
  ad3count: "K", // ad3
  ad4count: "N", // ad4
  ad5count: "Q", // ad5
  folderUrl: "A3", // フォルダのURL
  issueDate: "A7", // 発行日
};

// テンプレートのセル
const TEMPLATE_CELLS = {
  issueNumber: "Z1", // 発行番号
  issueDate: "Z2", // 発行日
  companyName: "A4", // 企業名
  contactName: "Y5", // 担当者名
  ad1: "A13", // ad1
  ad2: "A14", // ad2
  ad3: "A15", // ad3
  ad4: "A16", // ad4
  ad5: "A17", // ad5
  ad1count: "Y13", // ad1
  ad2count: "Y14", // ad2
  ad3count: "Y15", // ad3
  ad4count: "Y16", // ad4
  ad5count: "Y17", // ad5
};

// ファイル名のフォーマット
const FILE_NAME_FORMAT = "{issueNumber}_第70回理工展報告書_{companyName}";
