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

/**
 * デフォルトの開始・終了行でコピー処理を開始するダイアログを表示する。
 */
function showDialog() {
  const ui = SpreadsheetApp.getUi();
  ui.alert("報告書の作成を実行します。", ui.ButtonSet.OK_CANCEL);
  logSheet.clear();
  logToSheet("報告書の作成を実行します。");
  copyTemplateSheets();
}

/**
 * 行番号を入力してコピー処理を開始するダイアログを表示する。
 */
function showDialogAndInputNumbers() {
  const ui = SpreadsheetApp.getUi();
  const startRowDialog = ui.prompt(
    "初めの行を入力してください(4以上):",
    "例: 4",
    ui.ButtonSet.OK_CANCEL
  );
  const endRowDialog = ui.prompt(
    "終わりの行を入力してください:",
    "例: 10",
    ui.ButtonSet.OK_CANCEL
  );

  if (
    startRowDialog.getSelectedButton() == ui.Button.OK &&
    endRowDialog.getSelectedButton() == ui.Button.OK
  ) {
    const startRow = parseInt(startRowDialog.getResponseText());
    const endRow = parseInt(endRowDialog.getResponseText());

    if (isNaN(startRow) || isNaN(endRow)) {
      ui.alert("数字が入力されませんでした。");
      return;
    } else if (startRow < 4) {
      ui.alert("初めの行は4以上を入力してください");
      return;
    }

    ui.alert("報告書の作成を実行します。");
    logSheet.clear();
    logToSheet("報告書の作成を実行します。");
    copyTemplateSheets(startRow, endRow);
  }
}
