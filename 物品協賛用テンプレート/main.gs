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
  sponsorshipItem: "D", // 協賛品
  distributionCount: "E", // 配布数
  remainingCount: "F", // 残数
  distributionPlace: "G", // 配布場所
  distributionDetails: "H", // 配布の詳細
  otherDetails: "I", // その他
  folderUrl: "A3", // フォルダのURL
  issueDate: "A7", // 発行日
};

// テンプレートのセル
const TEMPLATE_CELLS_1 = {
  issueNumber: "Y1", // 発行番号
  issueDate: "Y2", // 発行日
  companyName: "A4", // 企業名
  contactName: "W5", // 担当者名
  sponsorshipItem: "G39", // 協賛品名
  distributionCount: "G41", // 配布数
  remainingCount: "P41", // 残数
  distributionPlace: "G43", // 配布場所
  distributionDetails: "G45", // 配布の詳細
  otherDetails: "G47", // その他
};

// ファイル名のフォーマット
const FILE_NAME_FORMAT = "{issueNumber}_第70回理工展実施報告書_{companyName}";

function addMenu() {
  spreadsheet.addMenu("プログラム", [
    { name: "すべての報告書を作成", functionName: "showDialog" },
    { name: "選択した企業を作成", functionName: "showDialogAndInputNumbers" },
  ]);
}

/**
 * スクリプト実行時に既存のトリガーを削除する
 */
function deleteAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    ScriptApp.deleteTrigger(trigger);
  }
}

/**
 * コピー処理のエントリーポイント。
 */
function copyTemplateSheets(startRow = 4, endRow = 10, chunkSize = 20) {
  deleteAllTriggers(); // 既存のトリガーを削除
  logToSheet("処理を開始します...");
  processChunk(startRow, endRow, chunkSize);
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
