/**
 * デフォルトの開始・終了行でコピー処理を開始するダイアログ表示
 */
function showDialog() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert("請求書の作成を実行します。", ui.ButtonSet.OK_CANCEL);

  if (response == ui.Button.OK) {
    logSheet.clear();
    logToSheet("請求書の作成を実行します。");
    copyTemplateSheets();
  }
}

/**
 * 行番号を入力してコピー処理を開始するダイアログ表示
 */
function showDialogAndInputNumbers() {
  const ui = SpreadsheetApp.getUi();

  const startRowDialog = ui.prompt(
    "初めの行を入力してください(4以上):",
    "例: 4",
    ui.ButtonSet.OK_CANCEL
  );

  if (startRowDialog.getSelectedButton() != ui.Button.OK) {
    return;
  }

  const startRow = parseInt(startRowDialog.getResponseText());

  if (isNaN(startRow) || startRow < 4) {
    ui.alert("初めの行は4以上の数字を入力してください。");
    return;
  }

  const endRowDialog = ui.prompt(
    "終わりの行を入力してください:",
    "例: 10",
    ui.ButtonSet.OK_CANCEL
  );

  if (endRowDialog.getSelectedButton() != ui.Button.OK) {
    return;
  }

  const endRow = parseInt(endRowDialog.getResponseText());

  if (isNaN(endRow)) {
    ui.alert("終わりの行は数字を入力してください。");
    return;
  }

  ui.alert("請求書の作成を実行します。");
  logSheet.clear();
  logToSheet("請求書の作成を実行します。");
  copyTemplateSheets(startRow, endRow);
}
