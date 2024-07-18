// グローバル変数
const ss = SpreadsheetApp.getActiveSpreadsheet();
const templateSheet1 = ss.getSheetByName("①報告書1");
const templateSheet2 = ss.getSheetByName("②報告書2");
const clientListSheet = ss.getSheetByName("③送付一覧");
const executeSheet = ss.getSheetByName("④実行");
const logSheet = ss.getSheetByName("実行ログ");

// セルの定義
const CELL_MAPPING = {
  number: "A",
  companyName: "B",
  name: "C",
  item: "D",
  itemCount: "E",
  itemRemain: "F",
  place: "G",
  detail: "H",
  other: "I",
  folderId: "A3",
  executionDate: "A8",
};

// テンプレートのセル
const TEMPLATE_CELLS_1 = {
  number: "Y1",
  date: "Y2",
  companyName: "A4",
  name: "W5",
  item: "G39",
  itemCount: "G41",
  itemRemain: "P41",
  place: "G43",
  detail: "G45",
  other: "G47",
};

// ファイル名のフォーマット
const FILE_NAME_FORMAT = "{number}_第70回理工展実施報告書_{companyName}";

// スクリプト実行時に既存のトリガーを削除する
function deleteAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    ScriptApp.deleteTrigger(trigger);
  }
}

/**
 * コピー処理のエントリーポイント。
 */
function copyTemplateSheets(start = 4, end = 400, chunkSize = 20) {
  deleteAllTriggers(); // 既存のトリガーを削除
  logToSheet("処理を開始します...");
  processChunk(start, end, chunkSize);
}

/**
 * 指定された範囲の行をチャンクで処理する。
 */
function processChunk(start, end, chunkSize) {
  const maxRetries = 3;
  const chunkEnd = Math.min(start + chunkSize - 1, end);

  for (let i = start; i <= chunkEnd; i++) {
    let retryCount = 0;
    while (retryCount < maxRetries) {
      try {
        const rowData = getRowData(i);
        if (rowData.companyName) {
          updateProgress(`処理中: ${i} 行目 (${rowData.companyName})`);
          fillTemplateSheets(rowData);
          createNewSpreadsheet(rowData);
        }
        break; // 成功したらループを抜ける
      } catch (error) {
        retryCount++;
        if (retryCount === maxRetries) {
          handleError(i, error, rowData.companyName);
        } else {
          Utilities.sleep(1000); // 1秒待ってから再試行
        }
      }
    }
  }

  if (chunkEnd < end) {
    continueProcessing(chunkEnd, end, chunkSize);
  } else {
    updateProgress("全ての処理が完了しました");
  }
}

/**
 * 行データをバッチで取得する。
 */
function getRowDataInBatch(start, end) {
  const range = clientListSheet.getRange(
    `${CELL_MAPPING.number}${start}:${CELL_MAPPING.other}${end}`
  );
  const values = range.getValues();

  return values.map((row) => ({
    number: row[0],
    companyName: row[1],
    name: row[2],
    item: row[3],
    itemCount: row[4],
    itemRemain: row[5],
    place: row[6],
    detail: row[7],
    other: row[8],
  }));
}

/**
 * テンプレートシートにデータを埋め込む。
 */
function fillTemplateSheets(data) {
  const templateData = [
    [data.number, executeSheet.getRange(CELL_MAPPING.executionDate).getValue()],
    [data.companyName, data.name],
    [data.item, data.itemCount, data.itemRemain],
    [data.place, data.detail, data.other],
  ];

  templateSheet1.getRange(TEMPLATE_CELLS_1.number).setValue(templateData[0][0]);
  templateSheet1.getRange(TEMPLATE_CELLS_1.date).setValue(templateData[0][1]);
  templateSheet1
    .getRange(TEMPLATE_CELLS_1.companyName)
    .setValue(templateData[1][0]);
  templateSheet1.getRange(TEMPLATE_CELLS_1.name).setValue(templateData[1][1]);
  templateSheet1.getRange(TEMPLATE_CELLS_1.item).setValue(templateData[2][0]);
  templateSheet1
    .getRange(TEMPLATE_CELLS_1.itemCount)
    .setValue(templateData[2][1]);
  templateSheet1
    .getRange(TEMPLATE_CELLS_1.itemRemain)
    .setValue(templateData[2][2]);
  templateSheet1.getRange(TEMPLATE_CELLS_1.place).setValue(templateData[3][0]);
  templateSheet1.getRange(TEMPLATE_CELLS_1.detail).setValue(templateData[3][1]);
  templateSheet1.getRange(TEMPLATE_CELLS_1.other).setValue(templateData[3][2]);

  templateSheet2.getRange(TEMPLATE_CELLS_1.date).setValue(templateData[0][1]);

  SpreadsheetApp.flush();
}

/**
 * 新しいスプレッドシートを作成し、テンプレートシートをコピーする。
 */
function createNewSpreadsheet(data) {
  const fileName = FILE_NAME_FORMAT.replace("{number}", data.number).replace(
    "{companyName}",
    data.companyName
  );
  const newSpreadsheet = SpreadsheetApp.create(fileName);
  const newSpreadsheetId = newSpreadsheet.getId();
  templateSheet1.copyTo(newSpreadsheet).setName(`${fileName}_1`);
  templateSheet2.copyTo(newSpreadsheet).setName(`${fileName}_2`);
  newSpreadsheet.deleteSheet(newSpreadsheet.getSheets()[0]);

  const folderId = executeSheet.getRange(CELL_MAPPING.folderId).getValue();
  const folder = DriveApp.getFolderById(folderId);
  const file = DriveApp.getFileById(newSpreadsheetId);

  removeExistingFiles(folder, file.getName());
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);

  logToSheet(`${fileName} を作成しました。`);
}

/**
 * 指定されたフォルダ内の同名の既存ファイルを削除する。
 */
function removeExistingFiles(folder, fileName) {
  const existingFiles = folder.getFilesByName(fileName);
  while (existingFiles.hasNext()) {
    const existingFile = existingFiles.next();
    folder.removeFile(existingFile);
    DriveApp.getRootFolder().removeFile(existingFile);
  }
}

/**
 * エラーを処理する。
 */
function handleError(row, error, companyName) {
  logToSheet(`エラー発生: ${row} 行目 (${companyName}) - ${error.message}`);
}

/**
 * 処理の続きを行う。
 */
function continueProcessing(chunkEnd, end, chunkSize) {
  ScriptApp.newTrigger("continueCopyTemplateSheets")
    .timeBased()
    .after(1000)
    .create();
  PropertiesService.getScriptProperties().setProperty(
    "nextStartRow",
    chunkEnd + 1
  );
  PropertiesService.getScriptProperties().setProperty("endRow", end);
  PropertiesService.getScriptProperties().setProperty("chunkSize", chunkSize);
}

/**
 * コピー処理の続行。
 */
function continueCopyTemplateSheets() {
  const nextStartRow = parseInt(
    PropertiesService.getScriptProperties().getProperty("nextStartRow")
  );
  const endRow = parseInt(
    PropertiesService.getScriptProperties().getProperty("endRow")
  );
  const chunkSize = parseInt(
    PropertiesService.getScriptProperties().getProperty("chunkSize")
  );
  processChunk(nextStartRow, endRow, chunkSize);
}

/**
 * 行番号を入力してコピー処理を開始するダイアログを表示する。
 */
function showDialogAndInputNumbers() {
  const ui = SpreadsheetApp.getUi();
  const result1 = ui.prompt(
    "初めの行を入力してください(4以上):",
    "例: 4",
    ui.ButtonSet.OK_CANCEL
  );
  const result2 = ui.prompt(
    "終わりの行を入力してください:",
    "例: 10",
    ui.ButtonSet.OK_CANCEL
  );

  if (
    result1.getSelectedButton() == ui.Button.OK &&
    result2.getSelectedButton() == ui.Button.OK
  ) {
    const number1 = parseInt(result1.getResponseText());
    const number2 = parseInt(result2.getResponseText());

    if (isNaN(number1) || isNaN(number2)) {
      ui.alert("数字が入力されませんでした。");
      return;
    } else if (number1 < 4) {
      ui.alert("初めの行は4以上を入力してください");
      return;
    }

    ui.alert("報告書の作成を実行します。");
    logSheet.clear();
    logToSheet("報告書の作成を実行します。");
    copyTemplateSheets(number1, number2);
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

/**
 * ログをシートに記録する
 */
function logToSheet(message) {
  const timestamp = new Date().toISOString();
  logSheet.appendRow([timestamp, message]);
}
