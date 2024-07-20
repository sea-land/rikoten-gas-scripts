/**
 * 指定された範囲の行をチャンクで処理する。
 */
function processChunk(startRow, endRow, chunkSize) {
  const chunkEndRow = Math.min(startRow + chunkSize - 1, endRow);

  // まとめて行データを取得
  const rowDataArray = getRowDataInBatch(startRow, chunkEndRow);

  rowDataArray.forEach((rowData, i) => {
    const rowIndex = startRow + i;
    if (rowData.companyName) {
      try {
        logToSheet(`処理中: ${rowIndex} 行目 (${rowData.companyName})`);
        fillTemplateSheets(rowData);
        createNewSpreadsheet(rowData);
      } catch (error) {
        handleError(rowIndex, error, rowData.companyName);
      }
    }
  });

  if (chunkEndRow < endRow) {
    continueProcessing(chunkEndRow, endRow, chunkSize);
  } else {
    logToSheet("全ての処理が完了しました");
  }
}

/**
 * 行データをバッチで取得する。
 */
function getRowDataInBatch(startRow, endRow) {
  const range = clientSheet.getRange(
    `${CELL_MAPPING.issueNumber}${startRow}:${CELL_MAPPING.otherDetails}${endRow}`
  );
  const values = range.getValues();

  return values.map((row) => ({
    issueNumber: row[0],
    companyName: row[1],
    contactName: row[2],
    sponsorshipItem: row[3],
    distributionCount: row[4],
    remainingCount: row[5],
    distributionPlace: row[6],
    distributionDetails: row[7],
    otherDetails: row[8],
  }));
}

/**
 * テンプレートシートにデータを埋め込む。
 */
function fillTemplateSheets(data) {
  const templateData = [
    [
      data.issueNumber,
      executionSheet.getRange(CELL_MAPPING.issueDate).getValue(),
    ],
    [data.companyName, data.contactName],
    [data.sponsorshipItem, data.distributionCount, data.remainingCount],
    [data.distributionPlace, data.distributionDetails, data.otherDetails],
  ];

  reportSheet1
    .getRange(TEMPLATE_CELLS_1.issueNumber)
    .setValue(templateData[0][0]);
  reportSheet1
    .getRange(TEMPLATE_CELLS_1.issueDate)
    .setValue(templateData[0][1]);
  reportSheet1
    .getRange(TEMPLATE_CELLS_1.companyName)
    .setValue(templateData[1][0]);
  reportSheet1
    .getRange(TEMPLATE_CELLS_1.contactName)
    .setValue(templateData[1][1]);
  reportSheet1
    .getRange(TEMPLATE_CELLS_1.sponsorshipItem)
    .setValue(templateData[2][0]);
  reportSheet1
    .getRange(TEMPLATE_CELLS_1.distributionCount)
    .setValue(templateData[2][1]);
  reportSheet1
    .getRange(TEMPLATE_CELLS_1.remainingCount)
    .setValue(templateData[2][2]);
  reportSheet1
    .getRange(TEMPLATE_CELLS_1.distributionPlace)
    .setValue(templateData[3][0]);
  reportSheet1
    .getRange(TEMPLATE_CELLS_1.distributionDetails)
    .setValue(templateData[3][1]);
  reportSheet1
    .getRange(TEMPLATE_CELLS_1.otherDetails)
    .setValue(templateData[3][2]);
  reportSheet2
    .getRange(TEMPLATE_CELLS_1.issueDate)
    .setValue(templateData[0][1]);

  SpreadsheetApp.flush();
}

/**
 * 新しいスプレッドシートを作成し、テンプレートシートをコピーする。
 */
function createNewSpreadsheet(data) {
  const fileName = FILE_NAME_FORMAT.replace(
    "{issueNumber}",
    data.issueNumber
  ).replace("{companyName}", data.companyName);
  const newSpreadsheet = SpreadsheetApp.create(fileName);
  const newSpreadsheetId = newSpreadsheet.getId();
  reportSheet1.copyTo(newSpreadsheet).setName(`${fileName}_1`);
  reportSheet2.copyTo(newSpreadsheet).setName(`${fileName}_2`);
  newSpreadsheet.deleteSheet(newSpreadsheet.getSheets()[0]);

  const folderUrl = executionSheet.getRange(CELL_MAPPING.folderUrl).getValue();
  const folderId = extractFolderIdFromUrl(folderUrl);
  if (!folderId) {
    throw new Error("フォルダのURLが無効です");
  }

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
 * エラーハンドリング
 */
function handleError(rowIndex, error, companyName) {
  logToSheet(`エラー: ${rowIndex} 行目 (${companyName}) - ${error.message}`);
}

/**
 * 次のチャンクを処理する
 */
function continueProcessing(chunkEndRow, endRow, chunkSize) {
  ScriptApp.newTrigger("processChunk")
    .timeBased()
    .after(1000) // 1秒後に次のチャンクを処理
    .create();
}
