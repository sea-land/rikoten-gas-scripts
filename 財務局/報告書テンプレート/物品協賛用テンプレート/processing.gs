/**
 * 処理開始の開始点
 */
function copyTemplateSheets(startRow = 4, endRow = 400, chunkSize = 20) {
  deleteAllTriggers();
  logToSheet("処理を開始します...");
  processChunk(startRow, endRow, chunkSize);
}

/**
 * 指定された範囲の行をチャンクで処理する。
 */
function processChunk(startRow, endRow, chunkSize) {
  const chunkEndRow = Math.min(startRow + chunkSize - 1, endRow);
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
        return;
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
    [data.issueNumber, executionSheet.getRange(CELL_MAPPING.issueDate).getValue()],
    [data.companyName, data.contactName],
    [data.sponsorshipItem, data.distributionCount, data.remainingCount],
    [data.distributionPlace, data.distributionDetails, data.otherDetails],
  ];

  reportSheet1.getRange(TEMPLATE_CELLS.issueNumber).setValue(templateData[0][0]);
  reportSheet1.getRange(TEMPLATE_CELLS.issueDate).setValue(templateData[0][1]);
  reportSheet1.getRange(TEMPLATE_CELLS.companyName).setValue(templateData[1][0]);
  reportSheet1.getRange(TEMPLATE_CELLS.contactName).setValue(templateData[1][1]);
  reportSheet1.getRange(TEMPLATE_CELLS.sponsorshipItem).setValue(templateData[2][0]);
  reportSheet1.getRange(TEMPLATE_CELLS.distributionCount).setValue(templateData[2][1]);
  reportSheet1.getRange(TEMPLATE_CELLS.remainingCount).setValue(templateData[2][2]);
  reportSheet1.getRange(TEMPLATE_CELLS.distributionPlace).setValue(templateData[3][0]);
  reportSheet1.getRange(TEMPLATE_CELLS.distributionDetails).setValue(templateData[3][1]);
  reportSheet1.getRange(TEMPLATE_CELLS.otherDetails).setValue(templateData[3][2]);
  reportSheet2.getRange(TEMPLATE_CELLS.issueDate).setValue(templateData[0][1]);

  SpreadsheetApp.flush();
}

/**
 * 新しいスプレッドシートを作成し、テンプレートシートをコピーする。
 */
function createNewSpreadsheet(data) {
  const stringValue = executionSheet.getRange(CELL_MAPPING.string).getValue();
  const fileName = FILE_NAME_FORMAT.replace("{issueNumber}", data.issueNumber).replace("{string}", stringValue).replace("{companyName}", data.companyName);
  const newSpreadsheet = SpreadsheetApp.create(fileName);
  reportSheet1.copyTo(newSpreadsheet).setName(reportSheet1.getName());
  reportSheet2.copyTo(newSpreadsheet).setName(reportSheet2.getName());
  newSpreadsheet.deleteSheet(newSpreadsheet.getSheets()[0]);

  const folderUrl = executionSheet.getRange(CELL_MAPPING.folderUrl).getValue();
  const folderId = extractFolderIdFromUrl(folderUrl);
  if (!folderId) {
    throw new Error("フォルダのURLが無効です");
  }

  const folder = DriveApp.getFolderById(folderId);

  removeExistingFile(folder, fileName);
  DriveApp.getFileById(newSpreadsheet.getId()).moveTo(folder);

  logToSheet(`${fileName} を作成しました。`);
}
