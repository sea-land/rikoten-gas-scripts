/**
 * 処理開始の開始点
 */
function copyTemplateSheets(startRow = 4, endRow = 50, chunkSize = 20) {
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
    `${CELL_MAPPING.issueNumber}${startRow}:${CELL_MAPPING.contactName}${endRow}`
  );
  const values = range.getValues();

  return values.map((row) => ({
    issueNumber: row[0],
    companyName: row[1],
    contactName: row[2],
  }));
}

/**
 * テンプレートシートにデータを埋め込む。
 */
function fillTemplateSheets(data) {
  const templateData = [
    [data.issueNumber, executionSheet.getRange(CELL_MAPPING.issueDate).getValue(),],
    [data.companyName, data.contactName],
    [data.ad1, data.ad1count],
    [data.ad2, data.ad2count],
    [data.ad3, data.ad3count],
    [data.ad4, data.ad4count],
    [data.ad5, data.ad5count],
  ];

  // 発行番号と発行日
  reportSheet1.getRange(TEMPLATE_CELLS.issueNumber).setValue(templateData[0][0]);
  reportSheet1.getRange(TEMPLATE_CELLS.issueDate).setValue(templateData[0][1]);

  // 企業名と担当者名
  reportSheet1.getRange(TEMPLATE_CELLS.companyName).setValue(templateData[1][0]);
  reportSheet1.getRange(TEMPLATE_CELLS.contactName).setValue(templateData[1][1]);

  SpreadsheetApp.flush();
}

/**
 * 新しいスプレッドシートを作成し、テンプレートシートをコピーする。
 */
function createNewSpreadsheet(data) {
  const fileName = FILE_NAME_FORMAT.replace("{issueNumber}", data.issueNumber).replace("{string}", CELL_MAPPING.string).replace("{companyName}", data.companyName);
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
