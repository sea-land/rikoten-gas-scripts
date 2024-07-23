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
        createPDF(rowData);
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
    `${CELL_MAPPING.issueNumber}${startRow}:${CELL_MAPPING.ad5count}${endRow}`
  );
  const values = range.getValues();

  return values.map((row) => ({
    issueNumber: row[0],
    companyName: row[1],
    contactName: row[2],
    ad1: row[3], ad1count: row[4],
    ad2: row[6], ad2count: row[7],
    ad3: row[9], ad3count: row[10],
    ad4: row[12], ad4count: row[13],
    ad5: row[15], ad5count: row[16],
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

  // 広告データとカウント
  for (let i = 2; i < templateData.length; i++) {
    reportSheet1.getRange(TEMPLATE_CELLS[`ad${i - 1}`]).setValue(templateData[i][0]);
    reportSheet1.getRange(TEMPLATE_CELLS[`ad${i - 1}count`]).setValue(templateData[i][1]);
  }

  SpreadsheetApp.flush();
}

/**
 * 新しいスプレッドシートを作成し、テンプレートシートをコピーしてPDFに変換する。
 */
function createPDF(data) {
  const fileName = FILE_NAME_FORMAT.replace("{issueNumber}", data.issueNumber).replace("{companyName}", data.companyName) + ".pdf";
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
  saveAsPDF(newSpreadsheet, fileName, folder);
  DriveApp.getFileById(newSpreadsheet.getId()).setTrashed(true);

  logToSheet(`${fileName} を作成しました。`);
}

/**
 * スプレッドシートをPDFに変換して指定フォルダに保存する。
 */
function saveAsPDF(sheet, fileName, folder) {
  const url = createUrlForPdf(sheet);
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  const blob = response.getBlob().setName(fileName);
  folder.createFile(blob);
}

/**
 * PDF出力用のURLを作成する.
 * @param {Spreadsheet} 出力対象のスプレッドシート.
 */
function createUrlForPdf(sheet) {
  const params = {
    'exportFormat': 'pdf',
    'format': 'pdf',
    'sheetnames': 'false',
    'printtitle': 'false',
    'pagenumbers': 'false',
    // 'gid': exportSheet.getSheetId(), // シート名を指定して出力対象シートのIDを指定
    'size': 'A4', // 用紙サイズ:A4
    'portrait': true, // 用紙向き:縦
    'fitw': true, // 幅を用紙に合わせる
    'horizontal_alignment': 'CENTER', // 水平方向:中央
    'gridlines': false, // グリッドライン:非表示
  }
  const query = Object.keys(params).map(function (key) {
    return encodeURIComponent(key) + '=' + encodeURIComponent(params[key]);
  }).join('&');
  return `https://docs.google.com/spreadsheets/d/${sheet.getId()}/export?${query}`;
}
