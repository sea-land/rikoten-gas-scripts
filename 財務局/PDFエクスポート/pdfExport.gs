/**
 * 指定フォルダ内のすべてのスプレッドシートを別のフォルダにPDF変換して保存する。
 */
function convertAllSpreadsheetsInFolderToPDF() {
  const sourceFolderUrl = executionSheet.getRange(CELL_MAPPING.folderUrl).getValue();
  const targetFolderUrl = executionSheet.getRange(CELL_MAPPING.folderUrlForPDF).getValue();
  const sourceFolder = DriveApp.getFolderById(extractFolderIdFromUrl(sourceFolderUrl));
  const targetFolder = DriveApp.getFolderById(extractFolderIdFromUrl(targetFolderUrl));

  if (!sourceFolder || !targetFolder) {
    throw new Error("フォルダのURLが無効です");
  }

  // 指定フォルダ内のすべてのファイルを取得し、スプレッドシートのみを名前順でソート
  const files = [];
  const fileIterator = sourceFolder.getFilesByType(MimeType.GOOGLE_SHEETS);

  while (fileIterator.hasNext()) {
    files.push(fileIterator.next());
  }

  files.sort((a, b) => a.getName().localeCompare(b.getName()));

  // ソートされたファイルを順番に処理
  files.forEach((file) => {
    const spreadsheet = SpreadsheetApp.openById(file.getId());
    const sheet = spreadsheet.getSheets()[0]; // 最初のシートを取得（必要に応じて変更可能）
    const fileName = `${file.getName()}.pdf`;

    try {
      logToSheet(`PDFを作成中: ${fileName}`);
      saveAsPDF(sheet, fileName, targetFolder);
    } catch (error) {
      Logger.log(`エラーが発生しました: ${fileName} - ${error.message}`);
    }
  });

  Logger.log("すべてのPDF変換が完了しました。");
}
/**
 * スプレッドシートをPDFに変換して指定フォルダに保存する。
 * @param {Sheet} sheet - PDFに変換するシート
 * @param {string} fileName - 作成するPDFファイル名
 * @param {Folder} folder - PDFを保存するフォルダ
 */
function saveAsPDF(sheet, fileName, folder) {
  const url = createUrlForPdf(sheet);
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  const blob = response.getBlob().setName(fileName);
  removeExistingFile(folder, fileName);
  folder.createFile(blob);
}

/**
 * PDF出力用のURLを作成する。
 * @param {Sheet} sheet - PDFに変換するシート
 * @return {string} PDF出力用のURL
 */
function createUrlForPdf(sheet) {
  const params = {
    'exportFormat': 'pdf',
    'format': 'pdf',
    'size': 'A4', // 用紙サイズ:A4
    'portrait': true, // 用紙向き:縦
    'fitw': true, // 幅を用紙に合わせる
    'sheetnames': false,
    'printtitle': false,
    'pagenumbers': false,
    'gridlines': false,
    'horizontal_alignment': 'CENTER'
  };

  const query = Object.keys(params).map((key) => {
    return encodeURIComponent(key) + '=' + encodeURIComponent(params[key]);
  }).join('&');

  return `https://docs.google.com/spreadsheets/d/${sheet.getParent().getId()}/export?${query}`;
}
