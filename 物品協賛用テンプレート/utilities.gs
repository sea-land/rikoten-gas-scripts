/**
 * エラーを処理する。
 */
function handleError(rowIndex, error, companyName) {
  logToSheet(
    `エラー発生: ${rowIndex} 行目 (${companyName}) - ${error.message}`
  );
}

/**
 * 処理の続きを行う。
 */
function continueProcessing(chunkEndRow, endRow, chunkSize) {
  ScriptApp.newTrigger("continueCopyTemplateSheets")
    .timeBased()
    .after(100)
    .create();
  PropertiesService.getScriptProperties().setProperty(
    "nextStartRow",
    chunkEndRow + 1
  );
  PropertiesService.getScriptProperties().setProperty("endRow", endRow);
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
 * ログをシートに出力する。
 */
function logToSheet(message) {
  const timestamp = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy/MM/dd HH:mm"
  );
  logSheet.appendRow([timestamp, message]);
  Logger.log(`${message}`);
}

/**
 * フォルダのURLからフォルダIDを抽出する。
 */
function extractFolderIdFromUrl(url) {
  const parts = url.split("/");
  return parts[parts.length - 1] || parts[parts.length - 2];
}
