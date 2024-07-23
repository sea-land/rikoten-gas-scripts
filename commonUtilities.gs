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
 * エラーハンドリング
 */
function handleError(rowIndex, error, companyName) {
  logToSheet(`エラー: ${rowIndex} 行目 (${companyName}) - ${error.message}`);
}

/**
 * 時間指定トリガーによる関数の実行設定
 */
function continueProcessing(chunkEndRow, endRow, chunkSize) {
  ScriptApp.newTrigger("continueCopyTemplateSheets").timeBased().after(1000).create();
  PropertiesService.getScriptProperties().setProperty("nextStartRow", chunkEndRow + 1);
  PropertiesService.getScriptProperties().setProperty("endRow", endRow);
  PropertiesService.getScriptProperties().setProperty("chunkSize", chunkSize);
}

/**
 * プロパティの取得、processChunkの実行
 */
function continueCopyTemplateSheets() {
  const nextStartRow = parseInt(PropertiesService.getScriptProperties().getProperty("nextStartRow"));
  const endRow = parseInt(PropertiesService.getScriptProperties().getProperty("endRow"));
  const chunkSize = parseInt(PropertiesService.getScriptProperties().getProperty("chunkSize"));
  processChunk(nextStartRow, endRow, chunkSize);
}

/**
 * ログをシートに出力
 */
function logToSheet(message) {
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm");
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

/**
 * 指定されたフォルダ内の同名の既存ファイルを削除
 */
function removeExistingFile(folder, fileName) {
  const existingFiles = folder.getFilesByName(fileName);
  while (existingFiles.hasNext()) {
    const existingFile = existingFiles.next();
    existingFile.setTrashed(true);
  }
}
