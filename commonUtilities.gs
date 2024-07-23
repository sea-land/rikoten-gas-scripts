/**
 * Deletes all existing triggers in the script.
 */
function deleteAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
}

/**
 * Handles errors by logging them to a sheet.
 * 
 * @param {number} rowIndex The index of the row where the error occurred.
 * @param {Error} error The error object.
 * @param {string} companyName The name of the company associated with the error.
 */
function handleError(rowIndex, error, companyName) {
  logToSheet(`エラー: ${rowIndex} 行目 (${companyName}) - ${error.message}`);
}

/**
 * Sets up a time-based trigger to continue processing.
 * 
 * @param {number} chunkEndRow The end row of the current chunk.
 * @param {number} endRow The end row of the entire processing range.
 * @param {number} chunkSize The size of each chunk to be processed.
 */
function continueProcessing(chunkEndRow, endRow, chunkSize) {
  ScriptApp.newTrigger("continueCopyTemplateSheets").timeBased().after(1000).create();
  PropertiesService.getScriptProperties().setProperty("nextStartRow", chunkEndRow + 1);
  PropertiesService.getScriptProperties().setProperty("endRow", endRow);
  PropertiesService.getScriptProperties().setProperty("chunkSize", chunkSize);
}

/**
 * Retrieves properties and continues the chunk processing.
 */
function continueCopyTemplateSheets() {
  const nextStartRow = parseInt(PropertiesService.getScriptProperties().getProperty("nextStartRow"));
  const endRow = parseInt(PropertiesService.getScriptProperties().getProperty("endRow"));
  const chunkSize = parseInt(PropertiesService.getScriptProperties().getProperty("chunkSize"));
  processChunk(nextStartRow, endRow, chunkSize);
}

/**
 * Logs a message to the log sheet with a timestamp.
 * 
 * @param {string} message The message to log.
 */
function logToSheet(message) {
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm");
  logSheet.appendRow([timestamp, message]);
  Logger.log(`${message}`);
}

/**
 * Extracts the folder ID from a given URL.
 * 
 * @param {string} url The URL of the folder.
 * @return {string} The extracted folder ID.
 */
function extractFolderIdFromUrl(url) {
  const parts = url.split("/");
  return parts[parts.length - 1] || parts[parts.length - 2];
}

/**
 * Removes existing files with the same name in the specified folder.
 * 
 * @param {Folder} folder The folder to check for existing files.
 * @param {string} fileName The name of the file to remove.
 */
function removeExistingFile(folder, fileName) {
  const existingFiles = folder.getFilesByName(fileName);
  while (existingFiles.hasNext()) {
    const existingFile = existingFiles.next();
    existingFile.setTrashed(true);
  }
}
