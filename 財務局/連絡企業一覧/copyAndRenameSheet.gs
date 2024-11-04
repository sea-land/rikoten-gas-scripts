/**
 * Main function: Copies and renames sheets, updates related sheets and summary sheet.
 */
function main() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var summarySheet = spreadsheet.getSheetByName("概要");
  var companySheet = spreadsheet.getSheetByName("企業一覧");

  // Retrieve necessary values from the summary sheet
  var eventName = summarySheet.getRange("C3").getValue();
  var generation = summarySheet.getRange("C4").getValue();
  var year = summarySheet.getRange("C5").getValue();

  // Validate input values
  if (!isNumeric(generation) || !isNumeric(year)) {
    SpreadsheetApp.getUi().alert("C4とC5には半角の数値を入力してください。");
    return;
  }

  // Retrieve information for the new sheet
  var sheetInfo = getSheetInfo(eventName, generation, year);
  if (!sheetInfo) {
    SpreadsheetApp.getUi().alert("C3の値が不正です。");
    return;
  }

  // Determine new sheet names
  var newSheetName1 = sheetInfo.newSheetName1;
  var newSheetName2 = sheetInfo.createSecondSheet ? sheetInfo.newSheetName2 : null;

  // Check for the existence of sheets with the same name
  var existingSheet1 = spreadsheet.getSheetByName(newSheetName1);
  var existingSheet2 = newSheetName2 ? spreadsheet.getSheetByName(newSheetName2) : null;

  if (existingSheet1 || existingSheet2) {
    SpreadsheetApp.getUi().alert(
      "同じ名前のシートがすでに存在します。シートの作成をスキップします。"
    );
    return;
  }

  // Get the source sheets to copy
  var sheetToCopy1 = spreadsheet.getSheetByName(sheetInfo.sheetToCopyName1);
  var sheetToCopy2 = sheetInfo.createSecondSheet ? spreadsheet.getSheetByName(sheetInfo.sheetToCopyName2) : null;

  // Check if the source sheets exist
  if (!sheetToCopy1 || (sheetInfo.createSecondSheet && !sheetToCopy2)) {
    SpreadsheetApp.getUi().alert("コピーするシートが見つかりません。");
    return;
  }

  // Ask the user to confirm the creation of the sheets
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    "シート作成確認",
    eventName + "シートを作成しますか？",
    ui.ButtonSet.YES_NO
  );
  if (response != ui.Button.YES) {
    ui.alert("シートの作成がキャンセルされました。");
    return;
  }

  // Copy and rename the sheets
  createAndRenameSheet(spreadsheet, sheetToCopy1, newSheetName1);
  if (sheetInfo.createSecondSheet) {
    createAndRenameSheet(spreadsheet, sheetToCopy2, newSheetName2);
  }

  // Update the company list sheet
  updateCompanySheet(companySheet, eventName, generation, year, sheetInfo);

  // Update the summary sheet
  updateSummarySheet(summarySheet, eventName, generation, year, sheetInfo);
}

/**
 * Checks if a value is a numeric string.
 * @param {string} value - The value to check.
 * @return {boolean} True if the value is numeric, false otherwise.
 */
function isNumeric(value) {
  return /^\d+$/.test(value);
}

/**
 * Retrieves information for the new sheet based on the event name, generation, and year.
 * @param {string} eventName - The name of the event.
 * @param {string} generation - The generation of the event.
 * @param {string} year - The year of the event.
 * @return {Object|null} An object containing sheet information, or null if the event name is invalid.
 */
function getSheetInfo(eventName, generation, year) {
  if (eventName === "理工展") {
    return {
      sheetToCopyName1: "〇期(20〇)広告",
      newSheetName1: generation + "期(" + year + ")広告",
      sheetToCopyName2: "〇期(20〇)物品",
      newSheetName2: generation + "期(" + year + ")物品",
      createSecondSheet: true,
    };
  } else if (eventName === "Welcome") {
    return {
      sheetToCopyName1: "〇期(20〇)Welcome広告",
      newSheetName1: generation + "期(" + year + ")Welcome広告",
      createSecondSheet: false,
    };
  }
  return null;
}

/**
 * Copies and renames a sheet.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The active spreadsheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheetToCopy - The sheet to copy.
 * @param {string} newSheetName - The name of the new sheet.
 */
function createAndRenameSheet(spreadsheet, sheetToCopy, newSheetName) {
  var newSheet = sheetToCopy.copyTo(spreadsheet);
  newSheet.setName(newSheetName);
  spreadsheet.setActiveSheet(newSheet);
  spreadsheet.moveActiveSheet(4);
}

/**
 * Updates the company list sheet with new information.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} companySheet - The company list sheet.
 * @param {string} eventName - The name of the event.
 * @param {string} generation - The generation of the event.
 * @param {string} year - The year of the event.
 * @param {Object} sheetInfo - Information about the sheets.
 */
function updateCompanySheet(companySheet, eventName, generation, year, sheetInfo) {
  var newSheetName = eventName === "理工展" ? generation + "期(" + year + ")" : "Welcome" + year;
  var existingValue = companySheet.getRange("G1").getValue();

  if (existingValue !== newSheetName) {
    companySheet.insertColumns(7);
    companySheet.getRange("H1:H2").copyTo(companySheet.getRange("G1:G2"));

    var formula = eventName === "理工展"
      ? "=if(ISBLANK($D2),,IFERROR(VLOOKUP($D2,'" + sheetInfo.newSheetName1 + "'!$B:$C,2,FALSE),IFERROR(VLOOKUP($D2,'" + sheetInfo.newSheetName2 + "'!$B:$C,2,FALSE),)))"
      : "=if(ISBLANK($D2),,IFERROR(VLOOKUP($D2,'" + sheetInfo.newSheetName1 + "'!$B:$C,2,FALSE),))";

    companySheet.getRange(2, 7, companySheet.getMaxRows() - 1).setValue(formula);
    companySheet.getRange("G1").setValue(newSheetName);
  } else {
    SpreadsheetApp.getUi().alert(
      "企業一覧シート " + newSheetName + " はすでに設定されているのでスキップします。"
    );
  }
}

/**
 * Updates the summary sheet with new information.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} summarySheet - The summary sheet.
 * @param {string} eventName - The name of the event.
 * @param {string} generation - The generation of the event.
 * @param {string} year - The year of the event.
 * @param {Object} sheetInfo - Information about the sheets.
 */
function updateSummarySheet(summarySheet, eventName, generation, year, sheetInfo) {
  var newSheetName = eventName === "理工展" ? generation + "期(" + year + ")" : "Welcome" + year;
  var existingValue = summarySheet.getRange("F3").getValue();

  if (existingValue !== newSheetName) {
    summarySheet.insertColumns(6);
    summarySheet.getRange("F3").setValue(newSheetName);
    summarySheet.getRange("F11").setValue(newSheetName);
    summarySheet.getRange("G4:G6").copyTo(summarySheet.getRange("F4:F6"));
    summarySheet.getRange("G12:G46").copyTo(summarySheet.getRange("F12:F46"));
  } else {
    SpreadsheetApp.getUi().alert(
      "概要シート " + newSheetName + " はすでに設定されているのでスキップします。"
    );
  }
}
