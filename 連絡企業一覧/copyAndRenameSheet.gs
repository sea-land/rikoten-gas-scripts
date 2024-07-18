/**
 * メイン関数: シートをコピーしてリネームし、関連するシートと要約シートを更新する。
 */
function main() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var summarySheet = spreadsheet.getSheetByName("概要");
  var companySheet = spreadsheet.getSheetByName("企業一覧");

  // 概要シートから必要な値を取得
  var eventName = summarySheet.getRange("C4").getValue();
  var generation = summarySheet.getRange("C5").getValue();
  var year = summarySheet.getRange("C6").getValue();

  // 入力値の検証
  if (!isNumeric(generation) || !isNumeric(year)) {
    SpreadsheetApp.getUi().alert("C5とC6には半角の数値を入力してください。");
    return;
  }

  // 新しいシートの情報を取得
  var sheetInfo = getSheetInfo(eventName, generation, year);
  if (!sheetInfo) {
    SpreadsheetApp.getUi().alert("C4の値が不正です。");
    return;
  }

  // 新しいシートの名前を決定
  var newSheetName1 = sheetInfo.newSheetName1;
  var newSheetName2 = sheetInfo.createSecondSheet
    ? sheetInfo.newSheetName2
    : null;

  // 既存のシートの存在を確認
  var existingSheet1 = spreadsheet.getSheetByName(newSheetName1);
  var existingSheet2 = newSheetName2
    ? spreadsheet.getSheetByName(newSheetName2)
    : null;

  if (existingSheet1 || existingSheet2) {
    SpreadsheetApp.getUi().alert(
      "同じ名前のシートがすでに存在します。シートの作成をスキップします。"
    );
    return;
  }

  // コピー元のシートを取得
  var sheetToCopy1 = spreadsheet.getSheetByName(sheetInfo.sheetToCopyName1);
  var sheetToCopy2 = sheetInfo.createSecondSheet
    ? spreadsheet.getSheetByName(sheetInfo.sheetToCopyName2)
    : null;

  // コピー元のシートが存在するか確認
  if (!sheetToCopy1 || (sheetInfo.createSecondSheet && !sheetToCopy2)) {
    SpreadsheetApp.getUi().alert("コピーするシートが見つかりません。");
    return;
  }

  // ユーザーにシート作成の確認を求める
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

  // シートをコピーしてリネーム
  createAndRenameSheet(spreadsheet, sheetToCopy1, newSheetName1);
  if (sheetInfo.createSecondSheet) {
    createAndRenameSheet(spreadsheet, sheetToCopy2, newSheetName2);
  }

  // 企業一覧シートを更新
  updateCompanySheet(companySheet, eventName, generation, year, sheetInfo);

  // 概要シートを更新
  updateSummarySheet(summarySheet, eventName, generation, year, sheetInfo);
}

/**
 * 半角数字かどうかを判定する関数
 */
function isNumeric(value) {
  return /^\d+$/.test(value);
}

/**
 * 新しいシートの情報を取得する関数
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
 * シートをコピーしてリネームする関数
 */
function createAndRenameSheet(spreadsheet, sheetToCopy, newSheetName) {
  var newSheet = sheetToCopy.copyTo(spreadsheet);
  newSheet.setName(newSheetName);
  spreadsheet.setActiveSheet(newSheet);
  spreadsheet.moveActiveSheet(4);
}

/**
 * 企業一覧シートを更新する関数
 */
function updateCompanySheet(
  companySheet,
  eventName,
  generation,
  year,
  sheetInfo
) {
  var newSheetName =
    eventName === "理工展" ? generation + "期(" + year + ")" : "Welcome" + year;
  var existingValue = companySheet.getRange("G1").getValue();

  if (existingValue !== newSheetName) {
    companySheet.insertColumns(7);
    companySheet.getRange("H1:H2").copyTo(companySheet.getRange("G1:G2"));

    var formula =
      eventName === "理工展"
        ? "=if(ISBLANK($D2),,IFERROR(VLOOKUP($D2,'" +
          sheetInfo.newSheetName1 +
          "'!$B:$C,2,FALSE),IFERROR(VLOOKUP($D2,'" +
          sheetInfo.newSheetName2 +
          "'!$B:$C,2,FALSE),)))"
        : "=if(ISBLANK($D2),,IFERROR(VLOOKUP($D2,'" +
          sheetInfo.newSheetName1 +
          "'!$B:$C,2,FALSE),))";

    companySheet
      .getRange(2, 7, companySheet.getMaxRows() - 1)
      .setValue(formula);
    companySheet.getRange("G1").setValue(newSheetName);
  } else {
    SpreadsheetApp.getUi().alert(
      "企業一覧シート " +
        newSheetName +
        " はすでに設定されているのでスキップします。"
    );
  }
}

/**
 * 概要シートを更新する関数
 */
function updateSummarySheet(
  summarySheet,
  eventName,
  generation,
  year,
  sheetInfo
) {
  var newSheetName =
    eventName === "理工展" ? generation + "期(" + year + ")" : "Welcome" + year;
  var existingValue = summarySheet.getRange("F3").getValue();

  if (existingValue !== newSheetName) {
    summarySheet.insertColumns(6);
    summarySheet.getRange("F3").setValue(newSheetName);
    summarySheet.getRange("F11").setValue(newSheetName);
    summarySheet.getRange("G4:G6").copyTo(summarySheet.getRange("F4:F6"));
    summarySheet.getRange("G12:G46").copyTo(summarySheet.getRange("F12:F46"));
  } else {
    SpreadsheetApp.getUi().alert(
      "概要シート " +
        newSheetName +
        " はすでに設定されているのでスキップします。"
    );
  }
}
