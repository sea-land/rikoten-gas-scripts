/**
 * Apply general style formatting to the entire sheet
 * @param {SpreadsheetApp.Spreadsheet} spreadsheet - The active spreadsheet
 */
function styleFormatting(spreadsheet) {
  var sheet = spreadsheet.getActiveSheet();
  var range = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns());

  range
    .setBorder(false, false, false, false, false, false)
    .setFontSize(9)
    .setFontFamily("Roboto");
}

/**
 * Set column widths for columns A and B
 */
function commonAtoB(sheet) {
  // Set column width for A to 160
  sheet.setColumnWidth(1, 160);

  // Set column width for B to 300
  sheet.setColumnWidth(2, 300);
}

/**
 * Set column widths for columns C to H
 */
function commonCtoH(sheet) {
  // Set column widths for C:G to 100
  sheet.setColumnWidths(3, 5, 100);

  // Set column width for H to 200
  sheet.setColumnWidth(8, 200);
}

/**
 * Set column widths for columns I to W for goods format
 */
function goodsItoW(sheet) {
  var totalColumns = sheet.getMaxColumns();

  // Ensure the columns exist before setting widths
  if (totalColumns >= 23) {
    // Set column widths for I:N to 100
    sheet.setColumnWidths(9, 6, 100);

    // Set column widths for O:P to 150
    sheet.setColumnWidths(15, 2, 150);

    // Set column widths for Q:U to 100
    sheet.setColumnWidths(17, 5, 100);

    // Set column widths for V:W to 200
    sheet.setColumnWidths(22, 2, 200);
  }
}

/**
 * Set column widths for columns I to U for ad format
 */
function adItoU(sheet) {
  var totalColumns = sheet.getMaxColumns();

  // Ensure the columns exist before setting widths
  if (totalColumns >= 21) {
    // Set column widths for I:N to 100
    sheet.setColumnWidths(9, 6, 100);

    // Set column widths for O:P to 150
    sheet.setColumnWidths(15, 2, 150);

    // Set column widths for Q:S to 100
    sheet.setColumnWidths(17, 3, 100);

    // Set column widths for T:U to 200
    sheet.setColumnWidths(20, 2, 200);
  }
}

/**
 * Apply summary format, setting column widths from F to the last column
 */
function summaryFormatting() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();

  // Set column widths for F:LAST to 100
  sheet.setColumnWidths(6, sheet.getMaxColumns() - 6, 100);
}

/**
 * Apply formatting for company list
 */
function companyFormatting() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();

  // Set column widths for A:B to 50
  sheet.setColumnWidths(1, 2, 50);

  // Set column width for C to 160
  sheet.setColumnWidth(3, 160);

  // Set column width for D to 300
  sheet.setColumnWidth(4, 300);

  // Set column widths for E:F to 100
  sheet.setColumnWidths(5, 2, 100);

  // Set column widths for G:LAST to 100
  sheet.setColumnWidths(7, sheet.getMaxColumns() - 7, 100);

  // Apply general style formatting
  styleFormatting(spreadsheet);
}

/**
 * Apply formatting for goods sponsorship
 */
function goodsFormatting() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  commonAtoB(sheet);
  commonCtoH(sheet);
  goodsItoW(sheet);
  styleFormatting(spreadsheet);
}

/**
 * Apply formatting for ad sponsorship
 */
function adFormat() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  commonAtoB(sheet);
  commonCtoH(sheet);
  adItoU(sheet);
  styleFormatting(spreadsheet);
}

/**
 * Applies conditional formatting colors to all cells in the active sheet.
 */
function setConditionalFormatRules() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  var range = sheet.getDataRange(); // 全セルの範囲を取得

  // 条件付き書式ルールの設定
  var rules = [
    {
      formula: "=$G1=TRUE",
      backgroundColor: "#FCE8B2",
    },
    {
      formula:
        "=OR(INDIRECT(\"'概要'!E41\")=$C1,INDIRECT(\"'概要'!E42\")=$C1,INDIRECT(\"'概要'!E43\")=$C1)",
      backgroundColor: "#F4C7C3",
    },
    {
      formula:
        "=OR(INDIRECT(\"'概要'!E38\")=$C1,INDIRECT(\"'概要'!E39\")=$C1,INDIRECT(\"'概要'!E40\")=$C1)",
      backgroundColor: "#B7E1CD",
    },
  ];

  // 条件付き書式ルールをクリアする
  var conditionalFormatRules = [];

  // 各ルールを設定して追加する
  rules.forEach(function (rule) {
    var newRule = SpreadsheetApp.newConditionalFormatRule()
      .setRanges([range])
      .whenFormulaSatisfied(rule.formula)
      .setBackground(rule.backgroundColor)
      .build();
    conditionalFormatRules.push(newRule);
  });

  // シートに設定した条件付き書式ルールを適用する
  sheet.setConditionalFormatRules(conditionalFormatRules);
}
