const A = 1;
const B = 2;
const C = 3;
const D = 4;
const E = 5;
const F = 6;
const G = 7;
const H = 8;
const I = 9;
const J = 10;
const K = 11;
const L = 12;
const M = 13;
const N = 14;
const O = 15;
const P = 16;
const Q = 17;
const R = 18;
const S = 19;
const T = 20;
const U = 21;
const V = 22;
const W = 23;
const X = 24;
const Y = 25;
const Z = 26;

/**
 * Applies general style formatting to the entire sheet.
 * @param {SpreadsheetApp.Spreadsheet} spreadsheet - The active spreadsheet
 */
function applyGeneralStyle(spreadsheet) {
  var sheet = spreadsheet.getActiveSheet();
  var range = sheet.getDataRange();

  range
    .setBorder(false, false, false, false, false, false)
    .setFontSize(9)
    .setFontFamily("Roboto");
}

/**
 * Sets column widths for columns A and B.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The active sheet
 */
function setColumnWidthsAB(sheet) {
  sheet.setColumnWidth(A, 160);
  sheet.setColumnWidth(B, 300);
}

/**
 * Sets column widths for columns C to H.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The active sheet
 */
function setColumnWidthsCH(sheet) {
  sheet.setColumnWidth(C, 100);
  sheet.setColumnWidth(D, 100);
  sheet.setColumnWidth(E, 100);
  sheet.setColumnWidth(F, 130);
  sheet.setColumnWidth(G, 100);
  sheet.setColumnWidth(H, 200);
}

/**
 * Sets column widths for columns I to W for goods format.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The active sheet
 */
function setColumnWidthsGoodsIW(sheet) {
  var totalColumns = sheet.getMaxColumns();

  if (totalColumns >= W) {
    sheet.setColumnWidth(I, 100);
    sheet.setColumnWidth(J, 100);
    sheet.setColumnWidth(K, 100);
    sheet.setColumnWidth(L, 100);
    sheet.setColumnWidth(M, 100);
    sheet.setColumnWidth(N, 100);
    sheet.setColumnWidth(O, 200);
    sheet.setColumnWidth(P, 350);
    sheet.setColumnWidth(Q, 100);
    sheet.setColumnWidth(R, 100);
    sheet.setColumnWidth(S, 100);
    sheet.setColumnWidth(T, 130);
    sheet.setColumnWidth(U, 200);
    sheet.setColumnWidth(V, 200);
  }
}

/**
 * Sets column widths for columns I to U for ad format.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The active sheet
 */
function setColumnWidthsAdIU(sheet) {
  var totalColumns = sheet.getMaxColumns();

  if (totalColumns >= U) {
    sheet.setColumnWidth(I, 100);
    sheet.setColumnWidth(J, 100);
    sheet.setColumnWidth(K, 100);
    sheet.setColumnWidth(L, 100);
    sheet.setColumnWidth(M, 100);
    sheet.setColumnWidth(N, 100);
    sheet.setColumnWidth(O, 150);
    sheet.setColumnWidth(P, 150);
    sheet.setColumnWidth(Q, 150);
    sheet.setColumnWidth(R, 150);
    sheet.setColumnWidth(S, 100);
    sheet.setColumnWidth(T, 100);
    sheet.setColumnWidth(U, 130);
    sheet.setColumnWidth(V, 200);
    sheet.setColumnWidth(W, 200);
  }
}

/**
 * Sets column widths for columns I to U for welcome format.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The active sheet
 */
function setColumnWidthsWelcomeIU(sheet) {
  sheet.setColumnWidth(I, 100);
  sheet.setColumnWidth(J, 100);
  sheet.setColumnWidth(K, 100);
  sheet.setColumnWidth(L, 100);
  sheet.setColumnWidth(M, 100);
  sheet.setColumnWidth(N, 100);
  sheet.setColumnWidth(O, 150);
  sheet.setColumnWidth(P, 100);
  sheet.setColumnWidth(Q, 100);
  sheet.setColumnWidth(R, 100);
  sheet.setColumnWidth(S, 130);
  sheet.setColumnWidth(T, 200);
  sheet.setColumnWidth(U, 200);
}

/**
 * Applies summary format, setting column widths from F to the last column.
 */
function applySummaryFormat() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  sheet.setColumnWidths(F, sheet.getMaxColumns() - F + 1, 100);
}

/**
 * Applies formatting for company list.
 */
function applyCompanyFormat() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();

  sheet.setColumnWidth(A, 50);
  sheet.setColumnWidth(B, 160);
  sheet.setColumnWidth(C, 300);
  sheet.setColumnWidths(D, sheet.getMaxColumns() - D + 1, 100);

  applyGeneralStyle(spreadsheet);
}

/**
 * Applies formatting for goods sponsorship.
 */
function applyGoodsFormat() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  setColumnWidthsAB(sheet);
  setColumnWidthsCH(sheet);
  setColumnWidthsGoodsIW(sheet);
  applyGeneralStyle(spreadsheet);
}

/**
 * Applies formatting for ad sponsorship.
 */
function applyAdFormat() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  setColumnWidthsAB(sheet);
  setColumnWidthsCH(sheet);
  setColumnWidthsAdIU(sheet);
  applyGeneralStyle(spreadsheet);
}

/**
 * Applies formatting for welcome format.
 */
function applyWelcomeFormat() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  setColumnWidthsAB(sheet);
  setColumnWidthsCH(sheet);
  setColumnWidthsWelcomeIU(sheet);
  applyGeneralStyle(spreadsheet);
}

/**
 * Applies conditional formatting colors to all cells in the active sheet.
 */
function applyConditionalFormatting() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  var range = sheet.getDataRange();

  var rules = [
    {
      formula: "=$G1=TRUE",
      backgroundColor: "#FCE8B2",
    },
    {
      formula: "=OR(INDIRECT(\"'概要'!E41\")=$C1)",
      backgroundColor: "#F4C7C3",
    },
    {
      formula: "=OR(INDIRECT(\"'概要'!E38\")=$C1,INDIRECT(\"'概要'!E39\")=$C1,INDIRECT(\"'概要'!E40\")=$C1)",
      backgroundColor: "#B7E1CD",
    },
  ];

  var conditionalFormatRules = [];

  rules.forEach(function (rule) {
    var newRule = SpreadsheetApp.newConditionalFormatRule()
      .setRanges([range])
      .whenFormulaSatisfied(rule.formula)
      .setBackground(rule.backgroundColor)
      .build();
    conditionalFormatRules.push(newRule);
  });

  sheet.setConditionalFormatRules(conditionalFormatRules);
}
