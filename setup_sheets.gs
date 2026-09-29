/**
 * Risk HKR — one-time sheet setup + cosmetic styling.
 *
 * Creates the `calls` and `config` tabs with the right headers, then applies
 * tab colours, header shading, row banding, number formats, column widths,
 * dropdowns and conditional highlighting.
 *
 * Reads FIELDS from Code.gs, so the header row can never drift out of sync
 * with what the app writes. Paste this alongside Code.gs in the same project.
 *
 * Safe to re-run: it never overwrites a data cell, only structure and format.
 *
 *   Run:  setupSheets()
 *   Undo: resetSheetStyles()      (formatting only — your rows stay)
 */

var HKR_STYLE = {
  calls:  { tab: '#B4443C', header: '#F3DDDA', band: '#FBF3F2' },
  config: { tab: '#7C8597', header: '#E6E9EE', band: '#F6F7F9' },
};

// Columns that hold rupee amounts / counts, addressed by FIELDS name so a
// reordered FIELDS array cannot mis-format a column.
var HKR_MONEY = ['hisPrice', 'myPrice', 'myValue'];
var HKR_COUNT = ['hisQty', 'myShares'];
var HKR_WIDE  = { hisThesis: 380, myView: 380, skipReason: 190, id: 240, createdAt: 160 };

function colOf_(field) {
  return FIELDS.indexOf(field) + 1;   // 1-based, 0 when absent
}

function setupSheets() {
  var ss = SpreadsheetApp.openById(SS_ID);

  var calls  = ensureSheet_(ss, 'calls',  FIELDS);
  var config = ensureSheet_(ss, 'config', ['key', 'value']);

  styleCalls_(calls);
  styleConfig_(config);

  // A brand-new spreadsheet ships with an empty "Sheet1" — drop it if untouched.
  var extra = ss.getSheetByName('Sheet1');
  if (extra && ss.getSheets().length > 1 && extra.getLastRow() === 0 && extra.getLastColumn() === 0) {
    ss.deleteSheet(extra);
  }

  SpreadsheetApp.flush();
  Logger.log('calls: %s cols · config ready · styling applied', FIELDS.length);
  return { calls: FIELDS.length, ok: true };
}

/** Create the tab if missing; always (re)write the header row text. */
function ensureSheet_(ss, name, header) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.getRange(1, 1, 1, header.length).setValues([header]);
  return sh;
}

function styleCalls_(sh) {
  var s = HKR_STYLE.calls;
  var nCols = FIELDS.length;

  sh.setTabColor(s.tab);
  sh.setFrozenRows(1);
  sh.setFrozenColumns(Math.min(4, nCols));       // id … ticker stay visible while scrolling

  // Header
  sh.getRange(1, 1, 1, nCols)
    .setBackground(s.header)
    .setFontWeight('bold')
    .setFontSize(10)
    .setVerticalAlignment('middle')
    .setWrap(false)
    .setBorder(null, null, true, null, null, null, '#9A6660', SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(1, 30);

  // Banding — remove any existing first, applyRowBanding throws on overlap
  sh.getBandings().forEach(function (b) { b.remove(); });
  var maxRows = Math.max(sh.getMaxRows() - 1, 1);
  sh.getRange(2, 1, maxRows, nCols)
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false)
    .setFirstRowColor(s.band)
    .setSecondRowColor('#FFFFFF')
    .setHeaderRowColor(null);

  // Number formats
  HKR_MONEY.forEach(function (f) {
    var c = colOf_(f);
    if (c) sh.getRange(2, c, maxRows, 1).setNumberFormat('₹#,##0.00');
  });
  HKR_COUNT.forEach(function (f) {
    var c = colOf_(f);
    if (c) sh.getRange(2, c, maxRows, 1).setNumberFormat('#,##0');
  });
  if (colOf_('date'))      sh.getRange(2, colOf_('date'), maxRows, 1).setNumberFormat('yyyy-mm-dd');
  if (colOf_('createdAt')) sh.getRange(2, colOf_('createdAt'), maxRows, 1).setNumberFormat('@');  // ISO text, not a date
  if (colOf_('id'))        sh.getRange(2, colOf_('id'), maxRows, 1).setNumberFormat('@');

  // Column widths
  for (var i = 1; i <= nCols; i++) sh.setColumnWidth(i, 110);
  Object.keys(HKR_WIDE).forEach(function (f) {
    var c = colOf_(f);
    if (c) sh.setColumnWidth(c, HKR_WIDE[f]);
  });

  // Long prose wraps; everything else clips so rows stay one line tall
  ['hisThesis', 'myView', 'skipReason'].forEach(function (f) {
    var c = colOf_(f);
    if (c) sh.getRange(2, c, maxRows, 1).setWrap(true).setVerticalAlignment('top');
  });

  // Dropdowns — allowInvalid so an API write can never be rejected by validation
  dropdown_(sh, 'side',     ['buy', 'sell'],      maxRows);
  dropdown_(sh, 'kind',     ['stock', 'etf'],     maxRows);
  dropdown_(sh, 'decision', ['copied', 'skipped'], maxRows);

  applyCallsRules_(sh, nCols, maxRows);
}

function dropdown_(sh, field, values, maxRows) {
  var c = colOf_(field);
  if (!c) return;
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(values, true)
    .setAllowInvalid(true)          // warn, never reject — the app is the writer
    .build();
  sh.getRange(2, c, maxRows, 1).setDataValidation(rule);
}

/** Row-level highlighting so buys, sells and skips are readable at a glance. */
function applyCallsRules_(sh, nCols, maxRows) {
  var body = sh.getRange(2, 1, maxRows, nCols);
  var sideCol     = columnLetter_(colOf_('side'));
  var decisionCol = columnLetter_(colOf_('decision'));
  var kindCol     = columnLetter_(colOf_('kind'));

  var rules = [];

  if (decisionCol) {
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$' + decisionCol + '2="skipped"')
      .setBackground('#F1F3F4').setFontColor('#80868B')
      .setRanges([body]).build());
  }
  if (sideCol) {
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$' + sideCol + '2="sell"')
      .setBackground('#FCE8E6')
      .setRanges([body]).build());
  }
  if (kindCol) {
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$' + kindCol + '2="etf"')
      .setFontColor('#5F6368').setItalic(true)
      .setRanges([sh.getRange(2, colOf_('ticker') || 1, maxRows, 1)]).build());
  }

  sh.setConditionalFormatRules(rules);
}

function styleConfig_(sh) {
  var s = HKR_STYLE.config;
  sh.setTabColor(s.tab);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, 2)
    .setBackground(s.header).setFontWeight('bold').setFontSize(10)
    .setBorder(null, null, true, null, null, null, '#6B7280', SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(1, 28);
  sh.setColumnWidth(1, 220);
  sh.setColumnWidth(2, 150);

  sh.getBandings().forEach(function (b) { b.remove(); });
  var maxRows = Math.max(sh.getMaxRows() - 1, 1);
  sh.getRange(2, 1, maxRows, 2)
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false)
    .setFirstRowColor(s.band).setSecondRowColor('#FFFFFF').setHeaderRowColor(null);

  sh.getRange(2, 1, maxRows, 1).setNumberFormat('@');
  // tank/mine/his are plain rupees; m:TICKER marks are prices — one format suits both
  sh.getRange(2, 2, maxRows, 1).setNumberFormat('#,##0.00');

  // Label the m: convention for anyone reading the sheet directly
  sh.getRange(1, 4).setValue('tank / mine / his are rupees. m:TICKER rows are marked prices.')
    .setFontColor('#80868B').setFontSize(9).setFontStyle('italic');
}

function columnLetter_(col) {
  if (!col) return '';
  var s = '';
  while (col > 0) {
    var m = (col - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    col = (col - m - 1) / 26;
  }
  return s;
}

/** Strip formatting only. Headers and data are left alone. */
function resetSheetStyles() {
  var ss = SpreadsheetApp.openById(SS_ID);
  ['calls', 'config'].forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh) return;
    sh.setTabColor(null);
    sh.getBandings().forEach(function (b) { b.remove(); });
    sh.setConditionalFormatRules([]);
    var rows = Math.max(sh.getMaxRows(), 1), cols = Math.max(sh.getMaxColumns(), 1);
    sh.getRange(1, 1, rows, cols)
      .setBackground(null).setFontColor(null).setFontWeight('normal')
      .setFontStyle('normal').setNumberFormat('General').setWrap(false)
      .setDataValidation(null);
  });
  Logger.log('styles reset');
}
