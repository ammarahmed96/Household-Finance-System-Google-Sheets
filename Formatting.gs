// ============================================================
// HOUSEHOLD FINANCE TRACKER — Formatting.gs
// All conditional formatting, UI utilities, and professional
// table styling (filters, named ranges, frozen panes, summary
// blocks, section dividers).
//
// Public entry points:
//   applyFormatting(ss)             — basic CF for all data sheets
//   applyProfessionalFormatting(ss) — full table styling layer
//   styleHeaderRow(range, bgColor)  — utility used by sheet builders
//   setDropdown(sheet, ...)         — utility used by sheet builders
// ============================================================

// ── BASIC CONDITIONAL FORMATTING ─────────────────────────────

function applyFormatting(ss) {
  formatTransactions(ss);
  formatBudget(ss);
  formatIncome(ss);
  formatSavings(ss);
  formatSubscriptions(ss);
  formatGroupSplits(ss);
  formatDashboard(ss);
}

function formatTransactions(ss) {
  var sh   = ss.getSheetByName(SHEETS.TRANSACTIONS);
  var data = sh.getRange(2, 1, 998, 16);  // A–P (O=Group Split?, P=# People)

  var altRow = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=ISEVEN(ROW())')
    .setBackground(COLORS.ALT_ROW).setRanges([data]).build();

  var income = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$C2="Income"')
    .setBackground(COLORS.LIGHT_GREEN).setRanges([data]).build();

  var shared = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($C2="Expense",$L2="Yes")')
    .setBackground('#e3f2fd').setRanges([data]).build();

  // Highlight international transactions (currency ≠ PKR)
  var intl = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($G2<>"",$G2<>"PKR")')
    .setFontColor(COLORS.PRIMARY_DARK).setRanges([sh.getRange(2, 7, 998, 3)]).build();

  sh.setConditionalFormatRules([altRow, income, shared, intl]);

  sh.getRange(1, 1, 1, 14)
    .setBorder(null, null, true, null, null, null, COLORS.PRIMARY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

function formatBudget(ss) {
  var sh   = ss.getSheetByName(SHEETS.BUDGET);
  var data = sh.getRange(2, 1, 198, 8);
  var pct  = sh.getRange(2, 7, 198, 1);

  var altRow = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=ISEVEN(ROW())')
    .setBackground(COLORS.ALT_ROW).setRanges([data]).build();

  var over = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(1)
    .setBackground(COLORS.LIGHT_RED).setFontColor(COLORS.DANGER)
    .setRanges([pct]).build();

  var near = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(0.8, 1)
    .setBackground(COLORS.LIGHT_YELLOW).setFontColor('#b45309')
    .setRanges([pct]).build();

  var under = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0.8)
    .setBackground(COLORS.LIGHT_GREEN).setFontColor('#166534')
    .setRanges([pct]).build();

  sh.setConditionalFormatRules([altRow, over, near, under]);
}

function formatIncome(ss) {
  var sh   = ss.getSheetByName(SHEETS.INCOME);
  var data = sh.getRange(2, 1, 498, 7);
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=ISEVEN(ROW())')
      .setBackground(COLORS.ALT_ROW).setRanges([data]).build()
  ]);
}

function formatSavings(ss) {
  var sh   = ss.getSheetByName(SHEETS.SAVINGS);
  var prog = sh.getRange(2, 5, 20, 1);
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .setGradientMinpoint('#ea4335')
      .setGradientMidpointWithValue('#fbbc04', SpreadsheetApp.InterpolationType.PERCENT, '50')
      .setGradientMaxpoint('#34a853')
      .setRanges([prog]).build()
  ]);
}

function formatSubscriptions(ss) {
  var sh   = ss.getSheetByName(SHEETS.SUBSCRIPTIONS);
  var data = sh.getRange(2, 1, 98, 13);

  var altRow = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=ISEVEN(ROW())')
    .setBackground(COLORS.ALT_ROW).setRanges([data]).build();

  var cancelled = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$K2="Cancelled"')
    .setFontColor('#9aa0a6').setRanges([data]).build();

  var renewal = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($K2="Active",$J2<=TODAY()+7)')
    .setBackground(COLORS.LIGHT_YELLOW).setRanges([data]).build();

  sh.setConditionalFormatRules([altRow, cancelled, renewal]);
}

function formatGroupSplits(ss) {
  var sh   = ss.getSheetByName(SHEETS.GROUP_SPLITS);
  var data = sh.getRange(4, 1, 496, 11);

  var altRow = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=ISEVEN(ROW())')
    .setBackground(COLORS.ALT_ROW).setRanges([data]).build();

  // Fully reimbursed → green
  var paid = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($F4<>"",$I4=0)')
    .setBackground(COLORS.LIGHT_GREEN).setRanges([data]).build();

  // Partially reimbursed → yellow
  var partial = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($H4>0,$I4>0)')
    .setBackground(COLORS.LIGHT_YELLOW).setRanges([data]).build();

  // Nothing received yet → light red on Outstanding cell only
  var none = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($F4<>"",$H4=0,$I4>0)')
    .setBackground(COLORS.LIGHT_RED).setFontColor(COLORS.DANGER)
    .setRanges([sh.getRange(4, 9, 496, 1)]).build();

  sh.setConditionalFormatRules([altRow, paid, partial, none]);
}

// ── UI UTILITIES ──────────────────────────────────────────────

function styleHeaderRow(range, bgColor) {
  range
    .setBackground(bgColor || COLORS.PRIMARY)
    .setFontColor(COLORS.HEADER_FG)
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);
  range.getSheet().setRowHeight(range.getRow(), 36);
}

function setDropdown(sheet, startRow, col, numRows, rangeRef) {
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(rangeRef, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(startRow, col, numRows, 1).setDataValidation(rule);
}

// ── PROFESSIONAL TABLE FORMATTING ─────────────────────────────
// Adds filters, warning-only header protection, HF_* named ranges,
// enhanced frozen panes, per-sheet summary blocks, and section dividers.
// Called from setup(). Safe to run standalone to refresh styling.

function applyProfessionalFormatting(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  addFilters(ss);
  protectHeaders(ss);
  addNamedRanges(ss);
  enhanceFrozenPanes(ss);
  addSummaryBlocks(ss);
  addSectionDividers(ss);
  ss.toast('Professional formatting applied!', 'Done', 5);
}

// ── FILTERS ───────────────────────────────────────────────────

function addFilters(ss) {
  var configs = [
    { name: SHEETS.TRANSACTIONS,  hdrRow: 1, cols: 16 },
    { name: SHEETS.BUDGET,        hdrRow: 1, cols: 8  },
    { name: SHEETS.INCOME,        hdrRow: 1, cols: 7  },
    { name: SHEETS.SAVINGS,       hdrRow: 1, cols: 8  },
    { name: SHEETS.SUBSCRIPTIONS, hdrRow: 1, cols: 13 },
    { name: SHEETS.GROUP_SPLITS,  hdrRow: 3, cols: 11 },
  ];
  configs.forEach(function(cfg) {
    var sh = ss.getSheetByName(cfg.name);
    if (!sh) return;
    var existing = sh.getFilter();
    if (existing) existing.remove();
    var last = Math.max(sh.getLastRow(), cfg.hdrRow + 1);
    sh.getRange(cfg.hdrRow, 1, last - cfg.hdrRow + 1, cfg.cols).createFilter();
  });
}

// ── HEADER PROTECTION ─────────────────────────────────────────

function protectHeaders(ss) {
  var targets = [
    { sheet: SHEETS.TRANSACTIONS,  row: 1, col: 1, rows: 1, cols: 16, desc: 'Transactions header' },
    { sheet: SHEETS.BUDGET,        row: 1, col: 1, rows: 1, cols: 8,  desc: 'Budget header' },
    { sheet: SHEETS.INCOME,        row: 1, col: 1, rows: 1, cols: 7,  desc: 'Income header' },
    { sheet: SHEETS.SAVINGS,       row: 1, col: 1, rows: 1, cols: 8,  desc: 'Savings Goals header' },
    { sheet: SHEETS.SUBSCRIPTIONS, row: 1, col: 1, rows: 1, cols: 13, desc: 'Subscriptions header' },
    { sheet: SHEETS.GROUP_SPLITS,  row: 1, col: 1, rows: 3, cols: 11, desc: 'Group Splits banner & header' },
    { sheet: SHEETS.DASHBOARD,     row: 1, col: 2, rows: 3, cols: 13, desc: 'Dashboard banner' },
  ];
  targets.forEach(function(t) {
    var sh = ss.getSheetByName(t.sheet);
    if (!sh) return;
    sh.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function(p) { p.remove(); });
    var prot = sh.getRange(t.row, t.col, t.rows, t.cols).protect();
    prot.setDescription(t.desc);
    prot.setWarningOnly(true);
  });
}

// ── NAMED RANGES ──────────────────────────────────────────────

function addNamedRanges(ss) {
  ss.getNamedRanges().forEach(function(nr) {
    if (nr.getName().indexOf('HF_') === 0) nr.remove();
  });
  var tx  = ss.getSheetByName(SHEETS.TRANSACTIONS);
  var bud = ss.getSheetByName(SHEETS.BUDGET);
  var inc = ss.getSheetByName(SHEETS.INCOME);
  var sav = ss.getSheetByName(SHEETS.SAVINGS);
  var sub = ss.getSheetByName(SHEETS.SUBSCRIPTIONS);
  var gs  = ss.getSheetByName(SHEETS.GROUP_SPLITS);
  var defs = [
    ['HF_Transactions',      tx,  'A1:P1000'],
    ['HF_Transactions_Data', tx,  'A2:P1000'],
    ['HF_TX_Amount_PKR',     tx,  'J2:J1000'],
    ['HF_TX_Month',          tx,  'M2:M1000'],
    ['HF_Budget',            bud, 'A1:H200'],
    ['HF_Budget_Data',       bud, 'A2:H200'],
    ['HF_Income',            inc, 'A1:G500'],
    ['HF_Income_Data',       inc, 'A2:G500'],
    ['HF_SavingsGoals',      sav, 'A1:H30'],
    ['HF_Subscriptions',     sub, 'A1:M100'],
    ['HF_GroupSplits',       gs,  'A3:L500'],
  ];
  defs.forEach(function(d) {
    if (d[1]) ss.setNamedRange(d[0], d[1].getRange(d[2]));
  });
}

// ── FROZEN PANES ENHANCEMENTS ─────────────────────────────────

function enhanceFrozenPanes(ss) {
  var freezes = [
    { sheet: SHEETS.BUDGET,        cols: 2 },
    { sheet: SHEETS.INCOME,        cols: 1 },
    { sheet: SHEETS.SAVINGS,       cols: 1 },
    { sheet: SHEETS.SUBSCRIPTIONS, cols: 1 },
    // Group Splits excluded — rows 1-2 are full-width merged banners, which
    // blocks column freezing. Its 3 frozen rows already provide good UX.
  ];
  freezes.forEach(function(f) {
    var sh = ss.getSheetByName(f.sheet);
    if (sh && sh.getFrozenColumns() === 0) sh.setFrozenColumns(f.cols);
  });
}

// ── SUMMARY BLOCKS ────────────────────────────────────────────
// Per-sheet stat panels placed outside the data area (to the right).
// Locations: Transactions W-X (23-24), Budget K-L (11-12),
//            Income J-K (10-11), Savings K-L (11-12), Subscriptions M-N (13-14)

function addSummaryBlocks(ss) {
  addTransactionsSummaryBlock(ss);
  addBudgetSummaryBlock(ss);
  addIncomeSummaryBlock(ss);
  addSavingsSummaryBlock(ss);
  addSubscriptionsSummaryBlock(ss);
}

function addTransactionsSummaryBlock(ss) {
  var sh = ss.getSheetByName(SHEETS.TRANSACTIONS);
  if (!sh) return;
  // Data cols are now A–Q (17). Summary block: spacer R(18), label S(19), value T(20).
  sh.setColumnWidth(18, 18);
  sh.setColumnWidth(19, 175);
  sh.setColumnWidth(20, 130);

  function cell(row, col) { return sh.getRange(row, col); }

  sh.getRange(1, 19, 1, 2).merge()
    .setValue('QUICK STATS')
    .setBackground(COLORS.PRIMARY).setFontColor(COLORS.HEADER_FG)
    .setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center').setVerticalAlignment('middle');

  cell(2, 19).setValue('Month:').setFontColor(COLORS.MID_TEXT).setFontStyle('italic');
  cell(2, 20).setFormula('=Dashboard!$C$2').setFontWeight('bold').setFontColor(COLORS.PRIMARY);

  var metrics = [
    [4, 'Spending (this month)',   '=SUMIFS(J:J,M:M,Dashboard!$C$2,C:C,"Expense")',           PKR_FORMAT, COLORS.DANGER],
    [5, 'Income (this month)',     '=SUMIFS(J:J,M:M,Dashboard!$C$2,C:C,"Income")',             PKR_FORMAT, COLORS.SUCCESS],
    [6, 'Shared expenses',         '=SUMIFS(J:J,M:M,Dashboard!$C$2,C:C,"Expense",L:L,"Yes")', PKR_FORMAT, COLORS.DARK_TEXT],
    [7, 'Personal expenses',       '=SUMIFS(J:J,M:M,Dashboard!$C$2,C:C,"Expense",L:L,"No")',  PKR_FORMAT, COLORS.DARK_TEXT],
    [8, 'Net (after group reimb.)',
     '=SUMIFS(J:J,M:M,Dashboard!$C$2,C:C,"Expense")' +
     '-IFERROR(SUMIF(\'Group Splits\'!J:J,Dashboard!$C$2,\'Group Splits\'!H:H),0)',
     PKR_FORMAT, COLORS.PRIMARY],
  ];
  metrics.forEach(function(m) {
    sh.getRange(m[0], 19, 1, 2).setBackground(COLORS.LIGHT_BLUE);
    cell(m[0], 19).setValue(m[1]).setFontWeight('bold').setFontColor(m[4]);
    cell(m[0], 20).setFormula(m[2]).setNumberFormat(m[3]);
  });

  var counts = [[10, 'Total entries', '=COUNTA(A2:A)']];
  counts.forEach(function(c) {
    sh.getRange(c[0], 19, 1, 2).setBackground(COLORS.ALT_ROW);
    cell(c[0], 19).setValue(c[1]).setFontWeight('bold');
    cell(c[0], 20).setFormula(c[2]);
  });

  sh.getRange(1, 19, 10, 2).setBorder(
    true, true, true, true, false, false, COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID
  );
}

function addBudgetSummaryBlock(ss) {
  var sh = ss.getSheetByName(SHEETS.BUDGET);
  if (!sh) return;
  sh.setColumnWidth(10, 18);
  sh.setColumnWidth(11, 185);
  sh.setColumnWidth(12, 130);

  sh.getRange(1, 11, 1, 2).merge()
    .setValue('BUDGET SUMMARY')
    .setBackground(COLORS.PRIMARY).setFontColor(COLORS.HEADER_FG)
    .setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center').setVerticalAlignment('middle');

  var sharedRows = [
    [3, 'Shared — Budget Total', '=SUMIF(C2:C200,"Shared",D2:D200)',  PKR_FORMAT],
    [4, 'Shared — Actual Total', '=SUMIF(C2:C200,"Shared",E2:E200)',  PKR_FORMAT],
    [5, 'Shared — Variance',     '=L3-L4',                            PKR_FORMAT],
  ];
  sharedRows.forEach(function(r) {
    sh.getRange(r[0], 11, 1, 2).setBackground(COLORS.LIGHT_BLUE);
    sh.getRange(r[0], 11).setValue(r[1]).setFontWeight('bold').setFontColor(COLORS.PRIMARY_DARK);
    sh.getRange(r[0], 12).setFormula(r[2]).setNumberFormat(r[3]);
  });

  var personalRows = [
    [7, 'Personal — Budget Total', '=SUMIF(C2:C200,"Personal",D2:D200)', PKR_FORMAT],
    [8, 'Personal — Actual Total', '=SUMIF(C2:C200,"Personal",E2:E200)', PKR_FORMAT],
    [9, 'Personal — Variance',     '=L7-L8',                             PKR_FORMAT],
  ];
  personalRows.forEach(function(r) {
    sh.getRange(r[0], 11, 1, 2).setBackground('#f3e8ff');
    sh.getRange(r[0], 11).setValue(r[1]).setFontWeight('bold').setFontColor('#4a148c');
    sh.getRange(r[0], 12).setFormula(r[2]).setNumberFormat(r[3]);
  });

  sh.getRange(11, 11, 1, 2).setBackground(COLORS.ALT_ROW);
  sh.getRange(11, 11).setValue('Categories over budget').setFontColor(COLORS.DANGER);
  sh.getRange(11, 12).setFormula('=COUNTIF(G2:G200,">1")').setFontColor(COLORS.DANGER);
  sh.getRange(12, 11, 1, 2).setBackground(COLORS.ALT_ROW);
  sh.getRange(12, 11).setValue('Categories near limit (>80%)').setFontColor('#b45309');
  sh.getRange(12, 12).setFormula('=COUNTIFS(G2:G200,">0.8",G2:G200,"<=1")').setFontColor('#b45309');

  var varCells = [sh.getRange(5, 12), sh.getRange(9, 12)];
  var rules = sh.getConditionalFormatRules();
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0).setBackground(COLORS.LIGHT_GREEN).setFontColor(COLORS.SUCCESS)
      .setRanges(varCells).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0).setBackground(COLORS.LIGHT_RED).setFontColor(COLORS.DANGER)
      .setRanges(varCells).build()
  );
  sh.setConditionalFormatRules(rules);

  sh.getRange(1, 11, 12, 2).setBorder(
    true, true, true, true, false, false, COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID
  );
}

function addIncomeSummaryBlock(ss) {
  var sh = ss.getSheetByName(SHEETS.INCOME);
  if (!sh) return;
  sh.setColumnWidth(9,  18);
  sh.setColumnWidth(10, 160);
  sh.setColumnWidth(11, 130);

  sh.getRange(1, 10, 1, 2).merge()
    .setValue('INCOME SUMMARY')
    .setBackground(COLORS.SUCCESS).setFontColor(COLORS.HEADER_FG)
    .setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center').setVerticalAlignment('middle');

  var personRows = [
    [3, PERSON1 + ' — Total', '=SUMIF(B:B,"' + PERSON1 + '",E:E)'],
    [4, PERSON2 + ' — Total', '=SUMIF(B:B,"' + PERSON2 + '",E:E)'],
    [5, 'Combined Total',     '=SUM(E2:E)'],
  ];
  personRows.forEach(function(r) {
    sh.getRange(r[0], 10, 1, 2).setBackground(COLORS.LIGHT_GREEN);
    sh.getRange(r[0], 10).setValue(r[1]).setFontWeight('bold');
    sh.getRange(r[0], 11).setFormula(r[2]).setNumberFormat(PKR_FORMAT);
  });
  sh.getRange(5, 10).setFontColor(COLORS.SUCCESS);

  sh.getRange(7, 10, 1, 2).setBackground(COLORS.ALT_ROW);
  sh.getRange(7, 10).setValue('Entry count');
  sh.getRange(7, 11).setFormula('=COUNTA(A2:A)');
  sh.getRange(8, 10, 1, 2).setBackground(COLORS.ALT_ROW);
  sh.getRange(8, 10).setValue('Sources (unique)');
  sh.getRange(8, 11).setFormula('=IFERROR(ROWS(UNIQUE(FILTER(C2:C,C2:C<>""))),0)');

  sh.getRange(1, 10, 8, 2).setBorder(
    true, true, true, true, false, false, COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID
  );
}

function addSavingsSummaryBlock(ss) {
  var sh = ss.getSheetByName(SHEETS.SAVINGS);
  if (!sh) return;
  sh.setColumnWidth(10, 18);
  sh.setColumnWidth(11, 175);
  sh.setColumnWidth(12, 130);

  sh.getRange(1, 11, 1, 2).merge()
    .setValue('SAVINGS SUMMARY')
    .setBackground(COLORS.PRIMARY_DARK).setFontColor(COLORS.HEADER_FG)
    .setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center').setVerticalAlignment('middle');

  var metricRows = [
    [3, 'Total Target',       '=SUM(B2:B)',                       PKR_FORMAT],
    [4, 'Total Saved So Far', '=SUM(C2:C)',                       PKR_FORMAT],
    [5, 'Total Remaining',    '=SUM(F2:F)',                       PKR_FORMAT],
    [6, 'Overall Progress',   '=IFERROR(SUM(C2:C)/SUM(B2:B),0)', PCT_FORMAT],
  ];
  metricRows.forEach(function(r) {
    sh.getRange(r[0], 11, 1, 2).setBackground(COLORS.LIGHT_BLUE);
    sh.getRange(r[0], 11).setValue(r[1]).setFontWeight('bold');
    sh.getRange(r[0], 12).setFormula(r[2]).setNumberFormat(r[3]);
  });

  sh.getRange(8, 11, 1, 2).setBackground(COLORS.ALT_ROW);
  sh.getRange(8, 11).setValue('Goals ≥ 50% complete');
  sh.getRange(8, 12).setFormula('=COUNTIF(E2:E,">=0.5")').setFontColor(COLORS.SUCCESS);
  sh.getRange(9, 11, 1, 2).setBackground(COLORS.ALT_ROW);
  sh.getRange(9, 11).setValue('Goals completed (100%)').setFontWeight('bold').setFontColor(COLORS.SUCCESS);
  sh.getRange(9, 12).setFormula('=COUNTIF(E2:E,">=1")').setFontWeight('bold').setFontColor(COLORS.SUCCESS);

  sh.getRange(1, 11, 9, 2).setBorder(
    true, true, true, true, false, false, COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID
  );
}

function addSubscriptionsSummaryBlock(ss) {
  var sh = ss.getSheetByName(SHEETS.SUBSCRIPTIONS);
  if (!sh) return;
  // Data cols A-M (13). Spacer at N(14), summary label at O(15), value at P(16).
  sh.setColumnWidth(14, 18);   // N spacer
  sh.setColumnWidth(15, 190);  // O label
  sh.setColumnWidth(16, 130);  // P value

  sh.getRange(1, 15, 1, 2).merge()
    .setValue('SUBSCRIPTIONS SUMMARY')
    .setBackground(COLORS.MID_TEXT).setFontColor(COLORS.HEADER_FG)
    .setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center').setVerticalAlignment('middle');

  // K = Status, H = Expected Reimb., I = Your Net Cost, J = Renewal Date
  var costRows = [
    [3, 'Active — Gross monthly cost', '=SUMIF(K2:K,"Active",C2:C)', PKR_FORMAT],
    [4, 'Active — Exp. reimbursed',    '=SUMIF(K2:K,"Active",H2:H)', PKR_FORMAT],
    [5, 'Active — Net monthly cost',   '=SUMIF(K2:K,"Active",I2:I)', PKR_FORMAT],
  ];
  costRows.forEach(function(r) {
    sh.getRange(r[0], 15, 1, 2).setBackground(COLORS.ALT_ROW);
    sh.getRange(r[0], 15).setValue(r[1]).setFontWeight('bold');
    sh.getRange(r[0], 16).setFormula(r[2]).setNumberFormat(r[3]);
  });

  sh.getRange(7, 15, 1, 2).setBackground(COLORS.LIGHT_YELLOW);
  sh.getRange(7, 15).setValue('Active count');
  sh.getRange(7, 16).setFormula('=COUNTIF(K2:K,"Active")');
  sh.getRange(8, 15, 1, 2).setBackground(COLORS.LIGHT_YELLOW);
  sh.getRange(8, 15).setValue('Active with splits');
  sh.getRange(8, 16).setFormula('=COUNTIFS(K2:K,"Active",D2:D,"<>"&"")');
  sh.getRange(9, 15, 1, 2).setBackground(COLORS.LIGHT_YELLOW);
  sh.getRange(9, 15).setValue('Renewals within 7 days');
  sh.getRange(9, 16).setFormula('=COUNTIFS(K2:K,"Active",J2:J,"<="&(TODAY()+7),J2:J,">="&TODAY())')
    .setFontColor(COLORS.WARNING);

  sh.getRange(1, 15, 9, 2).setBorder(
    true, true, true, true, false, false, COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID
  );
}

// ── SECTION DIVIDERS ──────────────────────────────────────────

function addSectionDividers(ss) {
  // Budget: color-code the Type column to separate Shared vs Personal rows
  var bud = ss.getSheetByName(SHEETS.BUDGET);
  var typeRange = bud.getRange('C2:C198');
  var budRules = bud.getConditionalFormatRules();
  budRules.unshift(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($C2<>"",$C2="Personal")')
      .setBackground('#f3e8ff').setRanges([typeRange]).build()
  );
  budRules.unshift(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($C2<>"",$C2="Shared")')
      .setBackground(COLORS.LIGHT_BLUE).setRanges([typeRange]).build()
  );
  bud.setConditionalFormatRules(budRules);

  // Thick bottom border on header rows across all data sheets
  var hdrBorders = [
    { sheet: SHEETS.TRANSACTIONS,  row: 1, cols: 16 },
    { sheet: SHEETS.BUDGET,        row: 1, cols: 8  },
    { sheet: SHEETS.INCOME,        row: 1, cols: 7  },
    { sheet: SHEETS.SAVINGS,       row: 1, cols: 8  },
    { sheet: SHEETS.SUBSCRIPTIONS, row: 1, cols: 13 },
    { sheet: SHEETS.GROUP_SPLITS,  row: 3, cols: 11 },
  ];
  hdrBorders.forEach(function(s) {
    var sh = ss.getSheetByName(s.sheet);
    if (sh) sh.getRange(s.row, 1, 1, s.cols).setBorder(
      null, null, true, null, null, null,
      COLORS.PRIMARY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM
    );
  });

  // Reimbursements: medium border around summary panel
  var reimb = ss.getSheetByName(SHEETS.REIMBURSEMENTS);
  if (reimb) reimb.getRange('K3:L6').setBorder(
    true, true, true, true, null, null,
    COLORS.PRIMARY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );

  // Group Splits: medium border around totals panel
  var gs = ss.getSheetByName(SHEETS.GROUP_SPLITS);
  if (gs) gs.getRange('M3:N6').setBorder(
    true, true, true, true, null, null,
    COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );
}
