// ============================================================
// HOUSEHOLD FINANCE TRACKER — Dashboard.gs
// Dashboard sheet layout, conditional formatting, and charts.
//
// Dashboard layout (v2):
//   Row  1  — Full-width header banner (B:N)
//   Row  2  — Month/Year selector (C2 = month dropdown, D2 = year dropdown, E2 = helper "May 2026")
//   Row  3  — 4px solid PRIMARY divider
//   Rows 4-8  — Card row 1: Income | Spending | Net Cash Flow | Savings Rate
//   Rows 9-13 — Card row 2: Shared | Personal | Reimb. | Subscriptions
//   Row  14 — Thin separator
//   Row  15 — Section titles: "BUDGET VS ACTUAL" | "SPENDING BY CATEGORY"
//   Row  16 — Table headers
//   Rows 17-27 — Budget QUERY | Category breakdown (10 rows)
//   Row  29 — Thin separator
//   Row  30 — "RECENT TRANSACTIONS" title
//   Row  31 — Transaction table headers
//   Rows 32-56 — QUERY (last 25 transactions)
//   Row  57+ — Charts (added by addDashboardCharts())
// ============================================================

function setupDashboard(ss) {
  var sh = ss.getSheetByName(SHEETS.DASHBOARD);
  sh.clear();
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations();

  var tz = Session.getScriptTimeZone();
  var now = new Date();
  var curMonthName = Utilities.formatDate(now, tz, 'MMMM');
  var curYear = Utilities.formatDate(now, tz, 'yyyy');
  var baseYear = parseInt(curYear, 10);
  var yearsList = [];
  for (var y = baseYear - 2; y <= baseYear + 3; y++) { yearsList.push(String(y)); }
  var lists = ss.getSheetByName(SHEETS.LISTS);

  // ── Column widths ─────────────────────────────────────────
  sh.setColumnWidth(1, 20);   // A left margin
  [120, 120, 120,  // B C D — card 1
   120, 120, 120,  // E F G — card 2
   120, 120, 120,  // H I J — card 3
   120, 120, 120,  // K L M — card 4
   24              // N     — right edge
  ].forEach(function(w, i) { sh.setColumnWidth(i + 2, w); });

  // ── Row heights ───────────────────────────────────────────
  sh.setRowHeight(1, 52);
  sh.setRowHeight(2, 36);
  sh.setRowHeight(3, 6);
  sh.setRowHeight(4, 10);
  sh.setRowHeight(5, 22);
  sh.setRowHeight(6, 44);
  sh.setRowHeight(7, 18);
  sh.setRowHeight(8, 10);
  sh.setRowHeight(9,  10);
  sh.setRowHeight(10, 22);
  sh.setRowHeight(11, 44);
  sh.setRowHeight(12, 18);
  sh.setRowHeight(13, 10);
  sh.setRowHeight(14, 8);
  sh.setRowHeight(15, 28);
  sh.setRowHeight(16, 36);

  // ── Banner (row 1) ────────────────────────────────────────
  sh.getRange('B1:N1').merge();
  sh.getRange('B1')
    .setValue('Household Finance Dashboard  •  PKR')
    .setFontSize(18).setFontWeight('bold')
    .setFontColor('#ffffff').setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground(COLORS.PRIMARY);

  // ── Month selector (row 2) ────────────────────────────────
  // C2 = month name dropdown ("May"), D2 = year dropdown ("2026"),
  // E2 = formula helper =C2&" "&D2 → "May 2026" (used by all reporting formulas).
  sh.getRange('B2:N2').setBackground('#f1f3f4');
  sh.getRange('B2').setValue('Month / Year:')
    .setFontWeight('bold').setFontColor(COLORS.MID_TEXT)
    .setFontSize(10).setVerticalAlignment('middle');
  sh.getRange('C2')
    .setValue(curMonthName)
    .setFontSize(12).setFontWeight('bold').setFontColor(COLORS.PRIMARY)
    .setVerticalAlignment('middle');
  setDropdown(sh, 2, 3, 1, lists.getRange('I2:I13'));
  sh.getRange('D2')
    .setValue(curYear)
    .setFontSize(12).setFontWeight('bold').setFontColor(COLORS.PRIMARY)
    .setVerticalAlignment('middle');
  sh.getRange('D2').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(yearsList, true)
      .setAllowInvalid(false)
      .build()
  );
  sh.getRange('E2')
    .setFormula('=C2&" "&D2')
    .setFontSize(11).setFontWeight('bold').setFontColor(COLORS.PRIMARY)
    .setVerticalAlignment('middle');
  sh.getRange('F2:H2').merge();
  sh.getRange('F2')
    .setValue('  ← Select month and year above')
    .setFontColor(COLORS.MID_TEXT).setFontStyle('italic').setFontSize(9)
    .setVerticalAlignment('middle');

  // ── Divider (row 3) ───────────────────────────────────────
  sh.getRange('B3:N3').setBackground(COLORS.PRIMARY);
  sh.setFrozenRows(3);

  // ── Card definitions ──────────────────────────────────────
  var cards1 = [
    { label: 'Monthly Income',    subLabel: 'Total for selected month',
      f: '=SUMIFS(Income!E:E,Income!F:F,E2)',
      fmt: PKR_FORMAT, col: 2, accentColor: COLORS.SUCCESS, bg: COLORS.LIGHT_GREEN },
    { label: 'Total Spending',    subLabel: 'Personal + shared, net of reimbursements',
      f: '=B11+E11',
      fmt: PKR_FORMAT, col: 5, accentColor: COLORS.DANGER, bg: COLORS.LIGHT_RED },
    { label: 'Net Cash Flow',     subLabel: 'Income minus net spending',
      f: '=B6-E6',
      fmt: PKR_FORMAT, col: 8, accentColor: COLORS.PRIMARY, bg: COLORS.LIGHT_BLUE },
    { label: 'Savings Rate',      subLabel: '% of income saved',
      f: '=IFERROR((B6-E6)/B6,0)',
      fmt: PCT_FORMAT, col: 11, accentColor: COLORS.PRIMARY_DARK, bg: '#f3e8ff' },
  ];

  var cards2 = [
    { label: 'Shared Spending',    subLabel: 'Net of group reimbursements',
      f: SHARED_NET_FORMULA,
      fmt: PKR_FORMAT, col: 2, accentColor: COLORS.PRIMARY, bg: COLORS.LIGHT_BLUE },
    { label: 'Personal Spending',  subLabel: 'Net of group reimbursements',
      f: PERSONAL_NET_FORMULA,
      fmt: PKR_FORMAT, col: 5, accentColor: '#e65100', bg: '#fff3e0' },
    { label: 'Group Outstanding',  subLabel: 'Friends & family owe you',
      f: "=IFERROR(SUM('Group Splits'!I:I),0)",
      fmt: PKR_FORMAT, col: 8, accentColor: COLORS.WARNING, bg: COLORS.LIGHT_YELLOW },
    { label: 'Subscriptions Cost', subLabel: 'Net monthly after splits',
      f: '=IFERROR(SUMIF(Subscriptions!I:I,">0",Subscriptions!I:I),0)',
      fmt: PKR_FORMAT, col: 11, accentColor: COLORS.MID_TEXT, bg: '#f8f9fa' },
  ];

  // ── Render card rows ──────────────────────────────────────
  function renderCardRow(cards, padTop, labelRow, valueRow, subRow, padBot) {
    cards.forEach(function(c) {
      var col = c.col;
      sh.getRange(padTop, col, padBot - padTop + 1, 3)
        .setBackground(c.bg);
      sh.getRange(padTop, col, padBot - padTop + 1, 3)
        .setBorder(
          true, true, true, true, false, false,
          COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID
        );
      sh.getRange(padTop, col, padBot - padTop + 1, 1)
        .setBorder(
          null, true, null, null, null, null,
          c.accentColor, SpreadsheetApp.BorderStyle.SOLID_MEDIUM
        );

      sh.getRange(labelRow, col, 1, 3).merge();
      sh.getRange(labelRow, col)
        .setValue(c.label)
        .setFontWeight('bold').setFontColor(COLORS.DARK_TEXT)
        .setFontSize(10).setHorizontalAlignment('center')
        .setVerticalAlignment('middle');

      sh.getRange(valueRow, col, 1, 3).merge();
      sh.getRange(valueRow, col)
        .setFormula(c.f)
        .setNumberFormat(c.fmt)
        .setFontSize(20).setFontWeight('bold')
        .setFontColor(c.accentColor).setHorizontalAlignment('center')
        .setVerticalAlignment('middle');

      sh.getRange(subRow, col, 1, 3).merge();
      sh.getRange(subRow, col)
        .setValue(c.subLabel)
        .setFontColor(COLORS.MID_TEXT).setFontSize(8)
        .setHorizontalAlignment('center').setVerticalAlignment('top')
        .setFontStyle('italic');
    });
  }

  renderCardRow(cards1, 4, 5, 6, 7, 8);
  renderCardRow(cards2, 9, 10, 11, 12, 13);

  // ── Thin separator before data tables ────────────────────
  sh.getRange('B14:N14').setBackground(COLORS.BORDER);

  // ── Section titles (row 15) ───────────────────────────────
  sh.getRange('B15:G15').merge();
  sh.getRange('B15')
    .setValue('BUDGET VS ACTUAL')
    .setFontSize(11).setFontWeight('bold').setFontColor(COLORS.PRIMARY_DARK)
    .setVerticalAlignment('middle');

  sh.getRange('I15:K15').merge();
  sh.getRange('I15')
    .setValue('SPENDING BY CATEGORY')
    .setFontSize(11).setFontWeight('bold').setFontColor(COLORS.PRIMARY_DARK)
    .setVerticalAlignment('middle');

  sh.getRange('L15:M15').merge();
  sh.getRange('L15')
    .setValue('SAVINGS GOALS')
    .setFontSize(11).setFontWeight('bold').setFontColor(COLORS.PRIMARY_DARK)
    .setVerticalAlignment('middle');

  sh.setColumnWidth(8, 20); // gutter between the two sections

  // ── Budget table headers (row 16) ─────────────────────────
  sh.getRange(16, 2, 1, 6).setValues([['Category', 'Type', 'Budget (PKR)', 'Actual (PKR)', 'Variance', '% Used']]);
  styleHeaderRow(sh.getRange(16, 2, 1, 6));

  // ── Budget QUERY (rows 17+) ───────────────────────────────
  sh.getRange('B17').setFormula(
    '=IFERROR(QUERY(Budget!A:G,' +
    '"SELECT B,C,D,E,F,G WHERE A=\'"&E2&"\' AND B<>\'\'",0),' +
    '{"No budget data for this month","","","","",""})'
  );

  // ── Category Breakdown headers (row 16, cols I-K) ─────────
  sh.getRange(16, 9, 1, 3).setValues([['Category', 'Amount (PKR)', 'Share %']]);
  styleHeaderRow(sh.getRange(16, 9, 1, 3));

  sh.setColumnWidth(9, 130);
  sh.setColumnWidth(10, 120);
  sh.setColumnWidth(11, 72);

  var cats = [
    'Housing', 'Food', 'Transportation', 'Health',
    'Entertainment', 'Subscriptions', 'Personal Care',
    'Travel', 'Gifts', 'Other'
  ];
  cats.forEach(function(cat, i) {
    var row = 17 + i;
    sh.getRange(row, 9).setValue(cat).setFontColor(COLORS.DARK_TEXT).setFontSize(10);
    sh.getRange(row, 10).setFormula(buildCategoryNetFormula(cat));
    sh.getRange(row, 11).setFormula('=IFERROR(J' + row + '/$E$6,0)');
    if (i % 2 === 1) {
      sh.getRange(row, 9, 1, 3).setBackground(COLORS.ALT_ROW);
    }
  });
  sh.getRange(17, 10, cats.length, 1).setNumberFormat(PKR_FORMAT);
  sh.getRange(17, 11, cats.length, 1).setNumberFormat(PCT_FORMAT);

  // ── Outer borders on table blocks ────────────────────────
  sh.getRange(16, 2, 12, 6)
    .setBorder(true, true, true, true, null, null, COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(16, 9, 11, 3)
    .setBorder(true, true, true, true, null, null, COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID);

  // ── Savings Goals section (cols L-M, rows 16-26) ─────────
  sh.getRange(16, 12, 1, 2).setValues([['Savings Goal', 'Progress']]);
  styleHeaderRow(sh.getRange(16, 12, 1, 2));
  sh.getRange(17, 12).setFormula(
    '=IFERROR(FILTER({\'Savings Goals\'!A2:A10,\'Savings Goals\'!E2:E10},\'Savings Goals\'!A2:A10<>""),{"No goals",""})'
  );
  sh.getRange(17, 13, 10, 1).setNumberFormat(PCT_FORMAT);
  sh.getRange(16, 12, 11, 2)
    .setBorder(true, true, true, true, null, null, COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID);

  // ── Recent Transactions section ───────────────────────────
  var txSepRow  = 29;
  var txTitleRow = 30;
  var txHdrRow  = 31;
  var txDataRow = 32;

  sh.getRange(txSepRow, 2, 1, 13).setBackground(COLORS.BORDER);
  sh.setRowHeight(txSepRow, 6);

  sh.getRange(txTitleRow, 2, 1, 13).merge();
  sh.getRange(txTitleRow, 2)
    .setValue('RECENT TRANSACTIONS  (this month, newest first)')
    .setFontSize(11).setFontWeight('bold').setFontColor(COLORS.PRIMARY_DARK)
    .setVerticalAlignment('middle');
  sh.setRowHeight(txTitleRow, 28);

  sh.getRange(txHdrRow, 2, 1, 8).setValues([
    ['Date', 'Person', 'Category', 'Description', 'Currency', 'Orig. Amt.', 'PKR Amt.', 'Type']
  ]);
  styleHeaderRow(sh.getRange(txHdrRow, 2, 1, 8));

  sh.getRange(txDataRow, 2).setFormula(
    '=IFERROR(QUERY(Transactions!A:N,' +
    '"SELECT A,B,D,F,G,H,J,C WHERE M=\'"&E2&"\' AND A IS NOT NULL ' +
    'ORDER BY A DESC LIMIT 25",0),' +
    '{"No transactions this month","","","","","","",""})'
  );

  sh.getRange(txDataRow, 2, 25, 1).setNumberFormat('DD/MM/YYYY');
  sh.getRange(txDataRow, 7, 25, 1).setNumberFormat(PKR_FORMAT);
  sh.getRange(txDataRow, 8, 25, 1).setNumberFormat(PKR_FORMAT);

  sh.getRange(txHdrRow, 2, 26, 8)
    .setBorder(true, true, true, true, null, null, COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID);

  // ── Final column width pass ────────────────────────────────
  [110, 80, 110, 110, 100, 80].forEach(function(w, i) { sh.setColumnWidth(i + 2, w); });
  sh.setColumnWidth(8,  20);
  sh.setColumnWidth(9,  130);
  sh.setColumnWidth(10, 120);
  sh.setColumnWidth(11, 72);
  sh.setColumnWidth(12, 130);
  sh.setColumnWidth(13, 120);
  sh.setColumnWidth(14, 80);
}

// ── DASHBOARD CONDITIONAL FORMATTING ─────────────────────────

function formatDashboard(ss) {
  var sh = ss.getSheetByName(SHEETS.DASHBOARD);

  // Net Cash Flow (H6): green if positive, red if negative
  var cashFlowCell = sh.getRange('H6');
  var cfPos = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0).setFontColor(COLORS.SUCCESS).setRanges([cashFlowCell]).build();
  var cfNeg = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0).setFontColor(COLORS.DANGER).setRanges([cashFlowCell]).build();

  // Savings Rate (K6): good ≥ 20%, warning 10-20%, danger < 10%
  var srCell = sh.getRange('K6');
  var srGood = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(0.2).setFontColor(COLORS.SUCCESS).setRanges([srCell]).build();
  var srWarn = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(0.1, 0.199).setFontColor(COLORS.WARNING).setRanges([srCell]).build();
  var srBad = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0.1).setFontColor(COLORS.DANGER).setRanges([srCell]).build();

  // Outstanding Reimb. (H11): red when > 0, green when 0
  var reimbCell = sh.getRange('H11');
  var reimbRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0).setFontColor(COLORS.DANGER).setRanges([reimbCell]).build();
  var reimbZero = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(0).setFontColor(COLORS.SUCCESS).setRanges([reimbCell]).build();

  // Budget % Used column (col G, rows 17-27)
  var pctRange = sh.getRange(17, 7, 11, 1);
  var budOver = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(1)
    .setBackground(COLORS.LIGHT_RED).setFontColor(COLORS.DANGER)
    .setRanges([pctRange]).build();
  var budNear = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(0.8, 1)
    .setBackground(COLORS.LIGHT_YELLOW).setFontColor('#b45309')
    .setRanges([pctRange]).build();
  var budOk = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0.8)
    .setBackground(COLORS.LIGHT_GREEN).setFontColor('#166534')
    .setRanges([pctRange]).build();

  // Savings Goals progress (col M, rows 17-26)
  var savRange = sh.getRange(17, 13, 10, 1);
  var savDone = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(1).setFontColor(COLORS.SUCCESS).setRanges([savRange]).build();
  var savNear = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(0.5, 0.999).setFontColor(COLORS.WARNING).setRanges([savRange]).build();
  var savLow = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0.5).setFontColor(COLORS.DANGER).setRanges([savRange]).build();

  // Recent Transactions (rows 32-56): income rows green, alternating otherwise
  var txData = sh.getRange(32, 2, 25, 8);
  var txIncome = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$I32="Income"')
    .setBackground(COLORS.LIGHT_GREEN).setRanges([txData]).build();
  var txAlt = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=ISEVEN(ROW())')
    .setBackground(COLORS.ALT_ROW).setRanges([txData]).build();

  sh.setConditionalFormatRules([
    cfPos, cfNeg,
    srGood, srWarn, srBad,
    reimbRule, reimbZero,
    budOver, budNear, budOk,
    savDone, savNear, savLow,
    txIncome, txAlt,
  ]);
}

// ── DASHBOARD CHARTS ──────────────────────────────────────────
// Run after entering a full month of data. Charts placed at row 57
// so they never overlap the Recent Transactions table.

function addDashboardCharts() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEETS.DASHBOARD);

  sh.getCharts().forEach(function(c) { sh.removeChart(c); });

  // Donut chart — Spending by Category (data at I17:J26)
  var pie = sh.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(sh.getRange('I17:J26'))
    .setOption('title', 'Spending by Category')
    .setOption('pieHole', 0.45)
    .setOption('legend', {position: 'right', textStyle: {fontSize: 11}})
    .setOption('chartArea', {left: 10, top: 30, width: '55%', height: '85%'})
    .setOption('backgroundColor', '#ffffff')
    .setPosition(57, 2, 0, 0)
    .build();
  sh.insertChart(pie);

  // Bar chart — Budget vs Actual (B17:B27 + D17:E27)
  var bar = sh.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(sh.getRange('B17:B27'))
    .addRange(sh.getRange('D17:E27'))
    .setOption('title', 'Budget vs Actual')
    .setOption('hAxis', {title: 'PKR', format: 'Rs #,##0'})
    .setOption('vAxis', {title: ''})
    .setOption('series', {0: {color: COLORS.PRIMARY}, 1: {color: COLORS.DANGER}})
    .setOption('legend', {position: 'top', textStyle: {fontSize: 11}})
    .setOption('chartArea', {left: 120, top: 40, width: '75%', height: '82%'})
    .setOption('backgroundColor', '#ffffff')
    .setPosition(57, 9, 0, 0)
    .build();
  sh.insertChart(bar);

  ss.toast('Charts added below the transactions table (row 57).', 'Done', 5);
}
