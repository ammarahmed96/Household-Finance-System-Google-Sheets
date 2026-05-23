// ============================================================
// HOUSEHOLD FINANCE TRACKER — Reimbursements.gs
// Reimbursements tab (3-section QUERY view) and Group Splits tab.
//
// Reimbursements aggregates live data from:
//   Section 1 — Transactions where N (Reimbursable?) = Yes
//   Section 2 — Subscriptions where Expected Reimb. > 0
//   Section 3 — Group Splits where Outstanding > 0
//
// Group Splits column map:
//   A=Date, B=Paid By, C=Event/Description, D=Category,
//   E=Total Bill(PKR), F=Owed By, G=Their Share(PKR),
//   H=Reimbursed(PKR), I=Outstanding(PKR)[FORMULA], J=Month[FORMULA], K=Notes
// ============================================================

function setupReimbursements(ss) {
  var sh = ss.getSheetByName(SHEETS.REIMBURSEMENTS);
  sh.clearContents();
  sh.clearFormats();

  // ── Column headers (row 1) ────────────────────────────────
  var hdrs = ['Date', 'Paid By', 'Event / Description', 'Owed By', 'Their Share (PKR)', 'Received (PKR)', 'Outstanding (PKR)'];
  sh.getRange(1, 1, 1, hdrs.length).setValues([hdrs]);
  styleHeaderRow(sh.getRange(1, 1, 1, hdrs.length), COLORS.PRIMARY_DARK);
  sh.setFrozenRows(1);

  // ── QUERY: outstanding group splits ──────────────────────
  sh.getRange('A2').setFormula(
    "=IFERROR(QUERY('Group Splits'!A:I," +
    "\"SELECT A,B,C,F,G,H,I WHERE F<>'' AND I>0 ORDER BY A DESC\",0)," +
    '{"No outstanding group splits","","","","","",""})'
  );

  // ── Summary panel (two-column table: J1:K4) ──────────────
  sh.getRange('J1').setValue('SUMMARY');
  sh.getRange('K1').setValue('Amount (PKR)');
  styleHeaderRow(sh.getRange(1, 10, 1, 2));

  var summaryRows = [
    ['Outstanding:',     "=IFERROR(SUM('Group Splits'!I:I),0)"],
    ['Total Received:',  "=IFERROR(SUM('Group Splits'!H:H),0)"],
    ['Total Expected:',  "=IFERROR(SUM('Group Splits'!G:G),0)"],
  ];
  summaryRows.forEach(function(r, i) {
    sh.getRange(2 + i, 10).setValue(r[0]).setFontWeight(i === 0 ? 'bold' : 'normal');
    sh.getRange(2 + i, 11).setFormula(r[1]).setNumberFormat(PKR_FORMAT)
      .setFontWeight(i === 0 ? 'bold' : 'normal')
      .setFontColor(i === 0 ? COLORS.DANGER : COLORS.DARK_TEXT);
  });

  // ── Column widths ─────────────────────────────────────────
  [100, 90, 220, 160, 130, 130, 140, 30, 30, 210, 120]
    .forEach(function(w, i) { sh.setColumnWidth(i + 1, w); });
}

function setupGroupSplits(ss) {
  var sh    = ss.getSheetByName(SHEETS.GROUP_SPLITS);
  var lists = ss.getSheetByName(SHEETS.LISTS);
  sh.clearContents();
  sh.clearFormats();

  // ── Column headers (row 1) ────────────────────────────────
  var headers = [
    'Date', 'Paid By', 'Event / Description', 'Category',
    'Total Bill (PKR)', 'Owed By', 'Their Share (PKR)',
    'Reimbursed (PKR)', 'Outstanding (PKR)', 'Month', 'Notes', 'Tx Key'
  ];
  var hdr = sh.getRange(1, 1, 1, headers.length);
  hdr.setValues([headers]);
  styleHeaderRow(hdr);
  sh.setFrozenRows(1);

  [100, 90, 220, 130, 130, 160, 130, 120, 130, 115, 220, 1]
    .forEach(function(w, i) { sh.setColumnWidth(i + 1, w); });
  sh.hideColumns(GS_COL_TX_KEY);  // L — internal link key, invisible to user

  // ── Dropdowns ─────────────────────────────────────────────
  var ROWS = 500;
  setDropdown(sh, 2, 2, ROWS, lists.getRange('A2:A3'));
  setDropdown(sh, 2, 4, ROWS, lists.getRange('C2:C15'));

  // ── Number formats ────────────────────────────────────────
  sh.getRange(2, 1, ROWS, 1).setNumberFormat('MM/DD/YYYY');
  sh.getRange(2, 5, ROWS, 1).setNumberFormat(PKR_FORMAT);
  sh.getRange(2, 7, ROWS, 1).setNumberFormat(PKR_FORMAT);
  sh.getRange(2, 8, ROWS, 1).setNumberFormat(PKR_FORMAT);
  sh.getRange(2, 9, ROWS, 1).setNumberFormat(PKR_FORMAT);

  // ── Sample data ───────────────────────────────────────────
  var today = new Date();
  // I(9,Outstanding) and J(10,Month) are ARRAYFORMULA columns — skip them in setValues.
  var samples = [
    [today, PERSON1, 'Dinner at Salt & Pepper', 'Food', 4800, 'Ali',    1200, 0, 'Karachi trip'],
    [today, PERSON1, 'Dinner at Salt & Pepper', 'Food', 4800, 'Hassan', 1200, 0, ''],
    [today, PERSON1, 'Dinner at Salt & Pepper', 'Food', 4800, 'Bilal',  1200, 0, ''],
    [today, PERSON2, 'Shopping trip',            'Personal Care', 6000, 'Sara', 2000, 0, 'Mall run'],
  ];
  sh.getRange(2, 1, samples.length, 8).setValues(samples.map(function(r) { return r.slice(0, 8); }));
  sh.getRange(2, 11, samples.length, 1).setValues(samples.map(function(r) { return [r[8]]; }));
  sh.getRange(2, 1, samples.length, 1).setNumberFormat('MM/DD/YYYY');
  sh.getRange(2, 5, samples.length, 1).setNumberFormat(PKR_FORMAT);
  sh.getRange(2, 7, samples.length, 2).setNumberFormat(PKR_FORMAT);

  // ── Totals panel (two-column table: M1:N4) ───────────────
  sh.getRange('M1').setValue('TOTALS');
  sh.getRange('N1').setValue('Amount (PKR)');
  styleHeaderRow(sh.getRange(1, 13, 1, 2));
  sh.getRange('M2').setValue('Total Outstanding:');
  sh.getRange('N2').setFormula('=IFERROR(SUM(I2:I1000),0)').setNumberFormat(PKR_FORMAT).setFontWeight('bold').setFontColor(COLORS.DANGER);
  sh.getRange('M3').setValue('Total Reimbursed:');
  sh.getRange('N3').setFormula('=IFERROR(SUM(H2:H1000),0)').setNumberFormat(PKR_FORMAT);
  sh.getRange('M4').setValue('Total Expected:');
  sh.getRange('N4').setFormula('=IFERROR(SUM(G2:G1000),0)').setNumberFormat(PKR_FORMAT);
  sh.setColumnWidth(13, 160);
  sh.setColumnWidth(14, 120);
}
