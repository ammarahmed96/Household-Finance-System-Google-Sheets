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
//
// Note: Group Splits rows 1-2 are full-width merged banner cells
// (A1:K1 and A2:K2). This blocks setFrozenColumns — do not add
// column freezes for this sheet in enhanceFrozenPanes().
// ============================================================

function setupReimbursements(ss) {
  var sh = ss.getSheetByName(SHEETS.REIMBURSEMENTS);
  sh.clearContents();
  sh.clearFormats();

  sh.setColumnWidth(1, 30);

  // ── Title ─────────────────────────────────────────────────
  sh.getRange('B1').setValue('REIMBURSEMENTS — Money Owed to Me')
    .setFontSize(16).setFontWeight('bold').setFontColor(COLORS.PRIMARY);

  // ── Column headers ────────────────────────────────────────
  var hdrs = ['Date', 'Paid By', 'Event / Description', 'Owed By', 'Their Share (PKR)', 'Received (PKR)', 'Outstanding (PKR)'];
  sh.getRange(3, 2, 1, hdrs.length).setValues([hdrs]);
  styleHeaderRow(sh.getRange(3, 2, 1, hdrs.length), COLORS.PRIMARY_DARK);
  sh.setFrozenRows(3);

  // ── QUERY: outstanding group splits ──────────────────────
  sh.getRange('B4').setFormula(
    "=IFERROR(QUERY('Group Splits'!A:I," +
    "\"SELECT A,B,C,F,G,H,I WHERE F<>'' AND I>0 ORDER BY A DESC\",0)," +
    '{"No outstanding group splits","","","","","",""})'
  );

  // ── Summary panel (two-column table: K3:L6) ──────────────
  sh.getRange('K3').setValue('SUMMARY');
  sh.getRange('L3').setValue('Amount (PKR)');
  styleHeaderRow(sh.getRange(3, 11, 1, 2));

  var summaryRows = [
    ['Outstanding:',     "=IFERROR(SUM('Group Splits'!I:I),0)"],
    ['Total Received:',  "=IFERROR(SUM('Group Splits'!H:H),0)"],
    ['Total Expected:',  "=IFERROR(SUM('Group Splits'!G:G),0)"],
  ];
  summaryRows.forEach(function(r, i) {
    sh.getRange(4 + i, 11).setValue(r[0]).setFontWeight(i === 0 ? 'bold' : 'normal');
    sh.getRange(4 + i, 12).setFormula(r[1]).setNumberFormat(PKR_FORMAT)
      .setFontWeight(i === 0 ? 'bold' : 'normal')
      .setFontColor(i === 0 ? COLORS.DANGER : COLORS.DARK_TEXT);
  });

  // ── Column widths ─────────────────────────────────────────
  [100, 90, 220, 160, 130, 130, 140, 30, 30, 210, 120]
    .forEach(function(w, i) { sh.setColumnWidth(i + 2, w); });
}

function setupGroupSplits(ss) {
  var sh    = ss.getSheetByName(SHEETS.GROUP_SPLITS);
  var lists = ss.getSheetByName(SHEETS.LISTS);
  sh.clearContents();
  sh.clearFormats();

  // ── Header banner (rows 1-2, full-width merged) ───────────
  // These merged cells block setFrozenColumns — see note at top of file.
  sh.getRange('A1:K1').merge();
  sh.getRange('A1')
    .setValue('GROUP SPLITS — Track group expenses with friends, family, and colleagues')
    .setFontSize(13).setFontWeight('bold')
    .setFontColor(COLORS.PRIMARY).setHorizontalAlignment('left');
  sh.setRowHeight(1, 34);

  sh.getRange('A2:K2').merge();
  sh.getRange('A2')
    .setValue('Add one row per person who owes you. Same event = same Date + Description. Update "Reimbursed" when they pay you back.')
    .setFontColor(COLORS.MID_TEXT).setFontStyle('italic').setFontSize(9);

  // ── Column headers (row 3) ────────────────────────────────
  var headers = [
    'Date', 'Paid By', 'Event / Description', 'Category',
    'Total Bill (PKR)', 'Owed By', 'Their Share (PKR)',
    'Reimbursed (PKR)', 'Outstanding (PKR)', 'Month', 'Notes', 'Tx Key'
  ];
  var hdr = sh.getRange(3, 1, 1, headers.length);
  hdr.setValues([headers]);
  styleHeaderRow(hdr);
  sh.setFrozenRows(3);

  [100, 90, 220, 130, 130, 160, 130, 120, 130, 115, 220, 1]
    .forEach(function(w, i) { sh.setColumnWidth(i + 1, w); });
  sh.hideColumns(GS_COL_TX_KEY);  // L — internal link key, invisible to user

  // ── Dropdowns ─────────────────────────────────────────────
  var ROWS = 500;
  setDropdown(sh, 4, 2, ROWS, lists.getRange('A2:A3'));
  setDropdown(sh, 4, 4, ROWS, lists.getRange('C2:C15'));

  // ── Number formats ────────────────────────────────────────
  sh.getRange(4, 1, ROWS, 1).setNumberFormat('MM/DD/YYYY');
  sh.getRange(4, 5, ROWS, 1).setNumberFormat(PKR_FORMAT);
  sh.getRange(4, 7, ROWS, 1).setNumberFormat(PKR_FORMAT);
  sh.getRange(4, 8, ROWS, 1).setNumberFormat(PKR_FORMAT);
  sh.getRange(4, 9, ROWS, 1).setNumberFormat(PKR_FORMAT);

  // ── Sample data ───────────────────────────────────────────
  var today = new Date();
  // I(9,Outstanding) and J(10,Month) are ARRAYFORMULA columns — skip them in setValues.
  var samples = [
    [today, PERSON1, 'Dinner at Salt & Pepper', 'Food', 4800, 'Ali',    1200, 0, 'Karachi trip'],
    [today, PERSON1, 'Dinner at Salt & Pepper', 'Food', 4800, 'Hassan', 1200, 0, ''],
    [today, PERSON1, 'Dinner at Salt & Pepper', 'Food', 4800, 'Bilal',  1200, 0, ''],
    [today, PERSON2, 'Shopping trip',            'Personal Care', 6000, 'Sara', 2000, 0, 'Mall run'],
  ];
  sh.getRange(4, 1, samples.length, 8).setValues(samples.map(function(r) { return r.slice(0, 8); }));
  sh.getRange(4, 11, samples.length, 1).setValues(samples.map(function(r) { return [r[8]]; }));
  sh.getRange(4, 1, samples.length, 1).setNumberFormat('MM/DD/YYYY');
  sh.getRange(4, 5, samples.length, 1).setNumberFormat(PKR_FORMAT);
  sh.getRange(4, 7, samples.length, 2).setNumberFormat(PKR_FORMAT);

  // ── Totals panel (two-column table: M3:N6) ───────────────
  sh.getRange('M3').setValue('TOTALS');
  sh.getRange('N3').setValue('Amount (PKR)');
  styleHeaderRow(sh.getRange(3, 13, 1, 2));
  sh.getRange('M4').setValue('Total Outstanding:');
  sh.getRange('N4').setFormula('=IFERROR(SUM(I4:I1000),0)').setNumberFormat(PKR_FORMAT).setFontWeight('bold').setFontColor(COLORS.DANGER);
  sh.getRange('M5').setValue('Total Reimbursed:');
  sh.getRange('N5').setFormula('=IFERROR(SUM(H4:H1000),0)').setNumberFormat(PKR_FORMAT);
  sh.getRange('M6').setValue('Total Expected:');
  sh.getRange('N6').setFormula('=IFERROR(SUM(G4:G1000),0)').setNumberFormat(PKR_FORMAT);
  sh.setColumnWidth(13, 160);
  sh.setColumnWidth(14, 120);
}
