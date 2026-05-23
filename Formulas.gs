// ============================================================
// HOUSEHOLD FINANCE TRACKER — Formulas.gs
// Writes all ARRAYFORMULA formulas into their designated cells.
// Called from setup(). Can be re-run standalone to restore
// any accidentally overwritten formula column.
//
// ARRAYFORMULA ownership (never write static values here):
//   Transactions: J2 (Amount PKR), M2 (Month)
//   Budget:       E2 (Actual), F2 (Variance), G2 (% Used)
//   Income:       F2 (Month)
//   Savings:      E2 (Progress%), F2 (Remaining)
//   Subscriptions:G2 (Per Person), H2 (Expected Reimb.), I2 (Your Net Cost)
//   Group Splits: I4 (Outstanding), J4 (Month)
// ============================================================

// ── BUDGET E2 FORMULA ─────────────────────────────────────
// LET pre-resolves the named ranges before ARRAYFORMULA
// expands them — fixes the SUMIFS broadcasting bug in
// Google Sheets where all rows return the same value.

var BUDGET_ACTUAL_FORMULA =
  '=LET(' +
  'months,TEXT(A2:A,"MMMM yyyy"),' +
  'cats,B2:B,' +
  'grouped,QUERY(Transactions!C:M,' +
  '"select Col11,Col2,sum(Col8) where Col1=\'Expense\' group by Col11,Col2 label sum(Col8) \'\'",0),' +
  'ARRAYFORMULA(IF(A2:A="","",' +
  'IFNA(VLOOKUP(months&"|"&cats,' +
  '{INDEX(grouped,,1)&"|"&INDEX(grouped,,2),INDEX(grouped,,3)},' +
  '2,FALSE),0))))';

// ── PERSONAL SPENDING NET FORMULA (Dashboard E11) ─────────────
// Two-part calculation:
//   normalPersonal — rows where Shared?=No and Group Split?≠Yes (no TX Key needed)
//   splitAmts/splitKeys — rows where Shared?=No AND Group Split?=Yes (TX Key present)
//     → subtract reimbursements received via Group Splits for those linked rows
// References E2 (the text helper "May 2026" from C2&" "&D2 dropdowns).
// Must remain A1-style — QUERY does not support structured table refs.

var PERSONAL_NET_FORMULA =
  '=IFERROR(LET(' +
  'normalPersonal,IFERROR(INDEX(QUERY(Transactions!A2:Q,' +
  '"SELECT SUM(Col10) WHERE Col3=\'Expense\' AND Col13=\'"&E2&"\' AND Col12=\'No\' AND (Col15<>\'Yes\' OR Col15 IS NULL) LABEL SUM(Col10) \'\'",0),1,1),0),' +
  'splitTxData,IFERROR(QUERY(Transactions!A2:Q,' +
  '"SELECT Col10, Col17 WHERE Col3=\'Expense\' AND Col13=\'"&E2&"\' AND Col12=\'No\' AND Col15=\'Yes\' AND Col17 IS NOT NULL LABEL Col10 \'\', Col17 \'\'",0),{0,""}),' +
  'splitAmts,INDEX(splitTxData,,1),' +
  'splitKeys,INDEX(splitTxData,,2),' +
  'reimb,IFERROR(QUERY(\'Group Splits\'!A2:L,' +
  '"SELECT Col12, SUM(Col8) WHERE Col12 IS NOT NULL GROUP BY Col12 LABEL SUM(Col8) \'\'",0),{"",0}),' +
  'normalPersonal+SUM(splitAmts-IFERROR(VLOOKUP(splitKeys,reimb,2,FALSE),0))),0)';

// ── SHARED SPENDING NET FORMULA (Dashboard B11) ────────────────
// Identical structure to PERSONAL_NET_FORMULA but for Shared?=Yes transactions.
//   normalShared — shared rows where Group Split?≠Yes (no TX Key needed)
//   splitTxData — shared rows where Group Split?=Yes (TX Key present, subtract reimb.)

var SHARED_NET_FORMULA =
  '=IFERROR(LET(' +
  'normalShared,IFERROR(INDEX(QUERY(Transactions!A2:Q,' +
  '"SELECT SUM(Col10) WHERE Col3=\'Expense\' AND Col13=\'"&E2&"\' AND Col12=\'Yes\' AND (Col15<>\'Yes\' OR Col15 IS NULL) LABEL SUM(Col10) \'\'",0),1,1),0),' +
  'splitTxData,IFERROR(QUERY(Transactions!A2:Q,' +
  '"SELECT Col10, Col17 WHERE Col3=\'Expense\' AND Col13=\'"&E2&"\' AND Col12=\'Yes\' AND Col15=\'Yes\' AND Col17 IS NOT NULL LABEL Col10 \'\', Col17 \'\'",0),{0,""}),' +
  'splitAmts,INDEX(splitTxData,,1),' +
  'splitKeys,INDEX(splitTxData,,2),' +
  'reimb,IFERROR(QUERY(\'Group Splits\'!A2:L,' +
  '"SELECT Col12, SUM(Col8) WHERE Col12 IS NOT NULL GROUP BY Col12 LABEL SUM(Col8) \'\'",0),{"",0}),' +
  'normalShared+SUM(splitAmts-IFERROR(VLOOKUP(splitKeys,reimb,2,FALSE),0))),0)';

// ── CATEGORY NET FORMULA BUILDER ──────────────────────────────
// Generates a net spending formula for one expense category.
// Identical two-part logic to PERSONAL/SHARED_NET_FORMULA:
//   1. Normal rows (Group Split?≠Yes) are summed directly — no TX Key needed.
//   2. Group-split rows (Group Split?=Yes, TX Key present) use TX Key to look up
//      and subtract linked reimbursements from Group Splits.
// Param: cat — category name string matching Transactions col D values.

function buildCategoryNetFormula(cat) {
  return '=IFERROR(LET(' +
    'normalAmt,IFERROR(INDEX(QUERY(Transactions!A2:Q,' +
    '"SELECT SUM(Col10) WHERE Col3=\'Expense\' AND Col13=\'"&E2&"\' AND Col4=\'' + cat + '\' AND (Col15<>\'Yes\' OR Col15 IS NULL) LABEL SUM(Col10) \'\'",0),1,1),0),' +
    'splitTxData,IFERROR(QUERY(Transactions!A2:Q,' +
    '"SELECT Col10, Col17 WHERE Col3=\'Expense\' AND Col13=\'"&E2&"\' AND Col4=\'' + cat + '\' AND Col15=\'Yes\' AND Col17 IS NOT NULL LABEL Col10 \'\', Col17 \'\'",0),{0,""}),' +
    'splitAmts,INDEX(splitTxData,,1),' +
    'splitKeys,INDEX(splitTxData,,2),' +
    'reimb,IFERROR(QUERY(\'Group Splits\'!A2:L,' +
    '"SELECT Col12, SUM(Col8) WHERE Col12 IS NOT NULL GROUP BY Col12 LABEL SUM(Col8) \'\'",0),{"",0}),' +
    'normalAmt+SUM(splitAmts-IFERROR(VLOOKUP(splitKeys,reimb,2,FALSE),0))),0)';
}

function addFormulas(ss) {
  // ── Transactions ─────────────────────────────────────────
  var tx = ss.getSheetByName(SHEETS.TRANSACTIONS);
  tx.getRange('J2:J').clearContent();
  tx.getRange('M2:M').clearContent();

  tx.getRange('J2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",H2:H*IF((I2:I="")+(I2:I=0),1,I2:I)))'
  );
  tx.getRange('M2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",TEXT(A2:A,"MMMM yyyy")))'
  );

  // ── Budget ───────────────────────────────────────────────
  var bud = ss.getSheetByName(SHEETS.BUDGET);
  bud.getRange('E2:G').clearContent();

  bud.getRange('E2').setFormula(BUDGET_ACTUAL_FORMULA);
  bud.getRange('F2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",D2:D-E2:E))'
  );
  bud.getRange('G2').setFormula(
    '=ARRAYFORMULA(IF((A2:A="")+(D2:D=0),"",E2:E/D2:D))'
  );

  // ── Income ───────────────────────────────────────────────
  var inc = ss.getSheetByName(SHEETS.INCOME);
  inc.getRange('F2:F').clearContent();
  inc.getRange('F2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",TEXT(A2:A,"MMMM yyyy")))'
  );

  // ── Savings Goals ────────────────────────────────────────
  var sav = ss.getSheetByName(SHEETS.SAVINGS);
  sav.getRange('E2:F').clearContent();
  sav.getRange('E2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",IF(B2:B=0,0,C2:C/B2:B)))'
  );
  sav.getRange('F2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",B2:B-C2:C))'
  );

  // ── Subscriptions ────────────────────────────────────────
  var sub = ss.getSheetByName(SHEETS.SUBSCRIPTIONS);
  sub.getRange('G2:I').clearContent();
  sub.getRange('G2').setFormula(
    '=ARRAYFORMULA(IF(C2:C="","",' +
    'IF((E2:E="")+(E2:E=0),0,C2:C/E2:E)))'
  );
  sub.getRange('H2').setFormula(
    '=ARRAYFORMULA(IF(C2:C="","",' +
    'IF((E2:E="")+(E2:E=0),0,C2:C/E2:E*IF(F2:F="",0,F2:F))))'
  );
  sub.getRange('I2').setFormula(
    '=ARRAYFORMULA(IF(C2:C="","",' +
    'C2:C-IF((E2:E="")+(E2:E=0),0,C2:C/E2:E*IF(F2:F="",0,F2:F))))'
  );

  // ── Group Splits ─────────────────────────────────────────
  var gs = ss.getSheetByName(SHEETS.GROUP_SPLITS);
  gs.getRange('I2:J').clearContent();
  gs.getRange('I2').setFormula(
    '=ARRAYFORMULA(IF(F2:F="","",G2:G-IF(H2:H="",0,H2:H)))'
  );
  gs.getRange('J2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",TEXT(A2:A,"MMMM yyyy")))'
  );
}

// ── STANDALONE BUDGET REPAIR ──────────────────────────────
// Run this directly from Apps Script if Budget shows 0 Actual
// without re-running full setup(). Clears and rewrites all
// ARRAYFORMULA columns for Budget, Transactions, and Savings.

function fixBudgetFormulas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Diagnose what's in Transactions before fixing
  var tx = ss.getSheetByName(SHEETS.TRANSACTIONS);
  var txLast = tx.getLastRow();
  Logger.log('Transactions last row: ' + txLast);
  if (txLast >= 2) {
    var sample = tx.getRange(2, 1, Math.min(3, txLast - 1), 13).getValues();
    sample.forEach(function(r, i) {
      Logger.log('Tx row ' + (i + 2) + ': Date=' + r[0] + ' Type=' + r[2] +
        ' Cat=' + r[3] + ' Amt(H)=' + r[7] + ' AmtPKR(J)=' + r[9] + ' Month(M)=' + r[12]);
    });
  }

  // Clear and rewrite Transactions ARRAYFORMULA cols
  tx.getRange('J2:J').clearContent();
  tx.getRange('M2:M').clearContent();
  tx.getRange('J2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",H2:H*IF((I2:I="")+(I2:I=0),1,I2:I)))'
  );
  tx.getRange('M2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",TEXT(A2:A,"MMMM yyyy")))'
  );

  // Clear and rewrite Budget ARRAYFORMULA cols
  var bud = ss.getSheetByName(SHEETS.BUDGET);
  bud.getRange('E2:G').clearContent();
  bud.getRange('E2').setFormula(BUDGET_ACTUAL_FORMULA);
  bud.getRange('F2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",D2:D-E2:E))'
  );
  bud.getRange('G2').setFormula(
    '=ARRAYFORMULA(IF((A2:A="")+(D2:D=0),"",E2:E/D2:D))'
  );

  // Log what Budget sees now
  SpreadsheetApp.flush();
  var budLast = bud.getLastRow();
  if (budLast >= 2) {
    var budSample = bud.getRange(2, 1, Math.min(3, budLast - 1), 7).getValues();
    budSample.forEach(function(r, i) {
      Logger.log('Budget row ' + (i + 2) + ': Month=' + r[0] + ' Cat=' + r[1] +
        ' Budget=' + r[3] + ' Actual=' + r[4] + ' Variance=' + r[5]);
    });
  }

  ss.toast('Budget formulas repaired. Check View → Logs for diagnostics.', 'Done', 10);
  Logger.log('fixBudgetFormulas complete.');
}