// ============================================================
// HOUSEHOLD FINANCE TRACKER — Tables.gs
// Google Sheets Table creation and structured-reference upgrade.
//
// Two-step process:
//   Step 1: createTables(ss)            — create formal tables via Sheets API
//   Step 2: upgradeToTableFormulas(ss)  — switch cross-sheet SUMIFS to
//                                         structured references
//
// Combined entry point:
//   setupWithTables()                   — runs setup() then both steps above
//
// Prerequisite for Step 1:
//   Apps Script editor → ⊕ Services → Google Sheets API → Add
//   Then add to appsscript.json (already done if you re-ran clasp push).
//
// Manual fallback for Step 1 (if API fails):
//   On each sheet: Insert → Table → select A1:last-column / last-row
//   Name the table in the Table Properties panel (right sidebar):
//     Transactions   → TransactionsTable
//     Budget         → BudgetTable
//     Income         → IncomeTable
//     Savings Goals  → SavingsGoalsTable
//     Subscriptions  → SubscriptionsTable
//   Then run upgradeToTableFormulas() alone.
//
// WHAT STAYS A1-STYLE (do not convert):
//   • All QUERY formulas — QUERY does not support structured references
//   • ARRAYFORMULAs within their own sheet's data range — more reliable
//   • Group Splits — rows 1-2 are merged banners; leave as named range
//   • Reimbursements — QUERY-driven; no table needed
//   • Summary blocks (within-sheet short refs like J:J, M:M) — fine as-is
// ============================================================

// ── TABLE NAME CONSTANTS ──────────────────────────────────────

var TABLE_NAMES = {
  TRANSACTIONS:  'TransactionsTable',
  BUDGET:        'BudgetTable',
  INCOME:        'IncomeTable',
  SAVINGS:       'SavingsGoalsTable',
  SUBSCRIPTIONS: 'SubscriptionsTable',
};

// ── TABLE DELETION ────────────────────────────────────────────
// Removes all existing table objects from the spreadsheet.
// Called automatically by setup() so re-running setup is always safe.
// No-ops silently if Sheets advanced service is not enabled.

function deleteTables(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  if (typeof Sheets === 'undefined') return;

  var ssId = ss.getId();
  try {
    var resp = Sheets.Spreadsheets.get(ssId, { fields: 'sheets(tables(tableId))' });
    var requests = [];
    (resp.sheets || []).forEach(function(sh) {
      (sh.tables || []).forEach(function(tbl) {
        requests.push({ deleteTable: { tableId: tbl.tableId } });
      });
    });
    if (requests.length > 0) {
      Sheets.Spreadsheets.batchUpdate({ requests: requests }, ssId);
      Logger.log('deleteTables: removed ' + requests.length + ' table(s).');
    }
  } catch (err) {
    Logger.log('deleteTables (non-fatal): ' + err.message);
  }
}

// ── TABLE CREATION ────────────────────────────────────────────
// Returns true on success, false if the Sheets service is unavailable.

function createTables(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();

  if (typeof Sheets === 'undefined') {
    Logger.log(
      'createTables: Sheets advanced service not enabled.\n' +
      'Go to Apps Script editor → ⊕ (Add service) → Google Sheets API → Add.\n' +
      'Then re-run createTables(), or follow the manual steps in Tables.gs header.'
    );
    ss.toast(
      'Enable "Google Sheets API" in Apps Script → Services, then re-run.',
      'Tables: Action Required', 20
    );
    return false;
  }

  var ssId = ss.getId();

  var defs = [
    { sheetName: SHEETS.TRANSACTIONS,  name: TABLE_NAMES.TRANSACTIONS,  cols: 17 },
    { sheetName: SHEETS.BUDGET,        name: TABLE_NAMES.BUDGET,        cols: 8  },
    { sheetName: SHEETS.INCOME,        name: TABLE_NAMES.INCOME,        cols: 7  },
    { sheetName: SHEETS.SAVINGS,       name: TABLE_NAMES.SAVINGS,       cols: 8  },
    { sheetName: SHEETS.SUBSCRIPTIONS, name: TABLE_NAMES.SUBSCRIPTIONS, cols: 13 },
  ];

  // Remove sheet filters before table creation.
  // The Sheets API rejects addTable when a filter overlaps the target range.
  // setup() → applyProfessionalFormatting() → addFilters() runs before this,
  // so filters always exist at this point. Tables provide their own filter UI.
  defs.forEach(function(def) {
    var sh = ss.getSheetByName(def.sheetName);
    if (!sh) return;
    var filter = sh.getFilter();
    if (filter) filter.remove();
  });

  var requests = defs.reduce(function(acc, def) {
    var sh = ss.getSheetByName(def.sheetName);
    if (!sh) return acc;
    var endRow = Math.max(sh.getLastRow() + 500, 1000);
    acc.push({
      addTable: {
        table: {
          name: def.name,
          range: {
            sheetId:          sh.getSheetId(),
            startRowIndex:    0,
            endRowIndex:      endRow,
            startColumnIndex: 0,
            endColumnIndex:   def.cols,
          },
        }
      }
    });
    return acc;
  }, []);

  try {
    Sheets.Spreadsheets.batchUpdate({ requests: requests }, ssId);
    Logger.log('createTables: created ' + requests.length + ' tables.');
    ss.toast(
      'Tables created: ' + defs.map(function(d) { return d.name; }).join(', '),
      'Done', 12
    );
    return true;
  } catch (err) {
    Logger.log('createTables error: ' + err.message +
      '\nFall back to manual creation (see Tables.gs header for steps).');
    ss.toast(
      'Table API failed — create tables manually then run upgradeToTableFormulas().',
      'See Logs', 15
    );
    return false;
  }
}

// ── FORMULA UPGRADE ───────────────────────────────────────────
// Replaces cross-sheet SUMIFS/SUMIF with structured table references.
// Safe to re-run. Does NOT touch QUERY formulas or within-sheet ARRAYFORMULAs.
// Call this after createTables() succeeds, OR after manual table creation.

function upgradeToTableFormulas(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  var tx  = TABLE_NAMES.TRANSACTIONS;
  var inc = TABLE_NAMES.INCOME;
  var sub = TABLE_NAMES.SUBSCRIPTIONS;

  // ── Budget E2 ────────────────────────────────────────────────
  // Uses BUDGET_ACTUAL_FORMULA (LET+QUERY+VLOOKUP defined in Formulas.gs).
  // QUERY does not support structured table refs, so A1-style is correct here.
  var bud = ss.getSheetByName(SHEETS.BUDGET);
  bud.getRange('E2:G').clearContent();
  bud.getRange('E2').setFormula(BUDGET_ACTUAL_FORMULA);
  bud.getRange('F2').setFormula('=ARRAYFORMULA(IF(A2:A="","",D2:D-E2:E))');
  bud.getRange('G2').setFormula('=ARRAYFORMULA(IF((A2:A="")+(D2:D=0),"",E2:E/D2:D))');

  // ── Dashboard card values ─────────────────────────────────────
  // Card row 1 valueRow=6 (renderCardRow(cards1, 4, 5, 6, 7, 8))
  // Card row 2 valueRow=11 (renderCardRow(cards2, 9, 10, 11, 12, 13))
  var db = ss.getSheetByName(SHEETS.DASHBOARD);

  // B6 — Monthly Income
  db.getRange('B6').setFormula(
    '=SUMIFS(' + inc + '[Amount (PKR)],' + inc + '[Month],E2)'
  );

  // E6 — Total Spending (= B11 + E11, net of reimbursements)
  db.getRange('E6').setFormula('=B11+E11');

  // H6 — Net Cash Flow (=B6-E6): no table ref needed, keep as-is
  // K6 — Savings Rate (=IFERROR((B6-E6)/B6,0)): no table ref needed

  // B11 — Shared Spending net of reimbursements
  // QUERY-based; must stay A1-style even in table mode
  db.getRange('B11').setFormula(SHARED_NET_FORMULA);

  // E11 — Personal Spending (net of Group Split reimbursements)
  // QUERY-based; must stay A1-style even in table mode (QUERY ignores structured refs)
  db.getRange('E11').setFormula(PERSONAL_NET_FORMULA);

  // H11 — Group Outstanding (Group Splits only; no Transactions structured refs needed)
  db.getRange('H11').setFormula("=IFERROR(SUM('Group Splits'!I:I),0)");

  // K11 — Subscriptions Net Cost
  db.getRange('K11').setFormula(
    '=IFERROR(SUMIF(' + sub + '[Your Net Cost (PKR)],">0",' + sub + '[Your Net Cost (PKR)]),0)'
  );

  // ── Dashboard category breakdown (rows 17-26, col J = col 10) ─
  var cats = [
    'Housing', 'Food', 'Transportation', 'Health',
    'Entertainment', 'Subscriptions', 'Personal Care',
    'Travel', 'Gifts', 'Other'
  ];
  cats.forEach(function(cat, i) {
    db.getRange(17 + i, 10).setFormula(buildCategoryNetFormula(cat));
  });

  Logger.log('upgradeToTableFormulas: structured references applied to Budget and Dashboard.');
  ss.toast('Formulas upgraded to table structured references.', 'Done', 8);
}

// ── COMBINED ENTRY POINT ──────────────────────────────────────
// Run this instead of setup() to get full table + formula upgrade in one shot.
// setup() is called internally — do not run setup() first.

function setupWithTables() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Running full table setup…', 'Please wait', -1);

  setup();

  var tablesOk = createTables(ss);
  if (tablesOk) {
    upgradeToTableFormulas(ss);
    ss.toast('Full table setup complete!', '✅ Done', 12);
  } else {
    ss.toast(
      'Base setup done. Create tables manually (Insert → Table on each sheet), ' +
      'then run upgradeToTableFormulas() from Apps Script.',
      'Tables Pending', 25
    );
  }
}

// ── DOWNGRADE (ROLLBACK) ──────────────────────────────────────
// Restores all formulas to A1-style if structured refs cause issues.
// Does not remove the table objects — only reverts the formulas.

function revertToA1Formulas(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();

  // Budget E2: revert to BUDGET_ACTUAL_FORMULA (A1-style LET+QUERY, defined in Formulas.gs)
  var bud = ss.getSheetByName(SHEETS.BUDGET);
  bud.getRange('E2:G').clearContent();
  bud.getRange('E2').setFormula(BUDGET_ACTUAL_FORMULA);
  bud.getRange('F2').setFormula('=ARRAYFORMULA(IF(A2:A="","",D2:D-E2:E))');
  bud.getRange('G2').setFormula('=ARRAYFORMULA(IF((A2:A="")+(D2:D=0),"",E2:E/D2:D))');

  // Dashboard cards
  var db = ss.getSheetByName(SHEETS.DASHBOARD);
  db.getRange('B6').setFormula('=SUMIFS(Income!E:E,Income!F:F,E2)');
  db.getRange('E6').setFormula('=B11+E11');
  db.getRange('B11').setFormula(SHARED_NET_FORMULA);
  db.getRange('E11').setFormula(PERSONAL_NET_FORMULA);
  db.getRange('H11').setFormula("=IFERROR(SUM('Group Splits'!I:I),0)");
  db.getRange('K11').setFormula('=IFERROR(SUMIF(Subscriptions!I:I,">0",Subscriptions!I:I),0)');

  // Dashboard category breakdown
  var cats = [
    'Housing', 'Food', 'Transportation', 'Health',
    'Entertainment', 'Subscriptions', 'Personal Care',
    'Travel', 'Gifts', 'Other'
  ];
  cats.forEach(function(cat, i) {
    db.getRange(17 + i, 10).setFormula(buildCategoryNetFormula(cat));
  });

  Logger.log('revertToA1Formulas: all formulas restored to A1-style.');
  ss.toast('Formulas reverted to A1-style.', 'Done', 8);
}
