// ============================================================
// HOUSEHOLD FINANCE TRACKER — Setup.gs
// Main entry point and sheet-creation functions.
// setup() is fully idempotent — safe to re-run, but wipes data.
// Run setupForm() separately (needs FormApp OAuth scope).
// Run setupRecurring() optionally to enable recurring transactions.
// ============================================================

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Setting up Household Finance Tracker…', 'Please wait', -1);

  deleteTables(ss);  // remove any existing table objects before rebuilding sheets
  setupSheets(ss);
  setupLists(ss);
  setupTransactions(ss);
  setupBudget(ss);
  setupIncome(ss);
  setupSavingsGoals(ss);
  setupSubscriptions(ss);
  setupReimbursements(ss);
  setupGroupSplits(ss);
  setupDashboard(ss);
  setupInstructions(ss);
  addFormulas(ss);
  applyFormatting(ss);
  applyProfessionalFormatting(ss);

  ss.setActiveSheet(ss.getSheetByName(SHEETS.DASHBOARD));
  ss.toast('✅ Setup complete! Open Dashboard to get started.', 'Done', 10);
  Logger.log('Setup done. Run setupForm() separately to create the Google Form.');
}

// ── SHEET ORDER ───────────────────────────────────────────────

function setupSheets(ss) {
  var order = [
    SHEETS.DASHBOARD, SHEETS.TRANSACTIONS, SHEETS.BUDGET,
    SHEETS.INCOME, SHEETS.SAVINGS, SHEETS.SUBSCRIPTIONS,
    SHEETS.REIMBURSEMENTS, SHEETS.GROUP_SPLITS, SHEETS.LISTS, SHEETS.INSTRUCTIONS
  ];

  order.forEach(function(name) {
    if (!ss.getSheetByName(name)) ss.insertSheet(name);
  });

  var def = ss.getSheetByName('Sheet1');
  if (def && def.getLastRow() === 0) ss.deleteSheet(def);

  order.forEach(function(name, i) {
    ss.setActiveSheet(ss.getSheetByName(name));
    ss.moveActiveSheet(i + 1);
  });
}

// ── LISTS SHEET ───────────────────────────────────────────────

function setupLists(ss) {
  var sh = ss.getSheetByName(SHEETS.LISTS);
  sh.clear();
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations();

  var people  = [PERSON1, PERSON2];
  var types   = ['Expense', 'Income', 'Transfer'];
  var cats    = [
    'Housing', 'Food', 'Transportation', 'Health', 'Entertainment',
    'Personal Care', 'Education', 'Savings', 'Travel', 'Gifts',
    'Subscriptions', 'Children', 'Pets', 'Other'
  ];
  var subcats = [
    'Rent/Mortgage', 'Utilities', 'Internet', 'Home Insurance',
    'Home Maintenance', 'Furniture', 'Decor',
    'Groceries', 'Dining Out', 'Coffee/Tea', 'Food Delivery',
    'Gas', 'Car Insurance', 'Car Maintenance', 'Parking',
    'Public Transit', 'Rideshare',
    'Medical', 'Dental', 'Vision', 'Pharmacy', 'Health Insurance', 'Gym/Fitness',
    'Streaming', 'Movies/Events', 'Hobbies', 'Games', 'Sports',
    'Hair/Beauty', 'Clothing', 'Accessories', 'Personal Items',
    'Online Courses', 'Books', 'School Supplies',
    'Emergency Fund', 'Retirement', 'Investments', 'Other Savings',
    'Flights', 'Hotels', 'Vacation Activities', 'Travel Insurance',
    'Birthday Gifts', 'Holiday Gifts', 'Donations/Charity',
    'Software', 'Other Subscription',
    'Childcare', 'Child Education', 'Child Clothing', 'Child Activities',
    'Pet Food', 'Veterinary', 'Pet Grooming', 'Pet Supplies',
    'Miscellaneous',
  ];
  var paidBy  = people.concat(['Joint Account', 'Credit Card', 'Cash']);
  var statuses= ['Active', 'Cancelled', 'Paused', 'Pending', 'Paid', 'Partial', 'Overdue'];
  var yesNo   = ['No', 'Yes'];
  var sources = [
    'Salary', 'Bonus', 'Freelance', 'Side Hustle',
    'Investment Returns', 'Rental Income', 'Tax Refund', 'Gift', 'Other'
  ];
  var months  = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
  ];

  var cols = [people, types, cats, subcats, paidBy, statuses, yesNo, sources, months, CURRENCIES];
  var hdrs = ['People','Types','Categories','Sub-Categories','Paid By','Status','Yes/No','Income Sources','Months','Currencies'];

  sh.getRange(1, 1, 1, hdrs.length).setValues([hdrs]);
  styleHeaderRow(sh.getRange(1, 1, 1, hdrs.length));

  cols.forEach(function(col, i) {
    if (col.length > 0) {
      sh.getRange(2, i + 1, col.length, 1).setValues(col.map(function(v) { return [v]; }));
    }
  });

  sh.hideSheet();
}

// ── TRANSACTIONS SHEET ────────────────────────────────────────

function setupTransactions(ss) {
  var sh    = ss.getSheetByName(SHEETS.TRANSACTIONS);
  var lists = ss.getSheetByName(SHEETS.LISTS);
  sh.clear();
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations();

  var headers = [
    'Date',         // A
    'Person',       // B
    'Type',         // C
    'Category',     // D
    'Subcategory',  // E
    'Description',  // F
    'Currency',     // G
    'Orig. Amount', // H
    'Exch. Rate',   // I
    'Amount (PKR)', // J  ← FORMULA
    'Paid By',      // K
    'Shared?',      // L  (No by default)
    'Month',        // M  ← FORMULA
    'Notes',        // N
    'Group Split?', // O  (Yes → auto-creates Group Split rows)
    '# People',    // P  (how many rows to create)
    'Tx Key',       // Q  ← hidden internal link key
  ];

  var hdr = sh.getRange(1, 1, 1, headers.length);
  hdr.setValues([headers]);
  styleHeaderRow(hdr);
  sh.setFrozenRows(1);
  sh.setFrozenColumns(1);

  var widths = [100, 85, 90, 120, 140, 200, 80, 110, 90, 110, 100, 70, 115, 200, 90, 75];
  widths.forEach(function(w, i) { sh.setColumnWidth(i + 1, w); });
  sh.setColumnWidth(TX_COL_TX_KEY, 1);  // collapse before hiding (avoids phantom width on unhide)
  sh.hideColumns(TX_COL_TX_KEY);         // Q — internal key, invisible to user

  var ROWS = 1000;

  setDropdown(sh, 2, 2,  ROWS, lists.getRange('A2:A3'));
  setDropdown(sh, 2, 3,  ROWS, lists.getRange('B2:B4'));
  setDropdown(sh, 2, 4,  ROWS, lists.getRange('C2:C15'));
  setDropdown(sh, 2, 7,  ROWS, lists.getRange('J2:J' + (CURRENCIES.length + 1)));
  setDropdown(sh, 2, 11, ROWS, lists.getRange('E2:E6'));
  setDropdown(sh, 2, 12, ROWS, lists.getRange('G2:G3'));           // Shared?
  setDropdown(sh, 2, TX_COL_GROUP_SPLIT, ROWS, lists.getRange('G2:G3'));  // Group Split?

  sh.getRange(2, 1,  ROWS, 1).setNumberFormat('MM/DD/YYYY');
  sh.getRange(2, 8,  ROWS, 1).setNumberFormat(PKR_FORMAT);
  sh.getRange(2, 9,  ROWS, 1).setNumberFormat(RATE_FORMAT);
  sh.getRange(2, 10, ROWS, 1).setNumberFormat(PKR_FORMAT);
  // # People: require a whole number >= 1.
  // Combined with Group Split? = Yes, this is the trigger for auto-row creation.
  sh.getRange(2, TX_COL_NUM_PEOPLE, ROWS, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireNumberGreaterThanOrEqualTo(1)
      .setAllowInvalid(false)
      .setHelpText('Whole number ≥ 1: how many Group Split rows to create')
      .build()
  );

  // Sample row — PKR domestic transaction
  var today = new Date();
  sh.getRange(2, 1, 1, 9).setValues([[
    today, PERSON1, 'Expense', 'Food', 'Groceries', 'Sample grocery run', 'PKR', 1500, 1
  ]]);
  sh.getRange(2, 11, 1, 2).setValues([[PERSON1, 'No']]);
  sh.getRange(2, 1).setNumberFormat('MM/DD/YYYY');
  sh.getRange(2, 8).setNumberFormat(PKR_FORMAT);
  sh.getRange(2, 9).setNumberFormat(RATE_FORMAT);

  // Second sample — USD international transaction
  sh.getRange(3, 1, 1, 9).setValues([[
    today, PERSON1, 'Expense', 'Travel', 'Flights', 'Sample international purchase', 'USD', 50, 280
  ]]);
  sh.getRange(3, 11, 1, 2).setValues([[PERSON1, 'No']]);
  sh.getRange(3, 1).setNumberFormat('MM/DD/YYYY');
  sh.getRange(3, 8).setNumberFormat(PKR_FORMAT);
  sh.getRange(3, 9).setNumberFormat(RATE_FORMAT);
}

// ── BUDGET SHEET ──────────────────────────────────────────────

function setupBudget(ss) {
  var sh    = ss.getSheetByName(SHEETS.BUDGET);
  var lists = ss.getSheetByName(SHEETS.LISTS);
  sh.clear();
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations();

  var headers = ['Month', 'Category', 'Type', 'Budget (PKR)', 'Actual (PKR)', 'Variance', '% Used', 'Notes'];
  var hdr = sh.getRange(1, 1, 1, headers.length);
  hdr.setValues([headers]);
  styleHeaderRow(hdr);
  sh.setFrozenRows(1);

  [120, 130, 100, 120, 120, 110, 80, 200].forEach(function(w, i) { sh.setColumnWidth(i + 1, w); });

  var tz = Session.getScriptTimeZone();
  var mo = Utilities.formatDate(new Date(), tz, 'MMMM yyyy');

  // E(Actual), F(Variance), G(%Used) are ARRAYFORMULA columns — never write '' into them.
  // Writing empty strings blocks ARRAYFORMULA expansion to those rows.
  var rows = [
    [mo, 'Housing',       'Shared',    80000, ''],
    [mo, 'Food',          'Shared',    30000, ''],
    [mo, 'Transportation','Shared',    15000, ''],
    [mo, 'Health',        'Shared',    10000, ''],
    [mo, 'Entertainment', 'Shared',     8000, ''],
    [mo, 'Subscriptions', 'Shared',     5000, ''],
    [mo, 'Personal Care', 'Personal',   5000, PERSON1],
    [mo, 'Personal Care', 'Personal',   5000, PERSON2],
    [mo, 'Travel',        'Shared',    15000, 'Monthly accrual'],
    [mo, 'Gifts',         'Shared',     5000, ''],
    [mo, 'Other',         'Shared',     8000, ''],
  ];

  sh.getRange(2, 1, rows.length, 4).setValues(rows.map(function(r) { return [r[0], r[1], r[2], r[3]]; }));
  sh.getRange(2, 1, rows.length, 1).setNumberFormat('@STRING@');
  sh.getRange(2, 8, rows.length, 1).setValues(rows.map(function(r) { return [r[4]]; }));

  var typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Shared', 'Personal'], true).build();
  sh.getRange(2, 3, 200, 1).setDataValidation(typeRule);
  setDropdown(sh, 2, 2, 200, lists.getRange('C2:C15'));

  [4, 5, 6].forEach(function(c) { sh.getRange(2, c, 200, 1).setNumberFormat(PKR_FORMAT); });
  sh.getRange(2, 7, 200, 1).setNumberFormat(PCT_FORMAT);
}

// ── INCOME SHEET ──────────────────────────────────────────────

function setupIncome(ss) {
  var sh    = ss.getSheetByName(SHEETS.INCOME);
  var lists = ss.getSheetByName(SHEETS.LISTS);
  sh.clear();
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations();

  var headers = ['Date', 'Person', 'Source', 'Description', 'Amount (PKR)', 'Month', 'Notes'];
  var hdr = sh.getRange(1, 1, 1, headers.length);
  hdr.setValues([headers]);
  styleHeaderRow(hdr);
  sh.setFrozenRows(1);

  [100, 90, 130, 200, 120, 115, 200].forEach(function(w, i) { sh.setColumnWidth(i + 1, w); });

  setDropdown(sh, 2, 2, 500, lists.getRange('A2:A3'));
  setDropdown(sh, 2, 3, 500, lists.getRange('H2:H10'));

  sh.getRange(2, 1, 500, 1).setNumberFormat('MM/DD/YYYY');
  sh.getRange(2, 5, 500, 1).setNumberFormat(PKR_FORMAT);

  var today = new Date();
  sh.getRange(2, 1, 2, 5).setValues([
    [today, PERSON1, 'Salary', 'Monthly salary', 150000],
    [today, PERSON2, 'Salary', 'Monthly salary', 150000],
  ]);
  sh.getRange(2, 1, 2, 1).setNumberFormat('MM/DD/YYYY');
}

// ── SAVINGS GOALS SHEET ───────────────────────────────────────

function setupSavingsGoals(ss) {
  var sh = ss.getSheetByName(SHEETS.SAVINGS);
  sh.clear();
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations();

  var headers = [
    'Goal Name', 'Target (PKR)', 'Current (PKR)', 'Monthly Contribution',
    'Progress %', 'Remaining', 'Target Date', 'Notes'
  ];
  var hdr = sh.getRange(1, 1, 1, headers.length);
  hdr.setValues([headers]);
  styleHeaderRow(hdr);
  sh.setFrozenRows(1);

  [160, 120, 120, 160, 100, 120, 115, 200].forEach(function(w, i) { sh.setColumnWidth(i + 1, w); });

  // E(Progress%), F(Remaining) are ARRAYFORMULA columns — skip them in setValues.
  var goals = [
    ['Emergency Fund',     600000,  150000,  50000, new Date(new Date().getFullYear(), 11, 31), '6 months expenses'],
    ['Vacation',           200000,   50000,  20000, new Date(new Date().getFullYear(), 7, 1),   'Summer trip'],
    ['New Car',           1500000,  100000,  40000, new Date(new Date().getFullYear() + 2, 0, 1), ''],
    ['Home Down Payment', 5000000,  500000, 100000, new Date(new Date().getFullYear() + 4, 0, 1), ''],
  ];
  sh.getRange(2, 1, goals.length, 4).setValues(goals.map(function(r) { return [r[0], r[1], r[2], r[3]]; }));
  sh.getRange(2, 7, goals.length, 2).setValues(goals.map(function(r) { return [r[4], r[5]]; }));

  [2, 3, 4, 6].forEach(function(c) { sh.getRange(2, c, 20, 1).setNumberFormat(PKR_FORMAT); });
  sh.getRange(2, 5, 20, 1).setNumberFormat(PCT_FORMAT);
  sh.getRange(2, 7, 20, 1).setNumberFormat('MM/DD/YYYY');
}

// ── SUBSCRIPTIONS SHEET ───────────────────────────────────────
//
// Column map (13 data cols + automation col):
//   A  1  Subscription
//   B  2  Payer
//   C  3  Monthly Cost (PKR)
//   D  4  Split With         — comma-separated names: "Ali, Hassan, Bilal"
//   E  5  Total Slots        — plan capacity (e.g. 5 for a 5-seat plan)
//   F  6  # Friends          — how many friends owe you (e.g. 3)
//   G  7  Per Person (PKR)   ← ARRAYFORMULA: C / E
//   H  8  Expected Reimb.    ← ARRAYFORMULA: (C/E) × F
//   I  9  Your Net Cost      ← ARRAYFORMULA: C − H
//   J 10  Renewal Date
//   K 11  Status
//   L 12  Notes
//   M 13  Split Last Posted  — written by processSubscriptionSplits()

function setupSubscriptions(ss) {
  var sh    = ss.getSheetByName(SHEETS.SUBSCRIPTIONS);
  var lists = ss.getSheetByName(SHEETS.LISTS);
  sh.clear();
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations();

  var headers = [
    'Subscription', 'Payer', 'Monthly Cost (PKR)', 'Split With',
    'Total Slots', '# Friends', 'Per Person (PKR)', 'Expected Reimb. (PKR)',
    'Your Net Cost (PKR)', 'Renewal Date', 'Status', 'Notes', 'Split Last Posted'
  ];
  var hdr = sh.getRange(1, 1, 1, headers.length);
  hdr.setValues([headers]);
  styleHeaderRow(hdr);
  sh.setFrozenRows(1);

  [160, 90, 150, 190, 90, 80, 120, 150, 130, 115, 90, 200, 120]
    .forEach(function(w, i) { sh.setColumnWidth(i + 1, w); });

  setDropdown(sh, 2, 2,  100, lists.getRange('A2:A3'));  // Payer
  setDropdown(sh, 2, 11, 100, lists.getRange('F2:F4'));  // Status

  var mo = new Date();
  // G(7), H(8), I(9) are ARRAYFORMULA columns — skip them; write A-F and J-M separately.
  var subs = [
    ['Netflix',        PERSON1,  1000, 'Ali, Hassan, Bilal', 5, 3, new Date(mo.getFullYear(), mo.getMonth(), 15), 'Active', 'Rs 200/slot; you use slots 4-5', ''],
    ['Spotify Family', PERSON1,  1400, 'Sara',               6, 1, new Date(mo.getFullYear(), mo.getMonth(), 1),  'Active', '', ''],
    ['iCloud 200GB',   PERSON1,   840, '',                   1, 0, new Date(mo.getFullYear(), mo.getMonth(), 5),  'Active', 'Personal', ''],
    ['Gym Membership', PERSON2, 12000, '',                   1, 0, new Date(mo.getFullYear(), mo.getMonth(), 1),  'Active', PERSON2 + ' personal', ''],
  ];
  sh.getRange(2, 1, subs.length, 6).setValues(subs.map(function(r) { return [r[0],r[1],r[2],r[3],r[4],r[5]]; }));
  sh.getRange(2, 10, subs.length, 4).setValues(subs.map(function(r) { return [r[6],r[7],r[8],r[9]]; }));

  sh.getRange(2, 3, 100, 1).setNumberFormat(PKR_FORMAT);   // C Monthly Cost
  sh.getRange(2, 7, 100, 3).setNumberFormat(PKR_FORMAT);   // G H I formula cols
  sh.getRange(2, 10, 100, 1).setNumberFormat('MM/DD/YYYY');
  sh.getRange(2, 13, 100, 1).setNumberFormat('MM/DD/YYYY');
}

// ── INSTRUCTIONS SHEET ────────────────────────────────────────

function setupInstructions(ss) {
  var sh = ss.getSheetByName(SHEETS.INSTRUCTIONS);
  sh.clear();
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations();
  sh.setColumnWidth(1, 30);
  sh.setColumnWidth(2, 800);

  var lines = [
    ['TITLE',   'HOUSEHOLD FINANCE TRACKER — USER GUIDE (PKR)'],
    ['', ''],
    ['SECTION', 'GETTING STARTED'],
    ['', '1. Update your names: Apps Script editor → change PERSON1/PERSON2 at top of Config.gs → re-run setup().'],
    ['', '2. Enter income in the Income tab.'],
    ['', '3. Enter transactions in the Transactions tab as you spend.'],
    ['', '4. Set monthly budgets in the Budget tab.'],
    ['', '5. Review the Dashboard each month.'],
    ['', '6. Run setupForm() from Apps Script to create a Google Form for mobile entry.'],
    ['', ''],
    ['SECTION', 'CURRENCY HANDLING'],
    ['', 'All amounts are stored and displayed in PKR (Pakistani Rupee, Rs).'],
    ['', 'For domestic PKR transactions: set Currency = PKR, Exch. Rate = 1 (or leave blank).'],
    ['', 'For international transactions (e.g. USD online purchase):'],
    ['', '  • Currency = USD (or whatever you paid in)'],
    ['', '  • Orig. Amount = the amount in that currency (e.g. 50 for $50)'],
    ['', '  • Exch. Rate = PKR rate at time of transaction (e.g. 280 means 1 USD = Rs 280)'],
    ['', '  • Amount (PKR) auto-calculates: 50 × 280 = Rs 14,000'],
    ['', 'All budgets, dashboard metrics, and reimbursements always work in PKR.'],
    ['', ''],
    ['SECTION', 'TRANSACTIONS TAB — COLUMN GUIDE'],
    ['', 'Date          — When the transaction occurred.'],
    ['', 'Person        — Who made the purchase.'],
    ['', 'Type          — Expense, Income, or Transfer.'],
    ['', 'Category      — High-level grouping (Housing, Food, etc.).'],
    ['', 'Subcategory   — Optional finer detail.'],
    ['', 'Description   — Brief description.'],
    ['', 'Currency      — Dropdown. PKR for local. USD/EUR/etc. for international.'],
    ['', 'Orig. Amount  — Amount in the transaction currency.'],
    ['', 'Exch. Rate    — PKR exchange rate at time of purchase. Leave blank (or 1) for PKR.'],
    ['', 'Amount (PKR)  — AUTO: Orig. Amount × Exch. Rate.'],
    ['', 'Paid By       — Who physically paid.'],
    ['', 'Shared?       — Yes if joint/household expense.'],
    ['', 'Month         — AUTO: from Date column.'],
    ['', 'Notes         — Optional free-text notes.'],
    ['', ''],
    ['SECTION', 'BUDGET TAB'],
    ['', '• Add a row per category per month. Format month as "May 2025".'],
    ['', '• Actual column auto-fills from Transactions (in PKR).'],
    ['', '• Variance = Budget − Actual. Positive = under budget (good).'],
    ['', '• % Used is color-coded: green <80%, yellow 80–100%, red >100%.'],
    ['', ''],
    ['SECTION', 'INCOME TAB'],
    ['', '• Log all income here in PKR. Month column fills automatically from Date.'],
    ['', ''],
    ['SECTION', 'SAVINGS GOALS TAB'],
    ['', '• Update "Current" column periodically. Progress % and Remaining auto-calculate.'],
    ['', ''],
    ['SECTION', 'SUBSCRIPTIONS TAB'],
    ['', '• All amounts in PKR. If a subscription is billed in foreign currency, convert first.'],
    ['', '• Set "Split With" to comma-separated friend names sharing the subscription.'],
    ['', '• Set "Total Slots" to the plan capacity (e.g. 5 for a 5-seat plan).'],
    ['', '• Set "# Friends" to how many friends owe you. Per Person and Net Cost auto-calculate.'],
    ['', '• On the billing day, daily automation auto-creates Group Split rows for each friend.'],
    ['', ''],
    ['SECTION', 'GROUP SPLITS TAB — Split a bill with multiple people'],
    ['', 'Use this tab any time you (or your spouse) pay a bill on behalf of a group.'],
    ['', 'Examples: restaurant dinner with friends, group gift, shared taxi, holiday trip.'],
    ['', ''],
    ['', 'How to enter a group expense:'],
    ['', '  1. Add one row per person who owes you (not one row for the whole bill).'],
    ['', '  2. Fill in: Date, Paid By, Event Description, Category, Total Bill (for reference).'],
    ['', '  3. Owed By = their name (free text — can be anyone, not just household members).'],
    ['', '  4. Their Share = the PKR amount they owe you.'],
    ['', '  5. Outstanding auto-calculates as Their Share − Reimbursed.'],
    ['', ''],
    ['', 'Example — Rs 4,800 dinner with 3 friends, each owes Rs 1,200:'],
    ['', '  Row 1: [Date, Ammar, "Dinner at XYZ", Food, 4800, Ali,    1200, 0] → Outstanding: Rs 1,200'],
    ['', '  Row 2: [Date, Ammar, "Dinner at XYZ", Food, 4800, Hassan, 1200, 0] → Outstanding: Rs 1,200'],
    ['', '  Row 3: [Date, Ammar, "Dinner at XYZ", Food, 4800, Bilal,  1200, 0] → Outstanding: Rs 1,200'],
    ['', ''],
    ['', 'When a friend pays you back: enter the amount in the "Reimbursed" column for their row.'],
    ['', 'The Outstanding column and Reimbursements tab update automatically.'],
    ['', 'Conditional formatting: green = fully paid, yellow = partial, red = nothing received.'],
    ['', ''],
    ['SECTION', 'REIMBURSEMENTS TAB'],
    ['', '• Live view of all outstanding group splits. All amounts in PKR.'],
    ['', '• Pulls from Group Splits where Outstanding > 0.'],
    ['', '• To mark settled: enter the amount in the Reimbursed column of Group Splits.'],
    ['', ''],
    ['SECTION', 'DASHBOARD TAB'],
    ['', '• Use the Month dropdown (C2) and Year dropdown (D2) to switch reporting months. Cell E2 auto-shows "May 2026" and is used by all formulas.'],
    ['', '• All metrics update automatically. All figures in PKR.'],
    ['', ''],
    ['SECTION', 'GOOGLE FORM (Mobile Entry)'],
    ['', '• Run setupForm() from the Apps Script editor.'],
    ['', '• The form includes Currency, Original Amount, and Exchange Rate fields.'],
    ['', '• Submissions auto-write to Transactions with PKR amount calculated.'],
    ['', '• Form URL printed to Apps Script Logs after running.'],
    ['', ''],
    ['SECTION', 'MONTHLY WORKFLOW'],
    ['', '1. Enter transactions throughout the month (form or directly).'],
    ['', '2. Add group expenses in the Group Splits tab as they happen.'],
    ['', '3. Start of next month: duplicate Budget rows, update month value.'],
    ['', '4. Review Dashboard: spending vs budget, savings rate, net cash flow.'],
    ['', '5. Settle reimbursements → update Reimbursed column in Transactions or Group Splits.'],
    ['', '6. Settle subscription splits → update Subscriptions tab.'],
    ['', '7. Update Savings Goals "Current" balances.'],
    ['', ''],
    ['SECTION', 'KEY FORMULAS'],
    ['', 'Transactions!J  = Orig. Amount × Exch. Rate (PKR equivalent)'],
    ['', 'Transactions!M  = TEXT(Date, "MMMM yyyy")'],
    ['', 'Budget!E        = SUMIFS from Transactions by month + category (uses PKR amount)'],
    ['', 'Income!F        = TEXT(Date, "MMMM yyyy")'],
    ['', 'SavingsGoals!E  = Current ÷ Target'],
    ['', 'Subscriptions!G = Monthly Cost / Total Slots (per-person share)'],
    ['', 'Subscriptions!H = Per Person × # Friends (expected reimbursement)'],
    ['', 'Subscriptions!I = Monthly Cost − Expected Reimb. (your net cost)'],
  ];

  lines.forEach(function(line, i) {
    sh.getRange(i + 1, 1).setValue(line[0]);
    sh.getRange(i + 1, 2).setValue(line[1]);
  });

  lines.forEach(function(line, i) {
    if (line[0] === 'TITLE') {
      sh.getRange(i + 1, 2).setFontSize(16).setFontWeight('bold').setFontColor(COLORS.PRIMARY);
    } else if (line[0] === 'SECTION') {
      sh.getRange(i + 1, 1, 1, 2)
        .setBackground(COLORS.LIGHT_BLUE)
        .setFontWeight('bold').setFontSize(11).setFontColor(COLORS.PRIMARY_DARK);
    }
  });
}

// ── RECURRING SHEET ───────────────────────────────────────────
// Optional. Run once to create the sheet, then add fixed monthly
// expenses. processRecurringTransactions() in Triggers.gs posts
// each row automatically on its configured day of the month.

function setupRecurring(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(SHEETS.RECURRING)) ss.insertSheet(SHEETS.RECURRING);

  var sh    = ss.getSheetByName(SHEETS.RECURRING);
  var lists = ss.getSheetByName(SHEETS.LISTS);
  sh.clear();
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations();

  var headers = [
    'Description', 'Category', 'Amount (PKR)', 'Day of Month',
    'Paid By', 'Shared?', 'Active?', 'Last Posted', 'Notes'
  ];
  var hdr = sh.getRange(1, 1, 1, headers.length);
  hdr.setValues([headers]);
  styleHeaderRow(hdr);
  sh.setFrozenRows(1);

  [200, 130, 120, 110, 110, 80, 80, 115, 200]
    .forEach(function(w, i) { sh.setColumnWidth(i + 1, w); });

  setDropdown(sh, 2, 2, 200, lists.getRange('C2:C15'));
  setDropdown(sh, 2, 5, 200, lists.getRange('E2:E6'));
  setDropdown(sh, 2, 6, 200, lists.getRange('G2:G3'));
  setDropdown(sh, 2, 7, 200, lists.getRange('G2:G3'));

  sh.getRange(2, 3, 200, 1).setNumberFormat(PKR_FORMAT);
  sh.getRange(2, 8, 200, 1).setNumberFormat('MM/DD/YYYY');

  sh.getRange(2, 1, 3, 9).setValues([
    ['House Rent',     'Housing',       80000, 1, PERSON1, 'Yes', 'Yes', '', 'Post on 1st of month'],
    ['Internet Bill',  'Subscriptions',  3500, 5, PERSON1, 'Yes', 'Yes', '', ''],
    ['Gym Membership', 'Health',        12000, 1, PERSON2, 'No',  'Yes', '', PERSON2 + ' personal'],
  ]);
  sh.getRange(2, 3, 3, 1).setNumberFormat(PKR_FORMAT);

  Logger.log('Recurring sheet created. Edit sample rows, then run installAllTriggers() to activate.');
  ss.toast('Recurring sheet created. Add fixed monthly expenses and run installAllTriggers().', 'Done', 10);
}
