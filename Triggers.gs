// ============================================================
// HOUSEHOLD FINANCE TRACKER — Triggers.gs
// Trigger management and all automated daily/monthly functions.
//
// Run installAllTriggers() once from the editor to enable:
//   • onFormSubmit  — fires on every form response
//   • dailyAutomation — fires at 7 AM every day
//
// dailyAutomation() calls:
//   • processRecurringTransactions  — every day
//   • checkAndSendAlerts            — every day
//   • autoNewMonthBudget            — 1st of month only
//   • sendMonthlyEmailSummary       — 1st of month only
// ============================================================

function setupTriggers(ss) {
  var target = ss || SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'onFormSubmit') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onFormSubmit').forSpreadsheet(target).onFormSubmit().create();
  Logger.log('Trigger installed: onFormSubmit');
}

function installAllTriggers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  ScriptApp.getProjectTriggers().forEach(function(t) {
    var fn = t.getHandlerFunction();
    if (fn === 'onFormSubmit' || fn === 'dailyAutomation') ScriptApp.deleteTrigger(t);
  });

  ScriptApp.newTrigger('onFormSubmit').forSpreadsheet(ss).onFormSubmit().create();
  ScriptApp.newTrigger('dailyAutomation').timeBased().atHour(7).everyDays(1).create();

  Logger.log('Triggers installed: onFormSubmit + dailyAutomation (daily 7 AM)');
  ss.toast('Triggers installed. Daily automation runs at 7 AM each day.', 'Done', 8);
}

// ── DAILY AUTOMATION ─────────────────────────────────────────
// Called automatically each day at 7 AM by the time-based trigger.

function dailyAutomation() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var dom = new Date().getDate();

  processRecurringTransactions(ss);
  processSubscriptionSplits(ss);
  checkAndSendAlerts(ss);

  if (dom === 1) {
    autoNewMonthBudget(ss);
    sendMonthlyEmailSummary(ss);
  }
}

// ── BUDGET AUTO-COPY ──────────────────────────────────────────
// On the 1st of each month, copies last month's budget rows to the new month.
// Idempotent — skips if the current month already has rows.

function autoNewMonthBudget(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  var sh  = ss.getSheetByName(SHEETS.BUDGET);
  var tz  = Session.getScriptTimeZone();
  var now = new Date();

  var curMonth  = Utilities.formatDate(now, tz, 'MMMM yyyy');
  var prevMonth = Utilities.formatDate(
    new Date(now.getFullYear(), now.getMonth() - 1, 1), tz, 'MMMM yyyy'
  );

  var last = sh.getLastRow();
  if (last < 2) return;

  var data = sh.getRange(2, 1, last - 1, 8).getValues();

  if (data.some(function(r) { return r[0] === curMonth; })) {
    Logger.log('autoNewMonthBudget: ' + curMonth + ' already exists — skipped.');
    return;
  }

  var prevRows = data.filter(function(r) { return r[0] === prevMonth; });
  if (prevRows.length === 0) {
    Logger.log('autoNewMonthBudget: no rows for ' + prevMonth);
    return;
  }

  var writeRow = sh.getLastRow() + 1;

  // Write ONLY cols A–D (Month, Category, Type, Budget)
  // Never touch E/F/G — those belong to ARRAYFORMULA
  var budgetCols = prevRows.map(function(r) {
    return [curMonth, r[1], r[2], r[3]];
  });
  sh.getRange(writeRow, 1, budgetCols.length, 4).setValues(budgetCols);
  sh.getRange(writeRow, 1, budgetCols.length, 1).setNumberFormat('@STRING@');
  sh.getRange(writeRow, 4, budgetCols.length, 1).setNumberFormat(PKR_FORMAT);

  // Write Notes (col H) separately, leaving E/F/G completely untouched
  var notesCols = prevRows.map(function(r) { return [r[7]]; });
  sh.getRange(writeRow, 8, notesCols.length, 1).setValues(notesCols);

  Logger.log('autoNewMonthBudget: wrote ' + budgetCols.length + ' rows for ' + curMonth);
}

// ── MONTHLY EMAIL SUMMARY ─────────────────────────────────────
// Sends an HTML email summary for the previous calendar month.
// Called automatically on the 1st by dailyAutomation().

function sendMonthlyEmailSummary(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  var tz  = Session.getScriptTimeZone();
  var now = new Date();
  var prevMonthDate = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  var month = Utilities.formatDate(prevMonthDate, tz, 'MMMM yyyy');

  var tx  = ss.getSheetByName(SHEETS.TRANSACTIONS);
  var inc = ss.getSheetByName(SHEETS.INCOME);

  var txLast = tx.getLastRow();
  var txData = txLast > 1 ? tx.getRange(2, 1, txLast - 1, 14).getValues() : [];

  var totalIncome = 0, totalSpending = 0;
  var categoryMap = {};

  txData.forEach(function(r) {
    var rowMonth = r[12]; // M = Month (col 13, 0-indexed = 12)
    var type     = r[2];  // C = Type
    var amtPKR   = r[9];  // J = Amount (PKR)
    var cat      = r[3];  // D = Category

    if (rowMonth !== month) return;

    if (type === 'Expense') {
      totalSpending += (amtPKR || 0);
      categoryMap[cat] = (categoryMap[cat] || 0) + (amtPKR || 0);
    }
  });

  var incLast = inc.getLastRow();
  if (incLast > 1) {
    inc.getRange(2, 1, incLast - 1, 7).getValues().forEach(function(r) {
      if (r[5] === month) totalIncome += (r[4] || 0); // F=Month, E=Amount
    });
  }

  var net         = totalIncome - totalSpending;
  var savingsRate = totalIncome > 0 ? (net / totalIncome * 100).toFixed(1) : '0.0';

  var cats = Object.keys(categoryMap)
    .map(function(k) { return {cat: k, amt: categoryMap[k]}; })
    .sort(function(a, b) { return b.amt - a.amt; })
    .slice(0, 5);

  function fmtPKR(n) { return 'Rs ' + Math.round(n).toLocaleString(); }

  var catRows = cats.map(function(c) {
    var pct = totalSpending > 0 ? (c.amt / totalSpending * 100).toFixed(1) : '0';
    return '<tr><td style="padding:5px 10px">' + c.cat + '</td>' +
           '<td style="padding:5px 10px;text-align:right">' + fmtPKR(c.amt) + '</td>' +
           '<td style="padding:5px 10px;text-align:right">' + pct + '%</td></tr>';
  }).join('');

  var netColor = net >= 0 ? '#34a853' : '#ea4335';

  var html =
    '<div style="font-family:sans-serif;max-width:580px;margin:auto;border:1px solid #dadce0;border-radius:8px;overflow:hidden">' +
    '<div style="background:#1a73e8;color:#fff;padding:18px 20px">' +
    '<h2 style="margin:0;font-size:18px">Monthly Finance Summary</h2>' +
    '<p style="margin:4px 0 0;opacity:.85">' + month + '</p></div>' +
    '<div style="padding:16px 20px;background:#f8f9fa;display:flex;gap:16px">' +
    '<div style="flex:1;text-align:center"><div style="font-size:11px;color:#5f6368">Income</div>' +
    '<div style="font-size:22px;font-weight:bold;color:#34a853">' + fmtPKR(totalIncome) + '</div></div>' +
    '<div style="flex:1;text-align:center"><div style="font-size:11px;color:#5f6368">Spending</div>' +
    '<div style="font-size:22px;font-weight:bold;color:#ea4335">' + fmtPKR(totalSpending) + '</div></div>' +
    '<div style="flex:1;text-align:center"><div style="font-size:11px;color:#5f6368">Net</div>' +
    '<div style="font-size:22px;font-weight:bold;color:' + netColor + '">' + fmtPKR(net) + '</div></div>' +
    '<div style="flex:1;text-align:center"><div style="font-size:11px;color:#5f6368">Savings Rate</div>' +
    '<div style="font-size:22px;font-weight:bold">' + savingsRate + '%</div></div></div>' +
    '<div style="padding:16px 20px">' +
    '<h3 style="margin:0 0 8px;font-size:13px;color:#1557b0">Top Spending Categories</h3>' +
    '<table width="100%" style="border-collapse:collapse;font-size:13px">' +
    '<tr style="background:#e8f0fe"><th style="padding:5px 10px;text-align:left">Category</th>' +
    '<th style="padding:5px 10px;text-align:right">Amount</th>' +
    '<th style="padding:5px 10px;text-align:right">Share</th></tr>' +
    catRows + '</table>' +
    '<p style="margin:16px 0 0;font-size:11px;color:#9aa0a6">Sent by Household Finance Tracker · ' +
    '<a href="' + ss.getUrl() + '">Open spreadsheet</a></p></div></div>';

  MailApp.sendEmail({ to: ALERT_EMAIL, subject: 'Finance Summary — ' + month, htmlBody: html });
  Logger.log('Monthly summary sent for ' + month);
}

// ── ALERT SYSTEM ──────────────────────────────────────────────
// Emails alerts for large transactions and overdue reimbursements.
// PropertiesService deduplicates — each row triggers at most one alert per type.

function checkAndSendAlerts(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  var props  = PropertiesService.getScriptProperties();
  var tz     = Session.getScriptTimeZone();
  var today  = new Date();
  var alerts = [];

  var tx     = ss.getSheetByName(SHEETS.TRANSACTIONS);
  var txLast = tx.getLastRow();
  if (txLast < 2) return;

  var txData = tx.getRange(2, 1, txLast - 1, 14).getValues();

  txData.forEach(function(r, i) {
    var date   = r[0];  // A Date
    var type   = r[2];  // C Type
    var cat    = r[3];  // D Category
    var desc   = r[5];  // F Description
    var amtPKR = r[9];  // J Amount (PKR)

    if (!date) return;
    var rowKey = 'row_' + (i + 2);

    if (type === 'Expense' && amtPKR > LARGE_TX_THRESHOLD) {
      var k = 'large_' + rowKey;
      if (!props.getProperty(k)) {
        alerts.push('Large expense on ' +
          Utilities.formatDate(date instanceof Date ? date : new Date(date), tz, 'dd MMM') +
          ': ' + (desc || cat) + ' — Rs ' + Math.round(amtPKR).toLocaleString());
        props.setProperty(k, '1');
      }
    }
  });

  if (alerts.length > 0) {
    MailApp.sendEmail({
      to:      ALERT_EMAIL,
      subject: '[Finance Tracker] ' + alerts.length + ' alert(s)',
      body:    alerts.join('\n\n') + '\n\nOpen spreadsheet: ' + ss.getUrl(),
    });
    Logger.log('checkAndSendAlerts: sent ' + alerts.length + ' alert(s)');
  }
}

// ── SUBSCRIPTION SPLITS ───────────────────────────────────────
// Runs daily via dailyAutomation(). On a subscription's billing day,
// creates one Group Split row per person listed in "Split With".
// Idempotent — checks Split Last Posted (col M) to avoid double-posting.
//
// Subscriptions column map used here:
//   A(0) Subscription, B(1) Payer, C(2) Monthly Cost,
//   D(3) Split With, E(4) Total Slots, J(9) Renewal Date,
//   K(10) Status, M(12) Split Last Posted

function processSubscriptionSplits(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  var subSh = ss.getSheetByName(SHEETS.SUBSCRIPTIONS);
  var grpSh = ss.getSheetByName(SHEETS.GROUP_SPLITS);
  if (!subSh || !grpSh) return;

  var tz       = Session.getScriptTimeZone();
  var now      = new Date();
  var dom      = now.getDate();
  var curMonth = Utilities.formatDate(now, tz, 'MMMM yyyy');

  var last = subSh.getLastRow();
  if (last < 2) return;

  var data = subSh.getRange(2, 1, last - 1, 13).getValues();

  data.forEach(function(r, i) {
    var subName     = r[0];   // A Subscription name
    var payer       = r[1];   // B Payer
    var monthly     = r[2];   // C Monthly Cost (PKR)
    var splitWith   = r[3];   // D Split With (comma-separated)
    var totalSlots  = r[4];   // E Total Slots
    var renewalDate = r[9];   // J Renewal Date
    var status      = r[10];  // K Status
    var lastPosted  = r[12];  // M Split Last Posted

    if (status !== 'Active') return;
    if (!splitWith || !String(splitWith).trim()) return;
    if (!monthly || monthly <= 0) return;

    // Fire only on the subscription's billing day
    var billingDay = (renewalDate instanceof Date) ? renewalDate.getDate() : 1;
    if (dom !== billingDay) return;

    // Idempotency: skip if already posted this calendar month
    if (lastPosted instanceof Date &&
        Utilities.formatDate(lastPosted, tz, 'MMMM yyyy') === curMonth) return;

    // Parse names list
    var names = String(splitWith).split(',')
      .map(function(n) { return n.trim(); })
      .filter(function(n) { return n.length > 0; });
    if (names.length === 0) return;

    // Per-person share based on Total Slots; fall back to splitting evenly
    var slots     = (totalSlots && totalSlots > 0) ? totalSlots : names.length + 1;
    var perPerson = monthly / slots;

    // Scan col A (Date) for true last data row — getLastRow() is inflated
    // by ARRAYFORMULA outputs in cols I/J which emit "" for all empty rows.
    var gsLast0 = grpSh.getLastRow();
    var nextRow = 4;
    if (gsLast0 >= 4) {
      var gsColA0 = grpSh.getRange(4, 1, gsLast0 - 3, 1).getValues();
      for (var si = gsColA0.length - 1; si >= 0; si--) {
        if (gsColA0[si][0] !== '') { nextRow = si + 5; break; }
      }
    }
    var rows = names.map(function(personName) {
      return [now, payer, subName + ' — ' + curMonth, 'Subscriptions',
              monthly, personName, perPerson, 0, '', '', 'Auto: subscription split'];
    });

    grpSh.getRange(nextRow, 1, rows.length, 11).setValues(rows);
    grpSh.getRange(nextRow, 1, rows.length, 1).setNumberFormat('MM/DD/YYYY');
    grpSh.getRange(nextRow, 5, rows.length, 1).setNumberFormat(PKR_FORMAT);
    grpSh.getRange(nextRow, 7, rows.length, 3).setNumberFormat(PKR_FORMAT);

    // Mark last posted so we don't double-post
    subSh.getRange(i + 2, 13).setValue(now).setNumberFormat('MM/DD/YYYY');

    Logger.log('Subscription split posted: "' + subName + '" → ' + names.join(', ') + ' for ' + curMonth);
  });
}

// ── TRANSACTION → GROUP SPLIT: onEdit handler ────────────────
// Simple trigger — fires automatically on every edit, no installation needed.
// Watches Transactions col O (Group Split?) and col P (# People).

function onEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  if (sheet.getName() !== SHEETS.TRANSACTIONS) return;

  var col = e.range.getColumn();
  var row = e.range.getRow();
  if (row < 2 || e.range.getNumRows() > 1 || e.range.getNumColumns() > 1) return;

  var ss = e.source;

  if (col === TX_COL_GROUP_SPLIT) {
    if (e.value === 'Yes') {
      createGroupSplitRowsFromTx(ss, sheet, row);
    } else {
      removeOrFlagGroupSplitRows(ss, sheet, row);
    }
  } else if (col === TX_COL_NUM_PEOPLE) {
    // Create rows when # People is filled and Group Split? is already Yes but no key yet
    if (sheet.getRange(row, TX_COL_GROUP_SPLIT).getValue() !== 'Yes') return;
    if (!sheet.getRange(row, TX_COL_TX_KEY).getValue()) {
      createGroupSplitRowsFromTx(ss, sheet, row);
    }
  }
}

function makeTxKey(txRow, date, amount) {
  var tz = Session.getScriptTimeZone();
  var d = (date instanceof Date) ? date : new Date(date);
  var dateStr = Utilities.formatDate(d, tz, 'yyyyMMdd');
  return 'TX_R' + txRow + '_' + dateStr + '_' + Math.round(Math.abs(amount || 0));
}

function createGroupSplitRowsFromTx(ss, txSheet, txRow) {
  var txData   = txSheet.getRange(txRow, 1, 1, TX_COL_NUM_PEOPLE).getValues()[0];
  var date     = txData[0];   // A Date
  var paidBy   = txData[1];   // B Person
  var cat      = txData[3];   // D Category
  var desc     = txData[5];   // F Description
  var origAmt  = txData[7];   // H Orig.Amount
  var exchRate = txData[8];   // I Exch.Rate
  var amtPKR   = txData[9];   // J Amount(PKR) — ARRAYFORMULA (may already be resolved)
  var numPpl   = parseInt(txData[TX_COL_NUM_PEOPLE - 1], 10);  // P # People

  var totalAmt = (amtPKR > 0) ? amtPKR : (origAmt || 0) * (exchRate > 0 ? exchRate : 1);

  if (!date || !(date instanceof Date)) {
    ss.toast('Enter the transaction date before enabling Group Split?.', 'Group Split', 6);
    return;
  }
  if (!totalAmt || totalAmt <= 0) {
    ss.toast('Enter the transaction amount before enabling Group Split?.', 'Group Split', 6);
    return;
  }
  if (!numPpl || numPpl < 1) {
    ss.toast('Set # People first, then set Group Split? = Yes.', 'Group Split', 6);
    return;
  }

  // Idempotency: if a key is already stored and matching rows exist, skip
  var storedKey = txSheet.getRange(txRow, TX_COL_TX_KEY).getValue();
  var grpSh = ss.getSheetByName(SHEETS.GROUP_SPLITS);
  if (storedKey) {
    var lastRow = grpSh.getLastRow();
    if (lastRow >= 4) {
      var existingKeys = grpSh.getRange(4, GS_COL_TX_KEY, lastRow - 3, 1).getValues();
      if (existingKeys.some(function(r) { return r[0] === storedKey; })) {
        ss.toast('Group Split rows already exist for this transaction.', 'Already Created', 4);
        return;
      }
    }
  }

  var txKey = makeTxKey(txRow, date, totalAmt);
  // Scan col A (Date) for true last data row — getLastRow() is inflated
  // by ARRAYFORMULA outputs in cols I/J which emit "" for all empty rows.
  var gsLast = grpSh.getLastRow();
  var nextRow = 4;
  if (gsLast >= 4) {
    var gsColA = grpSh.getRange(4, 1, gsLast - 3, 1).getValues();
    for (var j = gsColA.length - 1; j >= 0; j--) {
      if (gsColA[j][0] !== '') { nextRow = j + 5; break; }
    }
  }

  // Write A–H (cols 1–8); skip I=Outstanding and J=Month (ARRAYFORMULA columns)
  var rowsAH = [];
  for (var i = 0; i < numPpl; i++) {
    rowsAH.push([date, paidBy, desc, cat, totalAmt, '', 0, 0]);
  }
  var blank = Array.apply(null, {length: numPpl}).map(function() { return ['']; });
  var keys  = Array.apply(null, {length: numPpl}).map(function() { return [txKey]; });

  grpSh.getRange(nextRow, 1,              numPpl, 8).setValues(rowsAH);
  grpSh.getRange(nextRow, 11,             numPpl, 1).setValues(blank);  // K Notes
  grpSh.getRange(nextRow, GS_COL_TX_KEY, numPpl, 1).setValues(keys);   // L Tx Key

  grpSh.getRange(nextRow, 1, numPpl, 1).setNumberFormat('MM/DD/YYYY');
  grpSh.getRange(nextRow, 5, numPpl, 1).setNumberFormat(PKR_FORMAT);
  grpSh.getRange(nextRow, 7, numPpl, 2).setNumberFormat(PKR_FORMAT);

  txSheet.getRange(txRow, TX_COL_TX_KEY).setValue(txKey);

  ss.toast(
    numPpl + ' rows created in Group Splits for "' + (desc || 'this transaction') + '". ' +
    'Fill in friend names and their share amounts.',
    'Group Split Created', 8
  );
  Logger.log('createGroupSplitRowsFromTx: key=' + txKey + ' rows=' + numPpl + ' at GS row=' + nextRow);
}

function removeOrFlagGroupSplitRows(ss, txSheet, txRow) {
  var txKey = txSheet.getRange(txRow, TX_COL_TX_KEY).getValue();
  if (!txKey) return;

  var grpSh   = ss.getSheetByName(SHEETS.GROUP_SPLITS);
  var lastRow = grpSh.getLastRow();
  if (lastRow < 4) return;

  var data    = grpSh.getRange(4, 1, lastRow - 3, GS_COL_TX_KEY).getValues();
  var toClear = [];
  var flagged = [];

  data.forEach(function(r, i) {
    if (r[GS_COL_TX_KEY - 1] !== txKey) return;  // col L (0-indexed = 11)
    var owedBy     = r[5];   // F Owed By
    var share      = r[6];   // G Their Share
    var reimbursed = r[7];   // H Reimbursed
    var notes      = r[10];  // K Notes
    var hasEdits = (owedBy && String(owedBy).trim() !== '') ||
                   (share > 0) || (reimbursed > 0) ||
                   (notes && String(notes).trim() !== '');
    (hasEdits ? flagged : toClear).push(i + 4);
  });

  // Clear rows that have no user data (clearContent preserves ARRAYFORMULA at I4/J4)
  toClear.forEach(function(sheetRow) {
    grpSh.getRange(sheetRow, 1, 1, 8).clearContent();   // A–H
    grpSh.getRange(sheetRow, 11, 1, 2).clearContent();  // K–L
  });

  if (flagged.length > 0) {
    ss.toast(
      flagged.length + ' Group Split row(s) have data and were kept. ' +
      'Review Group Splits tab to manage them manually.',
      'Manual Review Needed', 10
    );
  } else if (toClear.length > 0) {
    ss.toast(toClear.length + ' empty Group Split row(s) cleared.', 'Group Split', 5);
  }

  if (flagged.length === 0) {
    txSheet.getRange(txRow, TX_COL_TX_KEY).clearContent();
  }
  Logger.log('removeOrFlagGroupSplitRows: key=' + txKey + ' cleared=' + toClear.length + ' flagged=' + flagged.length);
}

// ── RECURRING TRANSACTIONS ────────────────────────────────────
// Runs daily via dailyAutomation(). Posts each active recurring
// row into Transactions on its configured day of the month.
// Idempotent — checks Last Posted to avoid double-posting.

function processRecurringTransactions(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  var recSh = ss.getSheetByName(SHEETS.RECURRING);
  if (!recSh) return;

  var tx  = ss.getSheetByName(SHEETS.TRANSACTIONS);
  var tz  = Session.getScriptTimeZone();
  var now = new Date();
  var dom = now.getDate();
  var curMonth = Utilities.formatDate(now, tz, 'MMMM yyyy');

  var last = recSh.getLastRow();
  if (last < 2) return;

  var data = recSh.getRange(2, 1, last - 1, 9).getValues();

  data.forEach(function(r, i) {
    var desc       = r[0]; // A Description
    var cat        = r[1]; // B Category
    var amt        = r[2]; // C Amount (PKR)
    var dayOfMonth = r[3]; // D Day of Month
    var paidBy     = r[4]; // E Paid By
    var shared     = r[5]; // F Shared?
    var active     = r[6]; // G Active?
    var lastPosted = r[7]; // H Last Posted

    if (active !== 'Yes' || !desc || !amt) return;
    if (parseInt(dayOfMonth) !== dom) return;

    // Idempotency: skip if already posted this calendar month
    if (lastPosted instanceof Date) {
      if (Utilities.formatDate(lastPosted, tz, 'MMMM yyyy') === curMonth) return;
    }

    // Write raw columns only — ARRAYFORMULAs compute J and M automatically
    var newRow = tx.getLastRow() + 1;
    tx.getRange(newRow, 1, 1, 9).setValues([[
      now, paidBy, 'Expense', cat, '', desc, 'PKR', amt, 1
    ]]);
    tx.getRange(newRow, 11, 1, 2).setValues([[paidBy, shared]]);
    tx.getRange(newRow, 1).setNumberFormat('MM/DD/YYYY');
    tx.getRange(newRow, 8).setNumberFormat(PKR_FORMAT);
    tx.getRange(newRow, 9).setNumberFormat(RATE_FORMAT);

    // Mark last posted so we don't double-post
    recSh.getRange(i + 2, 8).setValue(now).setNumberFormat('MM/DD/YYYY');

    Logger.log('Recurring posted: "' + desc + '" for ' + curMonth);
  });
}
