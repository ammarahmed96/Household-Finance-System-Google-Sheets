// ============================================================
// HOUSEHOLD FINANCE TRACKER — FormUtils.gs
// Google Form creation and form submission handler.
//
// Run setupForm() separately from setup() — it requires the
// FormApp OAuth scope which is not needed for the main setup.
//
// IMPORTANT — onFormSubmit writes ONLY raw input columns:
//   A–I (1–9), K–L (11–12), N (14).
//   Columns J and M are ARRAYFORMULA outputs — writing
//   static values into them causes #REF! on the entire formula.
// ============================================================

function setupForm() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var form = FormApp.create('Household Expenses — Quick Entry');
  form.setDescription(
    'Quick expense entry. All amounts converted to PKR automatically.\n' +
    PERSON1 + ' and ' + PERSON2 + ' both use this form.'
  );
  form.setCollectEmail(false);
  form.setAllowResponseEdits(true);
  form.setLimitOneResponsePerUser(false);
  form.setConfirmationMessage('Saved! PKR equivalent auto-calculated. Check the Transactions sheet.');

  form.addDateItem().setTitle('Date').setRequired(true).setIncludesYear(true);

  form.addMultipleChoiceItem()
    .setTitle('Person').setRequired(true).setChoiceValues([PERSON1, PERSON2]);

  form.addMultipleChoiceItem()
    .setTitle('Type').setRequired(true).setChoiceValues(['Expense', 'Income']);

  form.addListItem()
    .setTitle('Category').setRequired(true)
    .setChoiceValues([
      'Housing', 'Food', 'Transportation', 'Health', 'Entertainment',
      'Personal Care', 'Education', 'Savings', 'Travel', 'Gifts',
      'Subscriptions', 'Children', 'Pets', 'Other'
    ]);

  form.addTextItem().setTitle('Subcategory (optional)').setRequired(false);
  form.addTextItem().setTitle('Description').setRequired(true);

  form.addMultipleChoiceItem()
    .setTitle('Currency')
    .setRequired(true)
    .setChoiceValues(CURRENCIES);

  form.addTextItem()
    .setTitle('Amount (in selected currency)')
    .setRequired(true)
    .setValidation(FormApp.createTextValidation().requireNumber().build());

  form.addTextItem()
    .setTitle('Exchange Rate to PKR (leave blank if PKR — e.g. enter 280 for 1 USD = Rs 280)')
    .setRequired(false)
    .setValidation(
      FormApp.createTextValidation()
        .requireNumberGreaterThan(0)
        .setHelpText('Leave blank for PKR transactions')
        .build()
    );

  form.addMultipleChoiceItem()
    .setTitle('Paid By').setRequired(true)
    .setChoiceValues([PERSON1, PERSON2, 'Joint Account', 'Credit Card', 'Cash']);

  form.addMultipleChoiceItem()
    .setTitle('Shared expense?').setRequired(true).setChoiceValues(['No', 'Yes']);

  form.addParagraphTextItem().setTitle('Notes (optional)').setRequired(false);

  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  setupTriggers(ss);

  Logger.log('✅ Form created!');
  Logger.log('Share URL: ' + form.getPublishedUrl());
  Logger.log('Edit URL:  ' + form.getEditUrl());

  ss.toast('Form created! URL is in Apps Script Logs (View → Logs).', 'Form Ready', 15);
}

function onFormSubmit(e) {
  try {
    var ss   = SpreadsheetApp.getActiveSpreadsheet();
    var tx   = ss.getSheetByName(SHEETS.TRANSACTIONS);
    if (!tx) throw new Error('Transactions sheet not found.');
    var resp = e.namedValues;

    function val(key) { return (resp[key] && resp[key][0]) ? resp[key][0].trim() : ''; }

    var dateStr   = val('Date');
    var person    = val('Person');
    var type      = val('Type');
    var category  = val('Category');
    var subcat    = val('Subcategory (optional)');
    var desc      = val('Description');
    var currency  = val('Currency') || 'PKR';
    var amtStr    = val('Amount (in selected currency)');
    var rateStr   = val('Exchange Rate to PKR (leave blank if PKR — e.g. enter 280 for 1 USD = Rs 280)');
    var paidBy    = val('Paid By');
    var shared    = val('Shared expense?') || 'No';
    var notes     = val('Notes (optional)');

    // Validate required fields before writing anything
    if (!person)   throw new Error('Person is required.');
    if (!category) throw new Error('Category is required.');
    if (!amtStr)   throw new Error('Amount is required.');

    var date    = dateStr ? new Date(dateStr) : new Date();
    var origAmt = parseFloat(amtStr);
    if (isNaN(origAmt) || origAmt <= 0) throw new Error('Invalid amount: ' + amtStr);
    var exchRate = (rateStr && parseFloat(rateStr) > 0) ? parseFloat(rateStr) : 1;

    // Write only raw input columns — ARRAYFORMULAs own J(10) and M(13).
    // Writing static values into those columns causes #REF! errors on the entire formula.
    var newRow = tx.getLastRow() + 1;

    // A–I (cols 1–9): date, person, type, category, subcategory, description, currency, amount, rate
    tx.getRange(newRow, 1, 1, 9).setValues([[
      date, person, type, category, subcat, desc, currency, origAmt, exchRate
    ]]);

    // K–L (cols 11–12): Paid By, Shared?
    tx.getRange(newRow, 11, 1, 2).setValues([[paidBy, shared]]);

    // N (col 14): Notes
    tx.getRange(newRow, 14).setValue(notes);

    // Number formats for raw input columns only
    tx.getRange(newRow, 1).setNumberFormat('MM/DD/YYYY');
    tx.getRange(newRow, 8).setNumberFormat(PKR_FORMAT);
    tx.getRange(newRow, 9).setNumberFormat(RATE_FORMAT);

  } catch (err) {
    Logger.log('onFormSubmit error: ' + err.toString());
    try {
      MailApp.sendEmail({
        to:      ALERT_EMAIL,
        subject: '[Finance Tracker] Form submission error',
        body:    'A form submission failed to write to the Transactions sheet.\n\n' +
                 'Error: ' + err.toString() + '\n\n' +
                 'Check Apps Script Logs (Extensions → Apps Script → Executions) for details.',
      });
    } catch (mailErr) {
      Logger.log('Failed to send error email: ' + mailErr.toString());
    }
  }
}
