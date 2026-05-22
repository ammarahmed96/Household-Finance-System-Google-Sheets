// ============================================================
// HOUSEHOLD FINANCE TRACKER — Config.gs
// Global constants and lookup objects shared across all files.
// GAS executes .gs files alphabetically, so Config.gs (C) loads
// before Dashboard.gs, Formatting.gs, etc. — all vars are global.
// ============================================================

// ── PERSONAL SETTINGS ─────────────────────────────────────────
// UPDATE THESE with your actual names before running setup()
var PERSON1 = 'Ammar';
var PERSON2 = 'Wife'; // ← Replace with actual name
var ALERT_EMAIL          = 'ammarahmed12321@gmail.com'; // receives monthly summaries and alerts
var LARGE_TX_THRESHOLD   = 50000; // alert when a single expense exceeds this PKR amount
var OVERDUE_REIMB_DAYS   = 30;    // alert when a reimbursement is outstanding longer than this (days)

// ── CURRENCIES ────────────────────────────────────────────────
var CURRENCIES = ['PKR', 'USD', 'EUR', 'GBP', 'AED', 'SAR', 'CAD', 'AUD', 'CNY', 'INR'];

// ── NUMBER FORMATS ────────────────────────────────────────────
var PKR_FORMAT    = '"Rs "#,##0.00';  // used everywhere for PKR amounts
var RATE_FORMAT   = '#,##0.0000';     // exchange rate (e.g. 280.5000)
var PCT_FORMAT    = '0.0%';

// ── COLOR PALETTE ─────────────────────────────────────────────
var COLORS = {
  PRIMARY:      '#1a73e8',
  PRIMARY_DARK: '#1557b0',
  SUCCESS:      '#34a853',
  WARNING:      '#fbbc04',
  DANGER:       '#ea4335',
  HEADER_FG:    '#ffffff',
  ALT_ROW:      '#f8f9fa',
  BORDER:       '#dadce0',
  LIGHT_BLUE:   '#e8f0fe',
  LIGHT_GREEN:  '#e6f4ea',
  LIGHT_RED:    '#fce8e6',
  LIGHT_YELLOW: '#fef7e0',
  MID_TEXT:     '#5f6368',
  DARK_TEXT:    '#202124',
};

// ── SHEET NAME MAP ────────────────────────────────────────────
var SHEETS = {
  DASHBOARD:      'Dashboard',
  TRANSACTIONS:   'Transactions',
  BUDGET:         'Budget',
  INCOME:         'Income',
  SAVINGS:        'Savings Goals',
  SUBSCRIPTIONS:  'Subscriptions',
  REIMBURSEMENTS: 'Reimbursements',
  GROUP_SPLITS:   'Group Splits',
  LISTS:          'Lists',
  INSTRUCTIONS:   'Instructions',
  RECURRING:      'Recurring',
};

// ── TRANSACTIONS COLUMN MAP (DO NOT CHANGE ORDER) ─────────────
// A=1  Date
// B=2  Person
// C=3  Type
// D=4  Category
// E=5  Subcategory
// F=6  Description
// G=7  Currency          (PKR default)
// H=8  Orig. Amount      (amount in the currency above)
// I=9  Exch. Rate        (rate to PKR; 1 if PKR)
// J=10 Amount (PKR)      ← ARRAYFORMULA: H × I  — DO NOT WRITE STATIC VALUES
// K=11 Paid By
// L=12 Shared?           (No by default; Yes for joint household expenses)
// M=13 Month             ← ARRAYFORMULA: TEXT(A)   — DO NOT WRITE STATIC VALUES
// N=14 Notes
//
// ARRAYFORMULA columns: J2, M2. NEVER write static values into these.
//
// O=15 Group Split?  — user dropdown Yes/No
// P=16 # People      — how many Group Split rows to auto-create
// Q=17 Tx Key        — hidden internal link key (auto-set by script, never edit manually)

// ── TRANSACTION ↔ GROUP SPLIT COLUMN INDICES ─────────────────
var TX_COL_GROUP_SPLIT = 15;  // O
var TX_COL_NUM_PEOPLE  = 16;  // P
var TX_COL_TX_KEY      = 17;  // Q  (hidden)
var GS_COL_TX_KEY      = 12;  // L  (hidden in Group Splits)
