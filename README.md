# Household Finance Tracker — Google Apps Script

A complete spreadsheet system for tracking household finances for two people. Covers transactions, budgets, income, savings goals, subscriptions, reimbursements, and a live dashboard.

---

## Currency Design

**Default currency: PKR (Pakistani Rupee — Rs)**

All budgets, dashboard metrics, and reimbursements are always calculated and displayed in PKR.

**International transactions** are fully supported:

| Field | Example (USD purchase) |
|-------|----------------------|
| Currency | `USD` |
| Orig. Amount | `50` |
| Exch. Rate | `280` (1 USD = Rs 280 at time of purchase) |
| Amount (PKR) | `Rs 14,000` ← auto-calculated |

The exchange rate is entered manually per-transaction so you capture the **actual rate at the time** — important for travel, online shopping, AED/SAR remittances, etc.

Supported currencies in the dropdown: `PKR, USD, EUR, GBP, AED, SAR, CAD, AUD, CNY, INR` (customizable in `Code.gs` → `CURRENCIES` array).

---

## Features

| Sheet | What it does |
|-------|-------------|
| **Dashboard** | Monthly summary in PKR — income, spending, net cash flow, savings rate, budget vs actual, category breakdown, recent transactions with currency shown |
| **Transactions** | Full ledger with Currency / Orig. Amount / Exch. Rate / Amount (PKR) columns; shared/personal split; auto-filled month and net cost |
| **Budget** | Monthly budgets by category in PKR with auto-calculated actuals and variance |
| **Income** | Income log by person in PKR with auto-filled month |
| **Savings Goals** | Goal tracking in PKR with auto-calculated progress % and remaining |
| **Subscriptions** | Recurring subscription tracker in PKR with reimbursement and net cost |
| **Reimbursements** | Live view of all pending reimbursements (PKR) from transactions and subscriptions |
| **Lists** | Hidden helper sheet for all dropdown values including currencies |
| **Instructions** | Built-in user guide including currency workflow |

---

## Setup Instructions

### Step 1 — Open the Script Editor

1. Open your Google Sheet
2. Click **Extensions → Apps Script**
3. Paste the contents of `Code.gs` into the editor (replace any existing code)

### Step 2 — Update Your Names

At the top of `Code.gs`, change:

```javascript
var PERSON1 = 'Ammar';
var PERSON2 = 'Wife'; // ← Replace with actual name
```

### Step 3 — Run `setup()`

1. In the Apps Script editor, select `setup` from the function dropdown
2. Click **Run**
3. Authorize the script when prompted (it needs access to your Google Sheet)
4. Wait ~20–30 seconds for setup to complete
5. You'll see a toast notification: "Setup complete!"

### Step 4 — Create the Google Form (optional, for mobile entry)

1. In the Apps Script editor, select `setupForm` from the function dropdown
2. Click **Run**
3. Authorize additional OAuth scopes when prompted (FormApp access)
4. The form URL is printed to the **Logs** (View → Logs)
5. Share the URL with both users for quick mobile expense entry

### Step 5 — Add Dashboard Charts (optional)

After entering some transactions:

1. Select `addDashboardCharts` in the function dropdown
2. Click **Run**
3. Two charts appear on the Dashboard: spending by category (donut) and budget vs actual (bar)

---

## Manual Steps After Setup

| Action | Where |
|--------|-------|
| Update spouse name | `Code.gs` → `PERSON2` variable → re-run `setup()` |
| Set monthly budgets | **Budget** tab → update amounts in column D |
| Add income entries | **Income** tab → add rows |
| Enter transactions | **Transactions** tab or Google Form |
| Mark reimbursement received | **Transactions** tab → update column N (Reimbursed Amt.) |
| Settle subscription split | **Subscriptions** tab → update column F (Reimbursed) |
| Switch dashboard month | **Dashboard** tab → change cell C2 (format: "May 2025") |

---

## Transactions Column Guide

| Col | Name | Auto? | Notes |
|-----|------|-------|-------|
| A | Date | — | Enter date |
| B | Person | — | Dropdown |
| C | Type | — | Expense / Income / Transfer |
| D | Category | — | Dropdown |
| E | Subcategory | — | Optional dropdown |
| F | Description | — | Free text |
| G | **Currency** | — | Dropdown — PKR default; pick USD/EUR/etc. for international |
| H | **Orig. Amount** | — | Amount in the selected currency |
| I | **Exch. Rate** | — | PKR rate at time of purchase; leave blank (or 1) for PKR |
| J | **Amount (PKR)** | ✅ Formula | `H × I` — the PKR equivalent used everywhere |
| K | Paid By | — | Dropdown |
| L | Shared? | — | Yes / No |
| M | Split % | — | Other person's share; blank = 50/50 |
| N | Reimbursable? | — | Yes / No |
| O | Reimb. From | — | Dropdown: who owes |
| P | Exp. Reimb. | ✅ Formula | `Amount(PKR) × Split%` |
| Q | Reimbursed | — | Enter manually when paid |
| R | Net Cost | ✅ Formula | `Amount(PKR) − Reimbursed` |
| S | Month | ✅ Formula | Auto from Date |
| T | Notes | — | Free text |

---

## Budget Workflow

1. The Budget tab is pre-populated with the current month's categories.
2. At the start of each new month: copy rows, update the Month column to the new month (e.g., `"June 2025"`).
3. The **Actual** column auto-fills via SUMIFS from Transactions.
4. Use the Dashboard's Budget vs Actual table to review at month-end.

---

## Reimbursement Workflow

**Transactions:**
- Mark `Reimbursable? = Yes`, set `Reimb. From`
- When paid: enter amount in `Reimbursed Amt.` (col N)
- Net Cost updates automatically

**Subscriptions:**
- Set `Expected Reimb.` = the other person's share
- When paid each month: update `Reimbursed` column
- Outstanding amount shows in the Reimbursements tab

---

## Formula Reference

```
Transactions!M  = Amount × (Split%/100), defaults to 50%
Transactions!O  = Amount − Reimbursed Amount
Transactions!P  = TEXT(Date, "MMMM yyyy")

Budget!E        = SUMIFS(Transactions by month + category)
Budget!F        = Budget − Actual
Budget!G        = Actual ÷ Budget

Income!F        = TEXT(Date, "MMMM yyyy")

SavingsGoals!E  = Current ÷ Target
SavingsGoals!F  = Target − Current

Subscriptions!G = Monthly Cost − Expected Reimb.
```

---

## File Structure

```
household-finance-system/
├── Code.gs      ← All Apps Script code (paste into editor)
└── README.md    ← This file
```

---

## Tips

- Use the Google Form for quick mobile entry — no need to open the sheet
- The Dashboard `C2` cell is the only thing you change to navigate months
- Freeze rows are set automatically; scroll right in Transactions for all columns
- The Lists sheet is hidden (it's a helper for dropdowns); unhide via right-click if needed
- Run `addDashboardCharts()` after your first full month of data for meaningful charts
- If a formula column (M, O, P in Transactions) gets overwritten, delete the cell and re-run `addFormulas()` from the editor
