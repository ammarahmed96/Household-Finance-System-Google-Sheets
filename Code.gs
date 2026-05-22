// ============================================================
// HOUSEHOLD FINANCE TRACKER
// Version: 2.0.0 — Refactored into modular files
// ============================================================
//
// This file has been replaced by the modular structure below.
// Paste ALL .gs files into the Apps Script editor — GAS runs
// them in a single shared scope ordered alphabetically:
//
//   Config.gs         — Global constants (COLORS, SHEETS, formats)
//   Setup.gs          — setup() entry point + all sheet builders
//   Reimbursements.gs — setupReimbursements() + setupGroupSplits()
//   Dashboard.gs      — setupDashboard() + formatDashboard() + charts
//   Formulas.gs       — addFormulas() — all ARRAYFORMULA writes
//   Formatting.gs     — CF, professional table formatting, utilities
//   FormUtils.gs      — setupForm() + onFormSubmit()
//   Triggers.gs       — setupTriggers() + daily automation functions
//
// USAGE:
//   1. Run setup()          — builds the full spreadsheet
//   2. Run setupForm()      — creates the Google Form (separate OAuth)
//   3. Run installAllTriggers() — enables daily automation
//   4. Run setupRecurring() — optional recurring transactions sheet
//   5. Run addDashboardCharts() — after first full month of data
// ============================================================
