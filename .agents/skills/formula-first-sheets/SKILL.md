---
name: formula-first-sheets
description: >-
  Use this skill whenever designing or modifying Google Apps Script and Google Sheets workflows,
  especially for reporting, balance sheets, and data pipelines. It enforces using native Google Sheets
  formulas for calculations and summaries instead of static script-calculated values that get wiped on edits.
---

# Formula-First Google Sheets Design

## Core Principle
**Anything that can be computed by a spreadsheet equation MUST use native Google Sheets formulas.**
Reserve Apps Script strictly for tasks that spreadsheets cannot do natively:
1. Fetching external data (Drive files, APIs, email, cross-file imports).
2. Parsing raw binary files (like `.xlsx` unzipping/XML parsing).
3. Structural synchronization (inserting new date columns, appending new driver rows, maintaining metadata).
4. Highly stateful or recursive logic that cannot be expressed cleanly in sheet formulas.

---

## Anti-Patterns to Avoid
❌ **The "Wipe & Write Static Values" Anti-Pattern:**
Clearing entire data ranges (`clearContent()`) and pasting static pre-calculated numbers on every trigger or minor edit. This destroys user formatting, wipes in-flight edits, breaks cell undo history, and prevents live spreadsheet reactivity.

❌ **Intrusive `onEdit` Re-Runs:**
Attaching heavy scripts that clear and regenerate entire sheets to `onEdit` triggers. When users edit a single cell, the sheet should update via its own formula engine instantly, NOT run a 10-second script that re-clears the tab.

❌ **Script-Calculated Averages & Totals:**
Calculating sums, averages, CPH, DEPH, or simple percentages inside JavaScript loops and writing hardcoded numbers, rather than injecting `=SUM(...)`, `=AVERAGE(...)`, `=IF(...)`, or `=VLOOKUP(...)`.

---

## Best Practices

### 1. Dynamic Formulas for Derived Columns
When the script outputs rows, write formulas (e.g. via `setFormulas()` or R1C1 notation) for derived metrics:
- **Cost Per Hour (CPH)**: `=IF(HoursCell>0, CreditCell/HoursCell, 0)`
- **Driver DEPH / Avg Per Hour**: `=IF(HoursCell>0, (CashCell+BalanceCell)/HoursCell, 0)`
- **Standard Balance**: `=ROUND((CreditCell - TollsCell) * VLOOKUP(DriverCell, Settings!$A:$B, 2, FALSE) + TollsCell - CashCell, 2)`
- **Totals / Summaries**: `=SUM(range)`, `=AVERAGE(range)`

### 2. Live User Editing
Because columns are formula-driven, when a manager or accountant adjusts Cash, Tolls, Credit, or Hours in a cell:
- All dependent cells (CPH, Avg/Hour, Balance, Totals) **update instantly** in the UI.
- No script run or network round-trip is required for normal day-to-day tweaks.

### 3. Non-Destructive Ingestion
- Ingestion scripts should write raw data (Dates, Driver names, Daily raw totals) and insert the standard formulas.
- Never wipe manual notes or auxiliary columns when syncing new rows.
- Use explicit manual triggers (e.g. menu item `⚙️ Run Balance Report`) or time-driven scheduled syncs rather than aggressive `onEdit` full-sheet rebuilds.

### 4. Settings Reference via VLOOKUP
- Store business variables (fare percentages, commission tiers, driver configurations) in a dedicated `Settings` tab.
- Data sheets should reference `Settings` via formulas (e.g. `VLOOKUP`) so changing a driver's fare rate updates all balances immediately without needing a script re-run.
