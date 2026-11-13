# Wal Mart Summary Single-Company Application

This folder is the active Wal Mart Summary application using file-only Python report calculations.

Current status:

- Wal Mart Summary, Amazon Accounts, and Regular Accounts are available from the same UI.
- Wal Mart, Amazon Accounts, and Regular Accounts all process uploaded files without opening a database connection.
- The app now reads the active Wal Mart Summary config files under `companies/wallmartsummary/`.
- The prior Overstock config remains available as an archive/reference profile under `companies/overstock/`.
- Adapter modules remain under `app/adapters/` for clean mapping and validation code.
- Wal Mart Summary uses the approved Wal Mart workbook template and Python calculations; its former SQL/report flow is not used by the upload/report path.

Run from this folder with:

```bat
start_wallmartsummary_app.bat
```

Runtime companies:

- `wallmartsummary`
- File-only output: `output/WAL_MART_STATUS_SUMMARY.xlsx` (no database)
- Required raw upload columns: `Acct`, `Div`, `Name`, `Ref#`, `Date`, `Due Date`, `Order#`, `Current`, `1-30 Days`, `31-60 Days`, `61-90 Days`, `Over 90 Days`, `Claim#`, and `Comment`. `Ref Type` is optional: Python creates or refreshes it as IN, DB, or CR from `Ref#`. Extra columns are kept unchanged in `WallMartSummary_Sheet`. SQL-safe age-column names such as `col_1_30_days` are also accepted.
- Python refreshes `total` on every row as `current + col_1_30_days + col_31_60_days + col_61_90_days + over_90_days`; it creates `total` when the upload does not contain it.
- Every generated Wal Mart data sheet formats `Total`, `Current`, `1-30 Days`, `31-60 Days`, `61-90 Days`, and `Over 90 Days` as currency with a `$` symbol.
- Excel serial dates such as `44548` in `Date` or `Due Date` are converted to real dates and display as `MM/DD/YYYY`.
- `Ref#` determines the type by case-insensitive contains matching: `IN`, `DB`, or `CR`; Python writes that generated value into `Ref Type` for every output sheet.
- The `DBs_Status` and `Invoices_Status` sheets retain the established Wal Mart template layout, headings, colors, and currency formatting. DB rows group by `acct`, distinct `order`, and `YEAR(date)`; `Dispute_Amount` uses comments containing `DISPUTE`; Write-up Sent uses approved Write-up variations without `approval`.
- Invoice rows group by `acct` and `YEAR(date)`. Paid# sums IN comments that start with a single `Paid` keyword, excluding `Paid Fully` and deduction wording. Paid & Fully Deduct uses comments containing `deduct`; Retail Link Status uses comments containing both No and Record. Dispute applies to both IN and DB rows.
- `PaidStatus` has no default Paid value. It is blank unless the explicit Unpaid or InvPaid rule applies, so `deduct` and `DISPUTE` comments do not become Paid.
- When an uploaded `Comment` cell is blank, all comment-based helper columns remain blank (`PaidStatus`, Paid & Fully Deduct, Dispute, Write-up, and Retail Link), while `InvYear` continues to be prepared from Date.
- Only `WC9085`, `WA9085`, `SA9200`, and `WM9263` are included in the status totals. Their display names come from the uploaded `name` column. A blank or unsupported `acct` remains in `WallMartSummary_Sheet` but is shown as a red warning in the log and excluded from status totals.
- `amazonaccounts`
- File-only output: `output/amazonaccounts/` (no database; report mapping awaits Amazon sample files)
- Amazon account/type workbook: reads every input sheet, uses `Acct` and `Ref Type` as the source of truth, and creates `AM9285`, `AM9286`, and `ME9145` Invoice, DB, and CR sheets.
- The first three Amazon workbook sheets are `DB's Status`, `Invoice's Status`, and `CR's Status`. DB and Invoice status rules are populated; CR remains available for its future rule.
- `Master Aging` is the fourth sheet and contains every nonblank data row from every uploaded source sheet before Amazon account or Ref Type filtering.
- Their formatting is copied from the approved workbook template: DB status uses the `DBs_Status` style; Invoice and CR status use the `Invoices_Status` style. No Wal Mart values are copied.
- `Process Amazon File` automatically downloads the completed workbook as `Amazon Status Summary YYYY-MM-DD HH-MM-SS.xlsx`.
- The Amazon process button uses the browser's native automatic download and remains `Processing...` until the server has finished sending the download response; no fixed reset timeout is used.
- `DB's Status` filters `Ref Type=DB`, groups by `Acct`, distinct `Order#`, and `YEAR(Date)`, then sums `Total` into dynamic year columns. Its `Write-up Sent` values sum `Total` only where `Comment` contains `MEMO` (case-insensitive); it has no Dispute_Amount column.
- `Invoice's Status` filters `Ref Type=IN`, groups `Acct` and `YEAR(Date)`, and sums `Total` into dynamic year columns plus `Total Amount`. `Paid @<latest YYYY>` uses the latest Date year only, groups `Acct`, `Comment`, and year, then sums `Total` where Comment has `Paid on`, a payment-number phrase, and payment-number digits; small payment-word typos such as `payemnt no` are accepted.
- `CR's Status` filters verified `Ref Type=CR` rows (the same rows written to `<Acct>_CRs`), groups `Acct` and `YEAR(Date)`, and sums `Total` into dynamic years and Total Amount. Its Status column displays `-`.
- Account status sections with a total amount of `0.00` are omitted; this is the universal Amazon status-summary rule.
- Account/type data sheets are created only when they contain rows; empty `Invoices`, `DBs`, and `CRs` sheets are omitted.
- Unused formatted rows below the final visible Amazon status account section are removed from the output workbook.
- `regularaccounts`
- Company workflow structures are stored with their account: `companies/wallmartsummary/WallMartSummary Structure.txt`, `companies/amazonaccounts/AmazonAccounts Structure.txt`, and `companies/regularaccounts/RegularAccounts Structure.txt`.
- Serverless XLSX processing; choosing a workbook reads its `Full Ageing` sheet locally, finds the header row containing `Assigned Person` and `Acct`, and fills the Person Name dropdown with `{Person Name} @{count(unique Acct values)}`. Selecting a person and clicking Process downloads `{Person Name} {Uploaded File Name}.xlsx`: it includes selected-person `AGEING SHEET` rows, Follow-Up Entries where Status is not blank, IN and DB ageing summaries, and one detail sheet per unique Acct. Every data sheet preserves every source column, original header wording/order, and source value—including `Status`. No database is used.
- App port: `8040`
- Every generated XLSX workbook uses `AnsSys & openpyxl` as its Creator and Last Modified By metadata.
- Generated source-data sheets in all three company flows use normal Excel Tables, so users can create PivotTables directly from the output data. Wal Mart `date` and `due_date` cells display as `MM/DD/YYYY`.
- Regular Accounts top year totals use Excel formulas linked to `AGEING SHEET`. The INV and DB ageing tables are native PivotTables sourced from `AGEING SHEET`, so their values remain interactive and recalculate through Excel's PivotCache.
- Regular Accounts CR values are reported only in the separate `CR'S SUMMARY` sheet. A source row is recognized only when its `Ref Type` is exactly `CR`.
- Regular Accounts `Date` and `Due Date` values display as `MM/DD/YYYY` in FOLLOW-UP ENTRIES, AGEING SHEET, and every account detail sheet.
- Regular Accounts creates native Excel PivotTables silently through Microsoft Excel installed on the app-host PC. All summary sheets use the light-blue type/year top box and title-case ageing titles. Their visible Grand Total rows use `Dark Blue, Text 2, Lighter 60%` (`#B8CCE4`) with bold black text. INV SUMMARY, CR'S SUMMARY, and DB'S BY AGEING keep Row 4 and Row 6 blank, use Row 5 for their ageing title, and use Row 7 for `Acct`, `Name`, `Ref Type`, Current, 1-30 Days, 31-60 Days, 61-90 Days, Over 90 Days, and Total. INV SUMMARY also keeps `TERMS#` and `SALEMAN` after Total. They are limited to IN, CR, and DB rows respectively. No rows or columns are frozen in these summary sheets. `DB'S BY REASON` remains inside `DB'S SUMMARY`, below DB'S BY AGEING, with the visible headers `Order#`, `Ref Type`, Current, 1-30 Days, 31-60 Days, 61-90 Days, Over 90 Days, and Total; it shows DB rows only.
- For every Regular Accounts person report, this is a separate pre-download process: the completed temporary XLSX is saved, hidden Excel adds/updates the PivotTables, saves and closes the same file, then the browser receives the finished download. Wall Mart Summary and Amazon do not use this process.

Main idea:

- keep the current workflow stable
- keep the code clean and human readable
- keep the runtime fixed to Wal Mart Summary unless multi-company support is requested again

## App Control Center - rules for creating or changing it

The `App Control Center/` folder holds the double-click files that run the app (00 terminal, 11 background, 22 PC boot, 33 login, 44 dev mode). Follow these rules when creating or editing it:

- Folder depth: the folder sits ONE level below the project root. Resolve the root with `%~dp0..` (bat), `Join-Path $PSScriptRoot '..'` (PowerShell) and `Path(__file__).resolve().parents[1]` (Python). Never `..\..`.
- Bat files must be saved without a UTF-8 BOM (first bytes `@ec`). A BOM breaks the first line (`@echo off`) in cmd.
- Names: ON files start `NN 1 ON - ...`, OFF files `NN 2 OFF - ...`. The PC boot pair is `22 1 ON - Install Auto-start at PC boot (no recovery).bat` and `22 2 OFF - UnInstall Auto-start at PC boot (no recovery).bat`. Keep `README.txt` in that folder in sync with the file names.
- Tray icon: it is visible only while the app URL (HOST:PORT from `app_runtime.env`) answers; it is hidden while the server is starting or stopped. Its menu has "Open app", "Restart server" (full tray only), "Open log folder" and "Stop the app" (stops the server and removes the icon, in every mode).
- Only one tray app runs at a time (named mutex). A leftover tray process makes a new start exit silently, so every start file must stop the old one first (`stop_background.ps1`).
- HOST, PORT and PYTHON_EXE come from `app_runtime.env`; only one run mode is active at a time.
- Server log: `logs/server.log`. Back up each file before editing and append the change to `edit log/ai edit history.txt`.
