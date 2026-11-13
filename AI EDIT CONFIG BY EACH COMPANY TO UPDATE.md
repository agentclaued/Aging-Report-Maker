# AI EDIT CONFIG BY EACH COMPANY TO UPDATE

## Purpose

This is the single reference file for cloning this app to a new company.
When creating a new company copy, read this file first and update these config points before editing deeper workflow logic.

## Primary Company Identity

1. `app_runtime.env`
- `APP_TITLE=...`
- `SQL_DATABASE=...`
- `HOST=...`
- `PORT=...`
- `DEFAULT_COMPANY_CODE=...`

2. `app/DbConfigLogger.py`
- `Settings.app_title`
- `Settings.sql_database`
- `Settings.default_company_code`
- `get_settings()` env mapping keys:
  - `APP_TITLE`
  - `SQL_SERVER`
  - `SQL_DATABASE`
  - `SQL_DRIVER`
  - `TRUSTED_CONNECTION`
  - `DEFAULT_COMPANY_CODE`

## Company Profile Files

3. `companies/<company_code>/company.json`
- `"code"`
- `"display_name"`
- `"sql_database"`
- `"notes"`

4. `companies/<company_code>/step_01.json`
- Step 01 alias and required-column rules

5. `companies/<company_code>/step_02.json`
- Step 02 sheet aliases
- PO details column aliases
- required columns

6. `companies/<company_code>/step_03.json`
- Step 03 sheet aliases
- detail column aliases
- required columns

## Runtime Launcher Files

7. `start_wallmartsummary_app.bat`
- fallback `HOST`
- fallback `PORT`
- launcher filename when cloning for a new company

8. `stop_wallmartsummary_app.bat`
- fallback `PORT`
- stop-launcher filename when cloning for a new company

## UI And Runtime Wiring

9. `app/main.py`
- `DEFAULT_COMPANY_CODE = settings.default_company_code`
- Step 01A, Step 02, Step 03, and render paths that depend on active company settings

10. `templates/index.html`
- visible company wording
- refresh button behavior
- button labels and page headings

Important:
- Refresh must not hard-code a stale host/port.
- Prefer `window.location.origin + "/"` or runtime-driven values.

## Main Documentation Files

11. `README.md`
- runtime company summary
- start command

12. `Docs/HUMAN_GUIDE.md`
- operator-facing runtime company references
- start/stop commands

13. `AGENTS.md`
- active company folder
- runtime filenames
- documentation filenames
- pre-edit read list

14. `output/wallmartsummary_workflow_documentation.html`
- runtime profile name
- company wording
- launcher/doc file names

## Shared Generated-Workbook Standard

For Wall Mart Summary, Amazon Accounts, Regular Accounts, and each future company:

- Preserve normal Excel data-sheet behavior and add an Excel Table to each generated
  source-data sheet so it can be used as a PivotTable source.
- Format `Date` and `Due Date` output cells as `MMDDYYYY`.
- Keep summary calculations as Excel formulas when the report has an included source-data
  sheet. The formula must remain visible when the user clicks its cell.
- For any ageing summary that has a CR column, calculate CR as
  `-ABS(SUM(CR Total))` per Acct.
- Regular Accounts requires installed Microsoft Excel on the app-host PC to silently
  create its native PivotTables. Keep the top yearly blocks and ageing titles intact; the
  INV/DB PivotTables begin at column A below the title with Ref Type as the filter row,
  then a spacer row and Acct/amount headers. Do not apply this Excel-COM dependency to
  other company flows unless explicitly requested.

## Clone Checklist

When cloning for a new company:

1. Create or copy `companies/<new_company_code>/`
2. Update `app_runtime.env`
3. Update `app/DbConfigLogger.py`
4. Update launcher filenames if the company name is part of them
5. Update `app/main.py` references that depend on the default company
6. Update `templates/index.html` wording and verify refresh behavior
7. Update `README.md`, `Docs/HUMAN_GUIDE.md`, and `AGENTS.md`
8. Update the active workflow documentation HTML filename/content
9. Append both history files after the change

## Current Wal Mart Summary Values

- Company code: `wallmartsummary`
- Display name: `Wal Mart Summary`
- SQL database: `WallMartSummary`
- Host: `192.168.2.123`
- Port: `8040`
- Start launcher: `start_wallmartsummary_app.bat`
- Stop launcher: `stop_wallmartsummary_app.bat`
- Main runtime documentation: `output/wallmartsummary_workflow_documentation.html`
- Processing mode: file-only Python report calculations; the Wal Mart upload/report flow does not open SQL Server.
- Workbook template: `companies/wallmartsummary/wallmart_status_style_template.xlsx`
- Required raw input columns: `Acct`, `Div`, `Name`, `Ref#`, `Date`, `Due Date`, `Order#`, `Current`, `1-30 Days`, `31-60 Days`, `61-90 Days`, `Over 90 Days`, `Claim#`, and `Comment`; preserve extras unchanged in `WallMartSummary_Sheet`. `Ref Type` is optional and is created/refreshed from `Ref#`. SQL-safe age-column variants remain accepted.
- Report type classification: case-insensitive `Ref#` contains `IN`, `DB`, or `CR`; write that result into `Ref Type` for every output sheet.
- Total rule: always refresh/create `total = current + col_1_30_days + col_31_60_days + col_61_90_days + over_90_days`.
- Currency rule: in every Wal Mart output data sheet, format Total, Current, 1-30 Days, 31-60 Days, 61-90 Days, and Over 90 Days with `$` currency formatting.
- Date rule: convert Excel serial values in Date and Due Date (for example `44548`) to real dates and display them as `MM/DD/YYYY`.
- Comment helper rule: blank Comment leaves all comment-based helper columns blank except InvYear. The status helper column is named `PaidStatus`. For IN rows, InvPaid requires Comment to start with a single `Paid` keyword, not `Paid Fully`, and without `deduc`; PaidnFullyDeduct uses `deduct` anywhere. Dispute applies to IN and DB rows; Write-up remains DB-only.
- PaidStatus rule: do not assign a default `Paid` value. Populate it only through the explicit Unpaid and InvPaid conditions.
- Status accounts: include only WC9085, WA9085, SA9200, and WM9263; use source `name` for display labels. Keep blank/unsupported account rows in master detail but show a red log warning and omit them from status totals.

## Current Amazon Accounts Values

- Company code: `amazonaccounts`
- Display name: `Amazon Accounts`
- Processing mode: file-only; no database connection
- Output folder: `output/amazonaccounts/`
- Account/type output: `<Acct>_Invoices`, `<Acct>_DBs`, and `<Acct>_CRs` for `AM9285`, `AM9286`, and `ME9145`; classify from `Acct` and `Ref Type`, not source worksheet names.
- First sheet positions: `0=DB's Status`, `1=Invoice's Status`, `2=CR's Status`; DB and Invoice use Amazon-only source rules, while CR remains ready for its future status logic.
- Master source sheet: position `3=Master Aging`; retain every nonblank uploaded row from every source sheet before account or Ref Type filtering.
- Status style asset: `companies/amazonaccounts/amazon_status_style_template.xlsx`; DB uses the approved DB status style and Invoice/CR use the approved invoice status style, with no source values copied.
- DB status sample: filter `Ref Type=DB`; group `Acct`, `Order#`, and `YEAR(Date)`; create dynamic YYYY columns; sum `Total`; set Write-up Sent to `SUM(Total)` where `Comment` contains `MEMO` case-insensitively; do not create Dispute_Amount.
- Invoice status: filter `Ref Type=IN`; group `Acct` and `YEAR(Date)`; create dynamic YYYY columns and sum `Total`; for the latest Date year only, group `Acct`, `Comment`, and year, then set `Paid @<latest YYYY>` to `SUM(Total)` when Comment contains `Paid on`, a payment-number phrase, and digits. Accept small payment-word typos such as `payemnt no`.
- CR status: filter verified `Ref Type=CR` rows used for `<Acct>_CRs`; group `Acct` and `YEAR(Date)`; create dynamic YYYY columns and sum `Total`; include Total Amount and a Status column with `-` placeholders.
- Universal status visibility: omit any account's entire status section if its final `Total Amount` equals `0.00`, independent of account order.
- Account/type workbook visibility: omit `<Acct>_Invoices`, `<Acct>_DBs`, and `<Acct>_CRs` when their filtered source rows are empty.
- Status formatting cleanup: remove unused rows and copied formatting below the final visible account status section.
- Current state: status uploads are normalized into an Excel workbook; Amazon column mapping and final report rules require sample files before automatic report generation is enabled.

## WalMartSummary Step 02 MCP Context

Use this as the current WalMartSummary-specific reference when updating code or cloning for the next company.

Current WalMartSummary PO-details runtime names:

- Order key: `po_number`
- Quantity key: `quantiy`
- Cost key: `wholesale_price`

WalMartSummary input-to-runtime mapping:

- Excel `PO Number` -> normalized `po_number` -> runtime `po_number`
- Excel `Quantity` -> normalized `quantity` -> runtime `quantiy`
- Excel `Wholesale Price` -> normalized `wholesale_price` -> runtime `wholesale_price`

Files that currently carry the active WalMartSummary PO-details naming (the Step 02 upload route and its workbook prep code were removed; only the SQL-side names remain):

1. `companies/wallmartsummary/step_02.json`
- Canonical PO-details keys:
  - `po_number`
  - `quantiy`
  - `wholesale_price`
- Accepted legacy aliases still allowed during transition:
  - `retail_order`
  - `quantity`
  - `cost`

2. `app/main.py`
- Step 02 upload type SQL
- Step 04 SQL joins and mismatch tables
- Step 04A PO status summaries
- Step 06 joins between invoice and payment tables
- Latest SQL tables workbook money-column formatting

When cloning for the next company:

1. Decide the company's canonical PO-details order, quantity, and cost keys first.
2. Update `companies/<company_code>/step_02.json` before changing SQL in `app/main.py`.
3. Update only the PO-details-side SQL references in `app/main.py` that point to the PO-details table.
4. Keep payment-table names separate unless that company also changes Step 03 naming.
5. Verify the SQL table no longer carries stale old column names from the previous company.
