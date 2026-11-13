# AGENTS.md

## RULES
- Backup important files before each edit.
- Keep runtime focused on WalMartSummary only.
- Prefer small, clear edits.
- Before any unrelated edit action or assumption-changing edit, ask for user permission first.
- Do not create duplicate template versions.
- Do not add multi-company logic unless requested.
- Do not move runtime folders unless code is updated first.
- Keep Step 07 export behavior SQL-driven unless a change is explicitly requested.
- Keep workbook-template wording aligned across UI, docs, and history notes.
- When documentation changes, update both the user-facing doc and the append-only AI history.

## ACTIVE
- `app/`
- `templates/index.html`
- `companies/wallmartsummary/`
- `output/`
- `edit log/`
- `Documention Both/`
- `Documention Both/DEEP FLOW DIAGRAM/`
- `Documention Both/SOFTWARE   PROCESS  FLOW  DIAGRAM/`
- `app_runtime.env`
- `start_wallmartsummary_app.bat`
- `stop_wallmartsummary_app.bat`
- `requirements.txt`
- `AI EDIT CONFIG BY EACH COMPANY TO UPDATE.md`

## EXTRA
- `extras/` contains non-runtime files.
- Do not move files back from `extras/` unless needed.
- Root-level `SOFTWARE   PROCESS  FLOW  DIAGRAM/` is not the current runtime documentation location.
- Prefer `Documention Both/DEEP FLOW DIAGRAM/` for current deep flow markdown and HTML files.
- Prefer `Documention Both/SOFTWARE   PROCESS  FLOW  DIAGRAM/` for current user-flow and developer-flow Mermaid outputs.

## STEP 07 EXPORT NOTES
- `Run step_07_crApplied` builds the SQL source table `step_07_Credit_Applied_Sheet`.
- `Build Latest Tables Workbook` and `/process/export-latest-tables` were removed on 2026-10-06; `/download/latest-tables` remains.
- The separate `Build 07 Applied Sheet` export (`/process/export-step-07-applied`, `/download/step-07-applied`) has been removed.
- Do not document Step 07 as depending on a previously saved xlsx merge unless the code changes to do that.

## DOCUMENTATION RULES
- Keep template-style export documentation in its own box when describing `Download Latest SQL Tables`.
- Keep `Docs/HUMAN_GUIDE.md` focused on operator behavior and run order.
- Keep `output/wallmartsummary_workflow_documentation.html` aligned with current UI wording and routes.
- Keep flow-diagram outputs under `Documention Both/` and note whether they are user flow or developer flow.

## EDIT HISTORY RULES
- `edit log/ai edit history.txt` is append-only.
- Add date, time, purpose, files changed, backups created, and line references when available.
- When documenting a multi-file change, group the history notes by file and include the route/button/doc impact.
- Do not remove older history blocks.

## BEFORE EDIT
- Read `README.md`
- Read `AI EDIT CONFIG BY EACH COMPANY TO UPDATE.md`
- Read `app/main.py`
- Read `app/workflow.py`
- Read `app/adapters/company_loader.py`

## COMPANY CLONE CONFIG
- Use `AI EDIT CONFIG BY EACH COMPANY TO UPDATE.md` as the single reference file for company rename/clone updates.
- When cloning this app for a new company, update the config lines listed there first, then update runtime docs and UI wording to match.

## BACKUP FORMAT
- `filename.bak`
- `filename.backup`
- `filename.YYYY-MM-DD.bak`

## RUN
- `start_wallmartsummary_app.bat`
- `stop_wallmartsummary_app.bat`

## CURRENT STEP 07 FILES
- `app/main.py`
- `templates/index.html`
- `Docs/HUMAN_GUIDE.md`
- `output/wallmartsummary_workflow_documentation.html`
- `Documention Both/SOFTWARE   PROCESS  FLOW  DIAGRAM/User Flow.md`
- `Documention Both/SOFTWARE   PROCESS  FLOW  DIAGRAM/Developer Flow.md`
- `edit log/ai edit history.txt`

## REMOVED ROUTES (2026-10-04)
- The Step 02 upload, Step 03 upload, Step 07 applied-sheet export, and Amazon grouped-sheet download routes were removed because the UI never exposed them.
- `/process/step-02`, `/process/step-03`, `/process/export-step-07-applied`, `/download/step-02`, `/download/step-03`, `/download/step-07-applied`, `/download/amazon-account-sheets`.
- SQL tables `step_02_po_details` and `step_03_payments_combined` are still read by the Step 04+ SQL, but the app no longer has an upload route that fills them.

## WALLMART SUMMARY TABLE FLOW REFERENCE
- Flow: Walmart inquiry file -> `dbo.step_01_inquiry_rows` -> `dbo.WallMartSummary_Sheet` -> account-specific invoice/DB tables -> final status reports.
- Upload process: the Walmart inquiry file creates `dbo.step_01_inquiry_rows`, preserving original upload data, invoice/DB rows, account codes, dates, amounts, statuses, and generating `InvYear` from `[date]`.
- Main summary process: `dbo.step_01_inquiry_rows` is prepared into `dbo.WallMartSummary_Sheet`, including `InvYear`, `WriteStatus`, `RetailLinkStatus`, invoice/DB classification, and other calculated status fields.
- Account table procedure: `dbo.sp__WallMartSummary_db_and_inv_tabels` creates these pairs from the summary source:
  - `dbo.WC9085_Invoices` and `dbo.WC9085_DBs` (WAL-MART CANADA, INC.)
  - `dbo.WA9085_Invoices` and `dbo.WA9085_DBs` (WAL-MART STORE, INC.)
  - `dbo.SA9200_Invoices` and `dbo.SA9200_DBs` (SAM'S CLUB)
  - `dbo.WM9263_Invoices` and `dbo.WM9263_DBs` (WAL-MART STORE #2306)
- Invoice status procedure: `dbo.sp_WallMartSummary_Invoice_Status` reads all four invoice tables and creates `dbo.Invoices_Status` with all-account totals, dynamic-year totals, account totals, total amounts, `Paid#`, Paid & Fully Deduct, and Retail Link Status.
- Open DB status procedure: `dbo.sp_WallMartSummary_gen_DBs_Status` reads all four DB tables and creates `dbo.DBs_Status` with open-aging summary, all-account totals, dynamic-year totals, account/reason headings, reason amounts, account totals, total amount, and Write-up Sent.
- Final report tables: `dbo.Invoices_Status` is the Invoice Summary Report; `dbo.DBs_Status` is the Walmart Open Aging DB Status Summary.
