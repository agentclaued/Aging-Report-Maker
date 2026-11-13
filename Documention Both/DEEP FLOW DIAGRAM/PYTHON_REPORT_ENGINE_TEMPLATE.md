# Python Report Engine Template

## Purpose

This is the reusable template for creating reports in Python for WalMartSummary, Amazon Accounts, Regular Accounts, and future companies. It describes a file-only processing flow: upload a workbook, normalize the data, apply company rules, generate report sheets, and return a workbook download. It does not require a SQL stored procedure to calculate the report.

## Reusable Flow

```text
[Select Company]
       |
       v
[Upload File]
       |
       v
[Read Source Workbook / CSV]
       |
       v
[Normalize Column Names]
       |
       v
[Validate Required Company Columns]
       |
       v
[Create Canonical DataFrame]
       |
       +--> Preserve original source columns
       +--> Add calculated helper columns
       +--> Keep invalid-row reasons in a validation log
       |
       v
[Build Company Report DataFrames]
       |
       +--> Detail sheets
       +--> Account/type sheets
       +--> Status summaries
       +--> Dynamic year summaries
       +--> Grand totals
       |
       v
[Write Formatted XLSX Workbook]
       |
       +--> Source-data Excel Tables
       +--> Currency/date formatting
       +--> Company status templates
       +--> Workbook metadata
       |
       v
[Download Company Report]
```

## Common Python Components

```text
app/adapters/
    company_report_processor.py     <- shared report interface/template
    wallmart_python_processor.py    <- future Python replacement for Walmart report logic
    amazon_file_processor.py        <- existing Amazon file-only processor
    regular_accounts_processor.py   <- existing Regular Accounts file-only processor
    validation_engine.py            <- shared validation helpers
    mapping_engine.py               <- shared column mapping helpers
    company_loader.py               <- company profile loading

companies/<company_code>/
    company.json                    <- company identity and processing mode
    report_mapping.json             <- source columns and canonical fields
    report_rules.json               <- company calculation/status rules
    report_layout.json              <- sheet names, headings, and visible columns
    report_style_template.xlsx      <- optional approved workbook style template
```

## Standard Company Template

Each future company report must define the following before its processor is enabled.

```text
1. INPUT
   - Accepted file types
   - Accepted source sheet names or header-detection rule
   - Required columns
   - Optional columns

2. CANONICAL FIELDS
   - account_code
   - document_type
   - order_number
   - invoice_number
   - transaction_date
   - due_date
   - amount
   - reason
   - comment
   - status

3. CALCULATION RULES
   - Dynamic year field
   - Amount aggregation rule
   - Account grouping rule
   - Invoice/DB/CR classification rule
   - Paid rule
   - Dispute rule
   - Write-up rule
   - Omit-zero-section rule

4. OUTPUT
   - Detail-sheet list
   - Summary-sheet list
   - Sheet order
   - Header wording
   - Currency/date format
   - Workbook filename rule

5. VALIDATION
   - Missing-column message
   - Invalid date handling
   - Invalid amount handling
   - Unknown account handling
   - Unknown document-type handling
```

## Walmart Python Migration Template

The existing Walmart SQL logic will be moved in small Python functions. The SQL remains the comparison source until every total and section matches.

```text
Source upload DataFrame
        |
        v
prepare_walmart_summary(source_df)
        |
        +--> normalize Date and Due Date
        +--> InvYear = year(Date)
        +--> WriteStatus from DB Comment
        +--> RetailLinkStatus from Comment
        +--> Paid from IN Comment
        +--> PaidnFullyDeduct from Comment
        +--> Dispute_Amount status from DB Comment
        |
        v
split_walmart_account_tables(summary_df)
        |
        +--> WC9085 Invoices / DBs
        +--> WA9085 Invoices / DBs
        +--> SA9200 Invoices / DBs
        +--> WM9263 Invoices / DBs
        |
        +--------------------------+
        |                          |
        v                          v
build_walmart_dbs_status()   build_walmart_invoices_status()
        |                          |
        +--> dynamic years          +--> dynamic years
        +--> reason/order rows      +--> account rows
        +--> dispute totals         +--> paid totals
        +--> write-up totals        +--> fully-deduct totals
        |                          +--> portal-status totals
        v                          v
write_walmart_workbook()
        |
        +--> DBs_Status sheet
        +--> Invoices_Status sheet
        +--> account detail sheets
        +--> source summary sheet
```

## Walmart Rules to Port from SQL

```text
InvYear
    YEAR(Date)

WriteStatus
    DB row whose Comment starts with a write-up phrase,
    excluding comments containing approval.

RetailLinkStatus
    Comment contains both No and Record.

Paid
    IN row where Comment starts with Unpaid, Paid, or Fully paid,
    using the existing deduction exclusion rule.

PaidnFullyDeduct
    Comment contains both paid and deduc.

Dispute_Amount status
    DB row where Comment starts with REDISPUTE OPEN or DISPUTE OPEN.

DBs Status
    Group DB data by account, reason/order, and dynamic year.
    Summarize Total Amount, Dispute Amount, and Write-up Sent.

Invoices Status
    Group IN data by account and dynamic year.
    Summarize Total Amount, Paid#, Paid & Fully Deduct,
    and Not Found On Portal / RetailLinkStatus amount.
```

## Safe Build Sequence

```text
PHASE 1  Create configuration files and Python function signatures.
         No live Walmart behavior changes.

PHASE 2  Build Python DataFrames from a copy of the Walmart upload.
         Do not write SQL tables.

PHASE 3  Generate a Python-only comparison workbook.
         Compare it with the existing SQL report:
             - row counts
             - dynamic years
             - account totals
             - DB reason/order totals
             - paid/dispute/write-up totals

PHASE 4  Correct only documented differences until outputs match.

PHASE 5  Connect the approved Python processor to the Walmart UI button.
         Retain the SQL route as a fallback until the user approves removal.
```

## Future Company Template

```text
New company upload
      |
      v
Company report_mapping.json
      |
      v
Shared Python reader and validator
      |
      v
Company-specific processor functions
      |
      v
Company-specific formatted XLSX report
```

## Non-Negotiable Standards

- Never mix one company's data into another company's report.
- Keep original source rows on an included detail/source sheet.
- Use typed dates and numeric amounts, not display strings, for calculations.
- Generate dynamic year columns from valid source dates.
- Keep report calculations in Python DataFrames before Excel formatting.
- Keep formatting separate from calculation functions.
- Verify a new processor against a known sample before enabling it in the UI.
- For Walmart, retain the current SQL report as the comparison baseline during migration.
