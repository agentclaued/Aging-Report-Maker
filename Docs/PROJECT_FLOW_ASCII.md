# Walmart & Amazon Status Summary — Project Flow

```text
WALMART & AMAZON STATUS SUMMARY
│
├── PURPOSE
│   ├── WalMartSummary
│   │   └── Process Walmart inquiry files through the existing SQL Server workflow.
│   │
│   └── Amazon Accounts
│       └── Process uploaded Amazon workbooks without a database.
│
├── START / RUNTIME
│   ├── start_wallmartsummary_app.bat
│   │   └── Starts the FastAPI web application.
│   │
│   ├── app/main.py
│   │   ├── Loads company selection and page UI.
│   │   ├── Receives uploads.
│   │   ├── Routes Amazon without opening SQL Server.
│   │   └── Routes Walmart to the established SQL workflow.
│   │
│   └── templates/index.html
│       ├── Marketplace selector.
│       ├── File chooser.
│       ├── Process button.
│       └── Download / busy-state user interface.
│
├── COMPANY SELECTOR
│   │
│   ├── wallmartsummary
│   │   ├── Database: WallMartSummary (SQL Server)
│   │   └── Uses Walmart stored procedures and SQL tables.
│   │
│   └── amazonaccounts
│       ├── Database: NONE
│       ├── Output folder: output/amazonaccounts/
│       └── Uses app/adapters/amazon_file_processor.py only.
│
├── WALMART FLOW (SQL SERVER — UNCHANGED BY AMAZON WORK)
│   │
│   ├── Upload Walmart inquiry workbook
│   │   └── Step 01 normalization
│   │
│   ├── SQL source table
│   │   └── dbo.step_01_inquiry_rows
│   │
│   ├── Main prepared summary
│   │   └── dbo.WallMartSummary_Sheet
│   │
│   ├── Account invoice / DB tables
│   │   ├── WA9085_Invoices / WA9085_DBs
│   │   ├── WC9085_Invoices / WC9085_DBs
│   │   ├── SA9200_Invoices / SA9200_DBs
│   │   └── WM9263_Invoices / WM9263_DBs
│   │
│   ├── Status procedures
│   │   ├── dbo.sp_WallMartSummary_Invoice_Status
│   │   │   └── dbo.Invoices_Status
│   │   │
│   │   └── dbo.sp_WallMartSummary_gen_DBs_Status
│   │       └── dbo.DBs_Status
│   │
│   └── Walmart exports
│       └── LATEST_SQL_TABLES.xlsx
│
└── AMAZON FLOW (FILE-ONLY — NO DATABASE)
    │
    ├── User selects "Amazon Accounts"
    │   └── Clicks "Process Amazon File"
    │
    ├── Upload workbook
    │   └── Read every worksheet and every nonblank row.
    │
    ├── Master source preservation
    │   └── Master Aging
    │       └── Contains all uploaded rows before account or Ref Type filtering.
    │
    ├── Validate / classify source rows
    │   ├── [Acct]
    │   │   ├── AM9285
    │   │   ├── AM9286
    │   │   └── ME9145
    │   │
    │   └── [Ref Type]
    │       ├── IN = Invoice
    │       ├── DB = Debit
    │       └── CR = Credit
    │
    ├── Create account data sheets only when rows exist
    │   ├── AM9285_Invoices / AM9285_DBs / AM9285_CRs
    │   ├── AM9286_Invoices / AM9286_DBs / AM9286_CRs
    │   └── ME9145_Invoices / ME9145_DBs / ME9145_CRs
    │
    ├── Create status sheets in fixed positions
    │   ├── [00] DB's Status
    │   ├── [01] Invoice's Status
    │   ├── [02] CR's Status
    │   └── [03] Master Aging
    │
    ├── DB's Status
    │   ├── Source: [Ref Type] = DB
    │   ├── Group: [Acct] + [Order#] + YEAR([Date])
    │   ├── Value: SUM([Total])
    │   ├── Dynamic year columns from [Date]
    │   ├── Write-up Sent
    │   │   └── SUM([Total]) where [Comment] contains MEMO
    │   ├── No Dispute_Amount column
    │   └── Do not show an account section with Total Amount = 0.00
    │
    ├── Invoice's Status
    │   ├── Source: [Ref Type] = IN
    │   ├── Group: [Acct] + YEAR([Date])
    │   ├── Value: SUM([Total])
    │   ├── Dynamic year columns from [Date]
    │   ├── Total Amount: SUM([Total]) for all years
    │   ├── Paid @latest-year
    │   │   ├── Latest year = maximum YEAR([Date])
    │   │   ├── Source: latest-year IN rows only
    │   │   ├── Group: [Acct] + [Comment] + YEAR([Date])
    │   │   └── SUM([Total]) where Comment contains:
    │   │       ├── Paid on
    │   │       ├── payment no / payemnt no / similar small typo
    │   │       └── payment-number digits
    │   └── Do not show an account row with Total Amount = 0.00
    │
    ├── CR's Status
    │   ├── Source: verified [Ref Type] = CR rows
    │   ├── Same source rows as <Acct>_CRs sheets
    │   ├── Group: [Acct] + YEAR([Date])
    │   ├── Value: SUM([Total])
    │   ├── Dynamic year columns and Total Amount
    │   ├── Status column: "-" placeholder
    │   └── Do not show an account row with Total Amount = 0.00
    │
    ├── Formatting rules
    │   ├── Status style template:
    │   │   companies/amazonaccounts/amazon_status_style_template.xlsx
    │   ├── DB uses approved DB style.
    │   ├── Invoice and CR use approved Invoice style.
    │   ├── No Walmart values are copied into Amazon workbooks.
    │   └── Remove unused template rows / formatting after final visible data.
    │
    └── Amazon download
        ├── Output: AMAZON_ACCOUNT_TYPE_SHEETS.xlsx
        ├── Browser filename:
        │   Amazon Status Summary YYYY-MM-DD HH-MM-SS.xlsx
        ├── Native browser download is retained.
        └── Button stays "Processing..." until server file sending completes.
```

## Important boundaries

```text
AMAZON
├── File-only processing
├── No SQL Server connection
├── No Walmart stored procedure execution
└── No changes to Walmart calculation rules

WALMART
├── SQL Server workflow
├── Existing stored procedures
└── Existing export behavior
```
