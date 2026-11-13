/*
Wayfair runtime schema draft
Source: inferred from app/main.py, app/workflow.py, app/DbConfigLogger.py,
and companies/wayfair/*.json on 2026-06-15.

This file is a draft reference schema for the active Wayfair runtime.
Some step tables are dynamic and pick up upload columns at runtime, so the
column lists below combine fixed repository columns with currently expected
Wayfair upload and derived columns.
*/

USE [wayfairDBc];
GO

/* ============================================================
   CORE REPOSITORY TABLES
============================================================ */

IF OBJECT_ID('dbo.upload_batches', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.upload_batches (
        id INT IDENTITY(1,1) PRIMARY KEY,
        created_at DATETIME2 NOT NULL DEFAULT SYSUTCDATETIME(),
        raw_filename NVARCHAR(MAX) NOT NULL,
        combine_file_count INT NOT NULL,
        status NVARCHAR(50) NOT NULL,
        notes NVARCHAR(MAX) NULL
    );
END;
GO

IF OBJECT_ID('dbo.step_01_raw_rows', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.step_01_raw_rows (
        id INT IDENTITY(1,1) PRIMARY KEY,
        batch_id INT NOT NULL,
        row_number INT NOT NULL,
        ref_value NVARCHAR(255) NULL,
        payload_json NVARCHAR(MAX) NOT NULL,
        created_at DATETIME2 NOT NULL DEFAULT SYSUTCDATETIME()
    );
END;
GO

IF OBJECT_ID('dbo.combine_rows', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.combine_rows (
        id INT IDENTITY(1,1) PRIMARY KEY,
        batch_id INT NOT NULL,
        source_kind NVARCHAR(100) NOT NULL DEFAULT 'legacy',
        source_file NVARCHAR(255) NOT NULL,
        row_number INT NOT NULL,
        payload_json NVARCHAR(MAX) NOT NULL,
        created_at DATETIME2 NOT NULL DEFAULT SYSUTCDATETIME()
    );
END;
GO

IF OBJECT_ID('dbo.generated_files', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.generated_files (
        id INT IDENTITY(1,1) PRIMARY KEY,
        batch_id INT NOT NULL,
        kind NVARCHAR(50) NOT NULL,
        file_name NVARCHAR(255) NOT NULL,
        file_path NVARCHAR(1000) NOT NULL,
        created_at DATETIME2 NOT NULL DEFAULT SYSUTCDATETIME()
    );
END;
GO

/* ============================================================
   STEP 01 DYNAMIC TABLES
============================================================ */

IF OBJECT_ID('dbo.step_01_inquiry_rows', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.step_01_inquiry_rows (
        id INT IDENTITY(1,1) PRIMARY KEY,
        batch_id INT NOT NULL,
        row_number INT NOT NULL,
        ref NVARCHAR(MAX) NULL,
        [order] NVARCHAR(MAX) NULL,
        claim NVARCHAR(MAX) NULL,
        [current] DECIMAL(18,2) NULL,
        col_1_30_days DECIMAL(18,2) NULL,
        col_31_60_days DECIMAL(18,2) NULL,
        col_61_90_days DECIMAL(18,2) NULL,
        over_90_days DECIMAL(18,2) NULL,
        Ndatee DATE NULL,
        pay_day DECIMAL(18,2) NULL,
        Cclaim VARCHAR(50) NULL
    );
END;
GO

/*
Step 01A tables are rebuilt from step_01_inquiry_rows with SELECT * INTO.
They inherit the current Step 01 inquiry shape.
*/

IF OBJECT_ID('dbo.step_01a_cr_sheet', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.step_01a_cr_sheet (
        id INT IDENTITY(1,1) PRIMARY KEY,
        batch_id INT NOT NULL,
        row_number INT NOT NULL,
        ref NVARCHAR(MAX) NULL,
        [order] NVARCHAR(MAX) NULL,
        claim NVARCHAR(MAX) NULL,
        [current] DECIMAL(18,2) NULL,
        col_1_30_days DECIMAL(18,2) NULL,
        col_31_60_days DECIMAL(18,2) NULL,
        col_61_90_days DECIMAL(18,2) NULL,
        over_90_days DECIMAL(18,2) NULL,
        Ndatee DATE NULL,
        pay_day DECIMAL(18,2) NULL,
        Cclaim VARCHAR(50) NULL
    );
END;
GO

IF OBJECT_ID('dbo.step_01a_in_sheet', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.step_01a_in_sheet (
        id INT IDENTITY(1,1) PRIMARY KEY,
        batch_id INT NOT NULL,
        row_number INT NOT NULL,
        ref NVARCHAR(MAX) NULL,
        [order] NVARCHAR(MAX) NULL,
        claim NVARCHAR(MAX) NULL,
        [current] DECIMAL(18,2) NULL,
        col_1_30_days DECIMAL(18,2) NULL,
        col_31_60_days DECIMAL(18,2) NULL,
        col_61_90_days DECIMAL(18,2) NULL,
        over_90_days DECIMAL(18,2) NULL,
        Ndatee DATE NULL,
        pay_day DECIMAL(18,2) NULL,
        Cclaim VARCHAR(50) NULL
    );
END;
GO

IF OBJECT_ID('dbo.step_01a_db_sheet', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.step_01a_db_sheet (
        id INT IDENTITY(1,1) PRIMARY KEY,
        batch_id INT NOT NULL,
        row_number INT NOT NULL,
        ref NVARCHAR(MAX) NULL,
        [order] NVARCHAR(MAX) NULL,
        claim NVARCHAR(MAX) NULL,
        [current] DECIMAL(18,2) NULL,
        col_1_30_days DECIMAL(18,2) NULL,
        col_31_60_days DECIMAL(18,2) NULL,
        col_61_90_days DECIMAL(18,2) NULL,
        over_90_days DECIMAL(18,2) NULL,
        Ndatee DATE NULL,
        pay_day DECIMAL(18,2) NULL,
        Cclaim VARCHAR(50) NULL
    );
END;
GO

/* ============================================================
   STEP 02 DYNAMIC TABLES
============================================================ */

IF OBJECT_ID('dbo.step_02_farooq', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.step_02_farooq (
        id INT IDENTITY(1,1) PRIMARY KEY,
        batch_id INT NOT NULL,
        row_number INT NOT NULL,
        source_file NVARCHAR(MAX) NULL,
        qty DECIMAL(18,2) NULL,
        unit_price DECIMAL(18,2) NULL,
        total DECIMAL(18,2) NULL
    );
END;
GO

IF OBJECT_ID('dbo.step_02_po_details', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.step_02_po_details (
        id INT IDENTITY(1,1) PRIMARY KEY,
        batch_id INT NOT NULL,
        row_number INT NOT NULL,
        source_file NVARCHAR(MAX) NULL,
        po_number NVARCHAR(MAX) NULL,
        quantiy DECIMAL(18,2) NULL,
        wholesale_price DECIMAL(18,2) NULL,
        ship_date NVARCHAR(MAX) NULL,
        cRetailOrder NVARCHAR(MAX) NULL,
        qxc DECIMAL(18,2) NULL,
        Odatee DATE NULL,
        invoice_number NVARCHAR(255) NULL,
        cr_issue NVARCHAR(255) NULL,
        db_number VARCHAR(50) NULL,
        inv_db VARCHAR(50) NULL,
        Po_Count VARCHAR(50) NULL,
        Total_Invoice DECIMAL(18,2) NULL,
        Total_Payment DECIMAL(18,2) NULL,
        diff DECIMAL(18,2) NULL,
        po_status VARCHAR(50) NULL
    );
END;
GO

/* ============================================================
   STEP 03 DYNAMIC TABLES
============================================================ */

IF OBJECT_ID('dbo.step_03_payments_combined', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.step_03_payments_combined (
        id INT IDENTITY(1,1) PRIMARY KEY,
        batch_id INT NOT NULL,
        row_number INT NOT NULL,
        source_file NVARCHAR(MAX) NULL,
        sofs_order NVARCHAR(MAX) NULL,
        quantity DECIMAL(18,2) NULL,
        unit_price DECIMAL(18,2) NULL,
        total DECIMAL(18,2) NULL,
        Sum_Total NVARCHAR(MAX) NULL,
        invoice_number NVARCHAR(255) NULL,
        cr_issue NVARCHAR(255) NULL,
        db_number VARCHAR(50) NULL,
        inv_db VARCHAR(50) NULL,
        Total_Payment DECIMAL(18,2) NULL
    );
END;
GO

/* ============================================================
   STEP 04 OUTPUT TABLES
============================================================ */

IF OBJECT_ID('dbo.step_04a_FullMatched', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.step_04a_FullMatched (
        id INT IDENTITY(1,1) PRIMARY KEY
    );
END;
GO

IF OBJECT_ID('dbo.step_04b_order_not_matched', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.step_04b_order_not_matched (
        id INT IDENTITY(1,1) PRIMARY KEY
    );
END;
GO

IF OBJECT_ID('dbo.step_04c_qty_not_matched', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.step_04c_qty_not_matched (
        id INT IDENTITY(1,1) PRIMARY KEY
    );
END;
GO

IF OBJECT_ID('dbo.step_04d_cost_not_matched', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.step_04d_cost_not_matched (
        id INT IDENTITY(1,1) PRIMARY KEY
    );
END;
GO

IF OBJECT_ID('dbo.step_04e_total_not_matched', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.step_04e_total_not_matched (
        id INT IDENTITY(1,1) PRIMARY KEY
    );
END;
GO

/* ============================================================
   STEP 05 / STEP 06 / STEP 07 OUTPUT TABLES
============================================================ */

IF OBJECT_ID('dbo.ndatee', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.ndatee (
        id INT IDENTITY(1,1) PRIMARY KEY
    );
END;
GO

IF OBJECT_ID('dbo.unapplied', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.unapplied (
        id INT IDENTITY(1,1) PRIMARY KEY
    );
END;
GO

IF OBJECT_ID('dbo.NO_CR_Number', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.NO_CR_Number (
        id INT IDENTITY(1,1) PRIMARY KEY
    );
END;
GO

IF OBJECT_ID('dbo.Deductions_Summary', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.Deductions_Summary (
        id INT IDENTITY(1,1) PRIMARY KEY
    );
END;
GO

IF OBJECT_ID('dbo.step_07_Credit_Applied_Sheet', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.step_07_Credit_Applied_Sheet (
        id INT IDENTITY(1,1) PRIMARY KEY
    );
END;
GO

IF OBJECT_ID('dbo.Step_09A_Deductions_BreakDown', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.Step_09A_Deductions_BreakDown (
        id INT IDENTITY(1,1) PRIMARY KEY
    );
END;
GO

IF OBJECT_ID('dbo.Credit_CrDedNet_Summary', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.Credit_CrDedNet_Summary (
        id INT IDENTITY(1,1) PRIMARY KEY
    );
END;
GO

IF OBJECT_ID('dbo.Po_Status_Sum_Summary', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.Po_Status_Sum_Summary (
        id INT IDENTITY(1,1) PRIMARY KEY
    );
END;
GO

/*
Notes:
- Step 04 through Step 07 output tables are rebuilt by SQL SELECT INTO / DROP + rebuild flows.
- Their exact column shapes depend on the current SQL logic in app/main.py.
- Use this file as a stable starting reference, not as a guaranteed full-fidelity migration.
*/
