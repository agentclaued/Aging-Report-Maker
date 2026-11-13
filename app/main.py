import logging
import os
import re
import shutil
import tempfile
from html import escape
from contextlib import suppress
from dataclasses import replace
from datetime import datetime
from decimal import Decimal
from pathlib import Path

import pandas as pd
from fastapi import FastAPI, File, Form, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Color, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.worksheet import Worksheet
from starlette.background import BackgroundTask

from app.adapters.company_loader import load_company_profile
from app.adapters.amazon_file_processor import (
    AMAZON_GROUPED_OUTPUT_FILE,
    AmazonAccountSheetProcessor,
)
from app.adapters.regular_accounts_processor import (
    RegularAccountsProcessor,
)
from app.adapters.wallmart_python_processor import (
    WALLMART_PYTHON_OUTPUT_FILE,
    WallMartPythonReportProcessor,
)
from app.workbook_metadata import prepare_pivot_source_sheet, set_workbook_author_metadata
from app.excel_pivot_builder import pivot_skipped_types, run_regular_accounts_pivot_process
from app.DbConfigLogger import (
    Logger,
    SQLServerRepository,
    STEP_01A_CR_TABLE,
    STEP_01A_DB_TABLE,
    STEP_01A_IN_TABLE,
    STEP_01_COMPLETED_STATUS,
    STEP_01_INQUIRY_TABLE,
    STEP_01_NORMALIZED_FILE_KIND,
    get_settings,
)
from app.edit_history import append_app_run_log, append_history_event
from app import progress as job_progress
from app.workflow import (
    InquiryProcessor,
    STEP_01_OUTPUT_FILE,
    STEP_01A_OUTPUT_FILE,
    WorkflowOrchestrator,
)

settings = get_settings()
templates = Jinja2Templates(directory=str(settings.templates_dir))

DEFAULT_COMPANY_CODE = settings.default_company_code
SUPPORTED_COMPANY_CODES = ("wallmartsummary", "amazonaccounts", "regularaccounts", "newageing")
FILE_ONLY_COMPANY_CODES = ("wallmartsummary", "amazonaccounts", "regularaccounts", "newageing")

app = FastAPI(title=settings.app_title)
app.mount("/static", StaticFiles(directory=str(settings.base_dir / "app" / "static")), name="static")


class _QuietProgressPolling(logging.Filter):
    """The page checks /progress/status about once a second; keep those lines out of the console log."""

    def filter(self, record: logging.LogRecord) -> bool:
        return "/progress/status" not in record.getMessage()


logging.getLogger("uvicorn.access").addFilter(_QuietProgressPolling())
amazon_download_jobs: dict[str, dict[str, object]] = {}





@app.on_event("startup")
async def startup_history_log() -> None:
    append_app_run_log(
        "Automatic startup snapshot refresh",
        [
            "App startup completed.",
            "Startup snapshot logging was written to app run logs.txt instead of ai edit history.txt.",
        ],
    )


STEP_01_SQL_CR_IN_DB_GEN = f"""
                            EXEC dbo.sp_WallMartSummary_db_and_inv_tabels;
"""

RUN_WALLMART_SUMMARY_SQL = """
                            EXEC dbo.sp_Run_WallMartSummary;
"""

STEP_01A_INV_DB_CONTROL_TABLE = "Inv_DB_Control_Sheet"
STEP_01A_CREDIT_CONTROL_TABLE = "Credit_Control_Sheet"

STEP_04_FULL_MATCHED_TABLE = "step_04a_FullMatched"
STEP_04_ORDER_NOT_MATCHED_TABLE = "step_04b_order_not_matched"
STEP_04_QTY_NOT_MATCHED_TABLE = "step_04c_qty_not_matched"
STEP_04_COST_NOT_MATCHED_TABLE = "step_04d_cost_not_matched"
STEP_04_TOTAL_NOT_MATCHED_TABLE = "step_04e_total_not_matched"
STEP_04_TABLES = (
    STEP_04_FULL_MATCHED_TABLE,
    STEP_04_ORDER_NOT_MATCHED_TABLE,
    STEP_04_QTY_NOT_MATCHED_TABLE,
    STEP_04_COST_NOT_MATCHED_TABLE,
    STEP_04_TOTAL_NOT_MATCHED_TABLE,
)
LATEST_EXPORT_TABLES = (
    STEP_01_INQUIRY_TABLE,
    STEP_01A_CR_TABLE,
    STEP_01A_IN_TABLE,
    STEP_01A_DB_TABLE,
    "step_02_farooq",
    "step_02_po_details",
    "step_03_payments_combined",
    STEP_04_FULL_MATCHED_TABLE,
    STEP_04_ORDER_NOT_MATCHED_TABLE,
    STEP_04_QTY_NOT_MATCHED_TABLE,
    STEP_04_COST_NOT_MATCHED_TABLE,
    STEP_04_TOTAL_NOT_MATCHED_TABLE,
)
AUTOMATED_UPLOAD_REPORT_TABLES = (
    "DBs_Status",
    "Invoices_Status",
    "SA9200_DBs",
    "SA9200_Invoices",
    "WA9085_DBs",
    "WA9085_Invoices",
    "WallMartSummary_Sheet",
    "WC9085_DBs",
    "WC9085_Invoices",
    "WM9263_DBs",
    "WM9263_Invoices",
)
UPLOAD_DATA_TYPES_SQL = """
IF OBJECT_ID('dbo.step_02_po_details', 'U') IS NOT NULL
BEGIN
    -- Skipped to avoid date conversion errors from source text values.
    -- ALTER TABLE dbo.step_02_po_details ALTER COLUMN order_date DATE;
    -- ALTER TABLE dbo.step_02_po_details ALTER COLUMN rec_stamp DATETIME2;
    -- ALTER TABLE dbo.step_02_po_details ALTER COLUMN feed_date DATE;
    -- ALTER TABLE dbo.step_02_po_details ALTER COLUMN ship_date DATE;
    IF COL_LENGTH('dbo.step_02_po_details', 'retail_order') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_02_po_details ALTER COLUMN retail_order VARCHAR(100);
        RAISERROR('Altered column: step_02_po_details.retail_order', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing column: step_02_po_details.retail_order', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_02_po_details', 'quantity') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_02_po_details ALTER COLUMN quantity DECIMAL(18,2);
        RAISERROR('Altered column: step_02_po_details.quantity', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing column: step_02_po_details.quantity', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_02_po_details', 'cost') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_02_po_details ALTER COLUMN cost DECIMAL(18,2);
        RAISERROR('Altered column: step_02_po_details.cost', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing column: step_02_po_details.cost', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_02_po_details', 'ship_date') IS NOT NULL
        RAISERROR('Skipped source-text date column: step_02_po_details.ship_date', 10, 1) WITH NOWAIT;
    ELSE
        RAISERROR('Skipped missing column: step_02_po_details.ship_date', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_02_po_details', 'order_date') IS NOT NULL
        RAISERROR('Skipped source-text date column: step_02_po_details.order_date', 10, 1) WITH NOWAIT;
    ELSE
        RAISERROR('Skipped missing column: step_02_po_details.order_date', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_02_po_details', 'rec_stamp') IS NOT NULL
        RAISERROR('Skipped source-text date column: step_02_po_details.rec_stamp', 10, 1) WITH NOWAIT;
    ELSE
        RAISERROR('Skipped missing column: step_02_po_details.rec_stamp', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_02_po_details', 'feed_date') IS NOT NULL
        RAISERROR('Skipped source-text date column: step_02_po_details.feed_date', 10, 1) WITH NOWAIT;
    ELSE
        RAISERROR('Skipped missing column: step_02_po_details.feed_date', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_02_po_details', 'qxc') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_02_po_details ALTER COLUMN qxc DECIMAL(18,2);
        RAISERROR('Altered column: step_02_po_details.qxc', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing column: step_02_po_details.qxc', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_02_po_details', 'retail_price') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_02_po_details ALTER COLUMN retail_price DECIMAL(18,2);
        RAISERROR('Altered column: step_02_po_details.retail_price', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing column: step_02_po_details.retail_price', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_02_po_details', 'order_total') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_02_po_details ALTER COLUMN order_total DECIMAL(18,2);
        RAISERROR('Altered column: step_02_po_details.order_total', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing column: step_02_po_details.order_total', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_02_po_details', 'order_total_invoiced') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_02_po_details ALTER COLUMN order_total_invoiced DECIMAL(18,2);
        RAISERROR('Altered column: step_02_po_details.order_total_invoiced', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing column: step_02_po_details.order_total_invoiced', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_02_po_details', 'hsn_drop_ship_fee') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_02_po_details ALTER COLUMN hsn_drop_ship_fee DECIMAL(18,2);
        RAISERROR('Altered column: step_02_po_details.hsn_drop_ship_fee', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing column: step_02_po_details.hsn_drop_ship_fee', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_02_po_details', 'manual_cost') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_02_po_details ALTER COLUMN manual_cost DECIMAL(18,2);
        RAISERROR('Altered column: step_02_po_details.manual_cost', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing column: step_02_po_details.manual_cost', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_02_po_details', 'invoice_db_number') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_02_po_details ALTER COLUMN invoice_db_number NVARCHAR(255);
        RAISERROR('Altered column: step_02_po_details.invoice_db_number', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing column: step_02_po_details.invoice_db_number', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_02_po_details', 'cr_issue') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_02_po_details ALTER COLUMN cr_issue NVARCHAR(255);
        RAISERROR('Altered column: step_02_po_details.cr_issue', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing column: step_02_po_details.cr_issue', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_02_po_details', 'Odatee') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_02_po_details ALTER COLUMN Odatee DATE;
        RAISERROR('Altered column: step_02_po_details.Odatee', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing column: step_02_po_details.Odatee', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_02_po_details', 'Po_Count') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_02_po_details ALTER COLUMN Po_Count VARCHAR(50);
        RAISERROR('Altered column: step_02_po_details.Po_Count', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing column: step_02_po_details.Po_Count', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_02_po_details', 'Total_Invoice') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_02_po_details ALTER COLUMN Total_Invoice DECIMAL(18,2);
        RAISERROR('Altered column: step_02_po_details.Total_Invoice', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing column: step_02_po_details.Total_Invoice', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_02_po_details', 'Total_Payment') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_02_po_details ALTER COLUMN Total_Payment DECIMAL(18,2);
        RAISERROR('Altered column: step_02_po_details.Total_Payment', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing column: step_02_po_details.Total_Payment', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_02_po_details', 'diff') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_02_po_details ALTER COLUMN diff DECIMAL(18,2);
        RAISERROR('Altered column: step_02_po_details.diff', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing column: step_02_po_details.diff', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_02_po_details', 'po_status') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_02_po_details ALTER COLUMN po_status VARCHAR(50);
        RAISERROR('Altered column: step_02_po_details.po_status', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing column: step_02_po_details.po_status', 10, 1) WITH NOWAIT;
END;

IF OBJECT_ID('dbo.step_03_payments_combined', 'U') IS NOT NULL
BEGIN
    IF COL_LENGTH('dbo.step_03_payments_combined', 'total') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_03_payments_combined DROP COLUMN total;
        RAISERROR('Dropped stale column: step_03_payments_combined.total', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing stale column: step_03_payments_combined.total', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_03_payments_combined', 'quantity') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_03_payments_combined DROP COLUMN quantity;
        RAISERROR('Dropped stale column: step_03_payments_combined.quantity', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing stale column: step_03_payments_combined.quantity', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_03_payments_combined', 'unit_price') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_03_payments_combined DROP COLUMN unit_price;
        RAISERROR('Dropped stale column: step_03_payments_combined.unit_price', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing stale column: step_03_payments_combined.unit_price', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_03_payments_combined', 'db_number') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_03_payments_combined DROP COLUMN db_number;
        RAISERROR('Dropped stale column: step_03_payments_combined.db_number', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing stale column: step_03_payments_combined.db_number', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_03_payments_combined', 'inv_db') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_03_payments_combined DROP COLUMN inv_db;
        RAISERROR('Dropped stale column: step_03_payments_combined.inv_db', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing stale column: step_03_payments_combined.inv_db', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_03_payments_combined', 'payment_amount') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_03_payments_combined DROP COLUMN payment_amount;
        RAISERROR('Dropped stale column: step_03_payments_combined.payment_amount', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing stale column: step_03_payments_combined.payment_amount', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_03_payments_combined', 'Total_Payment') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_03_payments_combined DROP COLUMN Total_Payment;
        RAISERROR('Dropped stale column: step_03_payments_combined.Total_Payment', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing stale column: step_03_payments_combined.Total_Payment', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_03_payments_combined', 'description') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_03_payments_combined DROP COLUMN description;
        RAISERROR('Dropped stale column: step_03_payments_combined.description', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing stale column: step_03_payments_combined.description', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_03_payments_combined', 'invoice_amount') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_03_payments_combined DROP COLUMN invoice_amount;
        RAISERROR('Dropped stale column: step_03_payments_combined.invoice_amount', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing stale column: step_03_payments_combined.invoice_amount', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_03_payments_combined', 'invoice_currency') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_03_payments_combined DROP COLUMN invoice_currency;
        RAISERROR('Dropped stale column: step_03_payments_combined.invoice_currency', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing stale column: step_03_payments_combined.invoice_currency', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_03_payments_combined', 'withholding_amount') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_03_payments_combined DROP COLUMN withholding_amount;
        RAISERROR('Dropped stale column: step_03_payments_combined.withholding_amount', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing stale column: step_03_payments_combined.withholding_amount', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_03_payments_combined', 'terms_discount_taken') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_03_payments_combined DROP COLUMN terms_discount_taken;
        RAISERROR('Dropped stale column: step_03_payments_combined.terms_discount_taken', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing stale column: step_03_payments_combined.terms_discount_taken', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_03_payments_combined', 'amount_paid') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_03_payments_combined DROP COLUMN amount_paid;
        RAISERROR('Dropped stale column: step_03_payments_combined.amount_paid', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing stale column: step_03_payments_combined.amount_paid', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_03_payments_combined', 'remaining_amount') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_03_payments_combined DROP COLUMN remaining_amount;
        RAISERROR('Dropped stale column: step_03_payments_combined.remaining_amount', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing stale column: step_03_payments_combined.remaining_amount', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_03_payments_combined', 'order_date') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_03_payments_combined DROP COLUMN order_date;
        RAISERROR('Dropped stale column: step_03_payments_combined.order_date', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing stale column: step_03_payments_combined.order_date', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_03_payments_combined', 'invoice_number') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_03_payments_combined ALTER COLUMN invoice_number VARCHAR(50);
        RAISERROR('Altered column: step_03_payments_combined.invoice_number', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing column: step_03_payments_combined.invoice_number', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_03_payments_combined', 'keyrec_number') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_03_payments_combined ALTER COLUMN keyrec_number VARCHAR(50);
        RAISERROR('Altered column: step_03_payments_combined.keyrec_number', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing column: step_03_payments_combined.keyrec_number', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_03_payments_combined', 'doc_type') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_03_payments_combined ALTER COLUMN doc_type VARCHAR(20);
        RAISERROR('Altered column: step_03_payments_combined.doc_type', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing column: step_03_payments_combined.doc_type', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_03_payments_combined', 'transaction_value') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_03_payments_combined ALTER COLUMN transaction_value DECIMAL(18,2);
        RAISERROR('Altered column: step_03_payments_combined.transaction_value', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing column: step_03_payments_combined.transaction_value', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_03_payments_combined', 'cash_discount_amount') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_03_payments_combined ALTER COLUMN cash_discount_amount DECIMAL(18,2);
        RAISERROR('Altered column: step_03_payments_combined.cash_discount_amount', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing column: step_03_payments_combined.cash_discount_amount', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_03_payments_combined', 'clearing_document_number') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_03_payments_combined ALTER COLUMN clearing_document_number VARCHAR(50);
        RAISERROR('Altered column: step_03_payments_combined.clearing_document_number', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing column: step_03_payments_combined.clearing_document_number', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_03_payments_combined', 'payment_chargeback_date') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_03_payments_combined ALTER COLUMN payment_chargeback_date DATE;
        RAISERROR('Altered column: step_03_payments_combined.payment_chargeback_date', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing column: step_03_payments_combined.payment_chargeback_date', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_03_payments_combined', 'comments') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_03_payments_combined ALTER COLUMN comments NVARCHAR(500);
        RAISERROR('Altered column: step_03_payments_combined.comments', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing column: step_03_payments_combined.comments', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_03_payments_combined', 'reason_code') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_03_payments_combined ALTER COLUMN reason_code VARCHAR(50);
        RAISERROR('Altered column: step_03_payments_combined.reason_code', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing column: step_03_payments_combined.reason_code', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_03_payments_combined', 'sap_company_code') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_03_payments_combined ALTER COLUMN sap_company_code VARCHAR(20);
        RAISERROR('Altered column: step_03_payments_combined.sap_company_code', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing column: step_03_payments_combined.sap_company_code', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_03_payments_combined', 'po_number') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_03_payments_combined ALTER COLUMN po_number VARCHAR(100);
        RAISERROR('Altered column: step_03_payments_combined.po_number', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing column: step_03_payments_combined.po_number', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_03_payments_combined', 'reference_check_number') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_03_payments_combined ALTER COLUMN reference_check_number VARCHAR(100);
        RAISERROR('Altered column: step_03_payments_combined.reference_check_number', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing column: step_03_payments_combined.reference_check_number', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_03_payments_combined', 'invoice_date') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_03_payments_combined ALTER COLUMN invoice_date DATE;
        RAISERROR('Altered column: step_03_payments_combined.invoice_date', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing column: step_03_payments_combined.invoice_date', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_03_payments_combined', 'posting_date') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_03_payments_combined ALTER COLUMN posting_date DATE;
        RAISERROR('Altered column: step_03_payments_combined.posting_date', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing column: step_03_payments_combined.posting_date', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_03_payments_combined', 'payment_number') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_03_payments_combined ALTER COLUMN payment_number VARCHAR(50);
        RAISERROR('Altered column: step_03_payments_combined.payment_number', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing column: step_03_payments_combined.payment_number', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_03_payments_combined', 'source_file') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_03_payments_combined ALTER COLUMN source_file NVARCHAR(255);
        RAISERROR('Altered column: step_03_payments_combined.source_file', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing column: step_03_payments_combined.source_file', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_03_payments_combined', 'sum_total') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_03_payments_combined ALTER COLUMN sum_total DECIMAL(18,2);
        RAISERROR('Altered column: step_03_payments_combined.sum_total', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing column: step_03_payments_combined.sum_total', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_03_payments_combined', 'invoice_db_number') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_03_payments_combined ALTER COLUMN invoice_db_number VARCHAR(100);
        RAISERROR('Altered column: step_03_payments_combined.invoice_db_number', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing column: step_03_payments_combined.invoice_db_number', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_03_payments_combined', 'cr_issue') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_03_payments_combined ALTER COLUMN cr_issue NVARCHAR(255);
        RAISERROR('Altered column: step_03_payments_combined.cr_issue', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing column: step_03_payments_combined.cr_issue', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_03_payments_combined', 'total_payment') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_03_payments_combined ALTER COLUMN total_payment DECIMAL(18,2);
        RAISERROR('Altered column: step_03_payments_combined.total_payment', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing column: step_03_payments_combined.total_payment', 10, 1) WITH NOWAIT;
END;

IF OBJECT_ID('dbo.step_02_farooq', 'U') IS NOT NULL
BEGIN
    -- Skipped to avoid date conversion errors from source text values.
    -- ALTER TABLE dbo.step_02_farooq ALTER COLUMN [date] DATE;
    IF COL_LENGTH('dbo.step_02_farooq', 'qty') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_02_farooq ALTER COLUMN qty DECIMAL(18,2);
        RAISERROR('Altered column: step_02_farooq.qty', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing column: step_02_farooq.qty', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_02_farooq', 'unit_price') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_02_farooq ALTER COLUMN unit_price DECIMAL(18,2);
        RAISERROR('Altered column: step_02_farooq.unit_price', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing column: step_02_farooq.unit_price', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_02_farooq', 'total') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_02_farooq ALTER COLUMN total DECIMAL(18,2);
        RAISERROR('Altered column: step_02_farooq.total', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing column: step_02_farooq.total', 10, 1) WITH NOWAIT;
END;

IF OBJECT_ID('dbo.step_01_inquiry_rows', 'U') IS NOT NULL
BEGIN
    -- Skipped to avoid date conversion errors from source text values.
    -- ALTER TABLE dbo.step_01_inquiry_rows ALTER COLUMN [date] DATE;
    -- ALTER TABLE dbo.step_01_inquiry_rows ALTER COLUMN due_date DATE;
    IF COL_LENGTH('dbo.step_01_inquiry_rows', 'Cclaim') IS NULL
    BEGIN
        ALTER TABLE dbo.step_01_inquiry_rows
        ADD Cclaim VARCHAR(50);
        RAISERROR('Added column: step_01_inquiry_rows.Cclaim', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped existing column: step_01_inquiry_rows.Cclaim', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_01_inquiry_rows', 'current') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_01_inquiry_rows ALTER COLUMN [current] DECIMAL(18,2);
        RAISERROR('Altered column: step_01_inquiry_rows.current', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing column: step_01_inquiry_rows.current', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_01_inquiry_rows', 'col_1_30_days') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_01_inquiry_rows ALTER COLUMN col_1_30_days DECIMAL(18,2);
        RAISERROR('Altered column: step_01_inquiry_rows.col_1_30_days', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing column: step_01_inquiry_rows.col_1_30_days', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_01_inquiry_rows', 'col_31_60_days') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_01_inquiry_rows ALTER COLUMN col_31_60_days DECIMAL(18,2);
        RAISERROR('Altered column: step_01_inquiry_rows.col_31_60_days', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing column: step_01_inquiry_rows.col_31_60_days', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_01_inquiry_rows', 'col_61_90_days') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_01_inquiry_rows ALTER COLUMN col_61_90_days DECIMAL(18,2);
        RAISERROR('Altered column: step_01_inquiry_rows.col_61_90_days', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing column: step_01_inquiry_rows.col_61_90_days', 10, 1) WITH NOWAIT;

    IF COL_LENGTH('dbo.step_01_inquiry_rows', 'over_90_days') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.step_01_inquiry_rows ALTER COLUMN over_90_days DECIMAL(18,2);
        RAISERROR('Altered column: step_01_inquiry_rows.over_90_days', 10, 1) WITH NOWAIT;
    END
    ELSE
        RAISERROR('Skipped missing column: step_01_inquiry_rows.over_90_days', 10, 1) WITH NOWAIT;

    UPDATE dbo.step_01_inquiry_rows
    SET Cclaim =
        CASE
            WHEN LTRIM(RTRIM(claim)) LIKE 'DB%'
                OR LTRIM(RTRIM(claim)) LIKE 'CR%'
            THEN RIGHT(LTRIM(RTRIM(claim)), 6)
            ELSE LTRIM(RTRIM(claim))
        END;
    RAISERROR('Updated column: step_01_inquiry_rows.Cclaim', 10, 1) WITH NOWAIT;

    -- ALTER TABLE dbo.step_01_inquiry_rows ALTER COLUMN claim DECIMAL(18,2);
END;
"""
LATEST_TABLE_EXPORT_FILE = "LATEST_SQL_TABLES.xlsx"


STEP_07_HEADER_FILL = PatternFill(fill_type="solid", fgColor="B4C6E7")
STEP_07_DB_FILL = PatternFill(fill_type="solid", fgColor="FFC000")
STEP_07_BLANK_FILL = PatternFill(fill_type="solid", fgColor="F7F7F7")
STEP_07_WHITE_FILL = PatternFill(fill_type="solid", fgColor="FFFFFF")

STEP_04_SQL = """



-- 1. Rename column only if needed
IF EXISTS (
    SELECT 1 
    FROM sys.columns 
    WHERE name = 'quantity' 
      AND object_id = OBJECT_ID('STEP_02_PO_DETAILS')
)
AND NOT EXISTS (
    SELECT 1 
    FROM sys.columns 
    WHERE name = 'quantiy' 
      AND object_id = OBJECT_ID('STEP_02_PO_DETAILS')
)
BEGIN
    EXEC sp_rename 'STEP_02_PO_DETAILS.quantity', 'quantiy', 'COLUMN';
END;

-- 2. Add column only if it does not exist
IF NOT EXISTS (
    SELECT 1 
    FROM sys.columns 
    WHERE name = 'qxc' 
      AND object_id = OBJECT_ID('STEP_02_PO_DETAILS')
)
BEGIN
    ALTER TABLE STEP_02_PO_DETAILS
    ADD qxc DECIMAL(18,2);
END;

-- 3. Update safely (can run anytime)
UPDATE STEP_02_PO_DETAILS
SET qxc = 
    ROUND(
        TRY_CAST(quantiy AS DECIMAL(18,2)) *
        TRY_CAST(wholesale_price AS DECIMAL(18,2)),
        2
    );


--######################@@@@@@@@@@@@@@@@@222222222222222222222222


/* ---------- 04a  FULL MATCH -------------------------------------- */
DROP TABLE IF EXISTS step_04a_FullMatched;

SELECT  po.po_number,
        pay.po_number,
        po.quantiy       AS po_qty,
        pay.quantity     AS pay_qty,
        po.wholesale_price,
        pay.unit_price,
        po.qxc,
        pay.payment_amount,
        po.source_file   AS po_source_file,     -- audit: PO file
        pay.source_file  AS pay_source_file     -- audit: Payment file
INTO    step_04a_FullMatched
FROM    step_02_po_details        AS po
JOIN    step_03_payments_combined AS pay
        ON  po.po_number = pay.po_number
WHERE   po.quantiy = pay.quantity
    AND po.wholesale_price = pay.unit_price
    AND po.qxc     = pay.payment_amount;


/* ---------- 04b  ORDER NOT MATCHED ------------------------------- */
DROP TABLE IF EXISTS dbo.step_04b_order_not_matched;

SELECT  
    po.po_number,
    NULL AS po_number,
    po.quantiy AS qty,
    po.wholesale_price AS cost_or_unit_price,
    po.qxc AS qxc,
    po.invoice_db_number AS invoice_number,
    NULL AS db_number,
    NULL AS cr_issue,
    NULL AS payment_amount,
    po.source_file,
    'PO only (no payment)' AS source
INTO dbo.step_04b_order_not_matched
FROM dbo.step_02_po_details po
WHERE NOT EXISTS
(
    SELECT 1
    FROM dbo.step_03_payments_combined pay
    WHERE LTRIM(RTRIM(pay.po_number)) = LTRIM(RTRIM(po.po_number))
)

UNION ALL

SELECT  
    NULL AS po_number,
    pay.po_number,
    pay.quantity AS qty,
    pay.unit_price AS cost_or_unit_price,
    NULL AS qxc,
    NULL AS invoice_number,
    NULL AS db_number,
    pay.cr_issue,
    pay.payment_amount,
    pay.source_file,
    'Payment only (no PO)' AS source
FROM dbo.step_03_payments_combined pay
WHERE NOT EXISTS
(
    SELECT 1
    FROM dbo.step_02_po_details po
    WHERE LTRIM(RTRIM(po.po_number)) = LTRIM(RTRIM(pay.po_number))
);


ALTER TABLE dbo.step_04b_order_not_matched
ADD sumByInv_unpaid DECIMAL(18,2),
    sumByCr_unapplied DECIMAL(18,2);


/* sumByInv_unpaid = SUM(qxc) BY invoice_number    for 07 final row */
UPDATE t
SET t.sumByInv_unpaid = x.total_unpaid
FROM dbo.step_04b_order_not_matched t
INNER JOIN
(
    SELECT
        invoice_number,
        SUM(TRY_CAST(qxc AS DECIMAL(18,2))) AS total_unpaid
    FROM dbo.step_04b_order_not_matched
    WHERE invoice_number IS NOT NULL
    GROUP BY invoice_number
) x
    ON LTRIM(RTRIM(t.invoice_number)) = LTRIM(RTRIM(x.invoice_number));


/* sumByCr_unapplied = SUM(payment_amount) BY cr_issue     for 07 final rows */
UPDATE t
SET t.sumByCr_unapplied = x.total_unapplied
FROM dbo.step_04b_order_not_matched t
INNER JOIN
(
    SELECT
        cr_issue,
        SUM(TRY_CAST(payment_amount AS DECIMAL(18,2))) AS total_unapplied
    FROM dbo.step_04b_order_not_matched
    WHERE cr_issue IS NOT NULL
    GROUP BY cr_issue
) x
    ON LTRIM(RTRIM(t.cr_issue)) = LTRIM(RTRIM(x.cr_issue));













/* ---------- 04c  QTY MISMATCH ------------------------------------ */
DROP TABLE IF EXISTS step_04c_qty_not_matched;

SELECT  po.po_number,
        pay.po_number,
        TRY_CAST(po.quantiy  AS INT) AS po_qty,
        TRY_CAST(pay.quantity AS INT) AS pay_qty,
        TRY_CAST(po.quantiy  AS INT)
      - TRY_CAST(pay.quantity AS INT) AS qty_variance,
        po.wholesale_price,
        pay.unit_price,
        po.qxc,
        pay.payment_amount,
        po.source_file   AS po_source_file,        -- audit: PO file
        pay.source_file  AS pay_source_file        -- audit: Payment file
INTO    step_04c_qty_not_matched
FROM    step_02_po_details        AS po
JOIN    step_03_payments_combined AS pay
        ON  po.po_number = pay.po_number
WHERE   po.quantiy < > pay.quantity ;


/* ---------- 04d  COST vs UNIT_PRICE MISMATCH --------------------- */
DROP TABLE IF EXISTS step_04d_cost_not_matched;

SELECT  po.po_number,
        pay.po_number,
        TRY_CAST(po.quantiy   AS INT)           AS po_qty,
        TRY_CAST(pay.quantity AS INT)           AS pay_qty,
        TRY_CAST(po.wholesale_price AS DECIMAL(18,4)) AS po_cost,
        TRY_CAST(pay.unit_price AS DECIMAL(18,4)) AS pay_unit_price,
        TRY_CAST(po.wholesale_price AS DECIMAL(18,4))
      - TRY_CAST(pay.unit_price AS DECIMAL(18,4)) AS unit_variance,
        po.qxc,
        pay.payment_amount,
        po.source_file   AS po_source_file,        -- audit: PO file
        pay.source_file  AS pay_source_file        -- audit: Payment file
INTO    step_04d_cost_not_matched
FROM    step_02_po_details        AS po
JOIN    step_03_payments_combined AS pay
        ON  po.po_number = pay.po_number
WHERE   TRY_CAST(po.quantiy  AS INT)
      = TRY_CAST(pay.quantity AS INT)
    AND po.wholesale_price < > pay.unit_price ;


/* ---------- 04e  QXC vs TOTAL MISMATCH --------------------------- */
DROP TABLE IF EXISTS step_04e_total_not_matched;

SELECT  po.po_number,
        pay.po_number,
        TRY_CAST(po.quantiy   AS INT)            AS po_qty,
        TRY_CAST(pay.quantity AS INT)            AS pay_qty,
        TRY_CAST(po.wholesale_price AS DECIMAL(18,4)) AS po_cost,
        TRY_CAST(pay.unit_price AS DECIMAL(18,4)) AS pay_unit_price,
        TRY_CAST(po.qxc         AS DECIMAL(18,2)) AS po_qxc,
        TRY_CAST(pay.payment_amount AS DECIMAL(18,2)) AS pay_total,
        TRY_CAST(po.qxc         AS DECIMAL(18,2))
      - TRY_CAST(pay.payment_amount AS DECIMAL(18,2)) AS total_variance,
        po.source_file   AS po_source_file,        -- audit: PO file
        pay.source_file  AS pay_source_file        -- audit: Payment file
INTO    step_04e_total_not_matched
FROM    step_02_po_details        AS po
JOIN    step_03_payments_combined AS pay
        ON  po.po_number = pay.po_number
WHERE   TRY_CAST(po.quantiy  AS INT)
      = TRY_CAST(pay.quantity AS INT)
    AND TRY_CAST(po.wholesale_price AS DECIMAL(18,4))
      = TRY_CAST(pay.unit_price AS DECIMAL(18,4))
    AND TRY_CAST(po.qxc         AS DECIMAL(18,2))
     < > TRY_CAST(pay.payment_amount AS DECIMAL(18,2));


/* ---------- Summary ---------------------------------------------- */
SELECT 'step_04a_FullMatched'          AS bucket, COUNT(*) AS row_count FROM step_04a_FullMatched
UNION ALL SELECT 'step_04b_order_not_matched', COUNT(*) FROM step_04b_order_not_matched
UNION ALL SELECT 'step_04c_qty_not_matched',   COUNT(*) FROM step_04c_qty_not_matched
UNION ALL SELECT 'step_04d_cost_not_matched',  COUNT(*) FROM step_04d_cost_not_matched
UNION ALL SELECT 'step_04e_total_not_matched', COUNT(*) FROM step_04e_total_not_matched;







"""

STEP_04A_PO_STATUS_SQL = """
/* STEP 04A PO status refresh.
   Updates PO count, invoice/payment totals, diff, and po_status in existing upload tables. */

IF OBJECT_ID('dbo.step_02_po_details', 'U') IS NULL
BEGIN
    RAISERROR('Missing table: dbo.step_02_po_details', 16, 1);
    RETURN;
END;

IF OBJECT_ID('dbo.step_03_payments_combined', 'U') IS NULL
BEGIN
    RAISERROR('Missing table: dbo.step_03_payments_combined', 16, 1);
    RETURN;
END;

IF COL_LENGTH('dbo.step_02_po_details', 'Po_Count') IS NULL
BEGIN
    ALTER TABLE dbo.step_02_po_details
    ADD Po_Count VARCHAR(50) NULL;
END;

;WITH po_count AS (
    SELECT
        LTRIM(RTRIM(po_number)) AS po_number,
        COUNT(*) AS Po_Count
    FROM dbo.step_02_po_details
    WHERE po_number IS NOT NULL
    GROUP BY LTRIM(RTRIM(po_number))
)
UPDATE p
SET p.Po_Count = CAST(c.Po_Count AS VARCHAR(50))
FROM dbo.step_02_po_details AS p
INNER JOIN po_count AS c
    ON LTRIM(RTRIM(p.po_number)) = c.po_number;

RAISERROR('Step 04A Po_status: Po_Count updated in step_02_po_details.', 10, 1) WITH NOWAIT;

IF COL_LENGTH('dbo.step_02_po_details', 'Total_Invoice') IS NULL
BEGIN
    ALTER TABLE dbo.step_02_po_details
    ADD Total_Invoice DECIMAL(18,2) NULL;
END;

;WITH po_total AS (
    SELECT
        LTRIM(RTRIM(po_number)) AS po_number,
        SUM(ISNULL(TRY_CAST(qxc AS DECIMAL(18,2)), 0)) AS Total_Invoice
    FROM dbo.step_02_po_details
    WHERE po_number IS NOT NULL
    GROUP BY LTRIM(RTRIM(po_number))
)
UPDATE p
SET p.Total_Invoice = t.Total_Invoice
FROM dbo.step_02_po_details AS p
INNER JOIN po_total AS t
    ON LTRIM(RTRIM(p.po_number)) = t.po_number;

RAISERROR('Step 04A Po_status: Total_Invoice updated in step_02_po_details.', 10, 1) WITH NOWAIT;

IF COL_LENGTH('dbo.step_03_payments_combined', 'Total_Payment') IS NULL
BEGIN
    ALTER TABLE dbo.step_03_payments_combined
    ADD Total_Payment DECIMAL(18,2) NULL;
END;

IF COL_LENGTH('dbo.step_02_po_details', 'Total_Payment') IS NULL
BEGIN
    ALTER TABLE dbo.step_02_po_details
    ADD Total_Payment DECIMAL(18,2) NULL;
END;

;WITH pay_total AS (
    SELECT
        LTRIM(RTRIM(po_number)) AS po_number,
        SUM(ISNULL(TRY_CAST(payment_amount AS DECIMAL(18,2)), 0)) AS Total_Payment
    FROM dbo.step_03_payments_combined
    WHERE po_number IS NOT NULL
      AND TRY_CAST(payment_amount AS DECIMAL(18,2)) > 0
    GROUP BY LTRIM(RTRIM(po_number))
)
UPDATE p
SET p.Total_Payment = t.Total_Payment
FROM dbo.step_03_payments_combined AS p
INNER JOIN pay_total AS t
    ON LTRIM(RTRIM(p.po_number)) = t.po_number
WHERE p.po_number IS NOT NULL;

;WITH pay AS (
    SELECT
        LTRIM(RTRIM(po_number)) AS po_number,
        MAX(TRY_CAST(Total_Payment AS DECIMAL(18,2))) AS Total_Payment
    FROM dbo.step_03_payments_combined
    WHERE po_number IS NOT NULL
    GROUP BY LTRIM(RTRIM(po_number))
)
UPDATE p
SET p.Total_Payment = pay.Total_Payment
FROM dbo.step_02_po_details AS p
INNER JOIN pay
    ON LTRIM(RTRIM(p.po_number)) = pay.po_number;

RAISERROR('Step 04A Po_status: Total_Payment updated in payment and PO tables.', 10, 1) WITH NOWAIT;

IF COL_LENGTH('dbo.step_02_po_details', 'diff') IS NOT NULL
BEGIN
    ALTER TABLE dbo.step_02_po_details
    DROP COLUMN diff;
END;

ALTER TABLE dbo.step_02_po_details
ADD diff DECIMAL(18,2) NULL;

UPDATE p
SET p.diff =
    CASE
        WHEN p.Total_Invoice IS NULL
          OR p.Total_Payment IS NULL
        THEN NULL
        ELSE TRY_CAST(p.Total_Invoice AS DECIMAL(18,2))
           - TRY_CAST(p.Total_Payment AS DECIMAL(18,2))
    END
FROM dbo.step_02_po_details AS p;

RAISERROR('Step 04A Po_status: diff updated in step_02_po_details.', 10, 1) WITH NOWAIT;

IF COL_LENGTH('dbo.step_02_po_details', 'po_status') IS NULL
BEGIN
    ALTER TABLE dbo.step_02_po_details
    ADD po_status VARCHAR(50) NULL;
END;

UPDATE po
SET po.po_status =
    CASE
        WHEN TRY_CAST(po.Total_Invoice AS DECIMAL(18,2)) IS NOT NULL
         AND TRY_CAST(po.Total_Payment AS DECIMAL(18,2)) IS NULL
            THEN 'No_Pay'
        WHEN TRY_CAST(po.Total_Invoice AS DECIMAL(18,2)) =
             TRY_CAST(po.Total_Payment AS DECIMAL(18,2))
            THEN 'Full Matched'
        WHEN TRY_CAST(po.Total_Invoice AS DECIMAL(18,2)) <
             TRY_CAST(po.Total_Payment AS DECIMAL(18,2))
            THEN 'Over_Pay'
        WHEN TRY_CAST(po.Total_Invoice AS DECIMAL(18,2)) >
             TRY_CAST(po.Total_Payment AS DECIMAL(18,2))
            THEN 'Short_Pay'
        ELSE NULL
    END
FROM dbo.step_02_po_details AS po;

RAISERROR('Step 04A Po_status: po_status updated in step_02_po_details.', 10, 1) WITH NOWAIT;
"""

STEP_05_DATE_MATCH_TABLE = "Inv_DB_Match_Sheet"
STEP_07_CREDIT_APPLIED_TABLE = "step_07_Credit_Applied_Sheet"
STEP_07A_DEDUCTIONS_BREAKDOWN_TABLE = "Step_09A_Deductions_BreakDown"
STEP_07A_CREDIT_SUMMARY_TABLE = "Credit_CrDedNet_Summary"
STEP_07A_PO_STATUS_SUMMARY_TABLE = "Po_Status_Sum_Summary"
STEP_07A_SUMMARY_TABLES = (
    STEP_07A_DEDUCTIONS_BREAKDOWN_TABLE,
    STEP_07A_CREDIT_SUMMARY_TABLE,
    STEP_07A_PO_STATUS_SUMMARY_TABLE,
)


STEP_05_SQL = """
         /* -- EXEC dbo.sp_Inv_DB_Match_Sheet_Update;  OLD*/
           EXEC dbo.sp_Step_05_WallMartSummary_Inv_Credit_Full_Update;
"""

STEP_05_06_COMBINED_SQL = """
EXEC dbo.sp_WallMartSummary_gen_DBs_Status;
"""

STEP_07A_SUMMARY_SQL = """
/*==============================================================
 STEP 07A - THREE SECTION TABLES ONLY

 PURPOSE:
 - Section 1: Create deduction breakdown table
 - Section 2: Create gross / deduction / net pay table
 - Section 3: Create PO status summary table
 - Do NOT combine all three into one final sheet

 OUTPUT TABLES:
 - dbo.Step_09A_Deductions_BreakDown
 - dbo.Credit_CrDedNet_Summary
 - dbo.Po_Status_Sum_Summary

 APP SAFE:
 - No GO used
 - Dynamic SQL safe
 - Human readable
==============================================================*/

DECLARE @cols NVARCHAR(MAX);
DECLARE @selectCols NVARCHAR(MAX);
DECLARE @step_03_deducktionBrkDwn NVARCHAR(MAX);

IF OBJECT_ID('dbo.step_03_payments_combined', 'U') IS NULL
BEGIN
    RAISERROR('Missing table: dbo.step_03_payments_combined', 16, 1);
    RETURN;
END;

IF OBJECT_ID('dbo.step_02_po_details', 'U') IS NULL
BEGIN
    RAISERROR('Missing table: dbo.step_02_po_details', 16, 1);
    RETURN;
END;

DROP TABLE IF EXISTS dbo.Step_09A_Deductions_BreakDown;
DROP TABLE IF EXISTS dbo.Credit_CrDedNet_Summary;
DROP TABLE IF EXISTS dbo.Po_Status_Sum_Summary;

SELECT
    @cols = STRING_AGG(QUOTENAME(col_name), ','),
    @selectCols = STRING_AGG(
        'ISNULL(' + QUOTENAME(col_name) + ', 0) AS ' + QUOTENAME(col_name),
        ','
    )
FROM (
    SELECT DISTINCT
        LTRIM(RTRIM(line_type)) + '_' +
        CASE
            WHEN description LIKE '%Marketing Allowance%' THEN 'Marketing_Allowance'
            WHEN description LIKE '%Sponsored Product Ads%' THEN 'Sponsored_Ads'
            WHEN description LIKE '%Overbill Charge-Back%' THEN 'Overbill_Charge_Back'
            WHEN description LIKE '%Audit Fee%' THEN 'Audit_Fees'
            WHEN description LIKE '%Expected%' AND description LIKE '%Billed%' THEN 'Shipping_Corrections'
            WHEN description LIKE '%General Allowance%' THEN 'General_Allowance'
            WHEN description LIKE '%Supplier Oasis%' THEN 'Supplier_Oasis_Fees'
            ELSE 'Other'
        END AS col_name
    FROM dbo.step_03_payments_combined
    WHERE line_type IN ('Adjustments', 'Returns', 'Supplier Oasis Fees')
      AND cr_issue IS NOT NULL
      AND payment_amount IS NOT NULL
) AS x;

IF @cols IS NULL OR LTRIM(RTRIM(@cols)) = ''
BEGIN
    RAISERROR('No deduction rows found in step_03_payments_combined.', 16, 1);
    RETURN;
END;

SET @step_03_deducktionBrkDwn = '
/*==============================
 SECTION 1 - DEDUCTION BREAKDOWN
==============================*/
;WITH base AS (
    SELECT
        LTRIM(RTRIM(cr_issue)) AS cr_issue,
        LTRIM(RTRIM(line_type)) + ''_'' +
        CASE
            WHEN description LIKE ''%Marketing Allowance%'' THEN ''Marketing_Allowance''
            WHEN description LIKE ''%Sponsored Product Ads%'' THEN ''Sponsored_Ads''
            WHEN description LIKE ''%Overbill Charge-Back%'' THEN ''Overbill_Charge_Back''
            WHEN description LIKE ''%Audit Fee%'' THEN ''Audit_Fees''
            WHEN description LIKE ''%Expected%'' AND description LIKE ''%Billed%'' THEN ''Shipping_Corrections''
            WHEN description LIKE ''%General Allowance%'' THEN ''General_Allowance''
            WHEN description LIKE ''%Supplier Oasis%'' THEN ''Supplier_Oasis_Fees''
            ELSE ''Other''
        END AS deduction_column,
        TRY_CAST(payment_amount AS DECIMAL(18,2)) AS total_amount
    FROM dbo.step_03_payments_combined
    WHERE line_type IN (''Adjustments'', ''Returns'', ''Supplier Oasis Fees'')
      AND cr_issue IS NOT NULL
      AND payment_amount IS NOT NULL
)
SELECT
    cr_issue,
    ' + @selectCols + '
INTO dbo.Step_09A_Deductions_BreakDown
FROM base
PIVOT (
    SUM(total_amount)
    FOR deduction_column IN (' + @cols + ')
) AS p;

/*==============================
 SECTION 2 - GROSS / DEDUCTION / NET PAY SUMMARY
==============================*/
SELECT
    LTRIM(RTRIM(cr_issue)) AS cr_issue,
    SUM(
        CASE
            WHEN TRY_CAST(payment_amount AS DECIMAL(18,2)) > 0
            THEN TRY_CAST(payment_amount AS DECIMAL(18,2))
            ELSE 0
        END
    ) AS Gross_total_by_Cr,
    SUM(
        CASE
            WHEN TRY_CAST(payment_amount AS DECIMAL(18,2)) < 0
            THEN TRY_CAST(payment_amount AS DECIMAL(18,2))
            ELSE 0
        END
    ) AS Deduction_total_by_Cr,
    SUM(
        CASE
            WHEN TRY_CAST(payment_amount AS DECIMAL(18,2)) > 0
            THEN TRY_CAST(payment_amount AS DECIMAL(18,2))
            ELSE 0
        END
    )
    +
    SUM(
        CASE
            WHEN TRY_CAST(payment_amount AS DECIMAL(18,2)) < 0
            THEN TRY_CAST(payment_amount AS DECIMAL(18,2))
            ELSE 0
        END
    ) AS Net_Pay_Received
INTO dbo.Credit_CrDedNet_Summary
FROM dbo.step_03_payments_combined
WHERE cr_issue IS NOT NULL
GROUP BY LTRIM(RTRIM(cr_issue));

/*==============================
 SECTION 3 - PO STATUS SUMMARY
==============================*/
SELECT
    LTRIM(RTRIM(invoice_db_number)) AS invoice_number,
    LTRIM(RTRIM(po_status)) AS po_status,
    SUM(TRY_CAST(diff AS DECIMAL(18,2))) AS total_diff_by_Po_Status,
    COUNT(diff) AS total_count
INTO dbo.Po_Status_Sum_Summary
FROM dbo.step_02_po_details
WHERE po_status IS NOT NULL
  AND invoice_db_number IS NOT NULL
  AND diff IS NOT NULL
  AND UPPER(LTRIM(RTRIM(po_status))) <> ''FULL MATCHED''
GROUP BY
    LTRIM(RTRIM(invoice_db_number)),
    LTRIM(RTRIM(po_status));
';

EXEC sp_executesql @step_03_deducktionBrkDwn;

RAISERROR('Step 07A: three summary tables created.', 10, 1) WITH NOWAIT;
"""

STEP_07A_DEDUCTION_BREAKDOWN_SQL = """
DECLARE @cols NVARCHAR(MAX);
DECLARE @selectCols NVARCHAR(MAX);
DECLARE @deductionSql NVARCHAR(MAX);

IF OBJECT_ID('dbo.step_03_payments_combined', 'U') IS NULL
BEGIN
    RAISERROR('Missing table: dbo.step_03_payments_combined', 16, 1);
    RETURN;
END;

DROP TABLE IF EXISTS dbo.Step_09A_Deductions_BreakDown;

SELECT
    @cols = STRING_AGG(QUOTENAME(col_name), ','),
    @selectCols = STRING_AGG(
        'ISNULL(' + QUOTENAME(col_name) + ', 0) AS ' + QUOTENAME(col_name),
        ','
    )
FROM (
    SELECT DISTINCT
        LTRIM(RTRIM(line_type)) + '_' +
        CASE
            WHEN description LIKE '%Marketing Allowance%' THEN 'Marketing_Allowance'
            WHEN description LIKE '%Sponsored Product Ads%' THEN 'Sponsored_Ads'
            WHEN description LIKE '%Overbill Charge-Back%' THEN 'Overbill_Charge_Back'
            WHEN description LIKE '%Audit Fee%' THEN 'Audit_Fees'
            WHEN description LIKE '%Expected%' AND description LIKE '%Billed%' THEN 'Shipping_Corrections'
            WHEN description LIKE '%General Allowance%' THEN 'General_Allowance'
            WHEN description LIKE '%Supplier Oasis%' THEN 'Supplier_Oasis_Fees'
            ELSE 'Other'
        END AS col_name
    FROM dbo.step_03_payments_combined
    WHERE line_type IN ('Adjustments', 'Returns', 'Supplier Oasis Fees')
      AND cr_issue IS NOT NULL
      AND payment_amount IS NOT NULL
) AS x;

IF @cols IS NULL OR LTRIM(RTRIM(@cols)) = ''
BEGIN
    RAISERROR('No deduction rows found in step_03_payments_combined.', 16, 1);
    RETURN;
END;

SET @deductionSql = '
;WITH base AS (
    SELECT
        LTRIM(RTRIM(cr_issue)) AS cr_issue,
        LTRIM(RTRIM(line_type)) + ''_'' +
        CASE
            WHEN description LIKE ''%Marketing Allowance%'' THEN ''Marketing_Allowance''
            WHEN description LIKE ''%Sponsored Product Ads%'' THEN ''Sponsored_Ads''
            WHEN description LIKE ''%Overbill Charge-Back%'' THEN ''Overbill_Charge_Back''
            WHEN description LIKE ''%Audit Fee%'' THEN ''Audit_Fees''
            WHEN description LIKE ''%Expected%'' AND description LIKE ''%Billed%'' THEN ''Shipping_Corrections''
            WHEN description LIKE ''%General Allowance%'' THEN ''General_Allowance''
            WHEN description LIKE ''%Supplier Oasis%'' THEN ''Supplier_Oasis_Fees''
            ELSE ''Other''
        END AS deduction_column,
        TRY_CAST(payment_amount AS DECIMAL(18,2)) AS total_amount
    FROM dbo.step_03_payments_combined
    WHERE line_type IN (''Adjustments'', ''Returns'', ''Supplier Oasis Fees'')
      AND cr_issue IS NOT NULL
      AND payment_amount IS NOT NULL
)
SELECT
    cr_issue,
    ' + @selectCols + '
INTO dbo.Step_09A_Deductions_BreakDown
FROM base
PIVOT (
    SUM(total_amount)
    FOR deduction_column IN (' + @cols + ')
) AS p;';

EXEC sp_executesql @deductionSql;
RAISERROR('Step 07A Section 1: deduction breakdown table created.', 10, 1) WITH NOWAIT;
"""

STEP_07A_CREDIT_SUMMARY_SQL = """
IF OBJECT_ID('dbo.step_03_payments_combined', 'U') IS NULL
BEGIN
    RAISERROR('Missing table: dbo.step_03_payments_combined', 16, 1);
    RETURN;
END;

DROP TABLE IF EXISTS dbo.Credit_CrDedNet_Summary;

SELECT
    LTRIM(RTRIM(cr_issue)) AS cr_issue,
    SUM(
        CASE
            WHEN TRY_CAST(payment_amount AS DECIMAL(18,2)) > 0
            THEN TRY_CAST(payment_amount AS DECIMAL(18,2))
            ELSE 0
        END
    ) AS Gross_total_by_Cr,
    SUM(
        CASE
            WHEN TRY_CAST(payment_amount AS DECIMAL(18,2)) < 0
            THEN TRY_CAST(payment_amount AS DECIMAL(18,2))
            ELSE 0
        END
    ) AS Deduction_total_by_Cr,
    SUM(
        CASE
            WHEN TRY_CAST(payment_amount AS DECIMAL(18,2)) > 0
            THEN TRY_CAST(payment_amount AS DECIMAL(18,2))
            ELSE 0
        END
    )
    +
    SUM(
        CASE
            WHEN TRY_CAST(payment_amount AS DECIMAL(18,2)) < 0
            THEN TRY_CAST(payment_amount AS DECIMAL(18,2))
            ELSE 0
        END
    ) AS Net_Pay_Received
INTO dbo.Credit_CrDedNet_Summary
FROM dbo.step_03_payments_combined
WHERE cr_issue IS NOT NULL
GROUP BY LTRIM(RTRIM(cr_issue));

RAISERROR('Step 07A Section 2: credit gross/deduction/net table created.', 10, 1) WITH NOWAIT;
"""

STEP_07A_PO_STATUS_SUMMARY_SQL = """
IF OBJECT_ID('dbo.step_02_po_details', 'U') IS NULL
BEGIN
    RAISERROR('Missing table: dbo.step_02_po_details', 16, 1);
    RETURN;
END;

DROP TABLE IF EXISTS dbo.Po_Status_Sum_Summary;

SELECT
    LTRIM(RTRIM(invoice_db_number)) AS invoice_number,
    LTRIM(RTRIM(po_status)) AS po_status,
    SUM(TRY_CAST(diff AS DECIMAL(18,2))) AS total_diff_by_Po_Status,
    COUNT(diff) AS total_count
INTO dbo.Po_Status_Sum_Summary
FROM dbo.step_02_po_details
WHERE po_status IS NOT NULL
  AND invoice_db_number IS NOT NULL
  AND diff IS NOT NULL
  AND UPPER(LTRIM(RTRIM(po_status))) <> 'FULL MATCHED'
GROUP BY
    LTRIM(RTRIM(invoice_db_number)),
    LTRIM(RTRIM(po_status));

RAISERROR('Step 07A Section 3: PO status summary table created.', 10, 1) WITH NOWAIT;
"""

STEP_07_SQL = """



EXEC dbo.sp_WallMartSummary_Gen_step_07_Credit_Applied_Sheet;


"""

PROC_QUERY_SAMPLE_SQL = """
PRINT 'Proc query test started';
RAISERROR('Proc query progress message from SQL Server', 10, 1) WITH NOWAIT;
SELECT
    DB_NAME() AS current_database,
    SUSER_SNAME() AS login_name,
    USER_NAME() AS database_user;
PRINT 'Proc query test finished';

-- Replace the sample with your proc call, for example:
-- EXEC dbo.YourStoredProcedureName;
"""

STEP_05_STEP_06_SPLIT_MARKER = "--@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ 6666666666666666666666666666666"

STEP_06_SQL = """
EXEC dbo.sp_step_06_WallMartSummary_Generating_Credit_Match_Sheet;

"""

STEP_06B_DB_SQL = """
/*
UPDATE po
SET po.invoice_db_number = nm.invoice_number
FROM dbo.step_02_po_details po
INNER JOIN dbo.step_04b_order_not_matched nm
    ON LTRIM(RTRIM(po.po_number)) = LTRIM(RTRIM(nm.po_number))
WHERE nm.invoice_number IS NOT NULL;

UPDATE t
SET t.invoice_number = i.ref
FROM dbo.step_04b_order_not_matched t
INNER JOIN dbo.step_01_inquiry_rows i
    ON
        CASE
            WHEN t.invoice_number LIKE 'DB#%'
                THEN RIGHT(LTRIM(RTRIM(t.invoice_number)), 6)

            WHEN CHARINDEX('#', t.invoice_number) > 0
                THEN LTRIM(RTRIM(
                    RIGHT(
                        t.invoice_number,
                        LEN(t.invoice_number) - CHARINDEX('#', t.invoice_number)
                    )
                ))

            ELSE LTRIM(RTRIM(t.invoice_number))
        END
        =
        CASE
            WHEN i.claim LIKE 'DB%'
                THEN RIGHT(LTRIM(RTRIM(i.claim)), 6)

            ELSE LTRIM(RTRIM(i.claim))
        END
WHERE t.cr_issue IS NULL
  AND t.invoice_number IS NOT NULL
  AND LTRIM(RTRIM(i.[order])) = 'PARTIAL PAYMENT';  
  
  -- =============
 ---  OLD  METHOD
---  ===========
  */

  --===== NEW  METHOD   ======



  -- Run this only if db_number is INT and you want to store DB#528648 text
ALTER TABLE dbo.step_04b_order_not_matched
ALTER COLUMN db_number NVARCHAR(50);


UPDATE t
SET t.db_number = LTRIM(RTRIM(i.ref))
FROM dbo.step_04b_order_not_matched t
INNER JOIN dbo.step_01_inquiry_rows i
    ON
        CASE
            WHEN CHARINDEX('#', t.invoice_number) > 0
                THEN LTRIM(RTRIM(
                    RIGHT(
                        t.invoice_number,
                        LEN(t.invoice_number) - CHARINDEX('#', t.invoice_number)
                    )
                ))
            ELSE LTRIM(RTRIM(t.invoice_number))
        END
        =
        LTRIM(RTRIM(i.Cclaim))
WHERE t.cr_issue IS NULL
  AND t.invoice_number IS NOT NULL;




--    SHIFITNG    DB#   FROM  STEP 04  TO STEP 02 PO DETAIL

UPDATE po
SET po.db_number = nm.db_number
FROM dbo.step_02_po_details po
INNER JOIN dbo.step_04b_order_not_matched nm
    ON LTRIM(RTRIM(po.po_number)) = LTRIM(RTRIM(nm.po_number))
WHERE nm.invoice_number IS NOT NULL;




--updating   inv_db  col   > Db#

  UPDATE po
SET po.inv_db = nm.db_number
FROM dbo.step_02_po_details po
INNER JOIN dbo.step_04b_order_not_matched nm
    ON LTRIM(RTRIM(po.po_number)) = LTRIM(RTRIM(nm.po_number))
WHERE nm.invoice_number IS NOT NULL;


  --   updating  inv_db   col  >inv #


  UPDATE dbo.step_02_po_details
SET inv_db  = invoice_db_number
where invoice_db_number IS NOT NULL
  AND cr_issue IS NOT NULL;













"""

STEP_06D_UNAPPLIED_SQL = """
  IF OBJECT_ID('dbo.unapplied', 'U') IS NOT NULL
    DROP TABLE dbo.unapplied;


SELECT
    cr_issue,
    SUM(payment_amount) AS unapplied_amount,
    SUM(SUM(payment_amount)) OVER () AS total_unapplied_amount
INTO unapplied
FROM step_04b_order_not_matched
WHERE invoice_number IS NULL
  AND cr_issue IS NOT NULL
  AND payment_amount > 0
GROUP BY cr_issue;







"""

STEP_08_DATABASE_OPTIMIZATION_SQL = """
EXEC dbo.sp_WallMartSummary_Truncate_Selected_Tables;
"""

DEDUCTIONS_SUMMARY_SQL = """
DROP TABLE IF EXISTS dbo.Deductions_Summary;

SELECT DISTINCT
    p.*,
    line_type + '_' +
    CASE
        WHEN description LIKE '%Marketing Allowance%' THEN 'Marketing_Allowance'
        WHEN description LIKE '%Sponsored Product Ads%' THEN 'Sponsored_Ads'
        WHEN description LIKE '%Overbill Charge-Back%' THEN 'Overbill_CB'
        WHEN description LIKE '%Audit Fee%' THEN 'Audit_Fees'
        WHEN description LIKE '%Expected%'
         AND description LIKE '%Billed%' THEN 'Shipping_Corrections'
        WHEN description LIKE '%General Allowance%' THEN 'General_Allowance'
        WHEN description LIKE '%Supplier Oasis%' THEN 'Supplier_Oasis'
        ELSE 'Other'
    END AS col_name,
    description AS Deduction_category
INTO dbo.Deductions_Summary
FROM dbo.step_03_payments_combined AS p
WHERE line_type IS NOT NULL
  AND TRY_CAST(payment_amount AS DECIMAL(18,2)) < 0;
"""

INVOICE_PAYMENT_BREAKDOWN_SQL = """ 
EXEC dbo.sp_WallMartSummary_Invoice_Status;

"""

STEP_06D02_APPLYING_CR_UPDATE_SQL = """
/*==============================================================
 STEP 06D - APP SAFE UPDATE UNAPPLIED CR ISSUE

 APP SAFE VERSION:
 - No GO
 - No transaction
 - No @@ROWCOUNT
 - Safe for app SQL runner
 - Uses p.invoice_number
 - Updates only:
      p.payment_amount > 0
      AND p.invoice_number IS NULL

 LOGIC:
 1. Match:
    ABS(i.pay_day) = ABS(u.unapplied_amount)

 2. Clean p.cr_issue:
    CR#576763 -> 576763

 3. Match:
    cleaned p.cr_issue = i.Cclaim

 4. Update:
    p.cr_issue = i.ref
==============================================================*/

---------------------------------------------------------------
-- 0. SAFETY CHECKS
---------------------------------------------------------------

IF OBJECT_ID('dbo.step_03_payments_combined', 'U') IS NULL
BEGIN
    RAISERROR('Missing table: dbo.step_03_payments_combined', 16, 1);
    RETURN;
END;

IF OBJECT_ID('dbo.step_01_inquiry_rows', 'U') IS NULL
BEGIN
    RAISERROR('Missing table: dbo.step_01_inquiry_rows', 16, 1);
    RETURN;
END;

IF OBJECT_ID('dbo.unapplied', 'U') IS NULL
BEGIN
    RAISERROR('Missing table: dbo.unapplied', 16, 1);
    RETURN;
END;

IF COL_LENGTH('dbo.step_03_payments_combined', 'cr_issue') IS NULL
BEGIN
    RAISERROR('Missing column: dbo.step_03_payments_combined.cr_issue', 16, 1);
    RETURN;
END;

IF COL_LENGTH('dbo.step_03_payments_combined', 'payment_amount') IS NULL
BEGIN
    RAISERROR('Missing column: dbo.step_03_payments_combined.payment_amount', 16, 1);
    RETURN;
END;

IF COL_LENGTH('dbo.step_03_payments_combined', 'invoice_number') IS NULL
BEGIN
    RAISERROR('Missing column: dbo.step_03_payments_combined.invoice_number', 16, 1);
    RETURN;
END;

IF COL_LENGTH('dbo.step_01_inquiry_rows', 'ref') IS NULL
BEGIN
    RAISERROR('Missing column: dbo.step_01_inquiry_rows.ref', 16, 1);
    RETURN;
END;

IF COL_LENGTH('dbo.step_01_inquiry_rows', 'Cclaim') IS NULL
BEGIN
    RAISERROR('Missing column: dbo.step_01_inquiry_rows.Cclaim', 16, 1);
    RETURN;
END;

IF COL_LENGTH('dbo.step_01_inquiry_rows', 'pay_day') IS NULL
BEGIN
    RAISERROR('Missing column: dbo.step_01_inquiry_rows.pay_day', 16, 1);
    RETURN;
END;

IF COL_LENGTH('dbo.unapplied', 'unapplied_amount') IS NULL
BEGIN
    RAISERROR('Missing column: dbo.unapplied.unapplied_amount', 16, 1);
    RETURN;
END;


---------------------------------------------------------------
-- 1. COUNT ROWS BEFORE UPDATE
---------------------------------------------------------------

SELECT
    COUNT(*) AS rows_ready_to_update_before
FROM dbo.step_03_payments_combined p
INNER JOIN dbo.step_01_inquiry_rows i
    ON LTRIM(RTRIM(
        REPLACE(
            REPLACE(
                REPLACE(
                    REPLACE(p.cr_issue, 'CR#', ''),
                'CR #', ''),
            'CR-', ''),
        'CR', '')
    ))
    =
    LTRIM(RTRIM(i.Cclaim))
INNER JOIN dbo.unapplied u
    ON ABS(TRY_CAST(i.pay_day AS DECIMAL(18,2)))
     = ABS(TRY_CAST(u.unapplied_amount AS DECIMAL(18,2)))
WHERE p.cr_issue IS NOT NULL
  AND LTRIM(RTRIM(p.cr_issue)) <> ''
  AND i.ref IS NOT NULL
  AND TRY_CAST(p.payment_amount AS DECIMAL(18,2)) > 0
  AND p.invoice_number IS NULL
  AND LTRIM(RTRIM(p.cr_issue)) <> LTRIM(RTRIM(i.ref));


---------------------------------------------------------------
-- 2. PREVIEW BEFORE UPDATE
---------------------------------------------------------------

SELECT
    'PREVIEW BEFORE UPDATE' AS check_type,

    u.unapplied_amount,
    i.pay_day AS inquiry_pay_day,

    ABS(TRY_CAST(u.unapplied_amount AS DECIMAL(18,2))) AS abs_unapplied_amount,
    ABS(TRY_CAST(i.pay_day AS DECIMAL(18,2))) AS abs_inquiry_pay_day,

    p.cr_issue AS payment_cr_issue_before,

    LTRIM(RTRIM(
        REPLACE(
            REPLACE(
                REPLACE(
                    REPLACE(p.cr_issue, 'CR#', ''),
                'CR #', ''),
            'CR-', ''),
        'CR', '')
    )) AS cleaned_payment_cr_issue,

    i.Cclaim AS inquiry_Cclaim,
    i.ref AS new_cr_issue_to_apply,

    p.invoice_number,
    p.payment_amount AS payment_total,
    p.po_number,
    p.line_type,
    p.source_file

FROM dbo.step_03_payments_combined p
INNER JOIN dbo.step_01_inquiry_rows i
    ON LTRIM(RTRIM(
        REPLACE(
            REPLACE(
                REPLACE(
                    REPLACE(p.cr_issue, 'CR#', ''),
                'CR #', ''),
            'CR-', ''),
        'CR', '')
    ))
    =
    LTRIM(RTRIM(i.Cclaim))
INNER JOIN dbo.unapplied u
    ON ABS(TRY_CAST(i.pay_day AS DECIMAL(18,2)))
     = ABS(TRY_CAST(u.unapplied_amount AS DECIMAL(18,2)))
WHERE p.cr_issue IS NOT NULL
  AND LTRIM(RTRIM(p.cr_issue)) <> ''
  AND i.ref IS NOT NULL
  AND TRY_CAST(p.payment_amount AS DECIMAL(18,2)) > 0
  AND p.invoice_number IS NULL
  AND LTRIM(RTRIM(p.cr_issue)) <> LTRIM(RTRIM(i.ref))
ORDER BY
    ABS(TRY_CAST(u.unapplied_amount AS DECIMAL(18,2))),
    i.ref,
    p.cr_issue;


---------------------------------------------------------------
-- 3. UPDATE PAYMENT CR ISSUE
---------------------------------------------------------------

UPDATE p
SET p.cr_issue = i.ref
FROM dbo.step_03_payments_combined p
INNER JOIN dbo.step_01_inquiry_rows i
    ON LTRIM(RTRIM(
        REPLACE(
            REPLACE(
                REPLACE(
                    REPLACE(p.cr_issue, 'CR#', ''),
                'CR #', ''),
            'CR-', ''),
        'CR', '')
    ))
    =
    LTRIM(RTRIM(i.Cclaim))
INNER JOIN dbo.unapplied u
    ON ABS(TRY_CAST(i.pay_day AS DECIMAL(18,2)))
     = ABS(TRY_CAST(u.unapplied_amount AS DECIMAL(18,2)))
WHERE p.cr_issue IS NOT NULL
  AND LTRIM(RTRIM(p.cr_issue)) <> ''
  AND i.ref IS NOT NULL
  AND TRY_CAST(p.payment_amount AS DECIMAL(18,2)) > 0
  AND p.invoice_number IS NULL
  AND LTRIM(RTRIM(p.cr_issue)) <> LTRIM(RTRIM(i.ref));


---------------------------------------------------------------
-- 4. COUNT ROWS STILL READY AFTER UPDATE
-- If this shows 0, update is complete.
---------------------------------------------------------------

SELECT
    COUNT(*) AS rows_still_ready_to_update_after
FROM dbo.step_03_payments_combined p
INNER JOIN dbo.step_01_inquiry_rows i
    ON LTRIM(RTRIM(
        REPLACE(
            REPLACE(
                REPLACE(
                    REPLACE(p.cr_issue, 'CR#', ''),
                'CR #', ''),
            'CR-', ''),
        'CR', '')
    ))
    =
    LTRIM(RTRIM(i.Cclaim))
INNER JOIN dbo.unapplied u
    ON ABS(TRY_CAST(i.pay_day AS DECIMAL(18,2)))
     = ABS(TRY_CAST(u.unapplied_amount AS DECIMAL(18,2)))
WHERE p.cr_issue IS NOT NULL
  AND LTRIM(RTRIM(p.cr_issue)) <> ''
  AND i.ref IS NOT NULL
  AND TRY_CAST(p.payment_amount AS DECIMAL(18,2)) > 0
  AND p.invoice_number IS NULL
  AND LTRIM(RTRIM(p.cr_issue)) <> LTRIM(RTRIM(i.ref));


---------------------------------------------------------------
-- 5. CHECK AFTER UPDATE
---------------------------------------------------------------

SELECT
    'CHECK AFTER UPDATE' AS check_type,

    p.cr_issue AS payment_cr_issue_after,
    i.ref AS inquiry_ref,
    i.Cclaim,
    i.pay_day,
    u.unapplied_amount,

    p.invoice_number,
    p.payment_amount,
    p.po_number,
    p.line_type,
    p.source_file

FROM dbo.step_03_payments_combined p
INNER JOIN dbo.step_01_inquiry_rows i
    ON LTRIM(RTRIM(p.cr_issue)) = LTRIM(RTRIM(i.ref))
INNER JOIN dbo.unapplied u
    ON ABS(TRY_CAST(i.pay_day AS DECIMAL(18,2)))
     = ABS(TRY_CAST(u.unapplied_amount AS DECIMAL(18,2)))
WHERE TRY_CAST(p.payment_amount AS DECIMAL(18,2)) > 0
  AND p.invoice_number IS NULL
ORDER BY
    p.cr_issue,
    p.po_number;
"""


def _company_settings(company_code: str = DEFAULT_COMPANY_CODE):
    company_profile = load_company_profile(company_code)
    return company_profile, replace(settings, sql_database=company_profile.sql_database)


def build_services(company_code: str = DEFAULT_COMPANY_CODE) -> tuple[Logger, WorkflowOrchestrator, SQLServerRepository]:
    normalized_company_code = company_code.strip().lower()
    if normalized_company_code not in SUPPORTED_COMPANY_CODES:
        raise ValueError(f"Unsupported company code: {company_code}")
    logger = Logger()
    company_profile = load_company_profile(normalized_company_code)
    company_settings = replace(settings, sql_database=company_profile.sql_database)
    orchestrator = WorkflowOrchestrator(
        logger=logger,
        inquiry_processor=InquiryProcessor(logger, company_profile),
    )
    repository = SQLServerRepository(settings=company_settings, logger=logger)
    return logger, orchestrator, repository


def _company_output_dir(company_code: str = DEFAULT_COMPANY_CODE) -> Path:
    normalized_company_code = company_code.strip().lower()
    if normalized_company_code == DEFAULT_COMPANY_CODE:
        return settings.output_dir
    return settings.output_dir / normalized_company_code


def _automatic_report_sql(company_code: str) -> str | None:
    company_profile = load_company_profile(company_code)
    if not company_profile.report_workflow_ready:
        return None
    if company_profile.code == "wallmartsummary":
        return RUN_WALLMART_SUMMARY_SQL
    return None


def _get_step_05_in_sql() -> str:
    return STEP_05_SQL.split(STEP_05_STEP_06_SPLIT_MARKER, 1)[0].strip()




def _step_01_output_path() -> Path:
    return settings.output_dir / STEP_01_OUTPUT_FILE


def _step_01a_output_path() -> Path:
    return settings.output_dir / STEP_01A_OUTPUT_FILE




def _amazon_grouped_output_path() -> Path:
    return _company_output_dir("amazonaccounts") / AMAZON_GROUPED_OUTPUT_FILE


def _delete_file_after_download(path) -> None:
    """Remove a one-shot generated report from output/ once it has been sent to the browser."""
    try:
        Path(path).unlink(missing_ok=True)
    except OSError:
        pass


def _amazon_download_done(path, download_token: str) -> None:
    _mark_amazon_download_complete(download_token)
    _delete_file_after_download(path)


def _mark_amazon_download_complete(download_token: str) -> None:
    if download_token in amazon_download_jobs:
        amazon_download_jobs[download_token]["complete"] = True


def _amazon_download_filename() -> str:
    timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
    return f"Amazon Status Summary {timestamp}.xlsx"


def _wallmart_python_output_path() -> Path:
    return _company_output_dir("wallmartsummary") / WALLMART_PYTHON_OUTPUT_FILE


def _wallmart_python_download_filename() -> str:
    timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
    return f"Wal Mart Status Summary {timestamp}.xlsx"






def _latest_tables_output_path(company_code: str = DEFAULT_COMPANY_CODE) -> Path:
    return _company_output_dir(company_code) / LATEST_TABLE_EXPORT_FILE




def _safe_sheet_name(table_name: str, used_names: set[str]) -> str:
    # Sheet name cap
    base_name = table_name[:31] or "sheet"
    candidate = base_name
    suffix = 1
    while candidate.lower() in used_names:
        suffix_text = f"_{suffix}"
        candidate = f"{base_name[:31 - len(suffix_text)]}{suffix_text}"
        suffix += 1
    used_names.add(candidate.lower())
    return candidate


LATEST_TABLES_CURRENCY_FORMAT = "$#,##0.00"
LATEST_TABLES_STATUS_CURRENCY_SHEETS = {
    "invoices_status",
    "dbs_status",
}
INVOICES_STATUS_TABLE = "Invoices_Status"
DBS_STATUS_TABLE = "DBs_Status"
EXPORT_HIDDEN_COLUMNS = {"id", "batch_id", "row_number", "ndatee"}
LATEST_TABLES_MONEY_COLUMNS = {
    "current",
    "col_1_30_days",
    "col_31_60_days",
    "col_61_90_days",
    "over_90_days",
    "pay_day",
    "wholesale_price",
    "unit_price",
    "qxc",
    "payment_amount",
    "sum_total",
    "sumbyinv_unpaid",
    "sumbycr_unapplied",
    "retail_price",
    "order_total",
    "order_total_invoiced",
    "hsn_drop_ship_fee",
    "manual_cost",
    "cost_or_unit_price",
    "po_cost",
    "pay_unit_price",
    "unit_variance",
    "total_variance",
}


def _dynamic_year_columns(column_lookup: dict[str, object]) -> list[str]:
    """Return every four-digit year column supplied by the current SQL table."""
    year_columns = [
        str(source_column).strip()
        for source_column in column_lookup.values()
        if re.fullmatch(r"\d{4}", str(source_column).strip())
    ]
    return sorted(year_columns, key=int)


def _apply_currency_format_to_column(worksheet: Worksheet, excel_column_index: int) -> None:
    for row_cells in worksheet.iter_rows(
        min_row=2,
        min_col=excel_column_index,
        max_col=excel_column_index,
    ):
        cell = row_cells[0]
        if isinstance(cell.value, (int, float, Decimal)) and not isinstance(cell.value, bool):
            cell.number_format = LATEST_TABLES_CURRENCY_FORMAT


def _apply_currency_format_to_decimal_values(
    worksheet: Worksheet,
    convert_decimal_text: bool = False,
) -> None:
    """Format decimal report values, optionally converting decimal text for Excel."""
    for row_cells in worksheet.iter_rows(min_row=2):
        for cell in row_cells:
            if isinstance(cell.value, (float, Decimal)) and not isinstance(cell.value, bool):
                cell.number_format = LATEST_TABLES_CURRENCY_FORMAT
                continue

            if not convert_decimal_text or not isinstance(cell.value, str):
                continue

            text_value = cell.value.strip()
            if not re.fullmatch(r"[+-]?(?:\d{1,3}(?:,\d{3})+|\d+)\.\d+", text_value):
                continue

            numeric_value = _coerce_numeric_like_excel_value(text_value)
            if numeric_value is not None:
                cell.value = numeric_value
                cell.number_format = LATEST_TABLES_CURRENCY_FORMAT


def _coerce_numeric_like_excel_value(value: object) -> int | float | Decimal | None:
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float, Decimal)):
        return value
    if value is None:
        return None

    text_value = str(value).strip()
    if not text_value:
        return None

    normalized = text_value.replace(",", "").replace("$", "")
    try:
        decimal_value = Decimal(normalized)
    except Exception:
        return None

    if decimal_value == decimal_value.to_integral_value():
        return int(decimal_value)
    return float(decimal_value)


def _coerce_step_07_currency_columns(table_df: pd.DataFrame) -> pd.DataFrame:
    coerced_df = table_df.copy()
    for column_name in coerced_df.columns:
        series = coerced_df[column_name]
        if pd.api.types.is_numeric_dtype(series):
            continue

        nonblank_mask = series.notna() & series.astype(str).str.strip().ne("")
        if not nonblank_mask.any():
            continue

        numeric_values = pd.to_numeric(series, errors="coerce")
        if numeric_values[nonblank_mask].notna().all():
            coerced_df[column_name] = numeric_values

    return coerced_df


def _format_step_07_latest_tables_sheet(worksheet: Worksheet) -> None:
    for row_cells in worksheet.iter_rows(min_row=3):
        for cell in row_cells:
            numeric_value = _coerce_numeric_like_excel_value(cell.value)
            if numeric_value is None:
                continue
            cell.value = numeric_value
            cell.number_format = LATEST_TABLES_CURRENCY_FORMAT


def _step_07_export_header_labels(column_names: list[str]) -> list[str]:
    label_map = {
        "blank_1": "",
        "blank_2": "",
        "blank_3": "",
        "blank_4": "",
        "blank_5": "",
        "blank_6": "",
        "blank_7": "",
        "blank_8": "",
        "blank_9": "",
        "gross_total_by_cr": "Gross Total",
        "deduction_total_by_cr": "Deduction Total",
        "net_pay_received": "Net Pay Received",
        "adjustments_audit_fees": "Adjustment Audit Fees",
        "adjustments_marketing_allowance": "Adjustment Marketing",
        "adjustments_overbill_charge_back": "Adjustment Overbill CR",
        "adjustments_shipping_corrections": "Adjustment Shipping C",
        "adjustments_sponsored_ads": "Adjustment Sponsored",
        "returns_general_allowance": "Returns General A",
        "returns_other": "Returns Other",
        "supplier oasis fees_supplier_oasis_fees": "Supplier Oasis Fees",
        "net difference": "Net Difference",
        "net difference by cr": "Net Difference",
    }
    labels: list[str] = []
    for column_name in column_names:
        normalized = str(column_name).strip().lower()
        labels.append(label_map.get(normalized, str(column_name)))
    return labels


def _build_step_07_export_rows(table_df: pd.DataFrame) -> tuple[list[str], list[list[object]]]:
    header_labels = _step_07_export_header_labels(list(table_df.columns))
    data_rows = table_df.astype(object).where(pd.notna(table_df), None).values.tolist()
    if data_rows and str(data_rows[0][0]).strip().lower() == "credit #":
        data_rows = data_rows[1:]

    expanded_rows: list[list[object]] = []
    summary_break_labels = {
        "Total Credit Amount",
        "YCS Invoice+DB Amount",
        "Difference",
        "Unpaid",
        "Short_Pay",
        "Over_Pay",
    }
    for row in data_rows:
        label = "" if row[0] is None else str(row[0]).strip()
        expanded_rows.append(row)
        if label == "Invoice Month":
            expanded_rows.append([None] * len(header_labels))
        elif label in summary_break_labels:
            expanded_rows.append([None] * len(header_labels))

    return header_labels, expanded_rows


def _apply_step_07_export_sheet_format(worksheet: Worksheet, header_labels: list[str]) -> None:
    spacer_columns = {
        index for index, label in enumerate(header_labels, start=1) if str(label).strip() == ""
    }
    db_columns = {
        index for index, label in enumerate(header_labels, start=1) if str(label).strip().upper().startswith("DB#")
    }

    for column_index, label in enumerate(header_labels, start=1):
        cell = worksheet.cell(row=1, column=column_index)
        if column_index in spacer_columns:
            cell.fill = STEP_07_BLANK_FILL
        elif column_index in db_columns:
            cell.fill = STEP_07_DB_FILL
        else:
            cell.fill = STEP_07_HEADER_FILL

    for row_index in range(2, worksheet.max_row + 1):
        row_label = worksheet.cell(row=row_index, column=1).value
        row_label_text = "" if row_label is None else str(row_label).strip()
        is_blank_row = row_label_text == ""
        is_summary_row = row_label_text in {
            "Total Credit Amount",
            "YCS Invoice+DB Amount",
            "Difference",
            "Unpaid",
            "Short_Pay",
            "Over_Pay",
            "Net Difference",
        }

        for column_index in range(1, worksheet.max_column + 1):
            cell = worksheet.cell(row=row_index, column=column_index)
            if is_blank_row:
                cell.fill = STEP_07_BLANK_FILL
                cell.value = None
                continue

            if column_index == 1:
                cell.fill = STEP_07_HEADER_FILL
            elif column_index in spacer_columns:
                cell.fill = STEP_07_BLANK_FILL
            elif is_summary_row and row_label_text in {"Total Credit Amount", "Net Difference"} and column_index in db_columns:
                cell.fill = STEP_07_DB_FILL
            elif is_summary_row and row_label_text in {"Total Credit Amount", "Net Difference"}:
                cell.fill = STEP_07_HEADER_FILL
            else:
                cell.fill = STEP_07_WHITE_FILL

            numeric_value = _coerce_numeric_like_excel_value(cell.value)
            if numeric_value is not None:
                cell.value = numeric_value
                cell.number_format = LATEST_TABLES_CURRENCY_FORMAT


def _format_latest_tables_sheet(
    worksheet: Worksheet,
    table_name: str,
    table_df: pd.DataFrame,
) -> None:
    normalized_sheet_name = table_name.strip().lower()
    if normalized_sheet_name == STEP_07_CREDIT_APPLIED_TABLE.lower():
        _format_step_07_latest_tables_sheet(worksheet)
        return
    if normalized_sheet_name in LATEST_TABLES_STATUS_CURRENCY_SHEETS:
        _apply_currency_format_to_decimal_values(
            worksheet,
            convert_decimal_text=normalized_sheet_name == "dbs_status",
        )
        return

    for column_index, column_name in enumerate(table_df.columns, start=1):
        series = table_df[column_name]
        if not pd.api.types.is_numeric_dtype(series):
            continue

        normalized_column_name = str(column_name).strip().lower()
        should_format = normalized_sheet_name == STEP_07_CREDIT_APPLIED_TABLE.lower()
        if not should_format:
            should_format = normalized_column_name in LATEST_TABLES_MONEY_COLUMNS

        if should_format:
            _apply_currency_format_to_column(worksheet, column_index)


def _write_invoices_status_report(worksheet: Worksheet, table_df: pd.DataFrame) -> None:
    """Create the presentation layout requested for the invoice-status export sheet."""
    column_lookup = {str(column).strip().lower(): column for column in table_df.columns}

    def source_value(row: pd.Series, column_name: str) -> object:
        source_column = column_lookup.get(column_name.lower())
        if source_column is None:
            return None
        value = row.get(source_column)
        return None if pd.isna(value) else value

    row_type_column = column_lookup.get("row type")
    if row_type_column is None:
        raise ValueError("Invoices_Status is missing the required 'Row Type' column.")

    row_types = table_df[row_type_column].fillna("").astype(str).str.strip().str.upper()
    summary_rows = table_df.loc[row_types.eq("SUMMARY")]
    account_rows = table_df.loc[row_types.eq("ACCOUNT")]
    total_rows = table_df.loc[row_types.eq("TOTAL")]
    if summary_rows.empty or account_rows.empty:
        raise ValueError("Invoices_Status must contain SUMMARY and ACCOUNT rows for the report layout.")

    summary_row = summary_rows.iloc[0]
    years = _dynamic_year_columns(column_lookup)
    amount_report_columns = years + ["Total Amount", "Paid#", "Paid & Fully Deduct"]
    retail_link_column = "RetailLinkStatus" if "retaillinkstatus" in column_lookup else None
    report_columns = amount_report_columns + ([retail_link_column] if retail_link_column else [])
    table_headers = ["Account Name"] + [
        "Not Found On Portal" if column_name == retail_link_column else column_name
        for column_name in report_columns
    ]

    def account_has_positive_value(source_row: pd.Series) -> bool:
        for column_name in report_columns:
            numeric_value = _coerce_numeric_like_excel_value(source_value(source_row, column_name))
            if numeric_value is not None and numeric_value > 0:
                return True
        return False

    visible_account_rows = account_rows.loc[account_rows.apply(account_has_positive_value, axis=1)]
    title_fill = PatternFill("solid", fgColor="1F4E78")
    section_fill = PatternFill("solid", fgColor="2F75B5")
    header_fill = PatternFill("solid", fgColor="4472C4")
    total_fill = PatternFill("solid", fgColor="4472C4")
    thin_side = Side(style="thin", color="000000")
    grid_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    white_bold_font = Font(color="FFFFFF", bold=True)
    centered = Alignment(horizontal="center", vertical="center")

    worksheet.sheet_view.showGridLines = False
    worksheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(table_headers))
    title_cell = worksheet["A1"]
    title_cell.value = "WAL MART OPEN AGEING (INV) STATUS SUMMARY"
    title_cell.fill = title_fill
    title_cell.font = Font(color="FFFFFF", bold=True, size=12)
    title_cell.alignment = centered
    worksheet.row_dimensions[1].height = 22

    summary_labels = ["Total Amount"] + years
    for column_index, label in enumerate(summary_labels, start=2):
        value_column = "Total Amount" if label == "Total Amount" else label
        value_cell = worksheet.cell(row=3, column=column_index)
        value_cell.value = _coerce_numeric_like_excel_value(source_value(summary_row, value_column))
        value_cell.number_format = LATEST_TABLES_CURRENCY_FORMAT
        value_cell.fill = title_fill
        value_cell.font = Font(color="FFFFFF", size=12)
        value_cell.alignment = centered
        value_cell.border = grid_border

        label_cell = worksheet.cell(row=4, column=column_index)
        label_cell.value = label
        label_cell.fill = title_fill
        label_cell.font = white_bold_font
        label_cell.alignment = centered
        label_cell.border = grid_border

    worksheet.merge_cells(start_row=6, start_column=1, end_row=6, end_column=len(table_headers))
    section_cell = worksheet["A6"]
    section_cell.value = "YEAR-WISE AMOUNT BREAKDOWN"
    section_cell.fill = section_fill
    section_cell.font = white_bold_font
    section_cell.alignment = centered
    section_cell.border = grid_border

    for column_index, label in enumerate(table_headers, start=1):
        cell = worksheet.cell(row=7, column=column_index)
        cell.value = label
        cell.fill = header_fill
        cell.font = white_bold_font
        cell.alignment = centered
        cell.border = grid_border

    output_row = 8
    for _, source_row in visible_account_rows.iterrows():
        row_values = [source_value(source_row, "Summary / Account Name")]
        row_values.extend(source_value(source_row, column_name) for column_name in report_columns)
        for column_index, value in enumerate(row_values, start=1):
            cell = worksheet.cell(row=output_row, column=column_index)
            is_retail_link_column = retail_link_column is not None and column_index == len(table_headers)
            retail_link_numeric_value = _coerce_numeric_like_excel_value(value) if is_retail_link_column else None
            cell.value = (
                value
                if column_index == 1 or (is_retail_link_column and retail_link_numeric_value is None)
                else retail_link_numeric_value if is_retail_link_column else _coerce_numeric_like_excel_value(value)
            )
            if column_index > 1 and (not is_retail_link_column or retail_link_numeric_value is not None):
                cell.number_format = LATEST_TABLES_CURRENCY_FORMAT
            cell.border = grid_border
            cell.alignment = Alignment(
                horizontal="left" if column_index == 1 else "center" if is_retail_link_column else "right",
                vertical="center",
            )
        output_row += 1

    total_row = total_rows.iloc[0] if not total_rows.empty else None
    for column_index, label in enumerate(table_headers, start=1):
        cell = worksheet.cell(row=output_row, column=column_index)
        if column_index == 1:
            cell.value = source_value(total_row, "Summary / Account Name") if total_row is not None else "Total"
        else:
            source_column = report_columns[column_index - 2]
            raw_value = source_value(total_row, source_column) if total_row is not None else None
            is_retail_link_column = retail_link_column is not None and source_column == retail_link_column
            if raw_value is None and not is_retail_link_column:
                raw_value = visible_account_rows[column_lookup[source_column]].sum()
            retail_link_numeric_value = _coerce_numeric_like_excel_value(raw_value) if is_retail_link_column else None
            cell.value = (
                raw_value
                if is_retail_link_column and retail_link_numeric_value is None
                else retail_link_numeric_value if is_retail_link_column else _coerce_numeric_like_excel_value(raw_value)
            )
            if not is_retail_link_column or retail_link_numeric_value is not None:
                cell.number_format = LATEST_TABLES_CURRENCY_FORMAT
        cell.fill = total_fill
        cell.font = white_bold_font
        cell.border = grid_border
        cell.alignment = Alignment(
            horizontal="left" if column_index == 1 else "center" if is_retail_link_column else "right",
            vertical="center",
        )

    for column_index, label in enumerate(table_headers, start=1):
        if column_index == 1:
            width = 36
        elif label == "Paid & Fully Deduct":
            width = 19
        elif label == "Total Amount":
            width = 16
        elif retail_link_column is not None and label == "Not Found On Portal":
            width = 20
        else:
            width = 14
        worksheet.column_dimensions[get_column_letter(column_index)].width = width
    worksheet.freeze_panes = "B8"


def _write_dbs_status_report(worksheet: Worksheet, table_df: pd.DataFrame) -> None:
    """Create the presentation layout requested for the open-aging DB status export."""
    column_lookup = {str(column).strip().lower(): column for column in table_df.columns}
    report_line_column = column_lookup.get("report line")
    if report_line_column is None:
        raise ValueError("DBs_Status is missing the required 'Report Line' column.")

    def source_value(row: pd.Series, column_name: str) -> object:
        source_column = column_lookup.get(column_name.lower())
        if source_column is None:
            return None
        value = row.get(source_column)
        return None if pd.isna(value) else value

    def report_line(row: pd.Series) -> str:
        value = source_value(row, "Report Line")
        return "" if value is None else str(value).strip()

    years = _dynamic_year_columns(column_lookup)
    report_columns = years + ["Total Amount", "Dispute_Amount", "Write-up Sent"]
    table_headers = ["Reason / Order Column"] + report_columns
    all_accounts_rows = table_df.loc[
        table_df[report_line_column].fillna("").astype(str).str.strip().str.lower().eq("all accounts total")
    ]
    if all_accounts_rows.empty:
        raise ValueError("DBs_Status is missing the required 'All Accounts Total' row.")
    summary_row = all_accounts_rows.iloc[0]

    rows = [row for _, row in table_df.iterrows()]
    block_start_indexes = [
        index for index, row in enumerate(rows) if "(open db" in report_line(row).lower()
    ]
    if not block_start_indexes:
        raise ValueError("DBs_Status does not contain any open-DB account sections.")

    title_fill = PatternFill("solid", fgColor="1F4E78")
    summary_fill = PatternFill("solid", fgColor="D9EAF7")
    spacer_fill = PatternFill("solid", fgColor="2F75B5")
    account_fills = [
        PatternFill("solid", fgColor="1F4E78"),
        PatternFill("solid", fgColor="76933C"),
        PatternFill("solid", fgColor="C0504D"),
        PatternFill("solid", fgColor="8064A2"),
    ]
    thin_side = Side(style="thin", color="000000")
    grid_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    white_bold_font = Font(color="FFFFFF", bold=True)
    centered = Alignment(horizontal="center", vertical="center")

    worksheet.sheet_view.showGridLines = False
    worksheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(table_headers))
    title_cell = worksheet["A1"]
    title_cell.value = "WAL MART OPEN AGEING (DB'S) STATUS SUMMARY"
    title_cell.fill = title_fill
    title_cell.font = Font(color="FFFFFF", bold=True, size=12)
    title_cell.alignment = centered
    worksheet.row_dimensions[1].height = 22

    for column_index, label in enumerate(report_columns, start=1):
        value_cell = worksheet.cell(row=3, column=column_index)
        value_cell.value = _coerce_numeric_like_excel_value(source_value(summary_row, label))
        value_cell.number_format = LATEST_TABLES_CURRENCY_FORMAT
        value_cell.fill = summary_fill
        value_cell.border = grid_border
        value_cell.alignment = centered

        label_cell = worksheet.cell(row=4, column=column_index)
        label_cell.value = label
        label_cell.fill = title_fill
        label_cell.font = white_bold_font
        label_cell.border = grid_border
        label_cell.alignment = centered

    for column_index in range(1, len(table_headers) + 1):
        cell = worksheet.cell(row=5, column=column_index)
        cell.fill = spacer_fill
        cell.border = grid_border

    output_row = 8
    for block_index, start_index in enumerate(block_start_indexes):
        end_index = (
            block_start_indexes[block_index + 1]
            if block_index + 1 < len(block_start_indexes)
            else len(rows)
        )
        account_detail_rows = []
        for source_row in rows[start_index + 1 : end_index]:
            reason = report_line(source_row)
            if not reason or reason.lower() == "reason":
                continue
            account_detail_rows.append(source_row)

        has_account_total = any(
            report_line(source_row).upper() == "ACCOUNT TOTAL"
            for source_row in account_detail_rows
        )
        if not account_detail_rows or not has_account_total:
            continue

        section_fill = account_fills[block_index % len(account_fills)]
        worksheet.merge_cells(
            start_row=output_row,
            start_column=1,
            end_row=output_row,
            end_column=len(table_headers),
        )
        section_cell = worksheet.cell(row=output_row, column=1)
        section_cell.value = report_line(rows[start_index])
        section_cell.fill = section_fill
        section_cell.font = white_bold_font
        section_cell.border = grid_border
        section_cell.alignment = Alignment(horizontal="left", vertical="center")
        output_row += 1

        for column_index, label in enumerate(table_headers, start=1):
            cell = worksheet.cell(row=output_row, column=column_index)
            cell.value = label
            cell.fill = section_fill
            cell.font = white_bold_font
            cell.border = grid_border
            cell.alignment = centered
        output_row += 1

        for source_row in account_detail_rows:
            reason = report_line(source_row)
            is_total = reason.upper() == "ACCOUNT TOTAL"
            for column_index, label in enumerate(table_headers, start=1):
                cell = worksheet.cell(row=output_row, column=column_index)
                if column_index == 1:
                    cell.value = reason
                else:
                    cell.value = _coerce_numeric_like_excel_value(source_value(source_row, label))
                    cell.number_format = LATEST_TABLES_CURRENCY_FORMAT
                cell.border = grid_border
                cell.alignment = Alignment(horizontal="left" if column_index == 1 else "right", vertical="center")
                if is_total:
                    cell.fill = summary_fill
                    cell.font = Font(bold=True)
            output_row += 1

        output_row += 1

    for column_index, label in enumerate(table_headers, start=1):
        width = 28 if column_index == 1 else 16 if label in {"Total Amount", "Dispute_Amount", "Write-up Sent"} else 14
        worksheet.column_dimensions[get_column_letter(column_index)].width = width
    worksheet.freeze_panes = "B10"


def _remove_export_metadata_columns(table_name: str, table_df: pd.DataFrame) -> pd.DataFrame:
    """Hide upload helper columns from ordinary workbook sheets, preserving status reports."""
    if table_name.strip().lower() in {
        DBS_STATUS_TABLE.lower(),
        INVOICES_STATUS_TABLE.lower(),
    }:
        return table_df

    removable_columns = [
        column_name
        for column_name in table_df.columns
        if (
            str(column_name).strip().lower() in EXPORT_HIDDEN_COLUMNS
            or "unnamed" in str(column_name).strip().lower()
        )
    ]
    return table_df.drop(columns=removable_columns) if removable_columns else table_df


def _export_tables_to_workbook(
    repository: SQLServerRepository,
    table_names: list[str],
    output_path: Path,
    include_empty_tables: bool = False,
) -> tuple[list[str], list[str]]:
    # Export results
    exported_tables: list[str] = []
    skipped_tables: list[str] = []
    used_sheet_names: set[str] = set()

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        set_workbook_author_metadata(writer.book)
        for table_name in table_names:
            table_df = repository.read_table(table_name)
            table_df = _remove_export_metadata_columns(table_name, table_df)
            if table_df.empty:
                if include_empty_tables:
                    sheet_name = _safe_sheet_name(table_name, used_sheet_names)
                    table_df.to_excel(writer, sheet_name=sheet_name, index=False)
                    _format_latest_tables_sheet(writer.sheets[sheet_name], table_name, table_df)
                    prepare_pivot_source_sheet(writer.sheets[sheet_name], f"WallMart{sheet_name}")
                    exported_tables.append(table_name)
                    continue
                skipped_tables.append(table_name)
                continue

            if table_name.strip().lower() == STEP_07_CREDIT_APPLIED_TABLE.lower():
                table_df = _coerce_step_07_currency_columns(table_df)
                sheet_name = _safe_sheet_name(table_name, used_sheet_names)
                header_labels, export_rows = _build_step_07_export_rows(table_df)
                if (
                    writer.book.sheetnames == ["Sheet"]
                    and writer.book["Sheet"].max_row == 1
                    and writer.book["Sheet"].max_column == 1
                    and writer.book["Sheet"]["A1"].value is None
                ):
                    writer.book.remove(writer.book["Sheet"])
                worksheet = writer.book.create_sheet(title=sheet_name)
                writer.sheets[sheet_name] = worksheet
                for row in dataframe_to_rows(pd.DataFrame(export_rows, columns=header_labels), index=False, header=True):
                    worksheet.append(row)
                _apply_step_07_export_sheet_format(worksheet, header_labels)
                prepare_pivot_source_sheet(worksheet, f"WallMart{sheet_name}")
                exported_tables.append(table_name)
                continue

            if table_name.strip().lower() == INVOICES_STATUS_TABLE.lower():
                sheet_name = _safe_sheet_name(table_name, used_sheet_names)
                if (
                    writer.book.sheetnames == ["Sheet"]
                    and writer.book["Sheet"].max_row == 1
                    and writer.book["Sheet"].max_column == 1
                    and writer.book["Sheet"]["A1"].value is None
                ):
                    writer.book.remove(writer.book["Sheet"])
                worksheet = writer.book.create_sheet(title=sheet_name)
                writer.sheets[sheet_name] = worksheet
                _write_invoices_status_report(worksheet, table_df)
                exported_tables.append(table_name)
                continue

            if table_name.strip().lower() == DBS_STATUS_TABLE.lower():
                sheet_name = _safe_sheet_name(table_name, used_sheet_names)
                if (
                    writer.book.sheetnames == ["Sheet"]
                    and writer.book["Sheet"].max_row == 1
                    and writer.book["Sheet"].max_column == 1
                    and writer.book["Sheet"]["A1"].value is None
                ):
                    writer.book.remove(writer.book["Sheet"])
                worksheet = writer.book.create_sheet(title=sheet_name)
                writer.sheets[sheet_name] = worksheet
                _write_dbs_status_report(worksheet, table_df)
                exported_tables.append(table_name)
                continue

            # Unique sheet
            sheet_name = _safe_sheet_name(table_name, used_sheet_names)
            table_df.to_excel(writer, sheet_name=sheet_name, index=False)
            _format_latest_tables_sheet(writer.sheets[sheet_name], table_name, table_df)
            prepare_pivot_source_sheet(writer.sheets[sheet_name], f"WallMart{sheet_name}")
            exported_tables.append(table_name)

    return exported_tables, skipped_tables


def _available_automatic_report_tables(
    repository: SQLServerRepository,
    logger: Logger,
) -> list[str]:
    """Return existing automatic-report tables and log optional status-table notices."""
    available_table_names = {
        str(table_name).strip().lower()
        for table_name in repository.list_user_tables()
    }
    report_tables = [
        table_name
        for table_name in AUTOMATED_UPLOAD_REPORT_TABLES
        if table_name.lower() in available_table_names
    ]
    missing_tables = [
        table_name
        for table_name in AUTOMATED_UPLOAD_REPORT_TABLES
        if table_name.lower() not in available_table_names
    ]

    for status_table in (DBS_STATUS_TABLE, INVOICES_STATUS_TABLE):
        if status_table in missing_tables:
            logger.log(
                f"Notice: dbo.{status_table} was not found; proceeding without this report table.",
                "WARNING",
            )

    other_missing_tables = [
        table_name
        for table_name in missing_tables
        if table_name not in {DBS_STATUS_TABLE, INVOICES_STATUS_TABLE}
    ]
    if other_missing_tables:
        logger.log(
            "Notice: optional automatic-report tables were not found; proceeding with available tables: "
            f"{', '.join(other_missing_tables)}",
            "WARNING",
        )

    return report_tables


def _log_empty_status_report_tables(logger: Logger, skipped_tables: list[str]) -> None:
    """Make skipped empty DB/invoice status tables clear in the operator log."""
    skipped_table_names = {str(table_name).strip().lower() for table_name in skipped_tables}
    for status_table in (DBS_STATUS_TABLE, INVOICES_STATUS_TABLE):
        if status_table.lower() in skipped_table_names:
            logger.log(
                f"Notice: dbo.{status_table} is empty; proceeding without this report table.",
                "WARNING",
            )


def _render(
    request: Request,
    message: str | None,
    logs: list[dict[str, str]] | None = None,
    status_code: int = 200,
    selected_company_code: str = DEFAULT_COMPANY_CODE,
    preview_table_name: str | None = None,
    preview_columns: list[str] | None = None,
    preview_rows: list[list[object]] | None = None,
    preview_row_count: int | None = None,
) -> HTMLResponse:
    # Page defaults
    step_01a_overwrite_warning = False
    available_sql_tables: list[str] = []
    latest_batch_id: int | None = None
    sql_database_name = settings.sql_database
    try:
        selected_profile, company_settings = _company_settings(selected_company_code)
        sql_database_name = company_settings.sql_database
        if selected_profile.code not in FILE_ONLY_COMPANY_CODES:
            _, _, repository = build_services(selected_company_code)
            repository.ensure_database()
            step_01a_overwrite_warning = any(
                repository.table_has_rows(table_name)
                for table_name in (STEP_01A_CR_TABLE, STEP_01A_IN_TABLE, STEP_01A_DB_TABLE)
            )
            available_sql_tables = repository.list_user_tables()
            latest_batch_id = repository.get_latest_batch_id()
    except Exception:
        # The page should still load even if SQL Server is unavailable.
        step_01a_overwrite_warning = False
        available_sql_tables = []
        latest_batch_id = None

    formatted_message_html = None
    if message is not None:
        formatted_message_html = escape(message)
        formatted_message_html = re.sub(
            r"([A-Za-z0-9_]+)=([0-9]+)",
            lambda match: (
                f'{match.group(1)}=<span class="message-rowcount-hot">{match.group(2)}</span>'
                if int(match.group(2)) > 0
                else match.group(0)
            ),
            formatted_message_html,
        )
        formatted_message_html = formatted_message_html.replace(
            "[[red]]",
            '<span class="message-section-02">',
        ).replace(
            "[[/red]]",
            "</span>",
        ).replace("\n", "<br>")

    duplicate_upload_messages = [
        str(entry.get("message", ""))
        for entry in (logs or [])
        if str(entry.get("message", "")).startswith("Duplicate upload allowed:")
    ]

    return templates.TemplateResponse(
        request=request,
        name="index.html",
        context={
            "message": message,
            "formatted_message_html": formatted_message_html,
            "logs": logs,
            "sql_server": settings.sql_server,
            "sql_database": sql_database_name,
            "company_profiles": [load_company_profile(code) for code in SUPPORTED_COMPANY_CODES],
            "selected_company_name": selected_profile.display_name if 'selected_profile' in locals() else selected_company_code,
            "runtime_app_path": str(settings.base_dir / "app"),
            "runtime_templates_path": str(settings.base_dir / "templates"),
            "step_01_download_available": _step_01_output_path().exists(),
            "step_01a_download_available": _step_01a_output_path().exists(),
            "step_01a_overwrite_warning": step_01a_overwrite_warning,
            "available_sql_tables": available_sql_tables,
            "default_export_tables": list(LATEST_EXPORT_TABLES),
            "proc_query_sample_sql": PROC_QUERY_SAMPLE_SQL.strip(),
            "selected_company_code": selected_company_code,
            "latest_batch_id": latest_batch_id,
            "duplicate_upload_messages": duplicate_upload_messages,
            "preview_table_name": preview_table_name,
            "preview_columns": preview_columns or [],
            "preview_rows": preview_rows or [],
            "preview_row_count": preview_row_count,
        },
        status_code=status_code,
    )


def _save_upload(target_dir: str, upload: UploadFile, filename: str) -> str:
    path = os.path.join(target_dir, filename)
    with open(path, "wb") as file_handle:
        file_handle.write(upload.file.read())
    return path




def _sync_step_01_to_sql(repository: SQLServerRepository, inquiry_filename: str, result: dict) -> None:
    repository.ensure_database()
    repository.ensure_step_table(STEP_01_INQUIRY_TABLE, result["inquiry_df"])
    repository.log_prepared_columns(
        STEP_01_INQUIRY_TABLE,
        ["Ndatee", "pay_day", "Cclaim"],
        "Step 01",
    )
    batch_id = repository.create_batch(
        raw_filename=inquiry_filename,
        combine_file_count=0,
        status=STEP_01_COMPLETED_STATUS,
        notes="Step 01 inquiry upload.",
    )
    repository.store_raw_rows(batch_id, result["inquiry_df"])
    repository.replace_dynamic_table_rows(STEP_01_INQUIRY_TABLE, result["inquiry_df"], batch_id)
    repository.store_generated_file(
        batch_id,
        STEP_01_NORMALIZED_FILE_KIND,
        STEP_01_OUTPUT_FILE,
        result["output_path"],
    )
    repository.apply_upload_data_types(
        UPLOAD_DATA_TYPES_SQL,
        "step_01_inquiry_rows date and aging decimal columns",
    )
    repository.update_batch_status(batch_id, STEP_01_COMPLETED_STATUS, "Step 01 inquiry uploaded to SQL Server.")




def _run_step_01_flow(
    request: Request,
    logger: Logger,
    orchestrator: WorkflowOrchestrator,
    repository: SQLServerRepository,
    inquiry_path: str,
    inquiry_name: str,
    company_code: str,
    automatic_report: bool = False,
    download_token: str = "",
    amazon_accounts: list[str] | None = None,
):
    company_output_dir = _company_output_dir(company_code)
    company_output_dir.mkdir(parents=True, exist_ok=True)
    if company_code.strip().lower() == "wallmartsummary":
        logger.progress_token = download_token
        job_progress.update(download_token, 10, f"Reading {inquiry_name}")
        try:
            result = WallMartPythonReportProcessor(logger).process(inquiry_path, company_output_dir)
        except Exception as exc:
            logger.log(f"Wal Mart Python processing failed: {exc}", "ERROR")
            return _render(request, f"Wal Mart processing error: {exc}", logger.entries, 500, company_code)

        append_history_event(
            "Wal Mart Python report completed",
            [
                "File: app\\main.py",
                "Function: _run_step_01_flow()",
                f"Input file: {inquiry_name}",
                "Calculation method: Python file-only reporting.",
                "SQL Server was not opened, queried, or refreshed.",
            ],
        )
        logger.log("Wal Mart status workbook completed; no database connection was used.", "SUCCESS")
        job_progress.update(download_token, 100, "Report ready - downloading", done=True, file=_wallmart_python_download_filename())
        return FileResponse(
            path=result["output_path"],
            filename=_wallmart_python_download_filename(),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",

            background=BackgroundTask(_delete_file_after_download, result["output_path"]),
        )

    if company_code.strip().lower() == "amazonaccounts":
        logger.progress_token = download_token
        job_progress.update(download_token, 10, f"Reading {inquiry_name}")
        try:
            result = AmazonAccountSheetProcessor(logger).process(inquiry_path, company_output_dir, amazon_accounts)
        except Exception as exc:
            logger.log(f"Amazon file-only processing failed: {exc}", "ERROR")
            return _render(request, f"Amazon processing error: {exc}", logger.entries, 500, company_code)

        logger.log("Amazon account/type workbook completed; no database connection was used.", "SUCCESS")
        job_progress.update(download_token, 100, "Report ready - downloading", done=True, file=_amazon_download_filename())
        download_filename = _amazon_download_filename()
        if download_token:
            amazon_download_jobs[download_token] = {
                "filename": download_filename,
                "complete": False,
            }
            return RedirectResponse(
                url=f"/download/amazon-account-sheets?download_token={download_token}",
                status_code=303,
            )

        return FileResponse(
            path=result["output_path"],
            filename=download_filename,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",

            background=BackgroundTask(_delete_file_after_download, result["output_path"]),
        )

    success, result = orchestrator.run_step_01(inquiry_path, company_output_dir)
    if not success:
        return _render(request, f"Processing error: {result.get('error', 'Unknown error')}", logger.entries, 500, company_code)

    try:
        repository.ensure_database()
        repository.ensure_step_table(STEP_01_INQUIRY_TABLE, result["inquiry_df"])
        _sync_step_01_to_sql(repository, inquiry_name, result)
        logger.log("Step 01 inquiry uploaded to SQL Server", "SUCCESS")
        if not automatic_report:
            append_history_event(
                "Step 01 completed",
                [
                    "File: app\\main.py",
                    "Function: process_step_01()",
                    f"Input file: {inquiry_name}",
                    "Early prepared columns in step_01_inquiry_rows: Ndatee, pay_day, Cclaim.",
                    f"Refreshed SQL table: {STEP_01_INQUIRY_TABLE} (replace, no duplicate append).",
                    "Step 01 output saved and SQL upload completed.",
                ],
            )
            return _render(
                request,
                (
                    "Step 01 completed.\n"
                    f"Normalized inquiry saved to {result['output_path']}.\n"
                    f"Uploaded [[red]]{{{len(result['inquiry_df'])}}}[[/red]] rows from [[red]]{{1}}[[/red]] file to SQL table {STEP_01_INQUIRY_TABLE}.\n"
                    "Uploaded to SQL Server."
                ),
                logger.entries,
                selected_company_code=company_code,
            )

        report_sql = _automatic_report_sql(company_code)
        if report_sql is None:
            return _render(
                request,
                "Amazon Accounts upload is saved to AmazonAccountsDB. Upload/report column mapping is still required before its automatic report can be generated.",
                logger.entries,
                200,
                company_code,
            )
        automated_sql_steps = (("WalMartSummary report generation", report_sql),)
        for step_name, sql_script in automated_sql_steps:
            logger.log(f"Automatic report flow: {step_name}")
            logger.log("Starting stored procedure: EXEC dbo.sp_Run_WallMartSummary;")
            sql_messages = repository.execute_sql_script_with_messages(sql_script)
            for message_text in sql_messages:
                logger.log(f"{step_name} SQL message: {message_text}")
            logger.log("Stored procedure completed: EXEC dbo.sp_Run_WallMartSummary;", "SUCCESS")

        report_tables = _available_automatic_report_tables(repository, logger)
        if not report_tables:
            raise RuntimeError("No current SQL tables are available for the automatic report export.")

        report_output_path = _latest_tables_output_path(company_code)
        exported_tables, skipped_tables = _export_tables_to_workbook(
            repository,
            report_tables,
            report_output_path,
        )
        if not exported_tables:
            with suppress(Exception):
                report_output_path.unlink()
            raise RuntimeError("The automatic report export did not contain any rows.")

        logger.log(
            f"Automatic report workbook saved: {report_output_path} ({', '.join(exported_tables)})",
            "SUCCESS",
        )
        if skipped_tables:
            logger.log(f"Automatic report skipped empty tables: {', '.join(skipped_tables)}", "WARNING")
            _log_empty_status_report_tables(logger, skipped_tables)
        append_history_event(
            "Step 01 upload and automatic report completed",
            [
                "File: app\\main.py",
                "Function: process_step_01()",
                f"Input file: {inquiry_name}",
                "Early prepared columns in step_01_inquiry_rows: Ndatee, pay_day, Cclaim.",
                f"Refreshed SQL table: {STEP_01_INQUIRY_TABLE} (replace, no duplicate append).",
                "Executed dbo.sp_WallMartSummary_db_and_inv_tabels.",
                "Executed dbo.sp_WallMartSummary_gen_DBs_Status.",
                "Executed dbo.sp_WallMartSummary_Invoice_Status.",
                f"Auto-built latest workbook: {report_output_path}",
                f"Exported tables: {', '.join(exported_tables)}",
                f"Skipped empty tables: {', '.join(skipped_tables) if skipped_tables else 'none'}",
            ],
        )
    except Exception as sql_exc:
        logger.log(f"SQL Server skipped: {sql_exc}", "WARNING")
        append_history_event(
            "Step 01 SQL upload warning",
            [
                "File: app\\main.py",
                "Function: process_step_01()",
                f"Input file: {inquiry_name}",
                f"SQL warning: {sql_exc}",
            ],
        )
        return _render(request, "Step 01 processed but SQL upload failed. See logs.", logger.entries, 500, company_code)

    return FileResponse(
        path=str(report_output_path),
        filename=LATEST_TABLE_EXPORT_FILE,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )




@app.get("/", response_class=HTMLResponse)
async def index(request: Request, company_code: str = DEFAULT_COMPANY_CODE):
    return _render(request, None, selected_company_code=company_code)


@app.post("/preview-table", response_class=HTMLResponse)
async def preview_table(
    request: Request,
    table_name: str = Form(...),
    company_code: str = Form(DEFAULT_COMPANY_CODE),
):
    logger, _, repository = build_services(company_code)

    try:
        repository.ensure_database()
        selected_table_name = table_name.strip()
        if not selected_table_name:
            return _render(
                request,
                "Please select a SQL table to preview.",
                logger.entries,
                status_code=400,
                selected_company_code=company_code,
            )

        preview_df = repository.read_table(selected_table_name)
        full_row_count = len(preview_df.index)
        preview_df = preview_df.fillna("")
        preview_rows = preview_df.astype(str).values.tolist()
        preview_columns = [str(column_name) for column_name in preview_df.columns]
        logger.log(
            f"Previewed table {selected_table_name}: loaded {full_row_count} rows into the scrollable preview.",
            "SUCCESS",
        )
        return _render(
            request,
            f"Database preview loaded for {selected_table_name}.",
            logger.entries,
            selected_company_code=company_code,
            preview_table_name=selected_table_name,
            preview_columns=preview_columns,
            preview_rows=preview_rows,
            preview_row_count=full_row_count,
        )
    except Exception as exc:
        logger.log(f"Table preview failed: {exc}", "ERROR")
        return _render(
            request,
            f"Table preview failed: {exc}",
            logger.entries,
            status_code=500,
            selected_company_code=company_code,
        )


@app.get("/_stcore/health")  # the tools hub portal checks this address for its Online/Offline pill
@app.get("/health")
async def health():
    return {"status": "ok", "app": settings.app_title}


@app.get("/sql/health")
async def sql_health():
    logger, _, repository = build_services()
    try:
        repository.ensure_database()
        return repository.healthcheck()
    except Exception as exc:
        logger.log(f"SQL health check failed: {exc}", "ERROR")
        return JSONResponse({"status": "error", "message": str(exc)}, status_code=500)




@app.get("/download/step-01")
async def download_step_01():
    output_path = _step_01_output_path()
    if not output_path.exists():
        return JSONResponse({"error": "Step 01 output file not found."}, status_code=404)

    return FileResponse(
        path=str(output_path),
        filename=STEP_01_OUTPUT_FILE,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.get("/download/step-01a")
async def download_step_01a():
    output_path = _step_01a_output_path()
    if not output_path.exists():
        return JSONResponse({"error": "Step 01A output file not found."}, status_code=404)

    return FileResponse(
        path=str(output_path),
        filename=STEP_01A_OUTPUT_FILE,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )




@app.get("/download/amazon-account-sheets")
async def download_amazon_account_sheets(download_token: str = ""):
    output_path = _amazon_grouped_output_path()
    if not output_path.exists():
        return JSONResponse({"error": "Amazon account/type workbook not found."}, status_code=404)

    job = amazon_download_jobs.get(download_token) if download_token else None
    return FileResponse(
        path=str(output_path),
        filename=str(job.get("filename", AMAZON_GROUPED_OUTPUT_FILE)) if job else AMAZON_GROUPED_OUTPUT_FILE,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        background=BackgroundTask(_amazon_download_done, output_path, download_token),
    )


@app.post("/progress/cancel")
def progress_cancel(token: str = Form("")):
    """Stop every running process (Excel and PowerShell included); the page then refreshes itself."""
    return JSONResponse({"cancelled": job_progress.cancel_jobs(token)})


@app.get("/progress/status")
async def progress_status(token: str = ""):
    job = job_progress.get(token)
    if job:
        return JSONResponse({**job, "known": True})
    return JSONResponse({"percent": 0, "stage": "Waiting to start...", "done": False, "known": False})


@app.get("/download/amazon-download-status")
async def amazon_download_status(download_token: str):
    job = amazon_download_jobs.get(download_token)
    return JSONResponse({"complete": bool(job and job.get("complete"))})




@app.get("/download/latest-tables")
async def download_latest_tables(company_code: str = DEFAULT_COMPANY_CODE):
    output_path = _latest_tables_output_path(company_code)
    if not output_path.exists():
        return JSONResponse({"error": "Latest tables workbook not found."}, status_code=404)

    return FileResponse(
        path=str(output_path),
        filename=LATEST_TABLE_EXPORT_FILE,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )




def _regular_report_filename(assigned_person: str, source_filename: str) -> str:
    person_part = re.sub(r"[^A-Za-z0-9 _-]+", "_", assigned_person).strip() or "Selected Person"
    source_part = re.sub(r"[^A-Za-z0-9 _-]+", "_", Path(source_filename).stem).strip() or "Full Ageing"
    return f"{person_part} {source_part}.xlsx"


def _excel_cell_value(value: object) -> object:
    if pd.isna(value):
        return None
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()
    return value


def _regular_report_sheet_name(account: str, used_names: set[str]) -> str:
    cleaned = re.sub(r"[\\/*?:\[\]]", "_", account).strip()[:31] or "Account"
    candidate = cleaned
    suffix = 2
    while candidate.lower() in used_names:
        suffix_text = f"_{suffix}"
        candidate = f"{cleaned[:31 - len(suffix_text)]}{suffix_text}"
        suffix += 1
    used_names.add(candidate.lower())
    return candidate


def _style_regular_report_header(worksheet: Worksheet, row_number: int, column_count: int) -> None:
    fill = PatternFill("solid", fgColor="1F4E78")
    font = Font(bold=True, color="FFFFFF")
    border = Border(*(Side(style="thin", color="7F7F7F") for _ in range(4)))
    for column_number in range(1, column_count + 1):
        cell = worksheet.cell(row_number, column_number)
        cell.fill = fill
        cell.font = font
        cell.border = border
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _write_regular_dataframe_sheet(worksheet: Worksheet, data_frame: pd.DataFrame) -> None:
    # Keep the familiar Excel grid instead of a blank print/PDF-style sheet.
    worksheet.sheet_view.showGridLines = True
    for column_number, column_name in enumerate(data_frame.columns, start=1):
        worksheet.cell(1, column_number, column_name)
    _style_regular_report_header(worksheet, 1, len(data_frame.columns))

    for row_number, row_values in enumerate(data_frame.itertuples(index=False, name=None), start=2):
        for column_number, value in enumerate(row_values, start=1):
            worksheet.cell(row_number, column_number, _excel_cell_value(value))

    worksheet.freeze_panes = "A2"
    for column_number, column_name in enumerate(data_frame.columns, start=1):
        width = max(12, min(32, len(str(column_name)) + 3))
        if str(column_name) in {"Name", "Comment", "Status"}:
            width = 28
        worksheet.column_dimensions[get_column_letter(column_number)].width = width
    for column_name in ("Date", "Due Date"):
        if column_name in data_frame.columns:
            column_number = data_frame.columns.get_loc(column_name) + 1
            for row_number in range(2, worksheet.max_row + 1):
                worksheet.cell(row_number, column_number).number_format = "mm/dd/yyyy"
    currency_columns = ("Total", "Current", "1-30 Days", "31-60 Days", "61-90 Days", "Over 90 Days")
    for column_name in currency_columns:
        if column_name in data_frame.columns:
            column_number = data_frame.columns.get_loc(column_name) + 1
            for row_number in range(2, worksheet.max_row + 1):
                worksheet.cell(row_number, column_number).number_format = '$#,##0.00;[Red]($#,##0.00)'
    prepare_pivot_source_sheet(worksheet, f"Regular{worksheet.title}")
    # The shared Pivot-source helper applies its own date style.  Regular
    # Accounts requires MM/DD/YYYY, so set it last for this flow only.
    for column_name in ("Date", "Due Date"):
        if column_name in data_frame.columns:
            column_number = data_frame.columns.get_loc(column_name) + 1
            for row_number in range(2, worksheet.max_row + 1):
                worksheet.cell(row_number, column_number).number_format = "mm/dd/yyyy"
    # Paid rows turn green (fill only, no data change, so PivotTables are unaffected).
    paid_text_indexes = [
        index for index, name in enumerate(data_frame.columns)
        if _column_key_for_report(str(name)) in {"status", "comment"}
    ]
    if paid_text_indexes:
        for row_number, row_values in enumerate(data_frame.itertuples(index=False, name=None), start=2):
            if _row_marked_paid(row_values, paid_text_indexes):
                for column_number in range(1, len(data_frame.columns) + 1):
                    worksheet.cell(row_number, column_number).fill = PAID_ROW_FILL


PAST_DUE_STATUS_HEADER = "Due Status"
PAID_KEYWORD_PATTERN = re.compile(r"(?<!not )\bpaid\b", re.IGNORECASE)
PAID_ROW_FILL = PatternFill("solid", fgColor="C6EFCE")


def _row_marked_paid(row_values: tuple, text_indexes: list[int]) -> bool:
    """True when Status or Comment holds Paid as a standalone word (not Unpaid, not 'not paid')."""
    return any(
        isinstance(row_values[index], str) and PAID_KEYWORD_PATTERN.search(row_values[index])
        for index in text_indexes
    )


def _write_regular_past_due_sheet(worksheet: Worksheet, data_frame: pd.DataFrame) -> None:
    """Write the Past-due invoices sheet with a live Due Status formula and a date control box on the right."""
    due_name = next((name for name in data_frame.columns if _column_key_for_report(str(name)) == "due date"), None)
    due_index = list(data_frame.columns).index(due_name) if due_name is not None else len(data_frame.columns) - 1
    headers = list(data_frame.columns)
    headers.insert(due_index + 1, PAST_DUE_STATUS_HEADER)
    status_column = due_index + 2  # 1-based, right next to Due Date
    due_letter = get_column_letter(due_index + 1)
    last_column = len(headers)

    worksheet.sheet_view.showGridLines = True
    for column_number, header in enumerate(headers, start=1):
        worksheet.cell(1, column_number, header)
    _style_regular_report_header(worksheet, 1, last_column)

    box_label_column = last_column + 2
    box_value_column = last_column + 3
    yesterday_cell = f"${get_column_letter(box_value_column)}$3"
    today_cell = f"${get_column_letter(box_value_column)}$4"
    tomorrow_cell = f"${get_column_letter(box_value_column)}$5"
    cutoff_cell = f"${get_column_letter(box_value_column)}$6"

    paid_text_indexes = [
        index for index, name in enumerate(data_frame.columns)
        if _column_key_for_report(str(name)) in {"status", "comment"}
    ]
    for row_number, row_values in enumerate(data_frame.itertuples(index=False, name=None), start=2):
        source_values = iter(row_values)
        is_paid = _row_marked_paid(row_values, paid_text_indexes)
        for column_number in range(1, last_column + 1):
            if column_number == status_column and is_paid:
                worksheet.cell(row_number, column_number, "Paid")
            elif column_number == status_column:
                due_ref = f"{due_letter}{row_number}"
                worksheet.cell(
                    row_number,
                    column_number,
                    f'=IF({due_ref}="","",IF({due_ref}<{yesterday_cell},"Overdue",IF({due_ref}={yesterday_cell},"Passed Yesterday",'
                    f'IF({due_ref}={today_cell},"Due Today",IF({due_ref}={tomorrow_cell},"Due Tomorrow",'
                    f'IF({due_ref}<={cutoff_cell},"Due soon","Later"))))))',
                )
            else:
                worksheet.cell(row_number, column_number, _excel_cell_value(next(source_values)))

    last_row = max(2, len(data_frame.index) + 1)
    for row_number, row_values in enumerate(data_frame.itertuples(index=False, name=None), start=2):
        if _row_marked_paid(row_values, paid_text_indexes):
            for column_number in range(1, last_column + 1):
                worksheet.cell(row_number, column_number).fill = PAID_ROW_FILL
    worksheet.auto_filter.ref = f"A1:{get_column_letter(last_column)}{last_row}"
    worksheet.freeze_panes = "A2"
    for column_number, header in enumerate(headers, start=1):
        width = max(12, min(32, len(str(header)) + 3))
        if str(header) in {"Name", "Comment", "Status", PAST_DUE_STATUS_HEADER}:
            width = 28 if str(header) != PAST_DUE_STATUS_HEADER else 16
        worksheet.column_dimensions[get_column_letter(column_number)].width = width
    for column_name in ("Date", "Due Date"):
        if column_name in headers:
            column_number = headers.index(column_name) + 1
            for row_number in range(2, last_row + 1):
                worksheet.cell(row_number, column_number).number_format = "mm/dd/yyyy"
    for column_name in ("Total", "Current", "1-30 Days", "31-60 Days", "61-90 Days", "Over 90 Days"):
        if column_name in headers:
            column_number = headers.index(column_name) + 1
            for row_number in range(2, last_row + 1):
                worksheet.cell(row_number, column_number).number_format = '$#,##0.00;[Red]($#,##0.00)'

    # Control box: every date the Due Status formulas read lives here and refreshes whenever Excel opens the file.
    thin = Side(style="thin", color="808080")
    box_border = Border(left=thin, right=thin, top=thin, bottom=thin)
    box_fill = PatternFill("solid", fgColor="FFF2CC")
    worksheet.cell(2, box_label_column, "DUE STATUS DATES (auto-update)").font = Font(bold=True)
    worksheet.merge_cells(start_row=2, start_column=box_label_column, end_row=2, end_column=box_value_column)
    box_rows = (
        (3, "Yesterday", "=TODAY()-1"),
        (4, "Today", "=TODAY()"),
        (5, "Tomorrow", "=TODAY()+1"),
        (6, "Next Monday (cutoff)", "=TODAY()+MOD(8-WEEKDAY(TODAY(),2),7)"),
    )
    for row_number, label, formula in box_rows:
        worksheet.cell(row_number, box_label_column, label).font = Font(bold=True)
        value_cell = worksheet.cell(row_number, box_value_column, formula)
        value_cell.number_format = "mm/dd/yyyy"
        value_cell.alignment = Alignment(horizontal="center")
    worksheet.cell(7, box_label_column, "Passed Yesterday: due Yesterday | Due Today | Due Tomorrow | Due soon: after Tomorrow up to Next Monday | Overdue: due before Yesterday")
    worksheet.cell(7, box_label_column).alignment = Alignment(wrap_text=True, vertical="top")
    worksheet.merge_cells(start_row=7, start_column=box_label_column, end_row=7, end_column=box_value_column)
    worksheet.row_dimensions[7].height = 64
    for row_number in range(2, 8):
        for column_number in (box_label_column, box_value_column):
            cell = worksheet.cell(row_number, column_number)
            cell.border = box_border
            cell.fill = box_fill
    worksheet.column_dimensions[get_column_letter(box_label_column - 1)].width = 3
    worksheet.column_dimensions[get_column_letter(box_label_column)].width = 24
    worksheet.column_dimensions[get_column_letter(box_value_column)].width = 16


def _regular_source_column_letter(data_frame: pd.DataFrame, column_key: str) -> str | None:
    for column_number, column_name in enumerate(data_frame.columns, start=1):
        if _column_key_for_report(str(column_name)) == column_key:
            return get_column_letter(column_number)
    return None


def _regular_summary_formula(
    ageing_rows: pd.DataFrame,
    sum_column_key: str,
    group_column_key: str,
    group_cell: str,
    ref_marker: str,
) -> str | None:
    """Return an Excel SUMIFS formula linked to the PivotTable-ready AGEING SHEET."""
    sum_column = _regular_source_column_letter(ageing_rows, sum_column_key)
    group_column = _regular_source_column_letter(ageing_rows, group_column_key)
    if not sum_column or not group_column:
        return None
    final_row = len(ageing_rows) + 1
    sheet = "'AGEING SHEET'"
    if ref_marker.strip().casefold() == "cr":
        ref_type_column = _regular_source_column_letter(ageing_rows, "ref type")
        if ref_type_column:
            return (
                f'=SUMIFS({sheet}!${sum_column}$2:${sum_column}${final_row},'
                f'{sheet}!${group_column}$2:${group_column}${final_row},{group_cell},'
                f'{sheet}!${ref_type_column}$2:${ref_type_column}${final_row},"CR")'
            )
    ref_column = _regular_source_column_letter(ageing_rows, "ref#")
    if not ref_column:
        return None
    return (
        f'=SUMIFS({sheet}!${sum_column}$2:${sum_column}${final_row},'
        f'{sheet}!${group_column}$2:${group_column}${final_row},{group_cell},'
        f'{sheet}!${ref_column}$2:${ref_column}${final_row},"*{ref_marker.upper()}*")'
    )


def _regular_year_total_formula(ageing_rows: pd.DataFrame, year_cell: str, ref_marker: str) -> str | None:
    total_column = _regular_source_column_letter(ageing_rows, "total")
    date_column = _regular_source_column_letter(ageing_rows, "date")
    if not total_column or not date_column:
        return None
    final_row = len(ageing_rows) + 1
    sheet = "'AGEING SHEET'"
    if ref_marker.strip().casefold() == "cr":
        ref_type_column = _regular_source_column_letter(ageing_rows, "ref type")
        if not ref_type_column:
            return None
        return (
            f'=SUMIFS({sheet}!${total_column}$2:${total_column}${final_row},'
            f'{sheet}!${ref_type_column}$2:${ref_type_column}${final_row},"CR",'
            f'{sheet}!${date_column}$2:${date_column}${final_row},">="&DATE({year_cell},1,1),'
            f'{sheet}!${date_column}$2:${date_column}${final_row},"<"&DATE({year_cell}+1,1,1))'
        )
    ref_column = _regular_source_column_letter(ageing_rows, "ref#")
    if not ref_column:
        return None
    return (
        f'=SUMIFS({sheet}!${total_column}$2:${total_column}${final_row},'
        f'{sheet}!${ref_column}$2:${ref_column}${final_row},"*{ref_marker.upper()}*",'
        f'{sheet}!${date_column}$2:${date_column}${final_row},">="&DATE({year_cell},1,1),'
        f'{sheet}!${date_column}$2:${date_column}${final_row},"<"&DATE({year_cell}+1,1,1))'
    )




def _write_regular_invoice_summary_sheet(
    workbook: Workbook,
    assigned_person: str,
    year_totals: pd.DataFrame,
    ageing_summary: pd.DataFrame,
    ageing_rows: pd.DataFrame,
    account_analyst_lookup: dict[str, dict[str, str]],
    sheet_name: str = "INV SUMMARY",
    ref_marker: str = "in",
    ageing_title_text: str = "Invoice By Ageing",
    include_inv_companion_columns: bool = True,
    include_inv_cr_column: bool = True,
    compact_ageing_columns: bool = False,
) -> None:
    """Write the INV or CR summary sheet with matching native-Pivot source rows."""
    worksheet = workbook.create_sheet(sheet_name)
    worksheet.sheet_view.showGridLines = True
    reference_blue = PatternFill("solid", fgColor=Color(theme=4, tint=0.7999816888943144))
    dark_blue = PatternFill("solid", fgColor="1F4E78")
    grand_total_blue = PatternFill("solid", fgColor="B8CCE4")
    thin = Side(style="thin", color="000000")
    medium = Side(style="medium", color="000000")
    currency_format = '_(\\$* #,##0.00_);_(\\$* \\(#,##0.00\\);_(\\$* "-"??_);_(@_)'

    year_values = [int(year) for year in year_totals.get("year", pd.Series(dtype=int)).tolist()]
    person_cell = worksheet.cell(2, 1, assigned_person)
    person_cell.fill = reference_blue
    person_cell.font = Font(name="Aptos Narrow", size=16, bold=True)
    person_cell.alignment = Alignment(vertical="center")
    person_cell.border = Border(left=medium, top=medium)
    ref_type_cell = worksheet.cell(2, 2, ref_marker.upper())
    ref_type_cell.fill = reference_blue
    ref_type_cell.font = Font(name="Aptos Narrow", size=16, bold=True)
    ref_type_cell.alignment = Alignment(horizontal="center", vertical="center")
    ref_type_cell.border = Border(left=thin, right=thin, top=medium, bottom=thin)

    account_holder_cell = worksheet.cell(3, 1, "ACCOUNT HOLDER")
    account_holder_cell.fill = reference_blue
    account_holder_cell.font = Font(name="Aptos Narrow", size=10, bold=True)
    account_holder_cell.alignment = Alignment(vertical="center")
    account_holder_cell.border = Border(left=medium, bottom=medium)
    year_label_cell = worksheet.cell(3, 2, "Year")
    year_label_cell.fill = reference_blue
    year_label_cell.font = Font(name="Aptos Narrow", size=12, bold=True)
    year_label_cell.alignment = Alignment(horizontal="center", vertical="center")
    year_label_cell.border = Border(left=thin, right=thin, top=thin, bottom=medium)
    for column_number, year in enumerate(year_values, start=3):
        total = year_totals.loc[year_totals["year"] == year, "total"].iloc[0]
        total_cell = worksheet.cell(
            2,
            column_number,
            _regular_year_total_formula(ageing_rows, f"{get_column_letter(column_number)}3", ref_marker)
            or _excel_cell_value(total),
        )
        total_cell.fill = reference_blue
        total_cell.font = Font(name="Aptos Narrow", size=16, bold=True)
        total_cell.alignment = Alignment(vertical="center")
        total_cell.border = Border(left=thin, right=thin, top=medium, bottom=thin)
        total_cell.number_format = currency_format
        year_cell = worksheet.cell(3, column_number, year)
        year_cell.fill = reference_blue
        year_cell.font = Font(name="Aptos Narrow", size=12)
        year_cell.alignment = Alignment(horizontal="center", vertical="center")
        year_cell.border = Border(left=thin, right=thin, top=thin, bottom=medium)

    worksheet.merge_cells(start_row=5, start_column=2, end_row=5, end_column=9)
    ageing_title = worksheet.cell(5, 2, ageing_title_text)
    ageing_title.fill = dark_blue
    ageing_title.font = Font(name="Aptos Narrow", size=20, bold=True, color="FFFFFF")
    ageing_title.alignment = Alignment(horizontal="center", vertical="center")
    ageing_title.border = Border()
    headers = [
        "Acct",
        "Name",
        "Ref Type",
        "Sum of Current",
        "Sum of 1-30 Days",
        "Sum of 31-60 Days",
        "Sum of 61-90 Days",
        "Sum of Over 90 Days",
        "Sum of Total",
    ]
    if compact_ageing_columns:
        headers = [
            "Acct", "Sum of Current", "Sum of 1-30 Days", "Sum of 31-60 Days",
            "Sum of 61-90 Days", "Sum of Over 90 Days", "Sum of Total",
        ]
    elif include_inv_companion_columns:
        if include_inv_cr_column:
            headers.append("CR")
        headers.extend(["TERMS#", "SALEMAN"])
    summary_value_keys = {
        "Sum of Current": "current",
        "Sum of 1-30 Days": "1-30 days",
        "Sum of 31-60 Days": "31-60 days",
        "Sum of 61-90 Days": "61-90 days",
        "Sum of Over 90 Days": "over 90 days",
        "Sum of Total": "total",
    }
    if not compact_ageing_columns and include_inv_companion_columns and include_inv_cr_column:
        summary_value_keys["CR"] = "cr"
    header_row = 7
    data_start_row = header_row + 1
    for column_number, header in enumerate(headers, start=1):
        header_cell = worksheet.cell(header_row, column_number, header)
        header_cell.fill = dark_blue
        header_cell.font = Font(name="Aptos Narrow", size=11, bold=True, color="FFFFFF")
        header_cell.alignment = Alignment(vertical="center")
        header_cell.border = Border(
            left=medium if column_number == 1 else thin,
            right=thin,
            top=medium,
            bottom=thin,
        )

    for row_number, (_, row_values) in enumerate(ageing_summary.iterrows(), start=data_start_row):
        account = _excel_cell_value(row_values.get("acct", ""))
        account_cell = worksheet.cell(row_number, 1, account)
        account_cell.font = Font(name="Aptos Narrow", size=11, bold=True)
        account_cell.alignment = Alignment(horizontal="left")
        account_cell.border = Border(left=medium, right=thin, top=thin, bottom=thin)
        amount_start_column = 2 if compact_ageing_columns else 4
        if not compact_ageing_columns:
            name_cell = worksheet.cell(row_number, 2, _excel_cell_value(row_values.get("name", "")))
            name_cell.alignment = Alignment(horizontal="left")
            name_cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
            ref_type_cell = worksheet.cell(row_number, 3, ref_marker.upper())
            ref_type_cell.alignment = Alignment(horizontal="left")
            ref_type_cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
        for column_number, header in enumerate(headers[amount_start_column - 1:], start=amount_start_column):
            source_key = summary_value_keys.get(header)
            is_inv_cr_column = include_inv_companion_columns and include_inv_cr_column and header == "CR"
            formula = _regular_summary_formula(
                ageing_rows,
                "total" if is_inv_cr_column else source_key or "",
                "acct",
                f"$A{row_number}",
                "cr" if is_inv_cr_column else ref_marker,
            ) if source_key else None
            if is_inv_cr_column and formula:
                formula = f"=-ABS({formula[1:]})"
            value_cell = worksheet.cell(
                row_number,
                column_number,
                formula or (_excel_cell_value(row_values.get(source_key, 0)) if source_key else None),
            )
            value_cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
            if source_key:
                value_cell.number_format = currency_format
            if is_inv_cr_column and pd.to_numeric(row_values.get("cr", 0), errors="coerce") != 0:
                value_cell.font = Font(name="Aptos Narrow", size=11, color="FF0000")
        if not compact_ageing_columns and include_inv_companion_columns:
            analyst_values = account_analyst_lookup.get(str(account).strip().casefold(), {})
            terms_column = 11 if include_inv_cr_column else 10
            worksheet.cell(row_number, terms_column, analyst_values.get("terms", ""))
            worksheet.cell(row_number, terms_column + 1, analyst_values.get("slm", ""))

    total_row = max(data_start_row, worksheet.max_row + 1)
    total_label = worksheet.cell(total_row, 1, "Grand Total")
    total_label.font = Font(name="Aptos Narrow", size=11, bold=True)
    total_label.alignment = Alignment(horizontal="left")
    total_label.border = Border(left=medium, right=thin, top=thin, bottom=medium)
    for column_number in range(2, 2 if compact_ageing_columns else 4):
        worksheet.cell(total_row, column_number).border = Border(left=thin, right=thin, top=thin, bottom=medium)
        worksheet.cell(total_row, column_number).font = Font(name="Aptos Narrow", size=11, bold=True)
    last_amount_column = 7 if compact_ageing_columns else 10 if include_inv_companion_columns and include_inv_cr_column else 9
    amount_start_column = 2 if compact_ageing_columns else 4
    for column_number in range(amount_start_column, last_amount_column + 1):
        letter = get_column_letter(column_number)
        total_cell = worksheet.cell(total_row, column_number, f"=SUM({letter}{data_start_row}:{letter}{total_row - 1})")
        total_cell.border = Border(left=thin, right=thin, top=thin, bottom=medium)
        total_cell.number_format = currency_format
        total_cell.font = Font(name="Aptos Narrow", size=11, bold=True)
        if include_inv_companion_columns and include_inv_cr_column and column_number == 10:
            total_cell.fill = reference_blue
    for column_number in range(last_amount_column + 1, len(headers) + 1):
        total_cell = worksheet.cell(total_row, column_number)
        total_cell.border = Border(left=thin, right=thin, top=thin, bottom=medium)
        total_cell.font = Font(name="Aptos Narrow", size=11, bold=True)
    for column_number in range(1, len(headers) + 1):
        total_cell = worksheet.cell(total_row, column_number)
        total_cell.fill = grand_total_blue
        total_cell.font = Font(name="Aptos Narrow", size=11, bold=True)
    cr_total = pd.to_numeric(ageing_summary.get("cr", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()
    if include_inv_companion_columns and include_inv_cr_column and cr_total != 0:
        worksheet.cell(total_row, 10).font = Font(name="Aptos Narrow", size=11, bold=True, color="FF0000")

    # The interactive PivotTable header is on row 6.  Keep it visible without
    # freezing any account data rows below it.
    worksheet.freeze_panes = None
    for column_letter, width in {
        "A": 22.0, "B": 26.0, "C": 11.14, "D": 17.57, "E": 13.0,
        "F": 13.0, "G": 13.0, "H": 19.57, "I": 17.57, "J": 19.14,
        "K": 16.43, "L": 28.57,
    }.items():
        worksheet.column_dimensions[column_letter].width = width
    worksheet.row_dimensions[2].height = 24
    worksheet.row_dimensions[3].height = 20
    worksheet.row_dimensions[5].height = 28


def _write_regular_db_summary_template(
    workbook: Workbook,
    assigned_person: str,
    year_totals: pd.DataFrame,
    ageing_summary: pd.DataFrame,
    reason_summary: pd.DataFrame,
    ageing_rows: pd.DataFrame,
) -> None:
    """Write DB'S SUMMARY using the approved Ref# contains DB rules."""
    worksheet = workbook.create_sheet("DB'S SUMMARY")
    worksheet.sheet_view.showGridLines = True
    reference_blue = PatternFill("solid", fgColor=Color(theme=4, tint=0.7999816888943144))
    dark_blue = PatternFill("solid", fgColor="1F4E78")
    grand_total_blue = PatternFill("solid", fgColor="B8CCE4")
    thin = Side(style="thin", color="000000")
    medium = Side(style="medium", color="000000")
    currency_format = '_(\\$* #,##0.00_);_(\\$* \\(#,##0.00\\);_(\\$* "-"??_);_(@_)'
    years = [int(year) for year in year_totals.get("year", pd.Series(dtype=int)).tolist()]
    ageing_headers = [
        "Acct", "Name", "Ref Type", "Sum of Current", "Sum of 1-30 Days",
        "Sum of 31-60 Days", "Sum of 61-90 Days", "Sum of Over 90 Days", "Sum of Total",
    ]
    reason_headers = [
        "Order#", "Ref Type", "Sum of Current", "Sum of 1-30 Days",
        "Sum of 31-60 Days", "Sum of 61-90 Days", "Sum of Over 90 Days", "Sum of Total",
    ]

    for column_number in range(1, max(9, len(years) + 2) + 1):
        top_cell = worksheet.cell(2, column_number)
        top_cell.fill = reference_blue
        top_cell.border = Border(
            left=medium if column_number == 1 else thin,
            right=medium if column_number == max(9, len(years) + 2) else thin,
            top=medium,
            bottom=thin,
        )
        bottom_cell = worksheet.cell(3, column_number)
        bottom_cell.fill = reference_blue
        bottom_cell.border = Border(
            left=medium if column_number == 1 else thin,
            right=medium if column_number == max(9, len(years) + 2) else thin,
            top=thin,
            bottom=medium,
        )
    worksheet.cell(2, 1, assigned_person).font = Font(name="Aptos Narrow", size=16, bold=True)
    worksheet.cell(2, 1).alignment = Alignment(vertical="center")
    worksheet.cell(2, 2, "DB").font = Font(name="Aptos Narrow", size=16, bold=True)
    worksheet.cell(2, 2).alignment = Alignment(horizontal="center", vertical="center")
    worksheet.cell(3, 2, "Year").font = Font(name="Aptos Narrow", size=12, bold=True)
    worksheet.cell(3, 2).alignment = Alignment(horizontal="center", vertical="center")
    worksheet.cell(3, 1, "ACCOUNT HOLDER").font = Font(name="Aptos Narrow", size=10, bold=True)
    worksheet.cell(3, 1).alignment = Alignment(vertical="center")
    for column_number, year in enumerate(years, start=3):
        total = year_totals.loc[year_totals["year"] == year, "total"].iloc[0]
        total_cell = worksheet.cell(
            2,
            column_number,
            _regular_year_total_formula(ageing_rows, f"{get_column_letter(column_number)}3", "db")
            or _excel_cell_value(total),
        )
        total_cell.font = Font(name="Aptos Narrow", size=16, bold=True)
        total_cell.number_format = currency_format
        total_cell.alignment = Alignment(vertical="center")
        year_cell = worksheet.cell(3, column_number, year)
        year_cell.font = Font(name="Aptos Narrow", size=12)
        year_cell.alignment = Alignment(horizontal="center", vertical="center")

    def write_section(title_row: int, header_row: int, total_row: int, title: str, headers: list[str]) -> None:
        worksheet.merge_cells(start_row=title_row, start_column=2, end_row=title_row, end_column=len(headers))
        title_cell = worksheet.cell(title_row, 2, title)
        title_cell.fill = dark_blue
        title_cell.font = Font(name="Aptos Narrow", size=20, bold=True, color="FFFFFF")
        title_cell.alignment = Alignment(horizontal="center", vertical="center")
        title_cell.border = (
            Border()
            if title_row == 5
            else Border(left=medium, right=medium, top=medium, bottom=medium)
        )
        worksheet.row_dimensions[title_row].height = 28

        for column_number, header in enumerate(headers, start=1):
            header_cell = worksheet.cell(header_row, column_number, header)
            header_cell.fill = dark_blue
            header_cell.font = Font(name="Aptos Narrow", size=11, bold=True, color="FFFFFF")
            header_cell.border = Border(
                left=medium if column_number == 1 else thin,
                right=medium if column_number == len(headers) else thin,
                top=medium,
                bottom=thin,
            )
        total_cell = worksheet.cell(total_row, 1, "Grand Total")
        total_cell.fill = grand_total_blue
        total_cell.font = Font(name="Aptos Narrow", size=11, bold=True)
        total_cell.border = Border(left=medium, top=thin, bottom=medium)
        for column_number in range(2, len(headers) + 1):
            cell = worksheet.cell(total_row, column_number)
            cell.fill = grand_total_blue
            cell.font = Font(name="Aptos Narrow", size=11, bold=True)
            cell.border = Border(
                left=thin,
                right=medium if column_number == len(headers) else thin,
                top=thin,
                bottom=medium,
            )

    ageing_start_row = 8
    ageing_total_row = ageing_start_row + len(ageing_summary)
    write_section(5, 7, ageing_total_row, "DB's By Ageing", ageing_headers)
    summary_value_keys = {
        4: "current", 5: "1-30 days", 6: "31-60 days", 7: "61-90 days",
        8: "over 90 days", 9: "total",
    }
    for row_number, (_, values) in enumerate(ageing_summary.iterrows(), start=ageing_start_row):
        worksheet.cell(row_number, 1, _excel_cell_value(values.get("acct", ""))).font = Font(
            name="Aptos Narrow", size=11, bold=True
        )
        worksheet.cell(row_number, 2, _excel_cell_value(values.get("name", "")))
        worksheet.cell(row_number, 3, "DB")
        for column_number, source_key in summary_value_keys.items():
            formula = _regular_summary_formula(ageing_rows, source_key, "acct", f"$A{row_number}", "db")
            value_cell = worksheet.cell(row_number, column_number, formula or _excel_cell_value(values.get(source_key, 0)))
            value_cell.number_format = currency_format
        for column_number in range(1, len(ageing_headers) + 1):
            cell = worksheet.cell(row_number, column_number)
            cell.border = Border(
                left=medium if column_number == 1 else thin,
                right=medium if column_number == len(ageing_headers) else thin,
                top=thin,
                bottom=thin,
            )
    for column_number in range(4, len(ageing_headers) + 1):
        letter = get_column_letter(column_number)
        total_cell = worksheet.cell(ageing_total_row, column_number, f"=SUM({letter}{ageing_start_row}:{letter}{ageing_total_row - 1})")
        total_cell.number_format = currency_format

    reason_title_row = ageing_total_row + 2
    reason_header_row = reason_title_row + 3
    reason_start_row = reason_header_row + 1
    reason_total_row = reason_start_row + len(reason_summary)
    write_section(reason_title_row, reason_header_row, reason_total_row, "DB'S BY REASON", reason_headers)
    for row_number, (_, values) in enumerate(reason_summary.iterrows(), start=reason_start_row):
        worksheet.cell(row_number, 1, _excel_cell_value(values.get("order#", ""))).font = Font(
            name="Aptos Narrow", size=11, bold=True
        )
        worksheet.cell(row_number, 2, "DB")
        for column_number, source_key in {
            3: "current", 4: "1-30 days", 5: "31-60 days", 6: "61-90 days",
            7: "over 90 days", 8: "total",
        }.items():
            formula = _regular_summary_formula(ageing_rows, source_key, "order#", f"$A{row_number}", "db")
            value_cell = worksheet.cell(row_number, column_number, formula or _excel_cell_value(values.get(source_key, 0)))
            value_cell.number_format = currency_format
        for column_number in range(1, len(reason_headers) + 1):
            cell = worksheet.cell(row_number, column_number)
            cell.border = Border(
                left=medium if column_number == 1 else thin,
                right=medium if column_number == len(reason_headers) else thin,
                top=thin,
                bottom=thin,
            )
    for column_number in range(3, len(reason_headers) + 1):
        letter = get_column_letter(column_number)
        total_cell = worksheet.cell(reason_total_row, column_number, f"=SUM({letter}{reason_start_row}:{letter}{reason_total_row - 1})")
        total_cell.number_format = currency_format

    for column_letter, width in {
        "A": 22.0, "B": 26.0, "C": 11.14, "D": 17.57,
        "E": 13.0, "F": 13.0, "G": 13.0, "H": 19.57,
        "I": 17.57,
    }.items():
        worksheet.column_dimensions[column_letter].width = width
    worksheet.row_dimensions[2].height = 24
    worksheet.row_dimensions[3].height = 20


def _column_key_for_report(column_name: str) -> str:
    return " ".join(column_name.strip().casefold().split())


def _build_regular_accounts_report(
    processor: RegularAccountsProcessor,
    workbook_path: str,
    assigned_person: str,
    source_filename: str,
    token: str = "",
    browser_date: str = "",
) -> tuple[Path, int]:
    job_progress.update(token, 15, f"Reading Full Ageing sheet from {source_filename}")
    selected_rows = processor.selected_person_rows(workbook_path, assigned_person)
    job_progress.update(token, 35, f"Preparing summaries for {assigned_person} ({len(selected_rows.index)} rows)")
    follow_up_rows = processor.follow_up_entries(selected_rows)
    invoice_summary = processor.summary_rows(selected_rows, "in")
    invoice_year_totals = processor.summary_year_totals(selected_rows, "in", "INV SUMMARY")
    cr_summary = processor.summary_rows(selected_rows, "cr")
    cr_year_totals = processor.summary_year_totals(selected_rows, "cr", "CR'S SUMMARY")
    db_summary = processor.summary_rows(selected_rows, "db")
    db_year_totals = processor.summary_year_totals(selected_rows, "db", "DB'S SUMMARY")
    db_reason_summary = processor.db_reason_summary(selected_rows)
    account_analyst_lookup = processor.account_analyst_lookup(workbook_path)
    ageing_rows = processor.ageing_rows(selected_rows)
    try:
        report_today = pd.Timestamp(browser_date.strip()) if browser_date.strip() else pd.Timestamp.today()
    except ValueError:
        report_today = pd.Timestamp.today()
    past_due_rows = processor.past_due_rows(selected_rows, report_today)
    account_rows = processor.account_detail_rows(selected_rows)

    workbook = Workbook()
    workbook.remove(workbook.active)
    _write_regular_invoice_summary_sheet(
        workbook,
        assigned_person,
        invoice_year_totals,
        invoice_summary,
        ageing_rows,
        account_analyst_lookup,
        include_inv_cr_column=False,
    )
    _write_regular_invoice_summary_sheet(
        workbook,
        assigned_person,
        cr_year_totals,
        cr_summary,
        ageing_rows,
        account_analyst_lookup,
        sheet_name="CR'S SUMMARY",
        ref_marker="cr",
        ageing_title_text="CR By Ageing",
        include_inv_companion_columns=False,
    )
    _write_regular_db_summary_template(
        workbook,
        assigned_person,
        db_year_totals,
        db_summary,
        db_reason_summary,
        ageing_rows,
    )

    follow_up_sheet = workbook.create_sheet("FOLLOW-UP ENTRIES")
    _write_regular_dataframe_sheet(follow_up_sheet, follow_up_rows)
    past_due_sheet = workbook.create_sheet("Past-due invoices")
    _write_regular_past_due_sheet(past_due_sheet, past_due_rows)
    ageing_sheet = workbook.create_sheet("AGEING SHEET")
    _write_regular_dataframe_sheet(ageing_sheet, ageing_rows)

    used_names = {sheet_name.lower() for sheet_name in workbook.sheetnames}
    for account, rows in account_rows.items():
        account_sheet = workbook.create_sheet(_regular_report_sheet_name(account, used_names))
        _write_regular_dataframe_sheet(account_sheet, rows)

    # Excel cannot build a PivotTable from fewer than 2 source rows, so skip those pivots and say so in the workbook.
    skipped_pivots = pivot_skipped_types(ageing_rows)
    if skipped_pivots:
        skipped_text = ", ".join(f"{ref_type} ({row_count} row{'s' if row_count != 1 else ''})" for ref_type, row_count in skipped_pivots)
        notice_sheet = workbook.create_sheet("NOTICE")
        notice_sheet["A1"] = "Some PivotTables were not created"
        notice_sheet["A1"].font = Font(bold=True, size=14)
        notice_sheet["A2"] = (
            f"Excel needs at least 2 source rows to build a PivotTable. {assigned_person} has too few rows for: {skipped_text}. "
            "Those summary pivots were skipped; all other sheets are complete."
        )
        notice_sheet.column_dimensions["A"].width = 120

    output_dir = _company_output_dir("regularaccounts")
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / _regular_report_filename(assigned_person, source_filename)
    set_workbook_author_metadata(workbook)
    job_progress.update(token, 55, "Saving report workbook")
    workbook.save(output_path)
    workbook.close()
    if skipped_pivots:
        processor.logger.log(
            f"PivotTables skipped for {assigned_person}: {skipped_text} (Excel needs at least 2 rows). A NOTICE sheet was added.",
            "WARNING",
        )
    # Separate pre-download process: Excel adds native pivots, saves, and closes.
    job_progress.update(token, 62, "Building PivotTables in Excel (the longest step)", excel=True)
    run_regular_accounts_pivot_process(output_path, ageing_rows, token)
    job_progress.update(token, 97, "Finishing report")
    return output_path, len(follow_up_rows.index)


@app.post("/process/step-01", response_class=HTMLResponse)
def process_step_01(
    request: Request,
    inquiry_file: UploadFile = File(...),
    company_code: str = Form(DEFAULT_COMPANY_CODE),
    automatic_report: bool = Form(False),
    download_token: str = Form(""),
    amazon_accounts: list[str] = Form([]),
):
    if not inquiry_file.filename:
        return _render(request, "Please select an inquiry file.", status_code=400, selected_company_code=company_code)

    if company_code.strip().lower() == "regularaccounts":
        return _render(
            request,
            "Regular Accounts XLSX processing is awaiting its workflow rules.",
            selected_company_code=company_code,
        )

    logger, orchestrator, repository = build_services(company_code)
    temp_dir = tempfile.mkdtemp(prefix="wallmartsummary-")

    try:
        inquiry_name = os.path.basename(inquiry_file.filename)
        if company_code.strip().lower() not in FILE_ONLY_COMPANY_CODES:
            # Duplicate uploads are allowed; SQL refresh behavior remains unchanged.
            repository.ensure_database()
            existing_inquiry_files = repository.list_uploaded_batch_filenames()
            if inquiry_name in existing_inquiry_files:
                logger.log(f"Duplicate upload allowed: Step 01 file {inquiry_name}. Processing will continue.", "WARNING")

        inquiry_ext = os.path.splitext(inquiry_name)[1] or ".xlsx"
        inquiry_path = _save_upload(temp_dir, inquiry_file, f"inquiry{inquiry_ext}")
        return _run_step_01_flow(
            request,
            logger,
            orchestrator,
            repository,
            inquiry_path,
            inquiry_name,
            company_code,
            automatic_report=automatic_report,
            download_token=download_token,
            amazon_accounts=amazon_accounts,
        )
    except Exception as exc:
        logger.log(f"Fatal error: {exc}", "ERROR")
        append_history_event(
            "Step 01 failed",
            [
                "File: app\\main.py",
                "Function: process_step_01()",
                f"Error: {exc}",
            ],
        )
        return _render(request, f"Fatal error: {exc}", logger.entries, 500, company_code)
    finally:
        with suppress(Exception):
            shutil.rmtree(temp_dir)


@app.post("/process/regular-accounts/people")
async def regular_accounts_people(workbook_file: UploadFile = File(...)):
    if not workbook_file.filename:
        return JSONResponse({"error": "Please select a Regular Accounts workbook."}, status_code=400)

    workbook_name = os.path.basename(workbook_file.filename)
    if os.path.splitext(workbook_name)[1].lower() != ".xlsx":
        return JSONResponse({"error": "Regular Accounts supports XLSX workbooks only."}, status_code=400)

    logger = Logger()
    temp_dir = tempfile.mkdtemp(prefix="regularaccounts-")
    try:
        workbook_path = _save_upload(temp_dir, workbook_file, workbook_name)
        people = RegularAccountsProcessor(logger).person_options(workbook_path)
        return JSONResponse({"people": people, "message": f"Found {len(people)} Assigned Person name(s)."})
    except Exception as exc:
        logger.log(f"Regular Accounts person extraction failed: {exc}", "ERROR")
        return JSONResponse({"error": str(exc)}, status_code=400)
    finally:
        with suppress(Exception):
            shutil.rmtree(temp_dir)


@app.post("/process/regular-accounts/follow-up-report")
def regular_accounts_follow_up_report(
    inquiry_file: UploadFile = File(...),
    assigned_person: str = Form(...),
    download_token: str = Form(""),
    browser_date: str = Form(""),
):
    if not inquiry_file.filename:
        return JSONResponse({"error": "Please select a Regular Accounts workbook."}, status_code=400)
    if not assigned_person.strip():
        return JSONResponse({"error": "Please select a Person Name."}, status_code=400)

    source_filename = os.path.basename(inquiry_file.filename)
    if os.path.splitext(source_filename)[1].lower() != ".xlsx":
        return JSONResponse({"error": "Regular Accounts supports XLSX workbooks only."}, status_code=400)

    logger = Logger()
    temp_dir = tempfile.mkdtemp(prefix="regularaccounts-report-")
    try:
        job_progress.update(download_token, 5, f"Receiving {source_filename}")
        workbook_path = _save_upload(temp_dir, inquiry_file, source_filename)
        processor = RegularAccountsProcessor(logger)
        output_path, follow_up_count = _build_regular_accounts_report(
            processor,
            workbook_path,
            assigned_person.strip(),
            source_filename,
            download_token,
            browser_date,
        )
        logger.log(
            f"Regular Accounts report created for {assigned_person}: {follow_up_count} FOLLOW-UP ENTRIES row(s).",
            "SUCCESS",
        )
        job_progress.update(download_token, 100, "Report ready - downloading", done=True, file=output_path.name)
        return FileResponse(
            path=str(output_path),
            filename=output_path.name,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",

            background=BackgroundTask(_delete_file_after_download, output_path),
        )
    except Exception as exc:
        logger.log(f"Regular Accounts report generation failed: {exc}", "ERROR")
        job_progress.update(download_token, 100, "Report failed", done=True, error=str(exc))
        return JSONResponse({"error": str(exc)}, status_code=400)
    finally:
        with suppress(Exception):
            shutil.rmtree(temp_dir)




@app.post("/process/step-08-database-optimization", response_class=HTMLResponse)
async def process_step_08_database_optimization(request: Request, company_code: str = Form(DEFAULT_COMPANY_CODE)):
    logger, _, repository = build_services(company_code)

    try:
        repository.ensure_database()
        logger.log("Running step_08_DataBaseOptimization SQL: execute selected table maintenance")
        sql_messages = repository.execute_sql_script_with_messages(STEP_08_DATABASE_OPTIMIZATION_SQL)
        for message_text in sql_messages:
            logger.log(f"step_08_DataBaseOptimization SQL message: {message_text}")

        optimization_tables = (
            "Allowances",
            "Credit_Control_Sheet",
            "Credit_Match_Sheet",
            "Deduction",
            "Inv_DB_Control_Sheet",
            "Inv_DB_Match_Sheet",
            "Over_Pay",
            "Short_Pay",
            "step_01_inquiry_rows",
            "step_01_raw_rows",
            "step_02_farooq",
            "step_02_po_details",
            "step_03_payments_combined",
            "step_07_Credit_Applied_Sheet",
            "Unapplied",
            "Unpaid",
        )
        table_counts: dict[str, str] = {}
        for table_name in optimization_tables:
            try:
                table_counts[table_name] = str(len(repository.read_table(table_name).index))
            except Exception as table_exc:
                table_counts[table_name] = f"unavailable ({table_exc})"

        summary_text = ", ".join(
            f"{table_name}={row_count}" for table_name, row_count in table_counts.items()
        )
        logger.log(f"step_08_DataBaseOptimization completed: {summary_text}", "SUCCESS")
        append_history_event(
            "step_08_DataBaseOptimization completed",
            [
                "File: app\\main.py",
                "Function: process_step_08_database_optimization()",
                "SQL flow: step_08_database_optimization",
                "Executed dbo.sp_WallMartSummary_Truncate_Selected_Tables.",
                f"Rows: {summary_text}",
            ],
        )
        return _render(
            request,
            f"step_08_DataBaseOptimization completed. {summary_text}",
            logger.entries,
            selected_company_code=company_code,
        )
    except Exception as exc:
        logger.log(f"step_08_DataBaseOptimization failed: {exc}", "ERROR")
        append_history_event(
            "step_08_DataBaseOptimization failed",
            [
                "File: app\\main.py",
                "Function: process_step_08_database_optimization()",
                "SQL flow: step_08_database_optimization",
                f"Error: {exc}",
            ],
        )
        return _render(request, f"step_08_DataBaseOptimization failed: {exc}", logger.entries, 500, company_code)


@app.post("/process/deductions-summary", response_class=HTMLResponse)
async def process_deductions_summary(request: Request, company_code: str = Form(DEFAULT_COMPANY_CODE)):
    logger, _, repository = build_services(company_code)

    try:
        repository.ensure_database()
        logger.log("Running Deductions_Summary SQL: rebuild deduction category sheet")
        sql_messages = repository.execute_sql_script_with_messages(DEDUCTIONS_SUMMARY_SQL)
        for message_text in sql_messages:
            logger.log(f"Deductions_Summary SQL message: {message_text}")

        deductions_summary_count = len(repository.read_table("Deductions_Summary").index)
        logger.log(f"Deductions_Summary completed: Deductions_Summary={deductions_summary_count}", "SUCCESS")
        append_history_event(
            "Deductions_Summary completed",
            [
                "File: app\\main.py",
                "Function: process_deductions_summary()",
                "SQL flow: deductions_summary",
                "Rebuilt dbo.Deductions_Summary from step_03_payments_combined negative total rows.",
                f"Rows in Deductions_Summary: {deductions_summary_count}",
            ],
        )
        return _render(
            request,
            f"Deductions_Summary completed. Deductions_Summary={deductions_summary_count}",
            logger.entries,
            selected_company_code=company_code,
        )
    except Exception as exc:
        logger.log(f"Deductions_Summary failed: {exc}", "ERROR")
        append_history_event(
            "Deductions_Summary failed",
            [
                "File: app\\main.py",
                "Function: process_deductions_summary()",
                "SQL flow: deductions_summary",
                f"Error: {exc}",
            ],
        )
        return _render(request, f"Deductions_Summary failed: {exc}", logger.entries, 500, company_code)


@app.post("/process/invoice-payment-breakdown", response_class=HTMLResponse)
async def process_invoice_payment_breakdown(
    request: Request,
    company_code: str = Form(DEFAULT_COMPANY_CODE),
):
    logger, _, repository = build_services(company_code)

    breakdown_tables = (
        "Deduction",
        "Unapplied",
        "Over_Pay",
        "Short_Pay",
        "Allowances",
        "Unpaid",
    )

    try:
        repository.ensure_database()
        logger.log("Running Invoice Payment Breakdown SQL: rebuild deduction and payment status tables")
        sql_messages = repository.execute_sql_script_with_messages(INVOICE_PAYMENT_BREAKDOWN_SQL)
        for message_text in sql_messages:
            logger.log(f"Invoice Payment Breakdown SQL message: {message_text}")

        table_counts: dict[str, int] = {}
        for table_name in breakdown_tables:
            table_counts[table_name] = len(repository.read_table(table_name).index)

        summary_text = ", ".join(
            f"{table_name}={row_count}" for table_name, row_count in table_counts.items()
        )
        logger.log(f"Invoice Payment Breakdown completed: {summary_text}", "SUCCESS")
        append_history_event(
            "Invoice Payment Breakdown completed",
            [
                "File: app\\main.py",
                "Function: process_invoice_payment_breakdown()",
                "SQL flow: invoice_payment_breakdown",
                "Rebuilt Deduction, Unapplied, Over_Pay, Short_Pay, Allowances, and Unpaid.",
                f"Rows: {summary_text}",
            ],
        )
        return _render(
            request,
            f"Invoice Payment Breakdown completed. {summary_text}",
            logger.entries,
            selected_company_code=company_code,
        )
    except Exception as exc:
        logger.log(f"Invoice Payment Breakdown failed: {exc}", "ERROR")
        append_history_event(
            "Invoice Payment Breakdown failed",
            [
                "File: app\\main.py",
                "Function: process_invoice_payment_breakdown()",
                "SQL flow: invoice_payment_breakdown",
                f"Expected tables: {', '.join(breakdown_tables)}",
                f"Error: {exc}",
            ],
        )
        return _render(
            request,
            f"Invoice Payment Breakdown failed: {exc}",
            logger.entries,
            500,
            company_code,
        )


@app.post("/process/step-06d02-applying-cr-update", response_class=HTMLResponse)
async def process_step_06d02_applying_cr_update(request: Request, company_code: str = Form(DEFAULT_COMPANY_CODE)):
    logger, _, repository = build_services(company_code)

    try:
        repository.ensure_database()
        logger.log("Running step6 D02 applying cr update SQL")
        sql_messages = repository.execute_sql_script_with_messages(STEP_06D02_APPLYING_CR_UPDATE_SQL)
        for message_text in sql_messages:
            logger.log(f"step6 D02 applying cr update SQL message: {message_text}")

        logger.log("step6 D02 applying cr update executed.", "SUCCESS")
        append_history_event(
            "step6 D02 applying cr update completed",
            [
                "File: app\\main.py",
                "Function: process_step_06d02_applying_cr_update()",
                "SQL flow: step_06d02_applying_cr_update",
                "Executed the app-safe unapplied CR issue update query.",
            ],
        )
        return _render(
            request,
            "step6 D02 applying cr update completed.",
            logger.entries,
            selected_company_code=company_code,
        )
    except Exception as exc:
        logger.log(f"step6 D02 applying cr update failed: {exc}", "ERROR")
        append_history_event(
            "step6 D02 applying cr update failed",
            [
                "File: app\\main.py",
                "Function: process_step_06d02_applying_cr_update()",
                "SQL flow: step_06d02_applying_cr_update",
                f"Error: {exc}",
            ],
        )
        return _render(request, f"step6 D02 applying cr update failed: {exc}", logger.entries, 500, company_code)


async def _run_step_07a_section(
    request: Request,
    company_code: str,
    section_title: str,
    sql_script: str,
    table_name: str,
) -> HTMLResponse:
    logger, _, repository = build_services(company_code)

    try:
        repository.ensure_database()
        logger.log(f"Running {section_title} SQL")
        sql_messages = repository.execute_sql_script_with_messages(sql_script)
        for message_text in sql_messages:
            logger.log(f"{section_title} SQL message: {message_text}")

        row_count = len(repository.read_table(table_name).index)
        logger.log(f"{section_title} completed: {table_name}={row_count}", "SUCCESS")
        append_history_event(
            f"{section_title} completed",
            [
                "File: app\\main.py",
                "Function: _run_step_07a_section()",
                "SQL flow: step_07a_section",
                f"Created table: {table_name}",
                f"Rows in {table_name}: {row_count}",
            ],
        )
        return _render(
            request,
            f"{section_title} completed. {table_name}={row_count}",
            logger.entries,
            selected_company_code=company_code,
        )
    except Exception as exc:
        logger.log(f"{section_title} failed: {exc}", "ERROR")
        append_history_event(
            f"{section_title} failed",
            [
                "File: app\\main.py",
                "Function: _run_step_07a_section()",
                "SQL flow: step_07a_section",
                f"Expected table: {table_name}",
                f"Error: {exc}",
            ],
        )
        return _render(request, f"{section_title} failed: {exc}", logger.entries, 500, company_code)


@app.post("/process/step-07a-deduction-breakdown", response_class=HTMLResponse)
async def process_step_07a_deduction_breakdown(
    request: Request,
    company_code: str = Form(DEFAULT_COMPANY_CODE),
):
    return await _run_step_07a_section(
        request,
        company_code,
        "Step 07A Section 1 Deduction Breakdown",
        STEP_07A_DEDUCTION_BREAKDOWN_SQL,
        STEP_07A_DEDUCTIONS_BREAKDOWN_TABLE,
    )


@app.post("/process/step-07a-credit-summary", response_class=HTMLResponse)
async def process_step_07a_credit_summary(
    request: Request,
    company_code: str = Form(DEFAULT_COMPANY_CODE),
):
    return await _run_step_07a_section(
        request,
        company_code,
        "Step 07A Section 2 Gross / Deduction / Net Pay Summary",
        STEP_07A_CREDIT_SUMMARY_SQL,
        STEP_07A_CREDIT_SUMMARY_TABLE,
    )


@app.post("/process/step-07a-po-status-summary", response_class=HTMLResponse)
async def process_step_07a_po_status_summary(
    request: Request,
    company_code: str = Form(DEFAULT_COMPANY_CODE),
):
    return await _run_step_07a_section(
        request,
        company_code,
        "Step 07A Section 3 PO Status Summary",
        STEP_07A_PO_STATUS_SUMMARY_SQL,
        STEP_07A_PO_STATUS_SUMMARY_TABLE,
    )


@app.post("/process/step-07a-summary", response_class=HTMLResponse)
async def process_step_07a_summary(request: Request, company_code: str = Form(DEFAULT_COMPANY_CODE)):
    logger, _, repository = build_services(company_code)

    try:
        repository.ensure_database()
        logger.log("Running Step 07A SQL: deduction, gross/net, and PO status summaries")
        sql_messages = repository.execute_sql_script_with_messages(STEP_07A_SUMMARY_SQL)
        for message_text in sql_messages:
            logger.log(f"Step 07A SQL message: {message_text}")

        table_counts: dict[str, int] = {}
        for table_name in STEP_07A_SUMMARY_TABLES:
            table_counts[table_name] = len(repository.read_table(table_name).index)

        summary_text = ", ".join(
            f"{table_name}={row_count}" for table_name, row_count in table_counts.items()
        )
        logger.log(f"Step 07A summary tables refreshed: {summary_text}", "SUCCESS")
        append_history_event(
            "Step 07A summary completed",
            [
                "File: app\\main.py",
                "Function: process_step_07a_summary()",
                "SQL flow: step_07a_summary",
                "Section 1 table: Step_09A_Deductions_BreakDown.",
                "Section 2 table: Credit_CrDedNet_Summary.",
                "Section 3 table: Po_Status_Sum_Summary.",
                f"Rows: {summary_text}",
            ],
        )
        return _render(
            request,
            f"Step 07A summary completed. {summary_text}",
            logger.entries,
            selected_company_code=company_code,
        )
    except Exception as exc:
        logger.log(f"Step 07A summary failed: {exc}", "ERROR")
        append_history_event(
            "Step 07A summary failed",
            [
                "File: app\\main.py",
                "Function: process_step_07a_summary()",
                "SQL flow: step_07a_summary",
                f"Expected tables: {', '.join(STEP_07A_SUMMARY_TABLES)}",
                f"Error: {exc}",
            ],
        )
        return _render(request, f"Step 07A summary failed: {exc}", logger.entries, 500, company_code)


@app.post("/process/step-07", response_class=HTMLResponse)
async def process_step_07(request: Request, company_code: str = Form(DEFAULT_COMPANY_CODE)):
    logger, _, repository = build_services(company_code)

    try:
        repository.ensure_database()
        logger.log("Running Step 07 SQL: credit applied sheet build")
        sql_messages = repository.execute_sql_script_with_messages(STEP_07_SQL)
        for message_text in sql_messages:
            logger.log(f"Step 07 SQL message: {message_text}")

        step_07_count = len(repository.read_table(STEP_07_CREDIT_APPLIED_TABLE).index)
        logger.log(
            f"Step 07 table refreshed: {STEP_07_CREDIT_APPLIED_TABLE}={step_07_count}",
            "SUCCESS",
        )
        append_history_event(
            "Step 07 completed",
            [
                "File: app\\main.py",
                "Function: process_step_07()",
                "SQL flow: step_07",
                f"Created table: {STEP_07_CREDIT_APPLIED_TABLE}",
                f"Rows in {STEP_07_CREDIT_APPLIED_TABLE}: {step_07_count}",
            ],
        )
        return _render(
            request,
            f"Step 07 completed. {STEP_07_CREDIT_APPLIED_TABLE}={step_07_count}",
            logger.entries,
            selected_company_code=company_code,
        )
    except Exception as exc:
        logger.log(f"Step 07 failed: {exc}", "ERROR")
        append_history_event(
            "Step 07 failed",
            [
                "File: app\\main.py",
                "Function: process_step_07()",
                "SQL flow: step_07",
                f"Expected table: {STEP_07_CREDIT_APPLIED_TABLE}",
                f"Error: {exc}",
            ],
        )
        return _render(request, f"Step 07 failed: {exc}", logger.entries, 500, company_code)





