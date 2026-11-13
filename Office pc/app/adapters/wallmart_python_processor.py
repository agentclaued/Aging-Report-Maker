from __future__ import annotations

import shutil
from pathlib import Path
from typing import Protocol

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

from app.adapters.full_ageing_source import read_full_ageing_rows
from app.workbook_metadata import prepare_pivot_source_sheet, set_workbook_author_metadata


WALLMART_PYTHON_OUTPUT_FILE = "WAL_MART_STATUS_SUMMARY.xlsx"
WALLMART_STATUS_STYLE_TEMPLATE_PATH = (
    Path(__file__).resolve().parents[2] / "companies" / "wallmartsummary" / "wallmart_status_style_template.xlsx"
)
WALLMART_ACCOUNTS = {
    "WC9085": "WAL-MART CANADA, INC.",
    "WA9085": "WAL-MART STORE, INC.",
    "SA9200": "SAM'S CLUB",
    "WM9263": "WAL-MART STORE #2306",
}
WALLMART_COLUMN_ALIASES = {
    "acct": ("acct",),
    "div": ("div",),
    "name": ("name",),
    "ref": ("ref",),
    "ref_type": ("reftype",),
    "date": ("date",),
    "due_date": ("duedate",),
    "order": ("order",),
    "current": ("current",),
    "col_1_30_days": ("130days", "col130days"),
    "col_31_60_days": ("3160days", "col3160days"),
    "col_61_90_days": ("6190days", "col6190days"),
    "over_90_days": ("over90days",),
    "claim": ("claim",),
    "comment": ("comment", "comments"),
}
WALLMART_REQUIRED_COLUMN_LABELS = {
    "acct": "Acct",
    "div": "Div",
    "name": "Name",
    "ref": "Ref#",
    "ref_type": "Ref Type",
    "date": "Date",
    "due_date": "Due Date",
    "order": "Order#",
    "current": "Current",
    "col_1_30_days": "1-30 Days",
    "col_31_60_days": "31-60 Days",
    "col_61_90_days": "61-90 Days",
    "over_90_days": "Over 90 Days",
    "claim": "Claim#",
    "comment": "Comment",
}
WALLMART_AGEING_COLUMNS = (
    "current",
    "col_1_30_days",
    "col_31_60_days",
    "col_61_90_days",
    "over_90_days",
)

_CURRENCY = '$#,##0.00;[Red]($#,##0.00)'
_THIN_GRAY = Side(style="thin", color="A6A6A6")
_HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
_SUMMARY_FILL = PatternFill("solid", fgColor="D9EAF7")
_ACCOUNT_FILL = PatternFill("solid", fgColor="4F81BD")


class _Logger(Protocol):
    def log(self, message: str, level: str = "INFO") -> None: ...


def _key(value: object) -> str:
    return "".join(character for character in str(value).casefold() if character.isalnum())


class WallMartPythonReportProcessor:
    """Build the Wal Mart status workbook directly from an uploaded file.

    This deliberately contains the reporting rules formerly applied by the
    Wal Mart SQL procedures.  It reads the file locally and creates the final
    workbook without creating, querying, or refreshing SQL Server tables.
    """

    def __init__(self, logger: _Logger) -> None:
        self.logger = logger

    def process(self, file_path: str, output_dir: Path) -> dict:
        source = self._read_source(file_path)
        prepared = self._prepare_rows(source)
        output_dir.mkdir(parents=True, exist_ok=True)
        output_path = output_dir / WALLMART_PYTHON_OUTPUT_FILE

        if not WALLMART_STATUS_STYLE_TEMPLATE_PATH.exists():
            raise FileNotFoundError(
                f"Wal Mart report template is missing: {WALLMART_STATUS_STYLE_TEMPLATE_PATH}"
            )
        shutil.copyfile(WALLMART_STATUS_STYLE_TEMPLATE_PATH, output_path)
        workbook = load_workbook(output_path)
        for sheet_name in list(workbook.sheetnames):
            if sheet_name not in {"DBs_Status", "Invoices_Status"}:
                workbook.remove(workbook[sheet_name])
        self._clear_template_sheet(workbook["DBs_Status"])
        self._clear_template_sheet(workbook["Invoices_Status"])
        self._write_db_status(workbook, prepared)
        self._write_invoice_status(workbook, prepared)
        self._write_detail_sheet(workbook, "WallMartSummary_Sheet", prepared, "WalMartSummary")
        self._write_account_detail_sheets(workbook, prepared)
        set_workbook_author_metadata(workbook)
        workbook.save(output_path)
        workbook.close()

        self.logger.log(
            f"Wal Mart Python report created: {len(prepared)} source row(s); no database connection was used.",
            "SUCCESS",
        )
        return {"output_path": str(output_path), "source_row_count": len(prepared)}

    @staticmethod
    def _clear_template_sheet(sheet) -> None:
        """Clear data only so the established Wal Mart workbook style survives."""
        for merged_range in list(sheet.merged_cells.ranges):
            sheet.unmerge_cells(str(merged_range))
        for row in sheet.iter_rows():
            for cell in row:
                cell.value = None

    def _read_source(self, file_path: str) -> pd.DataFrame:
        source_path = Path(file_path)
        if source_path.suffix.lower() == ".csv":
            frame = pd.read_csv(source_path).dropna(how="all")
            if frame.empty:
                raise ValueError("The uploaded Wal Mart CSV has no data rows.")
            self.logger.log(f"Read Wal Mart CSV: {len(frame)} row(s).")
            return frame

        if source_path.suffix.lower() not in {".xlsx", ".xls"}:
            raise ValueError("Wal Mart supports .csv, .xlsx, and .xls files.")
        ageing_rows = read_full_ageing_rows(file_path, tuple(WALLMART_ACCOUNTS))
        if ageing_rows is not None:
            self.logger.log(
                f"Read Full Ageing sheet: {len(ageing_rows)} row(s) for Acct {', '.join(WALLMART_ACCOUNTS)}."
            )
            ageing_rows["__source_sheet"] = "Full Ageing"
            return ageing_rows

        workbook = pd.ExcelFile(source_path)
        frames: list[pd.DataFrame] = []
        header_scan = load_workbook(source_path, read_only=True) if source_path.suffix.lower() == ".xlsx" else None
        for sheet_name in workbook.sheet_names:
            usecols = self._headed_columns(header_scan[sheet_name], sheet_name) if header_scan else None
            frame = pd.read_excel(workbook, sheet_name=sheet_name, usecols=usecols).dropna(how="all")
            self._precheck_sheet(sheet_name, frame)
            # Stray formatting/cells far to the right can create thousands of empty columns.
            frame = frame.dropna(axis=1, how="all")
            if frame.empty:
                self.logger.log(f"Wal Mart source sheet '{sheet_name}' is empty and was skipped.", "WARNING")
                continue
            frame = frame.copy()
            frame["__source_sheet"] = sheet_name
            frames.append(frame)
            self.logger.log(f"Read Wal Mart source sheet '{sheet_name}': {len(frame)} row(s).")
        if not frames:
            raise ValueError("The uploaded Wal Mart workbook has no data rows.")
        if header_scan:
            header_scan.close()
        return pd.concat(frames, ignore_index=True, sort=False)

    def _headed_columns(self, sheet, sheet_name: str) -> list[int] | None:
        """Return only the column positions that have a header, so stray far-right cells are never loaded."""
        header_row = next(sheet.iter_rows(min_row=1, max_row=1, values_only=True), ())
        headed = [index for index, value in enumerate(header_row) if value not in (None, "")]
        if not headed:
            return None
        headerless_with_data: dict[int, str] = {}
        for row in sheet.iter_rows(min_row=2, values_only=True):
            for index, value in enumerate(row):
                if value not in (None, "") and index not in headed and index not in headerless_with_data:
                    headerless_with_data[index] = str(value).strip()
        for index, sample in headerless_with_data.items():
            self.logger.log(
                f"Pre-check: sheet '{sheet_name}' column {get_column_letter(index + 1)} has data but no header "
                f"(sample: '{sample}'); it is unrelated to the template and was not loaded.",
                "WARNING",
            )
        return headed


    def _precheck_sheet(self, sheet_name: str, frame: pd.DataFrame) -> None:
        """Log anything unrelated to or outside the template before processing (log only, never blocks)."""
        sheet_code = str(sheet_name).strip().split(" ")[0].upper()
        if sheet_code not in WALLMART_ACCOUNTS:
            self.logger.log(
                f"Pre-check: sheet '{sheet_name}' does not start with a known account code "
                f"({', '.join(WALLMART_ACCOUNTS)}); it is not part of the template.",
                "WARNING",
            )
        known_keys = {alias for aliases in WALLMART_COLUMN_ALIASES.values() for alias in aliases}
        present_keys = {_key(column) for column in frame.columns}
        missing = [
            WALLMART_REQUIRED_COLUMN_LABELS[field]
            for field, aliases in WALLMART_COLUMN_ALIASES.items()
            if field != "ref_type" and not any(alias in present_keys for alias in aliases)
        ]
        if missing:
            self.logger.log(f"Pre-check: sheet '{sheet_name}' is missing template column(s): {', '.join(missing)}.", "WARNING")

        empty_columns = [column for column in frame.columns if frame[column].isna().all()]
        if empty_columns:
            self.logger.log(
                f"Pre-check: sheet '{sheet_name}' has {len(empty_columns)} empty column(s) "
                f"(unused range up to column {len(frame.columns)}); they will be ignored.",
                "WARNING",
            )
        extra = [
            (position, column)
            for position, column in enumerate(frame.columns, start=1)
            if column not in empty_columns and _key(column) not in known_keys and _key(column) != "total"
        ]
        for position, column in extra:
            sample = frame[column].dropna().astype(str).head(2).str.strip().tolist()
            self.logger.log(
                f"Pre-check: sheet '{sheet_name}' column {get_column_letter(position)} "
                f"('{column}') is outside the template; sample values: {sample}.",
                "WARNING",
            )
        if not missing and not extra and not empty_columns and sheet_code in WALLMART_ACCOUNTS:
            self.logger.log(f"Pre-check: sheet '{sheet_name}' matches the template ({len(frame)} row(s)).")

    @staticmethod
    def _coerce_excel_dates(values: pd.Series) -> pd.Series:
        """Return real dates for text, datetime, and Excel serial-date values."""
        dates = pd.to_datetime(values, errors="coerce")
        numeric_values = pd.to_numeric(values, errors="coerce")
        serial_mask = numeric_values.between(1, 100000, inclusive="both")
        if serial_mask.any():
            dates.loc[serial_mask] = pd.to_datetime(
                numeric_values.loc[serial_mask],
                unit="D",
                origin="1899-12-30",
                errors="coerce",
            )
        return dates

    def _prepare_rows(self, source: pd.DataFrame) -> pd.DataFrame:
        result = source.copy()
        columns = list(result.columns)
        column_lookup = {_key(column): str(column) for column in columns}
        resolved_columns = {
            field_name: next(
                (column_lookup[alias] for alias in aliases if alias in column_lookup),
                None,
            )
            for field_name, aliases in WALLMART_COLUMN_ALIASES.items()
        }
        missing = [
            WALLMART_REQUIRED_COLUMN_LABELS[field_name]
            for field_name, source_column in resolved_columns.items()
            if source_column is None and field_name != "ref_type"
        ]
        if missing:
            raise ValueError(
                "Wal Mart report is missing required column(s): " + ", ".join(missing) + "."
            )
        acct = resolved_columns["acct"]
        name = resolved_columns["name"]
        ref = resolved_columns["ref"]
        date = resolved_columns["date"]
        order = resolved_columns["order"]
        comment = resolved_columns["comment"]
        due_date = resolved_columns["due_date"]
        ageing_columns = [resolved_columns[column] for column in WALLMART_AGEING_COLUMNS]
        total_column = column_lookup.get("total", "total")

        result["__wm_acct"] = result[acct].fillna("").astype(str).str.strip().str.upper()
        blank_acct_rows = result.index[result["__wm_acct"] == ""].tolist()
        if blank_acct_rows:
            display_rows = ", ".join(str(index + 2) for index in blank_acct_rows[:20])
            extra = "" if len(blank_acct_rows) <= 20 else " (first 20 shown)"
            self.logger.log(
                f"WARNING: blank acct value found at uploaded row(s) {display_rows}{extra}. "
                "Those rows remain in WallMartSummary_Sheet but are excluded from status totals.",
                "ERROR",
            )
        result["__wm_name"] = result[name].fillna("").astype(str).str.strip()
        result["__wm_ref"] = result[ref].fillna("").astype(str).str.strip()
        ref_upper = result["__wm_ref"].str.upper()
        result["__wm_ref_type"] = ""
        for ref_type in ("IN", "DB", "CR"):
            result.loc[
                (result["__wm_ref_type"] == "") & ref_upper.str.contains(ref_type, regex=False),
                "__wm_ref_type",
            ] = ref_type
        ref_type_column = resolved_columns["ref_type"] or "Ref Type"
        result[ref_type_column] = result["__wm_ref_type"]
        result[date] = self._coerce_excel_dates(result[date])
        result[due_date] = self._coerce_excel_dates(result[due_date])
        result["__wm_date"] = result[date]
        for column in ageing_columns:
            result[column] = pd.to_numeric(result[column], errors="coerce").fillna(0.0)
        result[total_column] = result[ageing_columns].sum(axis=1)
        result["__wm_amount"] = result[total_column]
        result["__wm_reason"] = result[order].fillna("(Blank)").astype(str).str.strip()
        result.loc[result["__wm_reason"] == "", "__wm_reason"] = "(Blank)"
        comment_text = result[comment].fillna("").astype(str)
        normalized_comment = comment_text.str.casefold().str.replace(r"[\s_-]+", " ", regex=True).str.strip()
        result["__wm_year"] = result["__wm_date"].dt.year
        result["InvYear"] = result["__wm_year"]
        result["WriteStatus"] = ""
        result.loc[
            (result["__wm_ref_type"] == "DB")
            & normalized_comment.str.startswith("write up")
            & ~normalized_comment.str.contains("approval", regex=False),
            "WriteStatus",
        ] = "Write-up Sent"
        result["RetailLinkStatus"] = ""
        result.loc[
            normalized_comment.str.contains("no", regex=False) & normalized_comment.str.contains("record", regex=False),
            "RetailLinkStatus",
        ] = "NoRetailLink"
        result["PaidStatus"] = ""
        result.loc[
            (result["__wm_ref_type"] == "IN") & normalized_comment.str.startswith("unpaid"),
            "PaidStatus",
        ] = "Unpaid"
        result.loc[
            (result["__wm_ref_type"] == "IN")
            & normalized_comment.str.startswith("paid")
            & ~normalized_comment.str.startswith("paid fully")
            & normalized_comment.str.count(r"\bpaid\b").eq(1)
            & ~normalized_comment.str.contains("deduc", regex=False),
            "PaidStatus",
        ] = "InvPaid"
        result["PaidnFullyDeduct"] = ""
        result.loc[
            normalized_comment.str.contains("deduct", regex=False),
            "PaidnFullyDeduct",
        ] = "Paid & Fully Deduct"
        result["Dispute_Amount"] = ""
        result.loc[
            normalized_comment.str.contains("dispute", regex=False),
            "Dispute_Amount",
        ] = "Disputed"
        invalid_date_count = int(result["__wm_year"].isna().sum())
        if invalid_date_count:
            self.logger.log(f"Skipped {invalid_date_count} row(s) without a usable Date from status summaries.", "WARNING")
        return result

    @staticmethod
    def _years(rows: pd.DataFrame) -> list[int]:
        return sorted(rows["__wm_year"].dropna().astype(int).unique().tolist())

    def _write_db_status(self, workbook: Workbook, rows: pd.DataFrame) -> None:
        sheet = workbook["DBs_Status"]
        db_rows = rows.loc[(rows["__wm_ref_type"] == "DB") & rows["__wm_year"].notna()].copy()
        self._log_unsupported_accounts(db_rows)
        db_rows = db_rows.loc[db_rows["__wm_acct"].isin(WALLMART_ACCOUNTS)].copy()
        years = self._years(db_rows)
        labels = [*years, "Total Amount", "Dispute_Amount", "Write-up Sent"]
        self._write_status_title(sheet, "WAL MART OPEN AGING (DB'S) STATUS SUMMARY", labels)
        overall = self._summary_values(db_rows, years)
        self._write_summary_row(sheet, 3, "All Accounts Total", labels, overall)
        row = 8
        account_codes = list(WALLMART_ACCOUNTS)
        for account_code in account_codes:
            account_rows = db_rows.loc[db_rows["__wm_acct"] == account_code]
            if account_rows.empty:
                continue
            account_name = self._source_account_name(account_rows, account_code)
            sheet.merge_cells(start_row=row, start_column=1, end_row=row, end_column=len(labels) + 1)
            cell = sheet.cell(row, 1, f"{account_code} - {account_name} (Open DB's)")
            cell.fill = _ACCOUNT_FILL
            cell.font = Font(color="FFFFFF", bold=True)
            row += 1
            self._write_heading(sheet, row, "Reason", labels)
            row += 1
            for reason, reason_rows in account_rows.groupby("__wm_reason", dropna=False, sort=True):
                self._write_summary_row(sheet, row, str(reason), labels, self._summary_values(reason_rows, years))
                row += 1
            self._write_summary_row(sheet, row, "ACCOUNT TOTAL", labels, self._summary_values(account_rows, years), total=True)
            row += 3
        self._finish_status_sheet(sheet, len(labels) + 1)

    def _write_invoice_status(self, workbook: Workbook, rows: pd.DataFrame) -> None:
        sheet = workbook["Invoices_Status"]
        invoice_rows = rows.loc[(rows["__wm_ref_type"] == "IN") & rows["__wm_year"].notna()].copy()
        self._log_unsupported_accounts(invoice_rows)
        invoice_rows = invoice_rows.loc[invoice_rows["__wm_acct"].isin(WALLMART_ACCOUNTS)].copy()
        years = self._years(invoice_rows)
        labels = [*years, "Total Amount", "Paid#", "Paid & Fully Deduct", "No Record Found"]
        end_column = len(labels) + 1
        sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=end_column)
        title = sheet.cell(1, 1, "WAL MART OPEN AGEING (INV) STATUS SUMMARY")
        title.fill = _HEADER_FILL
        title.font = Font(color="FFFFFF", bold=True)
        title.alignment = Alignment(horizontal="center")
        totals = self._invoice_values(invoice_rows, years)
        sheet.cell(3, 2, totals["Total Amount"])
        sheet.cell(4, 2, "Total Amount")
        for index, year in enumerate(years, start=3):
            sheet.cell(3, index, totals[year])
            sheet.cell(4, index, str(year))
        for column in range(2, len(years) + 3):
            sheet.cell(3, column).number_format = _CURRENCY
            sheet.cell(3, column).alignment = Alignment(horizontal="right")
        sheet.merge_cells(start_row=6, start_column=1, end_row=6, end_column=end_column)
        sheet.cell(6, 1, "YEAR-WISE AMOUNT BREAKDOWN").fill = _HEADER_FILL
        sheet.cell(6, 1).font = Font(color="FFFFFF", bold=True)
        sheet.cell(6, 1).alignment = Alignment(horizontal="center")
        self._write_heading(sheet, 7, "Account Name", labels)
        row = 8
        account_codes = list(WALLMART_ACCOUNTS)
        for account_code in account_codes:
            account_rows = invoice_rows.loc[invoice_rows["__wm_acct"] == account_code]
            if account_rows.empty:
                continue
            account_name = self._source_account_name(account_rows, account_code)
            self._write_summary_row(
                sheet,
                row,
                f"{account_code} ({account_name})",
                labels,
                self._invoice_values(account_rows, years),
            )
            row += 1
        self._write_summary_row(sheet, row, "Total", labels, self._invoice_values(invoice_rows, years), total=True)
        self._finish_status_sheet(sheet, end_column)

    @staticmethod
    def _summary_values(rows: pd.DataFrame, years: list[int]) -> dict[object, float]:
        values = {year: float(rows.loc[rows["__wm_year"] == year, "__wm_amount"].sum()) for year in years}
        values["Total Amount"] = float(rows["__wm_amount"].sum())
        values["Dispute_Amount"] = float(rows.loc[rows["Dispute_Amount"] == "Disputed", "__wm_amount"].sum())
        values["Write-up Sent"] = float(rows.loc[rows["WriteStatus"] == "Write-up Sent", "__wm_amount"].sum())
        return values

    @staticmethod
    def _invoice_values(rows: pd.DataFrame, years: list[int]) -> dict[object, float]:
        values = {year: float(rows.loc[rows["__wm_year"] == year, "__wm_amount"].sum()) for year in years}
        values["Total Amount"] = float(rows["__wm_amount"].sum())
        values["Paid#"] = float(rows.loc[rows["PaidStatus"] == "InvPaid", "__wm_amount"].sum())
        values["Paid & Fully Deduct"] = float(rows.loc[rows["PaidnFullyDeduct"] != "", "__wm_amount"].sum())
        values["No Record Found"] = float(rows.loc[rows["RetailLinkStatus"] == "NoRetailLink", "__wm_amount"].sum())
        return values

    @staticmethod
    def _write_status_title(sheet, title: str, labels: list[object]) -> None:
        end_column = len(labels) + 1
        sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=end_column)
        cell = sheet.cell(1, 1, title)
        cell.fill = _HEADER_FILL
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center")
        WallMartPythonReportProcessor._write_heading(sheet, 4, "Summary", labels)
        sheet.freeze_panes = "B5"

    @staticmethod
    def _write_heading(sheet, row: int, first_label: str, labels: list[object]) -> None:
        values = [first_label, *labels]
        for column, value in enumerate(values, start=1):
            cell = sheet.cell(row, column, value)
            cell.fill = _HEADER_FILL
            cell.font = Font(color="FFFFFF", bold=True)
            cell.alignment = Alignment(horizontal="center")
            cell.number_format = "General"
            cell.border = Border(top=_THIN_GRAY, bottom=_THIN_GRAY, left=_THIN_GRAY, right=_THIN_GRAY)

    @staticmethod
    def _write_summary_row(sheet, row: int, label: str, labels: list[object], values: dict[object, float], total: bool = False) -> None:
        row_fill = _SUMMARY_FILL if total else PatternFill(fill_type=None)
        for column, value in enumerate([label, *(values.get(item, 0.0) for item in labels)], start=1):
            cell = sheet.cell(row, column, value)
            cell.fill = row_fill
            cell.border = Border(top=_THIN_GRAY, bottom=_THIN_GRAY, left=_THIN_GRAY, right=_THIN_GRAY)
            # The copied template has white-font placeholder rows.  Every generated
            # summary value must reset to visible black text before writing data.
            cell.font = Font(color="000000", bold=total)
            if column > 1:
                cell.number_format = _CURRENCY
                cell.alignment = Alignment(horizontal="right")

    @staticmethod
    def _finish_status_sheet(sheet, last_column: int) -> None:
        sheet.sheet_view.showGridLines = False
        sheet.column_dimensions["A"].width = 34
        for column in range(2, last_column + 1):
            sheet.column_dimensions[get_column_letter(column)].width = 16

    def _log_unsupported_accounts(self, rows: pd.DataFrame) -> None:
        unsupported = sorted(set(rows["__wm_acct"]) - set(WALLMART_ACCOUNTS) - {""})
        if unsupported:
            self.logger.log(
                "WARNING: unsupported acct value(s) kept in WallMartSummary_Sheet but excluded from status totals: "
                + ", ".join(unsupported),
                "ERROR",
            )

    @staticmethod
    def _source_account_name(rows: pd.DataFrame, account_code: str) -> str:
        source_names = rows.loc[rows["__wm_name"] != "", "__wm_name"]
        return str(source_names.iloc[0]) if not source_names.empty else WALLMART_ACCOUNTS[account_code]

    def _write_detail_sheet(self, workbook: Workbook, name: str, frame: pd.DataFrame, table_name: str) -> None:
        sheet = workbook.create_sheet(name[:31])
        visible_columns = [column for column in frame.columns if not str(column).startswith("__wm_")]
        for row in dataframe_to_rows(frame.loc[:, visible_columns], index=False, header=True):
            sheet.append(list(row))
        prepare_pivot_source_sheet(sheet, table_name)
        for column_index, column_name in enumerate(visible_columns, start=1):
            column_key = _key(column_name)
            if column_key in {"date", "duedate"}:
                for row_index in range(2, sheet.max_row + 1):
                    sheet.cell(row_index, column_index).number_format = "mm/dd/yyyy"
            if column_key in {"total", "current", "130days", "col130days", "3160days", "col3160days", "6190days", "col6190days", "over90days"}:
                for row_index in range(2, sheet.max_row + 1):
                    sheet.cell(row_index, column_index).number_format = _CURRENCY
        sheet.freeze_panes = "A2"
        sheet.sheet_view.showGridLines = False
        for index, column in enumerate(visible_columns, start=1):
            sheet.column_dimensions[get_column_letter(index)].width = min(max(len(str(column)) + 2, 12), 28)

    def _write_account_detail_sheets(self, workbook: Workbook, rows: pd.DataFrame) -> None:
        used_names = {sheet.title.casefold() for sheet in workbook.worksheets}
        for account_code in WALLMART_ACCOUNTS:
            for ref_type, suffix in (("IN", "Invoices"), ("DB", "DBs"), ("CR", "CRs")):
                subset = rows.loc[(rows["__wm_acct"] == account_code) & (rows["__wm_ref_type"] == ref_type)]
                if subset.empty:
                    continue
                base_name = f"{account_code}_{suffix}"[:31]
                sheet_name = base_name
                suffix_number = 2
                while sheet_name.casefold() in used_names:
                    tail = f"_{suffix_number}"
                    sheet_name = f"{base_name[:31 - len(tail)]}{tail}"
                    suffix_number += 1
                used_names.add(sheet_name.casefold())
                self._write_detail_sheet(workbook, sheet_name, subset, f"Walmart{account_code}{suffix}")
