from __future__ import annotations

import re
import shutil
from copy import copy
from pathlib import Path
from typing import Protocol

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from app.adapters.full_ageing_source import read_full_ageing_rows
from app.workbook_metadata import prepare_pivot_source_sheet, set_workbook_author_metadata


AMAZON_ACCOUNT_CODES = ("AM9285", "AM9286", "ME9145")
REF_TYPE_SHEET_SUFFIXES = {"IN": "Invoices", "DB": "DBs", "CR": "CRs"}
AMAZON_GROUPED_OUTPUT_FILE = "AMAZON_ACCOUNT_TYPE_SHEETS.xlsx"
AMAZON_STATUS_STYLE_TEMPLATE_PATH = (
    Path(__file__).resolve().parents[2] / "companies" / "amazonaccounts" / "amazon_status_style_template.xlsx"
)


class _Logger(Protocol):
    def log(self, message: str, level: str = "INFO") -> None: ...


def _column_key(column_name: object) -> str:
    return "".join(character for character in str(column_name).upper() if character.isalnum())


class AmazonAccountSheetProcessor:
    """Create Amazon account/type sheets from all source workbook sheets."""

    def __init__(self, logger: _Logger) -> None:
        self.logger = logger

    def process(self, file_path: str, output_dir: Path, accounts: list[str] | None = None) -> dict:
        # Full Ageing input: only the accounts the user ticked (default AM9285 and AM9286; ME9145 is optional).
        selected_accounts = tuple(code for code in AMAZON_ACCOUNT_CODES if code in {a.strip().upper() for a in (accounts or [])}) or ("AM9285", "AM9286")
        source_workbook = pd.ExcelFile(file_path)
        accepted_frames: list[pd.DataFrame] = []
        uploaded_frames: list[pd.DataFrame] = []
        source_columns: list[str] | None = None
        invalid_ref_type_rows = 0
        unsupported_account_rows = 0
        acct_column = "Acct"
        ref_type_column = "Ref Type"

        ageing_rows = read_full_ageing_rows(file_path, selected_accounts)
        if ageing_rows is not None:
            self.logger.log(f"Read Full Ageing sheet: {len(ageing_rows)} row(s) for Acct {', '.join(selected_accounts)}.")
        sheet_sources = (
            [("Full Ageing", ageing_rows)]
            if ageing_rows is not None
            else [(name, None) for name in source_workbook.sheet_names]
        )

        for sheet_name, preloaded_frame in sheet_sources:
            source_frame = (
                preloaded_frame
                if preloaded_frame is not None
                else pd.read_excel(source_workbook, sheet_name=sheet_name).dropna(how="all")
            )
            if source_frame.empty:
                self.logger.log(f"Amazon source sheet '{sheet_name}' is empty and was skipped.", "WARNING")
                continue

            column_lookup = {_column_key(column_name): str(column_name) for column_name in source_frame.columns}
            acct_column = column_lookup.get("ACCT", "")
            ref_type_column = column_lookup.get("REFTYPE", "")
            if not acct_column or not ref_type_column:
                raise ValueError(f"Amazon source sheet '{sheet_name}' must contain both Acct and Ref Type columns.")

            if source_columns is None:
                source_columns = [str(column_name) for column_name in source_frame.columns]

            uploaded_frames.append(source_frame.copy())
            source_frame = source_frame.copy()
            source_frame[acct_column] = source_frame[acct_column].astype(str).str.strip().str.upper()
            source_frame[ref_type_column] = source_frame[ref_type_column].astype(str).str.strip().str.upper()
            valid_ref_type = source_frame[ref_type_column].isin(REF_TYPE_SHEET_SUFFIXES)
            valid_account = source_frame[acct_column].isin(AMAZON_ACCOUNT_CODES)
            invalid_ref_type_rows += int((~valid_ref_type).sum())
            unsupported_account_rows += int((~valid_account).sum())
            accepted_frames.append(source_frame.loc[valid_ref_type & valid_account])
            self.logger.log(
                f"Amazon source sheet '{sheet_name}' read: {len(source_frame)} rows; "
                f"accepted {int((valid_ref_type & valid_account).sum())} rows."
            )

        if not accepted_frames or source_columns is None:
            raise ValueError("The Amazon workbook has no non-empty sheet with Acct and Ref Type data.")

        combined = pd.concat(accepted_frames, ignore_index=True, sort=False)
        master_aging_frame = pd.concat(uploaded_frames, ignore_index=True, sort=False)
        output_dir.mkdir(parents=True, exist_ok=True)
        output_path = output_dir / AMAZON_GROUPED_OUTPUT_FILE
        if not AMAZON_STATUS_STYLE_TEMPLATE_PATH.exists():
            raise FileNotFoundError(f"Amazon status style template is missing: {AMAZON_STATUS_STYLE_TEMPLATE_PATH}")

        shutil.copyfile(AMAZON_STATUS_STYLE_TEMPLATE_PATH, output_path)
        self.logger.log("Copied Amazon status-sheet formatting template.", "SUCCESS")
        with pd.ExcelWriter(output_path, engine="openpyxl", mode="a") as writer:
            set_workbook_author_metadata(writer.book)
            master_aging_frame.to_excel(writer, sheet_name="Master Aging", index=False)
            prepare_pivot_source_sheet(writer.sheets["Master Aging"], "AmazonMasterAging")
            self.logger.log(f"Created Master Aging: {len(master_aging_frame)} uploaded rows.", "SUCCESS")
            for account_code in AMAZON_ACCOUNT_CODES:
                for ref_type, suffix in REF_TYPE_SHEET_SUFFIXES.items():
                    output_sheet_name = f"{account_code}_{suffix}"
                    account_frame = combined.loc[
                        (combined[acct_column] == account_code) & (combined[ref_type_column] == ref_type)
                    ].copy().reindex(columns=source_columns)
                    if account_frame.empty:
                        self.logger.log(f"Skipped empty {output_sheet_name} sheet.", "NOTICE")
                        continue
                    account_frame.to_excel(writer, sheet_name=output_sheet_name, index=False)
                    prepare_pivot_source_sheet(
                        writer.sheets[output_sheet_name],
                        f"Amazon{account_code}{suffix}",
                    )
                    self.logger.log(f"Created {output_sheet_name}: {len(account_frame)} rows.", "SUCCESS")

        self._build_db_status_sheet(output_path, combined, acct_column, ref_type_column)
        self._build_invoice_status_sheet(output_path, combined, acct_column, ref_type_column)
        self._build_credit_status_sheet(output_path, combined, acct_column, ref_type_column)

        if invalid_ref_type_rows:
            self.logger.log(f"Skipped {invalid_ref_type_rows} row(s) with Ref Type other than IN, DB, or CR.", "WARNING")
        if unsupported_account_rows:
            self.logger.log(f"Skipped {unsupported_account_rows} row(s) outside AM9285, AM9286, or ME9145.", "WARNING")

        return {"output_path": str(output_path), "accepted_row_count": len(combined)}

    @staticmethod
    def _find_column(data_frame: pd.DataFrame, normalized_name: str) -> str | None:
        return next(
            (
                str(column_name)
                for column_name in data_frame.columns
                if _column_key(column_name) == normalized_name
            ),
            None,
        )

    @staticmethod
    def _copy_cell_style(source_cell, target_cell) -> None:
        target_cell._style = copy(source_cell._style)

    @staticmethod
    def _is_paid_comment(comment: object) -> bool:
        """Recognize a paid-on/payment-number comment despite small payment-word typos."""
        comment_text = str(comment).lower()
        paid_on_match = re.search(r"\bpaid\s+on\b", comment_text)
        if not paid_on_match:
            return False

        tokens = re.findall(r"[a-z]+|\d+", comment_text[paid_on_match.end() :])
        for index, token in enumerate(tokens[:-1]):
            if (
                AmazonAccountSheetProcessor._levenshtein_distance(token, "payment") <= 2
                and tokens[index + 1] == "no"
                and any(candidate.isdigit() for candidate in tokens[index + 2 :])
            ):
                return True
        return False

    @staticmethod
    def _levenshtein_distance(first_value: str, second_value: str) -> int:
        previous_row = list(range(len(second_value) + 1))
        for first_index, first_character in enumerate(first_value, start=1):
            current_row = [first_index]
            for second_index, second_character in enumerate(second_value, start=1):
                current_row.append(
                    min(
                        current_row[-1] + 1,
                        previous_row[second_index] + 1,
                        previous_row[second_index - 1] + (first_character != second_character),
                    )
                )
            previous_row = current_row
        return previous_row[-1]

    def _build_db_status_sheet(
        self,
        output_path: Path,
        combined: pd.DataFrame,
        acct_column: str,
        ref_type_column: str,
    ) -> None:
        date_column = self._find_column(combined, "DATE")
        order_column = self._find_column(combined, "ORDER")
        total_column = self._find_column(combined, "TOTAL")
        comment_column = self._find_column(combined, "COMMENT")
        if not date_column or not order_column or not total_column or not comment_column:
            raise ValueError("Amazon DB status requires Date, Order#, Total, and Comment columns.")

        db_rows = combined.loc[combined[ref_type_column] == "DB"].copy()
        db_rows[date_column] = pd.to_datetime(db_rows[date_column], errors="coerce")
        db_rows["__amazon_year"] = db_rows[date_column].dt.year
        db_rows["__amazon_order"] = db_rows[order_column].fillna("(Blank)").astype(str).str.strip()
        db_rows.loc[db_rows["__amazon_order"] == "", "__amazon_order"] = "(Blank)"
        db_rows["__amazon_total"] = pd.to_numeric(db_rows[total_column], errors="coerce").fillna(0)
        db_rows["__amazon_writeup_sent"] = db_rows[comment_column].astype(str).str.contains(
            "MEMO",
            case=False,
            regex=False,
            na=False,
        )
        db_rows = db_rows.dropna(subset=["__amazon_year"])
        db_rows["__amazon_year"] = db_rows["__amazon_year"].astype(int)
        years = sorted(db_rows["__amazon_year"].unique().tolist())
        if not years:
            raise ValueError("Amazon DB status found no valid Date values in Ref Type DB rows.")

        name_column = self._find_column(combined, "NAME")
        style_workbook = load_workbook(AMAZON_STATUS_STYLE_TEMPLATE_PATH, read_only=False, data_only=False)
        output_workbook = load_workbook(output_path, read_only=False, data_only=False)
        try:
            style_sheet = style_workbook["DB's Status"]
            status_sheet = output_workbook["DB's Status"]
            for merged_range in list(status_sheet.merged_cells.ranges):
                status_sheet.unmerge_cells(str(merged_range))
            for row in status_sheet.iter_rows():
                for cell in row:
                    cell.value = None

            total_column_index = 2 + len(years)
            writeup_column_index = total_column_index + 1
            last_column_index = writeup_column_index

            status_sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_column_index)
            status_sheet.cell(1, 1).value = "AMAZON OPEN AGEING (DB'S) STATUS SUMMARY"
            self._copy_cell_style(style_sheet["A1"], status_sheet.cell(1, 1))
            status_sheet.freeze_panes = "A7"
            status_sheet.column_dimensions["A"].width = style_sheet.column_dimensions["A"].width
            for column_index in range(2, 2 + len(years)):
                status_sheet.column_dimensions[get_column_letter(column_index)].width = style_sheet.column_dimensions["B"].width
            status_sheet.column_dimensions[get_column_letter(total_column_index)].width = style_sheet.column_dimensions["G"].width
            status_sheet.column_dimensions[get_column_letter(writeup_column_index)].width = style_sheet.column_dimensions["I"].width
            self._copy_cell_style(style_sheet["H2"], status_sheet.cell(2, writeup_column_index))

            overall_year_totals = db_rows.groupby("__amazon_year")["__amazon_total"].sum()
            for offset, year in enumerate(years, start=2):
                status_sheet.cell(3, offset).value = float(overall_year_totals.get(year, 0))
                self._copy_cell_style(style_sheet["B3"], status_sheet.cell(3, offset))
                status_sheet.cell(4, offset).value = str(year)
                self._copy_cell_style(style_sheet["B4"], status_sheet.cell(4, offset))
            status_sheet.cell(3, total_column_index).value = float(db_rows["__amazon_total"].sum())
            self._copy_cell_style(style_sheet["G3"], status_sheet.cell(3, total_column_index))
            status_sheet.cell(3, writeup_column_index).value = float(
                db_rows.loc[db_rows["__amazon_writeup_sent"], "__amazon_total"].sum()
            )
            self._copy_cell_style(style_sheet["H3"], status_sheet.cell(3, writeup_column_index))
            status_sheet.cell(4, total_column_index).value = "Total Amount"
            self._copy_cell_style(style_sheet["G4"], status_sheet.cell(4, total_column_index))
            for column_index, label, style_coordinate in ((writeup_column_index, "Write-up Sent", "H4"),):
                status_sheet.cell(4, column_index).value = label
                self._copy_cell_style(style_sheet[style_coordinate], status_sheet.cell(4, column_index))

            current_row = 7
            for account_code in AMAZON_ACCOUNT_CODES:
                account_rows = db_rows.loc[db_rows[acct_column] == account_code].copy()
                account_total = float(account_rows["__amazon_total"].sum())
                if abs(account_total) < 0.000001:
                    self.logger.log(
                        f"Skipped {account_code} DB status section because its Total Amount is 0.00.",
                        "NOTICE",
                    )
                    continue

                account_name = "AMAZON"
                if name_column and not account_rows.empty:
                    name_values = account_rows[name_column].dropna().astype(str).str.strip()
                    if not name_values.empty:
                        account_name = name_values.iloc[0]

                status_sheet.merge_cells(
                    start_row=current_row,
                    start_column=1,
                    end_row=current_row,
                    end_column=last_column_index,
                )
                status_sheet.cell(current_row, 1).value = f"{account_code} - {account_name} (Open DB's)"
                self._copy_cell_style(style_sheet["A8"], status_sheet.cell(current_row, 1))
                current_row += 1

                status_sheet.cell(current_row, 1).value = "Reason / Order Column"
                self._copy_cell_style(style_sheet["A9"], status_sheet.cell(current_row, 1))
                for offset, year in enumerate(years, start=2):
                    status_sheet.cell(current_row, offset).value = str(year)
                    self._copy_cell_style(style_sheet["B9"], status_sheet.cell(current_row, offset))
                for column_index, label, style_coordinate in (
                    (total_column_index, "Total Amount", "G9"),
                    (writeup_column_index, "Write-up Sent", "I9"),
                ):
                    status_sheet.cell(current_row, column_index).value = label
                    self._copy_cell_style(style_sheet[style_coordinate], status_sheet.cell(current_row, column_index))
                current_row += 1

                pivot = account_rows.pivot_table(
                    index="__amazon_order",
                    columns="__amazon_year",
                    values="__amazon_total",
                    aggfunc="sum",
                    fill_value=0,
                ).reindex(columns=years, fill_value=0).sort_index()
                writeup_order_totals = account_rows.loc[account_rows["__amazon_writeup_sent"]].groupby(
                    "__amazon_order"
                )["__amazon_total"].sum()
                for order_value, values in pivot.iterrows():
                    status_sheet.cell(current_row, 1).value = order_value
                    self._copy_cell_style(style_sheet["A10"], status_sheet.cell(current_row, 1))
                    for offset, year in enumerate(years, start=2):
                        status_sheet.cell(current_row, offset).value = float(values[year])
                        self._copy_cell_style(style_sheet["B10"], status_sheet.cell(current_row, offset))
                    status_sheet.cell(current_row, total_column_index).value = float(values.sum())
                    self._copy_cell_style(style_sheet["G10"], status_sheet.cell(current_row, total_column_index))
                    status_sheet.cell(current_row, writeup_column_index).value = float(
                        writeup_order_totals.get(order_value, 0)
                    )
                    self._copy_cell_style(style_sheet["I10"], status_sheet.cell(current_row, writeup_column_index))
                    current_row += 1

                status_sheet.cell(current_row, 1).value = "ACCOUNT TOTAL"
                self._copy_cell_style(style_sheet["A20"], status_sheet.cell(current_row, 1))
                for offset, year in enumerate(years, start=2):
                    total_value = float(account_rows.loc[account_rows["__amazon_year"] == year, "__amazon_total"].sum())
                    status_sheet.cell(current_row, offset).value = total_value
                    self._copy_cell_style(style_sheet["B20"], status_sheet.cell(current_row, offset))
                status_sheet.cell(current_row, total_column_index).value = account_total
                self._copy_cell_style(style_sheet["G20"], status_sheet.cell(current_row, total_column_index))
                status_sheet.cell(current_row, writeup_column_index).value = float(
                    account_rows.loc[account_rows["__amazon_writeup_sent"], "__amazon_total"].sum()
                )
                self._copy_cell_style(style_sheet["I20"], status_sheet.cell(current_row, writeup_column_index))
                current_row += 2

            first_unused_row = current_row - 1
            if first_unused_row <= status_sheet.max_row:
                status_sheet.delete_rows(
                    first_unused_row,
                    status_sheet.max_row - first_unused_row + 1,
                )
            set_workbook_author_metadata(output_workbook)
            output_workbook.save(output_path)
        finally:
            style_workbook.close()
            output_workbook.close()
        self.logger.log(
            f"Built DB's Status from {len(db_rows)} Ref Type DB rows across {len(years)} Date year(s).",
            "SUCCESS",
        )

    def _build_credit_status_sheet(
        self,
        output_path: Path,
        combined: pd.DataFrame,
        acct_column: str,
        ref_type_column: str,
    ) -> None:
        """Build the Amazon credit year summary with a placeholder Status column."""
        date_column = self._find_column(combined, "DATE")
        total_column = self._find_column(combined, "TOTAL")
        if not date_column or not total_column:
            raise ValueError("Amazon CR status requires Date and Total columns.")

        credit_rows = combined.loc[combined[ref_type_column] == "CR"].copy()
        credit_rows[date_column] = pd.to_datetime(credit_rows[date_column], errors="coerce")
        credit_rows["__amazon_year"] = credit_rows[date_column].dt.year
        credit_rows["__amazon_total"] = pd.to_numeric(credit_rows[total_column], errors="coerce").fillna(0)
        credit_rows = credit_rows.dropna(subset=["__amazon_year"])
        credit_rows["__amazon_year"] = credit_rows["__amazon_year"].astype(int)
        years = sorted(credit_rows["__amazon_year"].unique().tolist())
        if not years:
            raise ValueError("Amazon CR status found no valid Date values in Ref Type CR rows.")

        name_column = self._find_column(combined, "NAME")
        style_workbook = load_workbook(AMAZON_STATUS_STYLE_TEMPLATE_PATH, read_only=False, data_only=False)
        output_workbook = load_workbook(output_path, read_only=False, data_only=False)
        try:
            style_sheet = style_workbook["CR's Status"]
            status_sheet = output_workbook["CR's Status"]
            for merged_range in list(status_sheet.merged_cells.ranges):
                status_sheet.unmerge_cells(str(merged_range))
            for row in status_sheet.iter_rows():
                for cell in row:
                    cell.value = None

            total_column_index = 2 + len(years)
            status_column_index = total_column_index + 1
            last_column_index = status_column_index
            status_sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_column_index)
            status_sheet.cell(1, 1).value = "AMAZON OPEN AGEING (CR) STATUS SUMMARY"
            self._copy_cell_style(style_sheet["A1"], status_sheet.cell(1, 1))
            status_sheet.freeze_panes = "A7"
            status_sheet.column_dimensions["A"].width = style_sheet.column_dimensions["A"].width
            for column_index in range(2, 2 + len(years)):
                status_sheet.column_dimensions[get_column_letter(column_index)].width = style_sheet.column_dimensions["B"].width
            status_sheet.column_dimensions[get_column_letter(total_column_index)].width = style_sheet.column_dimensions["H"].width
            status_sheet.column_dimensions[get_column_letter(status_column_index)].width = style_sheet.column_dimensions["I"].width

            overall_year_totals = credit_rows.groupby("__amazon_year")["__amazon_total"].sum()
            status_sheet.cell(3, 2).value = float(credit_rows["__amazon_total"].sum())
            self._copy_cell_style(style_sheet["B3"], status_sheet.cell(3, 2))
            status_sheet.cell(4, 2).value = "Total Amount"
            self._copy_cell_style(style_sheet["B4"], status_sheet.cell(4, 2))
            for offset, year in enumerate(years, start=3):
                status_sheet.cell(3, offset).value = float(overall_year_totals.get(year, 0))
                self._copy_cell_style(style_sheet["C3"], status_sheet.cell(3, offset))
                status_sheet.cell(4, offset).value = str(year)
                self._copy_cell_style(style_sheet["C4"], status_sheet.cell(4, offset))

            status_sheet.merge_cells(start_row=6, start_column=1, end_row=6, end_column=last_column_index)
            status_sheet.cell(6, 1).value = "YEAR-WISE AMOUNT BREAKDOWN"
            self._copy_cell_style(style_sheet["A6"], status_sheet.cell(6, 1))
            status_sheet.cell(7, 1).value = "Account Name"
            self._copy_cell_style(style_sheet["A7"], status_sheet.cell(7, 1))
            for offset, year in enumerate(years, start=2):
                status_sheet.cell(7, offset).value = str(year)
                self._copy_cell_style(style_sheet["B7"], status_sheet.cell(7, offset))
            status_sheet.cell(7, total_column_index).value = "Total Amount"
            self._copy_cell_style(style_sheet["H7"], status_sheet.cell(7, total_column_index))
            status_sheet.cell(7, status_column_index).value = "Status"
            self._copy_cell_style(style_sheet["I7"], status_sheet.cell(7, status_column_index))

            current_row = 8
            for account_code in AMAZON_ACCOUNT_CODES:
                account_rows = credit_rows.loc[credit_rows[acct_column] == account_code].copy()
                account_total = float(account_rows["__amazon_total"].sum())
                if abs(account_total) < 0.000001:
                    self.logger.log(
                        f"Skipped {account_code} CR status row because its Total Amount is 0.00.",
                        "NOTICE",
                    )
                    continue

                account_name = "AMAZON"
                if name_column and not account_rows.empty:
                    name_values = account_rows[name_column].dropna().astype(str).str.strip()
                    if not name_values.empty:
                        account_name = name_values.iloc[0]
                status_sheet.cell(current_row, 1).value = f"{account_code} ({account_name})"
                self._copy_cell_style(style_sheet["A8"], status_sheet.cell(current_row, 1))
                for offset, year in enumerate(years, start=2):
                    amount = float(account_rows.loc[account_rows["__amazon_year"] == year, "__amazon_total"].sum())
                    status_sheet.cell(current_row, offset).value = amount
                    self._copy_cell_style(style_sheet["B8"], status_sheet.cell(current_row, offset))
                status_sheet.cell(current_row, total_column_index).value = account_total
                self._copy_cell_style(style_sheet["H8"], status_sheet.cell(current_row, total_column_index))
                status_sheet.cell(current_row, status_column_index).value = "-"
                self._copy_cell_style(style_sheet["I8"], status_sheet.cell(current_row, status_column_index))
                current_row += 1

            status_sheet.cell(current_row, 1).value = "Total"
            self._copy_cell_style(style_sheet["A12"], status_sheet.cell(current_row, 1))
            for offset, year in enumerate(years, start=2):
                status_sheet.cell(current_row, offset).value = float(overall_year_totals.get(year, 0))
                self._copy_cell_style(style_sheet["B12"], status_sheet.cell(current_row, offset))
            status_sheet.cell(current_row, total_column_index).value = float(credit_rows["__amazon_total"].sum())
            self._copy_cell_style(style_sheet["H12"], status_sheet.cell(current_row, total_column_index))
            status_sheet.cell(current_row, status_column_index).value = "-"
            self._copy_cell_style(style_sheet["I12"], status_sheet.cell(current_row, status_column_index))

            if current_row < status_sheet.max_row:
                status_sheet.delete_rows(current_row + 1, status_sheet.max_row - current_row)
            set_workbook_author_metadata(output_workbook)
            output_workbook.save(output_path)
        finally:
            style_workbook.close()
            output_workbook.close()
        self.logger.log(
            f"Built CR's Status from {len(credit_rows)} Ref Type CR rows across {len(years)} Date year(s).",
            "SUCCESS",
        )

    def _build_invoice_status_sheet(
        self,
        output_path: Path,
        combined: pd.DataFrame,
        acct_column: str,
        ref_type_column: str,
    ) -> None:
        """Build the Amazon invoice year summary and latest-year paid totals."""
        date_column = self._find_column(combined, "DATE")
        total_column = self._find_column(combined, "TOTAL")
        comment_column = self._find_column(combined, "COMMENT")
        if not date_column or not total_column or not comment_column:
            raise ValueError("Amazon Invoice status requires Date, Total, and Comment columns.")

        invoice_rows = combined.loc[combined[ref_type_column] == "IN"].copy()
        invoice_rows[date_column] = pd.to_datetime(invoice_rows[date_column], errors="coerce")
        invoice_rows["__amazon_year"] = invoice_rows[date_column].dt.year
        invoice_rows["__amazon_total"] = pd.to_numeric(invoice_rows[total_column], errors="coerce").fillna(0)
        invoice_rows = invoice_rows.dropna(subset=["__amazon_year"])
        invoice_rows["__amazon_year"] = invoice_rows["__amazon_year"].astype(int)
        years = sorted(invoice_rows["__amazon_year"].unique().tolist())
        if not years:
            raise ValueError("Amazon Invoice status found no valid Date values in Ref Type IN rows.")

        latest_year = years[-1]
        invoice_rows["__amazon_paid_latest_year"] = (
            (invoice_rows["__amazon_year"] == latest_year)
            & invoice_rows[comment_column].map(self._is_paid_comment)
        )
        paid_comment_year_totals = (
            invoice_rows.loc[invoice_rows["__amazon_paid_latest_year"]]
            .groupby([acct_column, comment_column, "__amazon_year"], dropna=False)["__amazon_total"]
            .sum()
        )
        paid_account_totals = paid_comment_year_totals.groupby(level=0).sum()

        name_column = self._find_column(combined, "NAME")
        style_workbook = load_workbook(AMAZON_STATUS_STYLE_TEMPLATE_PATH, read_only=False, data_only=False)
        output_workbook = load_workbook(output_path, read_only=False, data_only=False)
        try:
            style_sheet = style_workbook["Invoice's Status"]
            status_sheet = output_workbook["Invoice's Status"]
            for merged_range in list(status_sheet.merged_cells.ranges):
                status_sheet.unmerge_cells(str(merged_range))
            for row in status_sheet.iter_rows():
                for cell in row:
                    cell.value = None

            total_column_index = 2 + len(years)
            paid_column_index = total_column_index + 1
            last_column_index = paid_column_index
            status_sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_column_index)
            status_sheet.cell(1, 1).value = "AMAZON OPEN AGEING (INV) STATUS SUMMARY"
            self._copy_cell_style(style_sheet["A1"], status_sheet.cell(1, 1))
            status_sheet.freeze_panes = "A7"
            status_sheet.column_dimensions["A"].width = style_sheet.column_dimensions["A"].width
            for column_index in range(2, 2 + len(years)):
                status_sheet.column_dimensions[get_column_letter(column_index)].width = style_sheet.column_dimensions["B"].width
            status_sheet.column_dimensions[get_column_letter(total_column_index)].width = style_sheet.column_dimensions["H"].width
            status_sheet.column_dimensions[get_column_letter(paid_column_index)].width = style_sheet.column_dimensions["I"].width

            overall_year_totals = invoice_rows.groupby("__amazon_year")["__amazon_total"].sum()
            status_sheet.cell(3, 2).value = float(invoice_rows["__amazon_total"].sum())
            self._copy_cell_style(style_sheet["B3"], status_sheet.cell(3, 2))
            status_sheet.cell(4, 2).value = "Total Amount"
            self._copy_cell_style(style_sheet["B4"], status_sheet.cell(4, 2))
            for offset, year in enumerate(years, start=3):
                status_sheet.cell(3, offset).value = float(overall_year_totals.get(year, 0))
                self._copy_cell_style(style_sheet["C3"], status_sheet.cell(3, offset))
                status_sheet.cell(4, offset).value = str(year)
                self._copy_cell_style(style_sheet["C4"], status_sheet.cell(4, offset))

            status_sheet.merge_cells(start_row=6, start_column=1, end_row=6, end_column=last_column_index)
            status_sheet.cell(6, 1).value = "YEAR-WISE AMOUNT BREAKDOWN"
            self._copy_cell_style(style_sheet["A6"], status_sheet.cell(6, 1))
            status_sheet.cell(7, 1).value = "Account Name"
            self._copy_cell_style(style_sheet["A7"], status_sheet.cell(7, 1))
            for offset, year in enumerate(years, start=2):
                status_sheet.cell(7, offset).value = str(year)
                self._copy_cell_style(style_sheet["B7"], status_sheet.cell(7, offset))
            status_sheet.cell(7, total_column_index).value = "Total Amount"
            self._copy_cell_style(style_sheet["H7"], status_sheet.cell(7, total_column_index))
            status_sheet.cell(7, paid_column_index).value = f"Paid @{latest_year}"
            self._copy_cell_style(style_sheet["I7"], status_sheet.cell(7, paid_column_index))

            current_row = 8
            for account_code in AMAZON_ACCOUNT_CODES:
                account_rows = invoice_rows.loc[invoice_rows[acct_column] == account_code].copy()
                account_total = float(account_rows["__amazon_total"].sum())
                if abs(account_total) < 0.000001:
                    self.logger.log(
                        f"Skipped {account_code} Invoice status row because its Total Amount is 0.00.",
                        "NOTICE",
                    )
                    continue

                account_name = "AMAZON"
                if name_column and not account_rows.empty:
                    name_values = account_rows[name_column].dropna().astype(str).str.strip()
                    if not name_values.empty:
                        account_name = name_values.iloc[0]
                status_sheet.cell(current_row, 1).value = f"{account_code} ({account_name})"
                self._copy_cell_style(style_sheet["A8"], status_sheet.cell(current_row, 1))
                for offset, year in enumerate(years, start=2):
                    amount = float(account_rows.loc[account_rows["__amazon_year"] == year, "__amazon_total"].sum())
                    status_sheet.cell(current_row, offset).value = amount
                    self._copy_cell_style(style_sheet["B8"], status_sheet.cell(current_row, offset))
                status_sheet.cell(current_row, total_column_index).value = account_total
                self._copy_cell_style(style_sheet["H8"], status_sheet.cell(current_row, total_column_index))
                status_sheet.cell(current_row, paid_column_index).value = float(paid_account_totals.get(account_code, 0))
                self._copy_cell_style(style_sheet["I8"], status_sheet.cell(current_row, paid_column_index))
                current_row += 1

            status_sheet.cell(current_row, 1).value = "Total"
            self._copy_cell_style(style_sheet["A12"], status_sheet.cell(current_row, 1))
            for offset, year in enumerate(years, start=2):
                status_sheet.cell(current_row, offset).value = float(overall_year_totals.get(year, 0))
                self._copy_cell_style(style_sheet["B12"], status_sheet.cell(current_row, offset))
            status_sheet.cell(current_row, total_column_index).value = float(invoice_rows["__amazon_total"].sum())
            self._copy_cell_style(style_sheet["H12"], status_sheet.cell(current_row, total_column_index))
            status_sheet.cell(current_row, paid_column_index).value = float(paid_account_totals.sum())
            self._copy_cell_style(style_sheet["I12"], status_sheet.cell(current_row, paid_column_index))

            if current_row < status_sheet.max_row:
                status_sheet.delete_rows(current_row + 1, status_sheet.max_row - current_row)
            set_workbook_author_metadata(output_workbook)
            output_workbook.save(output_path)
        finally:
            style_workbook.close()
            output_workbook.close()
        self.logger.log(
            f"Built Invoice's Status from {len(invoice_rows)} Ref Type IN rows across {len(years)} Date year(s); "
            f"Paid @{latest_year} groups Acct, Comment, and Date year where Comment contains Paid on, "
            "a payment-number phrase (including small spelling errors), and a payment number.",
            "SUCCESS",
        )
