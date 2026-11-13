from __future__ import annotations

import re

from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo


WORKBOOK_AUTHOR = "AnsSys & openpyxl"
EXCEL_DATE_FORMAT = "mmddyyyy"


def set_workbook_author_metadata(workbook: Workbook) -> None:
    """Apply the standard author metadata to every generated XLSX workbook."""
    workbook.properties.creator = WORKBOOK_AUTHOR
    workbook.properties.lastModifiedBy = WORKBOOK_AUTHOR
    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True


def prepare_pivot_source_sheet(worksheet, table_name: str, header_row: int = 1) -> None:
    """Format Date columns and expose a normal data sheet as an Excel Table.

    Excel Tables are the approved PivotTable-ready source for generated data sheets.
    They are deliberately not applied to presentation-style summary sheets.
    """
    if worksheet.max_row <= header_row or worksheet.max_column < 1:
        return

    header_lookup = {
        " ".join(str(worksheet.cell(header_row, column).value or "").strip().casefold().split()): column
        for column in range(1, worksheet.max_column + 1)
    }
    for header in ("date", "due date"):
        column = header_lookup.get(header)
        if column is None:
            continue
        for row in range(header_row + 1, worksheet.max_row + 1):
            worksheet.cell(row, column).number_format = EXCEL_DATE_FORMAT

    headers = [worksheet.cell(header_row, column).value for column in range(1, worksheet.max_column + 1)]
    if any(header is None or not str(header).strip() for header in headers):
        return
    if len({str(header).strip().casefold() for header in headers}) != len(headers):
        return

    normalized_name = re.sub(r"[^A-Za-z0-9_]", "_", table_name)
    if not normalized_name or normalized_name[0].isdigit():
        normalized_name = f"Table_{normalized_name}"
    existing_names = {table.displayName.casefold() for table in worksheet.tables.values()}
    candidate = normalized_name[:240]
    suffix = 2
    while candidate.casefold() in existing_names:
        suffix_text = f"_{suffix}"
        candidate = f"{normalized_name[:240 - len(suffix_text)]}{suffix_text}"
        suffix += 1

    table = Table(
        displayName=candidate,
        ref=f"A{header_row}:{get_column_letter(worksheet.max_column)}{worksheet.max_row}",
    )
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    worksheet.add_table(table)
