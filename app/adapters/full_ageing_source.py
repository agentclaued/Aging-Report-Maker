"""Read the Full Ageing sheet (Regular Accounts format) as a source for the Wal Mart and Amazon status reports."""
from __future__ import annotations

from pathlib import Path

import pandas as pd

FULL_AGEING_SHEET_NAME = "full ageing"
HEADER_SCAN_ROWS = 15


def _key(value: object) -> str:
    return "" if pd.isna(value) else " ".join(str(value).strip().casefold().split())


def read_full_ageing_rows(file_path: str, accounts: tuple[str, ...] | list[str]) -> pd.DataFrame | None:
    """Return Full Ageing rows whose Acct is in `accounts`, or None when the file has no Full Ageing sheet.

    The first rows may be titles, so the header row is the first row (within the top rows)
    that contains both Acct and Ref Type.
    """
    if Path(file_path).suffix.lower() not in {".xlsx", ".xlsm", ".xls"}:
        return None
    workbook = pd.ExcelFile(file_path)
    sheet_name = next((name for name in workbook.sheet_names if _key(name) == FULL_AGEING_SHEET_NAME), None)
    if sheet_name is None:
        return None

    raw = pd.read_excel(workbook, sheet_name=sheet_name, header=None).dropna(how="all")
    header_index = None
    for index in range(min(HEADER_SCAN_ROWS, len(raw))):
        keys = {_key(value) for value in raw.iloc[index]}
        if "acct" in keys and "ref type" in keys:
            header_index = index
            break
    if header_index is None:
        raise ValueError("Full Ageing sheet: no header row with both Acct and Ref Type was found in the top rows.")

    headers = raw.iloc[header_index]
    frame = raw.iloc[header_index + 1 :].copy()
    frame.columns = [str(value).strip() if not pd.isna(value) else f"Column {i + 1}" for i, value in enumerate(headers)]
    frame = frame.loc[:, ~frame.columns.duplicated()].dropna(how="all").reset_index(drop=True)
    acct_column = next(column for column in frame.columns if _key(column) == "acct")
    wanted = {code.upper() for code in accounts}
    frame = frame.loc[frame[acct_column].astype(str).str.strip().str.upper().isin(wanted)].reset_index(drop=True)
    if frame.empty:
        raise ValueError(f"Full Ageing sheet has no rows for Acct: {', '.join(sorted(wanted))}.")
    return frame
