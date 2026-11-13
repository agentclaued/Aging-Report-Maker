from __future__ import annotations

from typing import Protocol

import pandas as pd

from app.local_cache import cached_frame


FULL_AGEING_SHEET_NAME = "full ageing"
ACCOUNT_ANALYST_LIST_SHEET_NAME = "account analyst list"
REQUIRED_COLUMNS = ("assigned person", "acct")
AGEING_BUCKETS = ("Current", "1-30 Days", "31-60 Days", "61-90 Days", "Over 90 Days")


class _Logger(Protocol):
    def log(self, message: str, level: str = "INFO") -> None: ...


def _column_key(value: object) -> str:
    if pd.isna(value):
        return ""
    return " ".join(str(value).strip().casefold().split())


def _cell_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def _summary_name(value: object) -> str:
    """Return a meaningful source Name; zero is an empty placeholder, not a name."""
    text = _cell_text(value)
    return "" if text in {"0", "0.0", "0.00"} else text


def _first_meaningful_name(values: pd.Series) -> str:
    for value in values:
        name = _summary_name(value)
        if name:
            return name
    return ""


class RegularAccountsProcessor:
    """Read Full Ageing locally and prepare Assigned Person dropdown options."""

    def __init__(self, logger: _Logger) -> None:
        self.logger = logger

    def person_options(self, file_path: str) -> list[dict[str, object]]:
        raw_rows = self._full_ageing_rows(file_path)
        options = self._person_options_from_rows(raw_rows)
        self.logger.log(
            f"Regular Accounts read Full Ageing and prepared {len(options)} Assigned Person option(s).",
            "SUCCESS",
        )
        return options

    def selected_person_rows(self, file_path: str, assigned_person: str) -> pd.DataFrame:
        raw_rows = self._full_ageing_rows(file_path)
        header_row_index, _ = self._find_header_row(raw_rows)
        source_rows = raw_rows.iloc[header_row_index + 1 :].copy()
        source_rows.columns = [
            _column_key(raw_rows.iloc[header_row_index, column_index]) or f"column_{column_index + 1}"
            for column_index in range(raw_rows.shape[1])
        ]
        source_rows = source_rows.loc[:, ~source_rows.columns.duplicated()].copy()
        source_headers = {
            _column_key(raw_rows.iloc[header_row_index, column_index]) or f"column_{column_index + 1}": _cell_text(raw_rows.iloc[header_row_index, column_index])
            or f"Column {column_index + 1}"
            for column_index in range(raw_rows.shape[1])
        }
        source_rows["assigned person"] = source_rows["assigned person"].map(_cell_text)
        source_rows["acct"] = source_rows["acct"].map(_cell_text)
        selected_name = _cell_text(assigned_person)
        selected_rows = source_rows.loc[
            (source_rows["assigned person"] == selected_name)
            & (source_rows["acct"] != "")
        ].copy()
        if selected_rows.empty:
            raise ValueError(f"No Full Ageing rows were found for '{assigned_person}'.")
        selected_rows = selected_rows.reset_index(drop=True)
        if "total" not in selected_rows.columns:
            selected_rows["total"] = self._derived_total_from_ageing_buckets(selected_rows)
            self.logger.log(
                "Regular Accounts created calculation-only Total values from ageing buckets.",
                "SUCCESS",
            )
        selected_rows.attrs["source_headers"] = source_headers
        return selected_rows

    def follow_up_entries(self, selected_rows: pd.DataFrame) -> pd.DataFrame:
        self._ensure_columns(selected_rows, ("status",), "FOLLOW-UP ENTRIES")
        follow_up_rows = selected_rows.loc[selected_rows["status"].map(_cell_text) != ""].copy()
        return self._preserve_source_columns(follow_up_rows, selected_rows.attrs.get("source_headers", {}))

    def ageing_rows(self, selected_rows: pd.DataFrame) -> pd.DataFrame:
        return self._preserve_source_columns(selected_rows, selected_rows.attrs.get("source_headers", {}))

    @staticmethod
    def next_monday(today: pd.Timestamp) -> pd.Timestamp:
        """Return the coming Monday; when today is Monday, today itself."""
        today = pd.Timestamp(today).normalize()
        return today + pd.Timedelta(days=(7 - today.weekday()) % 7)

    def past_due_rows(self, selected_rows: pd.DataFrame, today: pd.Timestamp) -> pd.DataFrame:
        """Invoice (Ref# contains IN) ageing rows whose Due Date is on or before the next Monday (overdue, due today, due soon)."""
        self._ensure_columns(selected_rows, ("due date", "ref#"), "Past-due invoices")
        invoice_rows = self._reference_rows(selected_rows, "in")  # Ref# contains IN, same as the IN summaries
        due_dates = pd.to_datetime(invoice_rows["due date"], errors="coerce").dt.normalize()
        cutoff = self.next_monday(today)
        matching_rows = invoice_rows.loc[due_dates.notna() & (due_dates <= cutoff)].copy()
        # Order: Passed Yesterday, Due Today, Due Tomorrow, Due soon (soonest first), then older Overdue (most recent first).
        today_date = pd.Timestamp(today).normalize()
        matching_due = due_dates.loc[matching_rows.index]
        days_out = (matching_due - today_date).dt.days
        group = days_out.map(lambda days: 0 if days == -1 else 1 if days == 0 else 2 if days == 1 else 3 if days > 1 else 4)
        matching_rows = matching_rows.assign(
            _group=group,
            _days=days_out.where(days_out >= 0, -days_out),
        ).sort_values(["_group", "_days"], kind="stable").drop(columns=["_group", "_days"])
        return self._preserve_source_columns(matching_rows, selected_rows.attrs.get("source_headers", {}))

    def account_detail_rows(self, selected_rows: pd.DataFrame) -> dict[str, pd.DataFrame]:
        details: dict[str, pd.DataFrame] = {}
        for account, account_rows in selected_rows.groupby("acct", sort=True):
            details[str(account)] = self._preserve_source_columns(
                account_rows,
                selected_rows.attrs.get("source_headers", {}),
            )
        return details

    def account_analyst_lookup(self, file_path: str) -> dict[str, dict[str, str]]:
        """Map each Account Analyst List Acct to its Terms and SLM values."""
        workbook = pd.ExcelFile(file_path)
        sheet_name = next(
            (
                name
                for name in workbook.sheet_names
                if name.strip().casefold() == ACCOUNT_ANALYST_LIST_SHEET_NAME
            ),
            None,
        )
        if sheet_name is None:
            raise ValueError("The workbook must contain a worksheet named 'Account Analyst List'.")

        analyst_rows = pd.read_excel(workbook, sheet_name=sheet_name, dtype=object)
        header_lookup = {_column_key(column): str(column) for column in analyst_rows.columns}
        required = ("acct", "terms", "slm")
        missing = [column_name for column_name in required if column_name not in header_lookup]
        if missing:
            raise ValueError(
                "Account Analyst List requires these columns: " + ", ".join(missing) + "."
            )

        lookup: dict[str, dict[str, str]] = {}
        for _, row in analyst_rows.iterrows():
            account_key = _column_key(row[header_lookup["acct"]])
            if not account_key or account_key in lookup:
                continue
            lookup[account_key] = {
                "terms": _cell_text(row[header_lookup["terms"]]),
                "slm": _cell_text(row[header_lookup["slm"]]),
            }
        return lookup

    def summary_rows(self, selected_rows: pd.DataFrame, ref_marker: str, include_cr: bool = False) -> pd.DataFrame:
        self._ensure_columns(selected_rows, ("ref#", "total"), f"{ref_marker} summary")
        report_rows = self._reference_rows(selected_rows, ref_marker)
        bucket_values = self._ageing_bucket_values(report_rows).reset_index(drop=True)
        total_values = pd.to_numeric(report_rows["total"], errors="coerce").fillna(0).reset_index(drop=True)
        summary = pd.concat(
            [report_rows[["acct"]].reset_index(drop=True), bucket_values, total_values.rename("total")],
            axis=1,
        )
        if summary.empty:
            summary = pd.DataFrame(columns=["acct", *[_column_key(bucket) for bucket in AGEING_BUCKETS], "total"])
        summary = summary.groupby("acct", sort=True, as_index=False).sum(numeric_only=True)
        if include_cr:
            cr_rows = self._reference_rows(selected_rows, "cr")
            cr_rows["total"] = pd.to_numeric(cr_rows["total"], errors="coerce").fillna(0)
            cr_totals = cr_rows.groupby("acct", sort=True)["total"].sum().rename("cr")
            summary = summary.set_index("acct").join(cr_totals, how="outer").reset_index()

        source_names = selected_rows[["acct"]].copy()
        source_names["name"] = (
            selected_rows["name"].map(_summary_name)
            if "name" in selected_rows.columns
            else ""
        )
        names_by_account = source_names.groupby("acct", sort=True, as_index=False)["name"].agg(_first_meaningful_name)
        summary = summary.merge(names_by_account, on="acct", how="left")
        amount_columns = [column for column in summary.columns if column not in {"acct", "name"}]
        summary[amount_columns] = summary[amount_columns].apply(pd.to_numeric, errors="coerce").fillna(0)
        summary["name"] = summary["name"].map(_summary_name)
        return summary

    def summary_year_totals(self, selected_rows: pd.DataFrame, ref_marker: str, report_name: str) -> pd.DataFrame:
        """Return matching Ref# totals by the unique calendar years found in Date."""
        self._ensure_columns(selected_rows, ("ref#", "date", "total"), report_name)
        report_rows = self._reference_rows(selected_rows, ref_marker)
        report_rows["year"] = pd.to_datetime(report_rows["date"], errors="coerce").dt.year
        report_rows["total"] = pd.to_numeric(report_rows["total"], errors="coerce").fillna(0)
        report_rows = report_rows.dropna(subset=["year"])
        if report_rows.empty:
            return pd.DataFrame(columns=["year", "total"])
        report_rows["year"] = report_rows["year"].astype(int)
        return report_rows.groupby("year", as_index=False, sort=True)["total"].sum()

    def db_reason_summary(self, selected_rows: pd.DataFrame) -> pd.DataFrame:
        """Summarize DB rows by distinct Order# for DB'S BY REASON."""
        self._ensure_columns(selected_rows, ("ref#", "order#", "total"), "DB'S BY REASON")
        db_rows = self._reference_rows(selected_rows, "db")
        db_rows = db_rows.loc[db_rows["order#"].map(_cell_text) != ""].copy()
        db_rows["order#"] = db_rows["order#"].map(_cell_text)
        bucket_values = self._ageing_bucket_values(db_rows).reset_index(drop=True)
        total_values = pd.to_numeric(db_rows["total"], errors="coerce").fillna(0).reset_index(drop=True)
        names = (
            db_rows["name"].map(_summary_name).reset_index(drop=True)
            if "name" in db_rows.columns
            else pd.Series("", index=db_rows.index).reset_index(drop=True)
        )
        summary = pd.concat(
            [db_rows[["order#"]].reset_index(drop=True), names.rename("name"), bucket_values, total_values.rename("total")],
            axis=1,
        )
        if summary.empty:
            return pd.DataFrame(columns=["order#", "name", *[_column_key(bucket) for bucket in AGEING_BUCKETS], "total"])
        names_by_order = summary.groupby("order#", sort=True, as_index=False)["name"].agg(_first_meaningful_name)
        summary = summary.groupby("order#", sort=True, as_index=False).sum(numeric_only=True)
        return summary.merge(names_by_order, on="order#", how="left")

    def _full_ageing_rows(self, file_path: str) -> pd.DataFrame:
        def load() -> pd.DataFrame:
            workbook = pd.ExcelFile(file_path)
            matching_sheet = next(
                (sheet_name for sheet_name in workbook.sheet_names if sheet_name.strip().casefold() == FULL_AGEING_SHEET_NAME),
                None,
            )
            if matching_sheet is None:
                raise ValueError("The workbook must contain a worksheet named 'Full Ageing'.")
            return pd.read_excel(workbook, sheet_name=matching_sheet, header=None, dtype=object)

        # Parsed once and kept in app/local.db, so the people list and the report share one read.
        return cached_frame(file_path, "full_ageing", load)

    @staticmethod
    def _find_header_row(raw_rows: pd.DataFrame) -> tuple[int, dict[str, int]]:
        for row_index, row in raw_rows.iterrows():
            row_columns = {_column_key(value): column_index for column_index, value in enumerate(row.tolist())}
            if all(column_name in row_columns for column_name in REQUIRED_COLUMNS):
                return int(row_index), row_columns
        raise ValueError("Could not find a header row containing both 'Assigned Person' and 'Acct'.")

    @staticmethod
    def _ensure_columns(rows: pd.DataFrame, required_columns: tuple[str, ...], report_name: str) -> None:
        missing_columns = [column_name for column_name in required_columns if column_name not in rows.columns]
        if missing_columns:
            raise ValueError(f"{report_name} requires these Full Ageing columns: {', '.join(missing_columns)}.")

    @staticmethod
    def _preserve_source_columns(rows: pd.DataFrame, source_headers: dict[str, str]) -> pd.DataFrame:
        source_columns = [column_name for column_name in rows.columns if str(column_name) in source_headers]
        preserved = rows.loc[:, source_columns].copy().reset_index(drop=True)
        preserved.columns = [source_headers.get(str(column_name), str(column_name)) for column_name in preserved.columns]
        if "total" in rows.columns and "total" not in source_headers:
            current_column_index = next(
                (
                    index
                    for index, source_column in enumerate(source_columns)
                    if str(source_column) == "current"
                ),
                len(source_columns),
            )
            preserved.insert(
                current_column_index,
                "Total",
                pd.to_numeric(rows["total"], errors="coerce").fillna(0).reset_index(drop=True),
            )
        return preserved

    @staticmethod
    def _derived_total_from_ageing_buckets(rows: pd.DataFrame) -> pd.Series:
        """Use the first non-zero ageing bucket, scanning from Current to Over 90 Days."""
        bucket_keys = [_column_key(bucket) for bucket in AGEING_BUCKETS]
        RegularAccountsProcessor._ensure_columns(rows, tuple(bucket_keys), "Total fallback calculation")
        derived_total = pd.Series(0.0, index=rows.index, dtype=float)
        assigned_value = pd.Series(False, index=rows.index, dtype=bool)
        for bucket_key in bucket_keys:
            bucket_values = pd.to_numeric(rows[bucket_key], errors="coerce").fillna(0)
            choose_value = (~assigned_value) & bucket_values.ne(0)
            derived_total.loc[choose_value] = bucket_values.loc[choose_value]
            assigned_value.loc[choose_value] = True
        return derived_total

    @staticmethod
    def _ageing_bucket_values(rows: pd.DataFrame) -> pd.DataFrame:
        bucket_keys = [_column_key(bucket) for bucket in AGEING_BUCKETS]
        if all(bucket_key in rows.columns for bucket_key in bucket_keys):
            return rows[bucket_keys].apply(pd.to_numeric, errors="coerce").fillna(0)

        RegularAccountsProcessor._ensure_columns(rows, ("due date", "total"), "Ageing bucket calculation")
        due_dates = pd.to_datetime(rows["due date"], errors="coerce")
        totals = pd.to_numeric(rows["total"], errors="coerce").fillna(0)
        days_overdue = (pd.Timestamp.today().normalize() - due_dates.dt.normalize()).dt.days
        values = pd.DataFrame(0.0, index=rows.index, columns=bucket_keys)
        values.loc[days_overdue <= 0, "current"] = totals
        values.loc[days_overdue.between(1, 30), "1-30 days"] = totals
        values.loc[days_overdue.between(31, 60), "31-60 days"] = totals
        values.loc[days_overdue.between(61, 90), "61-90 days"] = totals
        values.loc[days_overdue > 90, "over 90 days"] = totals
        return values

    @staticmethod
    def _reference_rows(rows: pd.DataFrame, marker: str) -> pd.DataFrame:
        marker_text = marker.strip().casefold()
        if marker_text == "cr":
            if "ref type" not in rows.columns:
                return rows.iloc[0:0].copy()
            return rows.loc[rows["ref type"].map(_cell_text).str.casefold().eq("cr")].copy()
        ref_number_match = rows["ref#"].map(_cell_text).str.casefold().str.contains(marker_text, regex=False, na=False)
        return rows.loc[ref_number_match].copy()

    @staticmethod
    def _person_options_from_rows(raw_rows: pd.DataFrame) -> list[dict[str, object]]:
        header_row_index, header_columns = RegularAccountsProcessor._find_header_row(raw_rows)

        values = raw_rows.iloc[header_row_index + 1 :, [
            header_columns["assigned person"],
            header_columns["acct"],
        ]].copy()
        values.columns = ["assigned_person", "acct"]
        values["assigned_person"] = values["assigned_person"].map(_cell_text)
        values["acct"] = values["acct"].map(_cell_text)
        values = values.loc[(values["assigned_person"] != "") & (values["acct"] != "")]

        if values.empty:
            return []

        unique_person_accounts = values.drop_duplicates(["assigned_person", "acct"])
        person_counts = unique_person_accounts.groupby("assigned_person", sort=True)["acct"].count()
        return [
            {
                "name": str(person_name),
                "count": int(acct_count),
                "label": f"{person_name} @{int(acct_count)}",
            }
            for person_name, acct_count in person_counts.items()
        ]
