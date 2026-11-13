import os
import re
import traceback
import csv
from pathlib import Path
from typing import Optional

import pandas as pd

from app.adapters.company_loader import CompanyProfile
from app.adapters.mapping_engine import rename_with_aliases
from app.adapters.validation_engine import ensure_required_columns
from app.DbConfigLogger import Logger
from app.workbook_metadata import prepare_pivot_source_sheet, set_workbook_author_metadata


STEP_01_OUTPUT_FILE = "STEP_01_NORMALIZED_INQUIRY.xlsx"
STEP_01A_OUTPUT_FILE = "STEP_01A_CR_IN_DB_GEN.xlsx"


def _normalize_claim_value(value: object) -> str:
    if pd.isna(value):
        return ""

    if isinstance(value, (int, float)):
        numeric_value = float(value)
        if numeric_value.is_integer():
            return str(int(numeric_value))
        return str(value).strip()

    text_value = str(value).strip()
    integer_match = re.fullmatch(r"(\d+)\.0+", text_value)
    if integer_match:
        return integer_match.group(1)

    return text_value


def _prepare_step_01_inquiry_columns(data_frame: pd.DataFrame, logger: Logger) -> tuple[pd.DataFrame, list[str]]:
    prepared = data_frame.copy()
    added_columns: list[str] = []

    if "claim" in prepared.columns:
        prepared["claim"] = prepared["claim"].map(_normalize_claim_value)

    if "order" in prepared.columns and "Ndatee" not in prepared.columns:
        order_text = prepared["order"].astype(str).str.strip()
        extracted = order_text.str.slice(6, 12)
        parsed_dates = pd.to_datetime(extracted, format="%m%d%y", errors="coerce")
        prepared["Ndatee"] = parsed_dates.dt.date
        added_columns.append("Ndatee")

    if "pay_day" not in prepared.columns:
        aging_columns = [
            "current",
            "col_1_30_days",
            "col_31_60_days",
            "col_61_90_days",
            "over_90_days",
        ]
        selected_pay_day = pd.Series(0.0, index=prepared.index, dtype="float64")
        has_selected_value = pd.Series(False, index=prepared.index)

        for column_name in aging_columns:
            if column_name not in prepared.columns:
                continue

            column_values = pd.to_numeric(prepared[column_name], errors="coerce")
            column_has_value = column_values.notna() & column_values.ne(0)
            rows_to_update = (~has_selected_value) & column_has_value
            selected_pay_day.loc[rows_to_update] = column_values.loc[rows_to_update]
            has_selected_value.loc[rows_to_update] = True

        if "ref" in prepared.columns:
            ref_text = prepared["ref"].fillna("").astype(str).str.upper()
            cr_rows = ref_text.str.contains("CR", na=False)
            selected_pay_day.loc[cr_rows] = -selected_pay_day.loc[cr_rows].abs()

        prepared["pay_day"] = selected_pay_day.fillna(0.0)
        added_columns.append("pay_day")

    if "claim" in prepared.columns and "Cclaim" not in prepared.columns:
        claim_text = prepared["claim"].fillna("").astype(str).str.strip()
        prepared["Cclaim"] = claim_text.where(
            ~(claim_text.str.startswith("DB") | claim_text.str.startswith("CR")),
            claim_text.str[-6:],
        )
        added_columns.append("Cclaim")

    if "pay_day" in prepared.columns and "Cclaim" in prepared.columns:
        reordered_columns = list(prepared.columns)
        reordered_columns.remove("Cclaim")
        pay_day_index = reordered_columns.index("pay_day")
        reordered_columns.insert(pay_day_index + 1, "Cclaim")
        prepared = prepared[reordered_columns]

    if added_columns:
        logger.log(f"Step 01 early column preparation applied: {', '.join(added_columns)}", "SUCCESS")

    return prepared, added_columns


def _read_file(path: str) -> pd.DataFrame:
    extension = os.path.splitext(path)[1].lower()
    if extension == ".csv":
        return pd.read_csv(path)

    engine = "xlrd" if extension == ".xls" else "openpyxl"
    return pd.read_excel(path, engine=engine)


def _normalize_sql_column_name(column_name: object, used_names: set[str]) -> str:
    text = "" if pd.isna(column_name) else str(column_name).strip().lower()
    text = re.sub(r"[^a-z0-9]+", "_", text)
    text = re.sub(r"_+", "_", text).strip("_")

    if not text:
        text = "column"
    if text[0].isdigit():
        text = f"col_{text}"

    base_name = text
    suffix = 2
    while text in used_names:
        text = f"{base_name}_{suffix}"
        suffix += 1

    used_names.add(text)
    return text


def _normalize_dataframe_columns(data_frame: pd.DataFrame) -> pd.DataFrame:
    used_names: set[str] = set()
    normalized = data_frame.copy()
    normalized.columns = [
        _normalize_sql_column_name(column_name, used_names)
        for column_name in normalized.columns
    ]
    return normalized


class InquiryProcessor:
    def __init__(self, logger: Logger, company_profile: CompanyProfile | None = None) -> None:
        self.logger = logger
        self.company_profile = company_profile

    def process(self, file_path: str) -> Optional[pd.DataFrame]:
        try:
            self.logger.log(f"Reading inquiry file: {os.path.basename(file_path)}")
            normalized_data_frame = _normalize_dataframe_columns(_read_file(file_path).copy())
            step_config = self.company_profile.steps.get("step_01", {}) if self.company_profile else {}
            mapped_data_frame = rename_with_aliases(normalized_data_frame, step_config.get("column_aliases", {}))
            ensure_required_columns(
                mapped_data_frame,
                step_config.get("required_columns", []),
                "Step 01 inquiry file",
            )
            data_frame, prepared_columns = _prepare_step_01_inquiry_columns(mapped_data_frame, self.logger)
            self.logger.log(
                f"Inquiry ready: {len(data_frame)} rows, {len(data_frame.columns)} SQL-safe columns",
                "SUCCESS",
            )
            if prepared_columns:
                self.logger.log(
                    f"Step 01 inquiry columns added before save/upload: {', '.join(prepared_columns)}",
                    "SUCCESS",
                )
            return data_frame
        except Exception as exc:
            self.logger.log(f"Inquiry processing failed: {exc}", "ERROR")
            self.logger.log(traceback.format_exc(), "ERROR")
            return None


class WorkflowOrchestrator:
    def __init__(
        self,
        logger: Logger,
        inquiry_processor: InquiryProcessor,
    ) -> None:
        self.logger = logger
        self.inquiry_processor = inquiry_processor

    def run_step_01(self, inquiry_path: str, output_dir: Path) -> tuple[bool, dict]:
        try:
            self.logger.log("STEP 01: Normalize inquiry columns for SQL")
            inquiry_df = self.inquiry_processor.process(inquiry_path)
            if inquiry_df is None:
                return False, {"error": "Failed to process inquiry file."}

            output_path = output_dir / STEP_01_OUTPUT_FILE
            with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
                set_workbook_author_metadata(writer.book)
                inquiry_df.to_excel(writer, sheet_name="step_01_inquiry", index=False)
                prepare_pivot_source_sheet(writer.sheets["step_01_inquiry"], "Step01Inquiry")

            self.logger.log(f"Step 01 workbook saved: {output_path}", "SUCCESS")
            self.logger.log("Step 01 completed", "SUCCESS")
            return True, {"inquiry_df": inquiry_df, "output_path": str(output_path)}
        except Exception as exc:
            self.logger.log(f"Workflow failed: {exc}", "ERROR")
            self.logger.log(traceback.format_exc(), "ERROR")
            return False, {"error": str(exc)}


