from __future__ import annotations

import pandas as pd


def ensure_required_columns(data_frame: pd.DataFrame, required_columns: list[str], context: str) -> None:
    if not required_columns:
        return

    available = {str(column).strip().lower() for column in data_frame.columns}
    missing = [column for column in required_columns if column.strip().lower() not in available]
    if missing:
        raise ValueError(f"{context} is missing required columns: {', '.join(missing)}")

