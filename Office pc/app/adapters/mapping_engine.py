from __future__ import annotations

import pandas as pd


def rename_with_aliases(data_frame: pd.DataFrame, alias_map: dict[str, list[str] | str]) -> pd.DataFrame:
    if not alias_map:
        return data_frame

    renamed = data_frame.copy()
    reverse_lookup: dict[str, str] = {}
    for canonical_name, aliases in alias_map.items():
        alias_values = aliases if isinstance(aliases, list) else [aliases]
        for alias in alias_values:
            reverse_lookup[str(alias).strip().lower()] = canonical_name

    rename_map: dict[str, str] = {}
    for column_name in renamed.columns:
        mapped_name = reverse_lookup.get(str(column_name).strip().lower())
        if mapped_name:
            rename_map[str(column_name)] = mapped_name

    if rename_map:
        renamed = renamed.rename(columns=rename_map)
    return renamed




