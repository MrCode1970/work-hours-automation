from __future__ import annotations

import os
from typing import Iterable

import pandas as pd

EXPECTED_COLUMNS = ["תאריך", "כניסה", "יציאה", "אתר", "הערות"]


def _ensure_columns(df: pd.DataFrame, columns: Iterable[str]) -> pd.DataFrame:
    df = df.copy()
    for col in columns:
        if col not in df.columns:
            df[col] = ""
    return df[list(columns)]


def write_local_xlsx(df: pd.DataFrame, out_path: str) -> None:
    temp_path = f"{out_path}.tmp"
    df = _ensure_columns(df, EXPECTED_COLUMNS)
    try:
        with open(temp_path, "wb") as fh:
            with pd.ExcelWriter(fh, engine="openpyxl") as writer:
                df.to_excel(writer, index=False)
        os.replace(temp_path, out_path)
    except Exception:
        if os.path.exists(temp_path):
            os.remove(temp_path)
        raise
