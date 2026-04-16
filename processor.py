"""
processor.py
============
Generic data processing — clean, merge, filter, KPI, summarize.
No Streamlit imports.
"""

import io
import pandas as pd
from typing import Optional


def load_file(uploaded_file) -> pd.DataFrame:
    name = uploaded_file.name.lower()
    raw  = uploaded_file.read()
    uploaded_file.seek(0)

    if name.endswith(".csv"):
        for enc in ["utf-8-sig", "utf-8", "cp874", "tis-620", "latin-1"]:
            try:
                return pd.read_csv(io.BytesIO(raw), encoding=enc)
            except (UnicodeDecodeError, pd.errors.ParserError):
                continue
        raise ValueError(f"Cannot decode {uploaded_file.name}")

    if name.endswith((".xlsx", ".xls")):
        return pd.read_excel(io.BytesIO(raw))

    raise ValueError(f"Unsupported type: {uploaded_file.name}")


def standardize_columns(df, col_map, drop_duplicates=True, drop_empty_rows=True, strip_whitespace=True):
    df = df.copy()
    if col_map:
        df = df.rename(columns=col_map)
    df.columns = [str(c).strip() for c in df.columns]
    if strip_whitespace:
        for col in df.select_dtypes(include="object").columns:
            df[col] = df[col].astype(str).str.strip().replace({"nan": None, "": None})
    if drop_empty_rows:
        df = df.dropna(how="all")
    if drop_duplicates:
        df = df.drop_duplicates()
    return df.reset_index(drop=True)


def merge_files(frames, mode="append", key=None, how="outer"):
    if not frames:
        return pd.DataFrame()
    if mode == "append":
        return pd.concat(frames, ignore_index=True, sort=False)
    if mode == "join":
        result = frames[0]
        for df in frames[1:]:
            overlap = [c for c in df.columns if c in result.columns and c != key]
            df = df.rename(columns={c: c + "_dup" for c in overlap})
            result = pd.merge(result, df, on=key, how=how)
        return result
    raise ValueError(f"Unknown mode: {mode}")


def apply_filters(df, filters):
    result = df.copy()
    for col, value in filters.items():
        if col not in result.columns:
            continue
        if isinstance(value, list):
            result = result[result[col].isin(value)]
        elif isinstance(value, str) and value:
            result = result[result[col].astype(str).str.contains(value, case=False, na=False)]
    return result.reset_index(drop=True)


def calculate_kpis(df, columns):
    kpis = {}
    for col in columns:
        if col not in df.columns:
            continue
        s = pd.to_numeric(df[col], errors="coerce").dropna()
        kpis[col] = {
            "sum": s.sum(), "mean": s.mean(), "count": s.count(),
            "max": s.max(), "min": s.min(), "std": s.std(),
        }
    return kpis


def build_summary_table(df, group_by, agg_cols, agg_func="sum"):
    work = df.copy()
    for col in agg_cols:
        work[col] = pd.to_numeric(work[col], errors="coerce")
    summary = work.groupby(group_by, dropna=False)[agg_cols].agg(agg_func)
    if agg_func in ("sum", "mean"):
        totals = summary.agg(agg_func).rename("TOTAL")
        summary = pd.concat([summary, totals.to_frame().T])
    return summary.reset_index()
