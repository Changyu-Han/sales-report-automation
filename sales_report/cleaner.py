"""Load messy sales exports and turn them into one clean, consistent table."""
from __future__ import annotations

import logging
from pathlib import Path

import pandas as pd

log = logging.getLogger(__name__)

# Every column name we've seen in the wild, mapped to one standard name.
# Add new aliases here when a client's file uses a different header.
COLUMN_ALIASES = {
    "date": ["date", "order_date", "sale date", "transaction date"],
    "product": ["product", "item", "product name", "sku name"],
    "quantity": ["qty", "quantity", "units"],
    "unit_price": ["unit price", "price", "unit_price"],
}
STANDARD_COLUMNS = list(COLUMN_ALIASES)
# Optional: if a file has an order/receipt number we use it to detect duplicates.
OPTIONAL_ALIASES = {"order_id": ["order_id", "order #", "order no", "receipt_no", "receipt"]}
SUPPORTED_SUFFIXES = {".csv", ".xlsx", ".xls"}


def read_file(path: Path) -> pd.DataFrame:
    """Read a CSV or Excel file as raw strings so nothing is silently converted."""
    if path.suffix.lower() == ".csv":
        return pd.read_csv(path, dtype=str, encoding="utf-8-sig")
    return pd.read_excel(path, dtype=str)


def standardise_columns(df: pd.DataFrame, source: str) -> pd.DataFrame:
    """Rename known aliases to standard names; fail loudly if something is missing."""
    all_aliases = {**COLUMN_ALIASES, **OPTIONAL_ALIASES}
    lookup = {alias: std for std, aliases in all_aliases.items() for alias in aliases}
    renamed = {col: lookup[col.strip().lower()] for col in df.columns if col.strip().lower() in lookup}
    df = df.rename(columns=renamed)
    missing = [c for c in STANDARD_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"{source}: could not find column(s) {missing}. Columns present: {list(df.columns)}")
    keep = STANDARD_COLUMNS + [c for c in OPTIONAL_ALIASES if c in df.columns]
    return df[keep]


def parse_dates(series: pd.Series) -> pd.Series:
    """Handle ISO (2026-09-01), slash (2026/09/01) and Australian DD/MM/YYYY dates."""
    s = series.str.strip()
    iso_like = s.str.match(r"^\d{4}[-/]\d{1,2}[-/]\d{1,2}")
    parsed = pd.Series(pd.NaT, index=s.index)
    parsed[iso_like] = pd.to_datetime(s[iso_like].str.replace("/", "-"), errors="coerce", format="%Y-%m-%d")
    parsed[~iso_like] = pd.to_datetime(s[~iso_like], errors="coerce", dayfirst=True)
    return parsed


def parse_money(series: pd.Series) -> pd.Series:
    """'$1,234.50' -> 1234.5"""
    return pd.to_numeric(series.str.replace(r"[$,\s]", "", regex=True), errors="coerce")


def clean(df: pd.DataFrame, store: str) -> tuple[pd.DataFrame, dict]:
    """Clean one store's data. Returns the clean table and a summary of what was fixed."""
    stats = {"store": store, "rows_in": len(df)}

    # Remove duplicate exports first, before normalising, so two genuine sales of the
    # same product on the same day are never merged. Prefer the order number if present.
    before = len(df)
    if "order_id" in df.columns:
        has_id = df["order_id"].notna()
        df = pd.concat([df[has_id].drop_duplicates(subset="order_id"), df[~has_id]])
    else:
        log.warning("%s: no order number column; only exact duplicate rows are removed", store)
        df = df.drop_duplicates()
    stats["dropped_duplicates"] = before - len(df)

    df = df.copy()
    df["date"] = parse_dates(df["date"].astype("string"))
    df["product"] = df["product"].astype("string").str.strip().str.title()
    df["quantity"] = pd.to_numeric(df["quantity"], errors="coerce")
    df["unit_price"] = parse_money(df["unit_price"].astype("string"))

    before = len(df)
    df = df.dropna(subset=STANDARD_COLUMNS)
    df = df[df["product"] != ""]
    stats["dropped_incomplete"] = before - len(df)

    stats["refund_rows"] = int((df["quantity"] < 0).sum())  # kept: refunds reduce revenue

    df["store"] = store
    df["revenue"] = (df["quantity"] * df["unit_price"]).round(2)
    stats["rows_out"] = len(df)
    return df.reset_index(drop=True), stats


def store_name_from_file(path: Path) -> str:
    """store_b_chatswood.xlsx -> 'Chatswood'"""
    return path.stem.split("_")[-1].replace("-", " ").title()


def load_all(input_dir: Path) -> tuple[pd.DataFrame, list[dict]]:
    """Load and clean every supported file in a folder and combine them."""
    files = sorted(p for p in input_dir.iterdir() if p.suffix.lower() in SUPPORTED_SUFFIXES)
    if not files:
        raise FileNotFoundError(f"No CSV or Excel files found in {input_dir}")

    frames, all_stats = [], []
    for path in files:
        raw = standardise_columns(read_file(path), path.name)
        cleaned, stats = clean(raw, store_name_from_file(path))
        stats["file"] = path.name
        log.info("%s: %d rows in -> %d clean rows", path.name, stats["rows_in"], stats["rows_out"])
        frames.append(cleaned)
        all_stats.append(stats)
    combined = pd.concat(frames, ignore_index=True).sort_values("date", ignore_index=True)
    return combined, all_stats
