# -*- coding: utf-8 -*-
"""Doc file PO (Excel / CSV) theo mau PO Import_Template."""

from __future__ import annotations

import io
from pathlib import Path

import pandas as pd

from config import OPTIONAL_FIELDS, REQUIRED_FIELDS

ROOT_DIR = Path(__file__).resolve().parent.parent
PO_TEMPLATE_PATH = ROOT_DIR / "templates" / "PO Import_Template.csv"
VAT_TEMPLATE_PATH = ROOT_DIR / "templates" / "VAT_Template.csv"
# Bảng thuế người dùng — lưu local, giữ sau F5 / mở lại app
VAT_PERSIST_PATH = ROOT_DIR / "data" / "VAT_Tax.xlsx"

_TEXT_COLUMN_ALIASES = {
    alias.strip().lower()
    for aliases in {**REQUIRED_FIELDS, **OPTIONAL_FIELDS}.values()
    for alias in aliases
    if alias.strip().lower()
    not in {
        "order qty",
        "quantity",
        "qty",
        "so luong",
        "sl",
        "unit price",
        "price",
        "don gia",
    }
}


def _cast_text_columns(df: pd.DataFrame) -> pd.DataFrame:
    work = df.copy()
    for col in work.columns:
        if str(col).strip().lower() in _TEXT_COLUMN_ALIASES:
            work[col] = work[col].map(lambda v: "" if pd.isna(v) else str(v).strip())
    return work


def _read_csv_bytes(data: bytes) -> pd.DataFrame:
    for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            return pd.read_csv(io.BytesIO(data), encoding=encoding, dtype=str)
        except UnicodeDecodeError:
            continue
    return pd.read_csv(io.BytesIO(data), encoding="latin-1", dtype=str)


def load_po_dataframe(uploaded_file) -> pd.DataFrame:
    """Doc file PO upload (.xlsx, .xls, .csv)."""
    name = uploaded_file.name.lower()
    data = uploaded_file.getvalue()

    if name.endswith(".csv"):
        df = _read_csv_bytes(data)
    else:
        # dtype=str giu so 0 dau (Store code / Customer account)
        df = pd.read_excel(io.BytesIO(data), engine="openpyxl", dtype=str)
        df = _cast_text_columns(df)

    return df.fillna("")


def load_po_template() -> pd.DataFrame:
    """Nap mau PO Import_Template.csv."""
    for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            return pd.read_csv(PO_TEMPLATE_PATH, encoding=encoding, dtype=str).fillna("")
        except UnicodeDecodeError:
            continue
    return pd.read_csv(PO_TEMPLATE_PATH, encoding="latin-1", dtype=str).fillna("")


def load_vat_template() -> pd.DataFrame:
    """Nap mau VAT_Template.csv."""
    for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            return pd.read_csv(VAT_TEMPLATE_PATH, encoding=encoding, dtype=str).fillna("")
        except UnicodeDecodeError:
            continue
    return pd.read_csv(VAT_TEMPLATE_PATH, encoding="latin-1", dtype=str).fillna("")


def load_vat_persisted() -> pd.DataFrame | None:
    """Nap bang thue da luu (data/VAT_Tax.xlsx). Tra None neu chua co."""
    if not VAT_PERSIST_PATH.exists():
        return None
    try:
        df = pd.read_excel(VAT_PERSIST_PATH, engine="openpyxl", dtype=str)
        return df.fillna("")
    except Exception:
        return None


def save_vat_persisted(df: pd.DataFrame) -> Path:
    """Ghi bang thue ra data/VAT_Tax.xlsx (tu tao thu muc)."""
    VAT_PERSIST_PATH.parent.mkdir(parents=True, exist_ok=True)
    export = df.copy().fillna("")
    export.to_excel(VAT_PERSIST_PATH, index=False, sheet_name="Tax", engine="openpyxl")
    return VAT_PERSIST_PATH

