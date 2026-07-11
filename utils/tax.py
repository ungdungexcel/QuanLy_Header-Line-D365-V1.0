# -*- coding: utf-8 -*-
"""Tra cuu thue VAT theo ma hang."""

from __future__ import annotations

import pandas as pd


def _normalize_tax_table(tax_df: pd.DataFrame) -> pd.DataFrame:
    work = tax_df.copy()
    rename = {}
    cols_lower = {str(c).strip().lower(): c for c in work.columns}
    for alias, target in {
        "item number": "item",
        "item": "item",
        "unit": "uom",
        "vat": "vat",
    }.items():
        if alias in cols_lower:
            rename[cols_lower[alias]] = target

    work = work.rename(columns=rename)
    if "item" not in work.columns or "vat" not in work.columns:
        return pd.DataFrame(columns=["item", "uom", "vat"])

    for col in ("item", "uom", "vat"):
        work[col] = work[col].astype(str).str.strip()

    work = work[(work["item"] != "") & (work["vat"] != "")]
    if "uom" not in work.columns:
        work["uom"] = ""

    return work[["item", "uom", "vat"]].drop_duplicates()


def apply_vat_to_orders(df: pd.DataFrame, tax_df: pd.DataFrame) -> tuple[pd.DataFrame, list[dict]]:
    """Gan ma VAT cho tung dong PO tu bang thue."""
    work = df.copy()
    tax = _normalize_tax_table(tax_df)
    errors: list[dict] = []

    if tax.empty:
        for idx in work.index:
            errors.append(
                {
                    "row": int(idx) + 2,
                    "field": "vat",
                    "value": "",
                    "message": "Ch\u01b0a c\u00f3 danh m\u1ee5c thu\u1ebf \u2014 vui l\u00f2ng nh\u1eadp b\u1ea3ng thu\u1ebf \u1edf sidebar",
                }
            )
        work["vat"] = ""
        return work, errors

    work["item"] = work["item"].astype(str).str.strip()
    work["uom"] = work.get("uom", pd.Series([""] * len(work))).astype(str).str.strip()

    item_uom_map = {
        (row["item"], row["uom"]): row["vat"]
        for _, row in tax.iterrows()
        if row["uom"]
    }
    item_map = {row["item"]: row["vat"] for _, row in tax.iterrows()}

    vats = []
    for idx, row in work.iterrows():
        item = row["item"]
        uom = row["uom"]
        vat = item_uom_map.get((item, uom)) or item_map.get(item, "")
        if not vat:
            errors.append(
                {
                    "row": int(idx) + 2,
                    "field": "vat",
                    "value": item,
                    "message": f"Kh\u00f4ng t\u00ecm th\u1ea5y VAT cho m\u00e3 h\u00e0ng '{item}'",
                }
            )
        vats.append(vat)

    work["vat"] = vats
    return work, errors
