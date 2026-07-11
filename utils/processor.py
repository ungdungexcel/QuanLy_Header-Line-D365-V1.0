# -*- coding: utf-8 -*-
"""Tach don theo PO + ma KH + thue VAT."""

from __future__ import annotations

from datetime import datetime

import pandas as pd

from utils.tax import apply_vat_to_orders


def generate_so_numbers(count: int, prefix: str = "SO", date: datetime | None = None) -> list[str]:
    """Sinh danh sach so SO: SO-YYYYMMDD-001, ..."""
    dt = date or datetime.now()
    date_part = dt.strftime("%Y%m%d")
    return [f"{prefix}-{date_part}-{i:03d}" for i in range(1, count + 1)]


def split_orders_by_po_and_item(
    df: pd.DataFrame,
    so_prefix: str = "SO",
    tax_df: pd.DataFrame | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame, list[dict]]:
    """
    Tach don theo quy tac:
    - Cung PO#, cung ma KH, cung thue VAT -> 1 SO (nhieu mat hang tren cung SO)
    - Cung PO# nhung khac thue VAT -> tach SO rieng

    Tra ve (df_header, df_lines, vat_errors).
    """
    work, vat_errors = apply_vat_to_orders(df, tax_df if tax_df is not None else pd.DataFrame())

    work["po_number"] = work["po_number"].astype(str).str.strip()
    work["item"] = work["item"].astype(str).str.strip()
    work["customer"] = work["customer"].astype(str).str.strip()
    work["vat"] = work["vat"].astype(str).str.strip()

    work["_group_key"] = work["po_number"] + "||" + work["customer"] + "||" + work["vat"]

    unique_groups = (
        work.groupby("_group_key", sort=False)
        .agg(
            po_number=("po_number", "first"),
            customer=("customer", "first"),
            vat=("vat", "first"),
        )
        .reset_index(drop=True)
    )

    unique_groups["_group_key"] = (
        unique_groups["po_number"] + "||" + unique_groups["customer"] + "||" + unique_groups["vat"]
    )
    unique_groups["so_number"] = generate_so_numbers(len(unique_groups), prefix=so_prefix)

    key_to_so = dict(zip(unique_groups["_group_key"], unique_groups["so_number"]))
    work["so_number"] = work["_group_key"].map(key_to_so)
    work = work.drop(columns=["_group_key"])

    po_kh_vat_counts = work.groupby(["po_number", "customer"])["vat"].nunique()
    multi_vat_po_kh = {
        (po, kh) for (po, kh), cnt in po_kh_vat_counts.items() if cnt > 1
    }

    headers = []
    # SO dau tien cua moi PO+KH (ke ca khi tach VAT) — khong highlight;
    # cac SO tach them do khac thue — highlight.
    first_so_of_po_kh: set[tuple] = set()

    for so, grp in work.groupby("so_number", sort=False):
        first = grp.iloc[0]
        po_kh = (first["po_number"], first["customer"])
        is_vat_split = po_kh in multi_vat_po_kh
        is_primary_so = po_kh not in first_so_of_po_kh
        if is_primary_so:
            first_so_of_po_kh.add(po_kh)
        # Chi to mau SO tach them (khong phai SO chinh / cung thue dau tien)
        vat_split_hl = bool(is_vat_split and not is_primary_so)

        memo = ""
        if "notes" in grp.columns:
            memos = grp["notes"].dropna().astype(str).str.strip()
            memos = memos[memos != ""].unique()
            if len(memos):
                memo = "; ".join(memos)

        note_parts = []
        if is_vat_split:
            vat_count = po_kh_vat_counts[po_kh]
            note_parts.append(
                f"CANH BAO: PO {first['po_number']} (KH {first['customer']}) co {vat_count} muc thue VAT "
                f"khac nhau - tach SO rieng theo thue"
            )
        note_parts.append(f"Tach tu PO# {first['po_number']} - VAT {first['vat']}")
        items = grp["item"].unique()
        if len(items) > 1:
            note_parts.append(f"Gop {len(items)} mat hang cung thue ({', '.join(items)})")
        elif grp.shape[0] > 1:
            note_parts.append(f"Gop {grp.shape[0]} dong cung mat hang")

        headers.append(
            {
                "so_number": so,
                "po_number": first["po_number"],
                "customer": first["customer"],
                "vat": first["vat"],
                "line_count": grp.shape[0],
                "item_count": len(items),
                "total_qty": grp["quantity"].astype(float).sum(),
                "memo": memo,
                "notes": " | ".join(note_parts),
                "vat_split": is_vat_split,
                "vat_split_hl": vat_split_hl,
                "order_date": first.get("order_date", ""),
                "delivery_date": first.get("delivery_date", ""),
                "store_name": first.get("store_name", ""),
                "site": str(first["site"]).strip() if "site" in grp.columns and pd.notna(first.get("site")) else "",
                "warehouse": (
                    str(first["warehouse"]).strip()
                    if "warehouse" in grp.columns and pd.notna(first.get("warehouse"))
                    else ""
                ),
            }
        )

    df_header = pd.DataFrame(headers)

    lines = work.copy()
    lines["line_number"] = lines.groupby("so_number").cumcount() + 1

    return df_header, lines, vat_errors


VAT_SPLIT_REPORT_COLUMNS = [
    "po_number",
    "customer",
    "store_name",
    "vat_count",
    "vat",
    "so_number",
    "item_count",
    "line_count",
    "total_qty",
    "notes",
]


def build_vat_split_report(df_header: pd.DataFrame) -> pd.DataFrame:
    """Bao cao rieng cac SO bi tach do cung PO+KH nhung khac VAT."""
    empty = pd.DataFrame(columns=VAT_SPLIT_REPORT_COLUMNS)
    if df_header.empty or "vat" not in df_header.columns:
        return empty

    if "vat_split" in df_header.columns:
        report = df_header[df_header["vat_split"]].copy()
    else:
        multi_vat = (
            df_header.groupby(["po_number", "customer"])["vat"]
            .nunique()
            .reset_index(name="vat_count")
        )
        multi_keys = {
            (row["po_number"], row["customer"])
            for _, row in multi_vat[multi_vat["vat_count"] > 1].iterrows()
        }
        if not multi_keys:
            return empty
        mask = df_header.apply(
            lambda r: (r["po_number"], r["customer"]) in multi_keys, axis=1
        )
        report = df_header.loc[mask].copy()

    if report.empty:
        return empty

    vat_counts = (
        report.groupby(["po_number", "customer"])["vat"]
        .nunique()
        .reset_index(name="vat_count")
    )
    report = report.merge(vat_counts, on=["po_number", "customer"], how="left")

    cols = [c for c in VAT_SPLIT_REPORT_COLUMNS if c in report.columns]
    return report[cols].reset_index(drop=True)
