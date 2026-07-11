# -*- coding: utf-8 -*-
"""Import Header from D365 and merge SO numbers with Line."""

from __future__ import annotations

import io

import pandas as pd

HEADER_IMPORT_ALIASES = {
    "temp_so": ["S\u1ed1 SO#", "So SO#", "SO#", "Temp SO"],
    "d365_so": ["Sales order", "Sales Order", "D365 SO"],
    "customer": ["Customer account", "Customer Account", "Store code", "Ma KH", "M\u00e3 KH"],
    "po_number": ["Customer requisition", "Purchase order number", "PO#", "Customer reference"],
}

_COL_SO_TEMP = "SO t\u1ea1m"
_COL_MA_KH = "M\u00e3 KH"


def _find_column(df: pd.DataFrame, aliases: list[str]) -> str | None:
    cols_lower = {str(c).strip().lower(): c for c in df.columns}
    for alias in aliases:
        key = alias.strip().lower()
        if key in cols_lower:
            return cols_lower[key]
    return None


def _normalize_customer(value) -> str:
    text = str(value).strip()
    if text.isdigit():
        return text.lstrip("0") or "0"
    return text


def load_d365_header_file(uploaded_file) -> pd.DataFrame:
    """Read Header file returned from D365 (.csv / .xlsx)."""
    name = uploaded_file.name.lower()
    data = uploaded_file.getvalue()
    if name.endswith(".csv"):
        for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
            try:
                return pd.read_csv(io.BytesIO(data), encoding=encoding, dtype=str).fillna("")
            except UnicodeDecodeError:
                continue
        return pd.read_csv(io.BytesIO(data), encoding="latin-1", dtype=str).fillna("")

    xl = pd.ExcelFile(io.BytesIO(data), engine="openpyxl")
    sheet = "Header" if "Header" in xl.sheet_names else xl.sheet_names[0]
    return pd.read_excel(xl, sheet_name=sheet, dtype=str).fillna("").astype(str)


def parse_d365_header_import(df: pd.DataFrame) -> tuple[pd.DataFrame, list[dict]]:
    """Normalize Header file imported from D365."""
    errors: list[dict] = []
    col_map = {key: _find_column(df, aliases) for key, aliases in HEADER_IMPORT_ALIASES.items()}

    if not col_map["d365_so"]:
        errors.append(
            {
                "row": "-",
                "field": "d365_so",
                "message": "Thi\u1ebfu c\u1ed9t 'Sales order' (s\u1ed1 SO t\u1eeb D365)",
            }
        )
        return pd.DataFrame(), errors

    work = pd.DataFrame()
    work["temp_so"] = df[col_map["temp_so"]].astype(str).str.strip() if col_map["temp_so"] else ""
    work["d365_so"] = df[col_map["d365_so"]].astype(str).str.strip()
    work["customer"] = df[col_map["customer"]].astype(str).str.strip() if col_map["customer"] else ""
    work["po_number"] = df[col_map["po_number"]].astype(str).str.strip() if col_map["po_number"] else ""

    work = work[work["d365_so"] != ""].reset_index(drop=True)
    skip_words = ("trong", "empty", "de trong", "\u0111\u1ec3 tr\u1ed1ng")
    skip_mask = work["d365_so"].str.lower().apply(lambda v: any(w in str(v) for w in skip_words))
    work = work[~skip_mask].reset_index(drop=True)
    if work.empty:
        errors.append(
            {
                "row": "-",
                "field": "d365_so",
                "message": "Kh\u00f4ng c\u00f3 d\u00f2ng n\u00e0o c\u00f3 s\u1ed1 SO t\u1eeb D365 (c\u1ed9t Sales order)",
            }
        )
    return work, errors


def build_so_mapping(import_df: pd.DataFrame, df_header: pd.DataFrame | None = None) -> tuple[dict[str, str], list[dict]]:
    """Build map temp SO -> D365 Sales order."""
    mapping: dict[str, str] = {}
    warnings: list[dict] = []

    for _, row in import_df.iterrows():
        d365_so = str(row.get("d365_so", "")).strip()
        temp_so = str(row.get("temp_so", "")).strip()
        if d365_so and temp_so and temp_so not in mapping:
            mapping[temp_so] = d365_so

    if df_header is not None and not df_header.empty:
        header_lookup = {
            (_normalize_customer(r["customer"]), str(r["po_number"]).strip()): str(r["so_number"]).strip()
            for _, r in df_header.iterrows()
        }
        for _, row in import_df.iterrows():
            d365_so = str(row.get("d365_so", "")).strip()
            po = str(row.get("po_number", "")).strip()
            customer = _normalize_customer(row.get("customer", ""))
            temp_so = header_lookup.get((customer, po), "")
            if temp_so and temp_so not in mapping and d365_so:
                mapping[temp_so] = d365_so

    if df_header is not None:
        missing = [
            str(so)
            for so in df_header["so_number"].astype(str).str.strip().unique()
            if so and so not in mapping
        ]
        for so in missing:
            warnings.append(
                {
                    "row": "-",
                    "field": "so_number",
                    "message": f"Kh\u00f4ng map \u0111\u01b0\u1ee3c {_COL_SO_TEMP} '{so}' sang SO D365",
                    "level": "warning",
                }
            )

    return mapping, warnings


def apply_so_mapping_to_lines(df_lines: pd.DataFrame, mapping: dict[str, str]) -> tuple[pd.DataFrame, list[dict]]:
    """Apply D365 SO numbers to line rows."""
    errors: list[dict] = []
    work = df_lines.copy()
    work["temp_so_number"] = work["so_number"].astype(str).str.strip()

    d365_sos = []
    for idx, row in work.iterrows():
        temp_so = row["temp_so_number"]
        d365_so = mapping.get(temp_so, "")
        if not d365_so:
            errors.append(
                {
                    "row": int(idx) + 2,
                    "field": "so_number",
                    "message": f"Kh\u00f4ng t\u00ecm th\u1ea5y SO D365 cho {_COL_SO_TEMP} '{temp_so}'",
                }
            )
        d365_sos.append(d365_so)

    work["d365_so_number"] = d365_sos
    work["so_number"] = work["d365_so_number"]
    return work, errors


def merge_d365_header_with_lines(
    df_lines: pd.DataFrame,
    import_df: pd.DataFrame,
    df_header: pd.DataFrame | None = None,
) -> tuple[pd.DataFrame, dict[str, str], list[dict], list[dict]]:
    """Import D365 Header -> map SO -> return merged lines."""
    mapping, map_warnings = build_so_mapping(import_df, df_header)
    merged_lines, line_errors = apply_so_mapping_to_lines(df_lines, mapping)
    return merged_lines, mapping, line_errors, map_warnings


def build_mapping_preview(mapping: dict[str, str], df_header: pd.DataFrame | None = None) -> pd.DataFrame:
    rows = [{_COL_SO_TEMP: k, "SO D365": v} for k, v in mapping.items()]
    if not rows:
        return pd.DataFrame(columns=[_COL_SO_TEMP, "SO D365", "PO#", _COL_MA_KH])
    preview = pd.DataFrame(rows)
    if df_header is not None and not preview.empty:
        meta = df_header.set_index("so_number")[["po_number", "customer"]].to_dict("index")
        preview["PO#"] = preview[_COL_SO_TEMP].map(lambda s: meta.get(s, {}).get("po_number", ""))
        preview[_COL_MA_KH] = preview[_COL_SO_TEMP].map(lambda s: meta.get(s, {}).get("customer", ""))
        preview = preview[[_COL_SO_TEMP, "SO D365", "PO#", _COL_MA_KH]]
    return preview
