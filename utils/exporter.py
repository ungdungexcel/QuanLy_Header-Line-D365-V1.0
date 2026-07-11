# -*- coding: utf-8 -*-
"""Xuat file Excel Head/Line cho D365."""

from __future__ import annotations

import io
import re
from datetime import datetime

import pandas as pd

from config import (
    D365_HEADER_COLUMNS,
    D365_LINE_COLUMNS,
    D365_LINE_REQUIRED_COLUMNS,
    EXPORT_DATETIME_FORMAT,
)

_DATE_COLUMN_PATTERN = re.compile(
    r"date|datetime|ngay|ngày|ship|receipt|invoice|(?:^|[_\s])time|giờ|gio",
    re.IGNORECASE,
)


def _parse_datetime(value) -> datetime | None:
    """Chuyen gia tri ve datetime; ho tro chuoi ngay VN va Excel serial."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, datetime):
        return value
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()

    text = str(value).strip()
    if not text or text.lower() in {"nan", "nat", "none"}:
        return None

    if re.match(r"^\d{4}-\d{2}-\d{2}", text):
        parsed = pd.to_datetime(text, errors="coerce")
    else:
        parsed = pd.to_datetime(text, dayfirst=True, errors="coerce")
    if pd.notna(parsed):
        return parsed.to_pydatetime()

    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y %H:%M:%S", "%d/%m/%Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def format_datetime_value(value, *, use_now_if_empty: bool = False) -> str:
    """Dinh dang ngay: dd/mm/yyyy."""
    dt = _parse_datetime(value)
    if dt is None:
        if use_now_if_empty:
            return datetime.now().strftime(EXPORT_DATETIME_FORMAT)
        if value is None or (isinstance(value, float) and pd.isna(value)):
            return ""
        text = str(value).strip()
        return "" if text.lower() in {"nan", "nat", "none"} else text
    return dt.strftime(EXPORT_DATETIME_FORMAT)


def _is_date_column(col_name: str) -> bool:
    return bool(_DATE_COLUMN_PATTERN.search(str(col_name)))


def format_dataframe_for_export(df: pd.DataFrame) -> pd.DataFrame:
    """Chuan hoa cot ngay truoc khi xuat Excel."""
    if df.empty:
        return df.copy()

    out = df.copy()
    for col in out.columns:
        if _is_date_column(col):
            out[col] = out[col].map(format_datetime_value)
    return out


def _format_customer_account(value) -> str:
    """Giu nguyen Store code, bao gom so 0 dau (vd: 00232)."""
    return str(value).strip()


def _format_vat_purchaser_name(row: pd.Series) -> str:
    """Ten CH + so PO."""
    store_name = str(row.get("store_name", "")).strip()
    po_number = str(row.get("po_number", "")).strip()
    if store_name and po_number:
        return f"{store_name} {po_number}"
    return store_name or po_number


def _format_warning_notes(row: pd.Series) -> str:
    """Ghi chu tach SO, canh bao VAT va Memo tu PO."""
    parts: list[str] = []
    notes = str(row.get("notes", "")).strip()
    memo = str(row.get("memo", "")).strip()
    if notes:
        parts.append(notes)
    if memo and memo not in notes:
        parts.append(memo)
    return " | ".join(parts)


def _pick_header_date(row: pd.Series, *fields: str) -> str:
    """Lay ngay tu cac truong uu tien, dinh dang dd/mm/yyyy."""
    for field in fields:
        if field not in row.index:
            continue
        value = row.get(field)
        if value is None or (isinstance(value, float) and pd.isna(value)):
            continue
        if str(value).strip():
            return format_datetime_value(value)
    return ""


def _series_text(df: pd.DataFrame, col: str) -> pd.Series:
    """Lay cot text; neu thieu cot thi tra series rong."""
    if col not in df.columns:
        return pd.Series([""] * len(df), index=df.index, dtype=str)
    out = df[col].map(
        lambda v: "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v).strip()
    )
    return out.replace({"nan": "", "None": "", "NaT": ""})


def _site_warehouse_for_header(df_header: pd.DataFrame, df_lines: pd.DataFrame) -> tuple[pd.Series, pd.Series]:
    """Lay Site/Warehouse tu PO (Header); neu thieu tung dong thi lay tu Line theo SO."""
    site = _series_text(df_header, "site")
    warehouse = _series_text(df_header, "warehouse")

    need_site = site.eq("").any() and "site" in df_lines.columns
    need_wh = warehouse.eq("").any() and "warehouse" in df_lines.columns
    if not (need_site or need_wh) or "so_number" not in df_lines.columns:
        return site, warehouse

    lines = df_lines.copy()
    # uu tien map theo SO tam neu co (sau khi ghép D365)
    so_key = "temp_so_number" if "temp_so_number" in lines.columns else "so_number"
    if need_site:
        lines["_site"] = _series_text(lines, "site")
    if need_wh:
        lines["_warehouse"] = _series_text(lines, "warehouse")

    def _first_nonempty(series: pd.Series) -> str:
        for value in series:
            text = str(value).strip()
            if text and text.lower() not in {"nan", "none"}:
                return text
        return ""

    grouped = lines.groupby(lines[so_key].astype(str).str.strip(), sort=False)
    header_so = df_header["so_number"].astype(str).str.strip()
    if need_site:
        site_map = grouped["_site"].agg(_first_nonempty)
        from_lines = header_so.map(site_map).fillna("").astype(str)
        site = site.where(site.ne(""), from_lines)
    if need_wh:
        wh_map = grouped["_warehouse"].agg(_first_nonempty)
        from_lines = header_so.map(wh_map).fillna("").astype(str)
        warehouse = warehouse.where(warehouse.ne(""), from_lines)
    return site, warehouse


def fill_missing_site_warehouse(
    target: pd.DataFrame,
    source: pd.DataFrame,
    *,
    key_candidates: tuple[str, ...] = (
        "Sales order",
        "Số SO#",
        "so_number",
        "temp_so_number",
        "Customer requisition",
        "po_number",
    ),
    overwrite: bool = False,
) -> pd.DataFrame:
    """Bo sung Site/Warehouse tu source (PO) vao target.

    overwrite=True: uu tien gia tri tu source (file PO), ghi de len target.
    overwrite=False: chi dien khi target dang trong.
    """
    out = target.copy()
    if out.empty or source.empty:
        return out

    src = source.copy()
    if "Site" not in src.columns and "site" in src.columns:
        src["Site"] = src["site"]
    if "Warehouse" not in src.columns and "warehouse" in src.columns:
        src["Warehouse"] = src["warehouse"]

    if "Site" not in out.columns:
        out["Site"] = ""
    if "Warehouse" not in out.columns:
        out["Warehouse"] = ""

    site_vals = _series_text(src, "Site" if "Site" in src.columns else "site")
    wh_vals = _series_text(src, "Warehouse" if "Warehouse" in src.columns else "warehouse")
    if site_vals.eq("").all() and wh_vals.eq("").all():
        return out

    # Gop map tu moi cot khoa co gia tri (tranh dung cot Số SO# rong)
    site_map: dict[str, str] = {}
    wh_map: dict[str, str] = {}
    for key in key_candidates:
        if key not in src.columns:
            continue
        keys = src[key].astype(str).str.strip().replace({"nan": "", "None": ""})
        for k, s, w in zip(keys, site_vals, wh_vals):
            if not k:
                continue
            if s and k not in site_map:
                site_map[k] = s
            if w and k not in wh_map:
                wh_map[k] = w

    if not site_map and not wh_map:
        return out

    cur_site = out["Site"].astype(str).str.strip().replace({"nan": "", "None": ""})
    cur_wh = out["Warehouse"].astype(str).str.strip().replace({"nan": "", "None": ""})

    mapped_site = pd.Series([""] * len(out), index=out.index, dtype=str)
    mapped_wh = pd.Series([""] * len(out), index=out.index, dtype=str)
    for key in key_candidates:
        if key not in out.columns:
            continue
        keys = out[key].astype(str).str.strip().replace({"nan": "", "None": ""})
        hit_site = keys.map(site_map).fillna("")
        hit_wh = keys.map(wh_map).fillna("")
        mapped_site = mapped_site.where(mapped_site.ne(""), hit_site)
        mapped_wh = mapped_wh.where(mapped_wh.ne(""), hit_wh)

    if overwrite:
        # Uu tien Site/Warehouse tu file PO
        out["Site"] = mapped_site.where(mapped_site.ne(""), cur_site)
        out["Warehouse"] = mapped_wh.where(mapped_wh.ne(""), cur_wh)
    else:
        out["Site"] = cur_site.where(cur_site.ne(""), mapped_site)
        out["Warehouse"] = cur_wh.where(cur_wh.ne(""), mapped_wh)
    return out


def build_d365_files(df_header: pd.DataFrame, df_lines: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Chuyen df noi bo sang format cot Header/Line Template (theo mau xuat D365)."""
    now_text = datetime.now().strftime(EXPORT_DATETIME_FORMAT)
    site_vals, warehouse_vals = _site_warehouse_for_header(df_header, df_lines)

    d365_header = pd.DataFrame(
        {
            "Số SO#": _series_text(df_header, "so_number").to_numpy(),
            "Sales order": "",
            "Status": "",
            "Requested ship date": df_header.apply(
                lambda r: _pick_header_date(r, "order_date", "delivery_date"), axis=1
            ).to_numpy(),
            "Requested receipt date": df_header.apply(
                lambda r: _pick_header_date(r, "delivery_date", "order_date"), axis=1
            ).to_numpy(),
            "VAT Invoice date": now_text,
            "Customer account": df_header["customer"].map(_format_customer_account).to_numpy(),
            "Invoice account": "",
            "Delivery address name": "",
            "Pool": "",
            "VAT purchaser name": df_header.apply(_format_vat_purchaser_name, axis=1).to_numpy(),
            "Site": site_vals.to_numpy(),
            "Warehouse": warehouse_vals.to_numpy(),
            "Customer requisition": _series_text(df_header, "po_number").to_numpy(),
            "Customer reference": df_header.apply(_format_warning_notes, axis=1).to_numpy(),
            "Mode of delivery": "",
        }
    )
    d365_header = d365_header.reindex(columns=D365_HEADER_COLUMNS).fillna("")

    line_site = _series_text(df_lines, "site")
    line_wh = _series_text(df_lines, "warehouse")
    if line_site.eq("").any() or line_wh.eq("").any():
        # Lay theo SO tam neu Line da doi sang SO D365
        if "temp_so_number" in df_lines.columns:
            temp_keys = df_lines["temp_so_number"].astype(str).str.strip()
            site_by_temp = dict(zip(d365_header["Số SO#"].astype(str), d365_header["Site"]))
            wh_by_temp = dict(zip(d365_header["Số SO#"].astype(str), d365_header["Warehouse"]))
            fill_site = temp_keys.map(site_by_temp).fillna("")
            fill_wh = temp_keys.map(wh_by_temp).fillna("")
        else:
            so_keys = df_lines["so_number"].astype(str).str.strip()
            site_by_so = dict(zip(d365_header["Số SO#"].astype(str), d365_header["Site"]))
            wh_by_so = dict(zip(d365_header["Số SO#"].astype(str), d365_header["Warehouse"]))
            fill_site = so_keys.map(site_by_so).fillna("")
            fill_wh = so_keys.map(wh_by_so).fillna("")
        line_site = line_site.where(line_site.ne(""), fill_site)
        line_wh = line_wh.where(line_wh.ne(""), fill_wh)

    unit_vals = (
        _series_text(df_lines, "uom").replace("", "Ea").to_numpy()
        if "uom" in df_lines.columns
        else ["Ea"] * len(df_lines)
    )
    text_vals = (
        _series_text(df_lines, "product_name").to_numpy()
        if "product_name" in df_lines.columns
        else _series_text(df_lines, "notes").to_numpy()
    )

    d365_line = pd.DataFrame(
        {
            "Company": "",
            "Line status": "",
            "Site": line_site.to_numpy(),
            "Warehouse": line_wh.to_numpy(),
            "Sales order": _series_text(df_lines, "so_number").to_numpy(),
            "External item number": "",
            "Item number": _series_text(df_lines, "item").to_numpy(),
            "Delivery address name": "",
            "Text": text_vals,
            "Unit": unit_vals,
            "Quantity": df_lines["quantity"].to_numpy() if "quantity" in df_lines.columns else "",
            "Unit price": (
                _series_text(df_lines, "unit_price").to_numpy()
                if "unit_price" in df_lines.columns
                else ""
            ),
            "Net amount": "",
            "Requested receipt date": df_lines.apply(
                lambda r: _pick_header_date(r, "delivery_date", "order_date"), axis=1
            ).to_numpy(),
        }
    )
    d365_line = d365_line.reindex(columns=D365_LINE_COLUMNS).fillna("")

    return format_dataframe_for_export(d365_header), format_dataframe_for_export(d365_line)


def apply_sales_order_to_header(d365_header: pd.DataFrame, mapping: dict[str, str]) -> pd.DataFrame:
    """Dien cot Sales order tu map SO tam -> SO D365 (giu nguyen format Header mau)."""
    out = d365_header.copy()
    if "Số SO#" not in out.columns:
        return out
    out["Sales order"] = out["Số SO#"].astype(str).str.strip().map(mapping).fillna("")
    return out.reindex(columns=D365_HEADER_COLUMNS).fillna("")


def build_d365_export_filename(kind: str = "Header", suffix: str = "") -> str:
    """Ten file xuat: D365_Header_ddMMyyyy_HHmmss.xlsx"""
    stamp = datetime.now().strftime("%d%m%Y_%H%M%S")
    mid = f"_{suffix}" if suffix else ""
    return f"D365_{kind}{mid}_{stamp}.xlsx"


def _style_line_header(worksheet, columns: list[str]) -> None:
    """Dinh dang dong tieu de Line: bat buoc = vang/do, con lai = xanh/trang."""
    from openpyxl.styles import Alignment, Font, PatternFill

    fill_required = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    fill_optional = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    font_required = Font(color="FF0000", bold=True)
    font_optional = Font(color="FFFFFF", bold=True)
    align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for col_idx, col_name in enumerate(columns, start=1):
        cell = worksheet.cell(row=1, column=col_idx)
        if col_name in D365_LINE_REQUIRED_COLUMNS:
            cell.fill = fill_required
            cell.font = font_required
        else:
            cell.fill = fill_optional
            cell.font = font_optional
        cell.alignment = align


def to_excel_bytes(*sheets: tuple[str, pd.DataFrame]) -> bytes:
    """Ghi nhieu sheet vao bytes Excel — cot ngay da dinh dang dd/mm/yyyy."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for sheet_name, df in sheets:
            export_df = format_dataframe_for_export(df)
            export_df.to_excel(writer, sheet_name=sheet_name, index=False)

            worksheet = writer.sheets[sheet_name]
            if sheet_name == "Line":
                _style_line_header(worksheet, list(export_df.columns))
            for col_idx, col_name in enumerate(export_df.columns, start=1):
                if not _is_date_column(col_name):
                    continue
                for row_idx in range(2, len(export_df) + 2):
                    worksheet.cell(row=row_idx, column=col_idx).number_format = "DD/MM/YYYY"
    return buffer.getvalue()
