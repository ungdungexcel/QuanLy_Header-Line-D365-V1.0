"""Kiểm tra và chuẩn hóa dữ liệu đơn hàng."""

from __future__ import annotations

import pandas as pd

from config import OPTIONAL_FIELDS, REQUIRED_FIELDS


def _find_column(df: pd.DataFrame, aliases: list[str]) -> str | None:
    cols_lower = {c.strip().lower(): c for c in df.columns}
    for alias in aliases:
        key = alias.strip().lower()
        if key in cols_lower:
            return cols_lower[key]
    return None


def normalize_columns(df: pd.DataFrame) -> tuple[pd.DataFrame, dict[str, str]]:
    """Map cột Excel sang tên chuẩn nội bộ."""
    mapping: dict[str, str] = {}

    for std_name, aliases in {**REQUIRED_FIELDS, **OPTIONAL_FIELDS}.items():
        found = _find_column(df, aliases)
        if found:
            mapping[std_name] = found

    normalized = df.rename(columns={v: k for k, v in mapping.items()})
    return normalized, mapping


def validate_orders(df: pd.DataFrame) -> tuple[pd.DataFrame, list[dict]]:
    """
    Kiểm tra lỗi dữ liệu.
    Trả về (df đã chuẩn hóa, danh sách lỗi).
    """
    errors: list[dict] = []
    normalized, col_map = normalize_columns(df.copy())

    # Kiểm tra cột bắt buộc
    missing_cols = [k for k in REQUIRED_FIELDS if k not in normalized.columns]
    if missing_cols:
        for col in missing_cols:
            aliases = ", ".join(REQUIRED_FIELDS[col])
            errors.append(
                {
                    "row": "-",
                    "field": col,
                    "value": "",
                    "message": f"Thiếu cột bắt buộc. Cần một trong: {aliases}",
                }
            )
        return normalized, errors

    # Ghi nhận mapping cột
    for std, orig in col_map.items():
        if std in REQUIRED_FIELDS:
            continue

    # Kiểm tra từng dòng
    for idx, row in normalized.iterrows():
        row_num = int(idx) + 2  # Excel row (header = 1)

        po = row.get("po_number")
        if pd.isna(po) or str(po).strip() == "":
            errors.append(
                {"row": row_num, "field": "po_number", "value": po, "message": "PO# không được để trống"}
            )

        customer = row.get("customer")
        if pd.isna(customer) or str(customer).strip() == "":
            errors.append(
                {"row": row_num, "field": "customer", "value": customer, "message": "Khách hàng không được để trống"}
            )

        item = row.get("item")
        if pd.isna(item) or str(item).strip() == "":
            errors.append(
                {"row": row_num, "field": "item", "value": item, "message": "Mã hàng không được để trống"}
            )

        qty = row.get("quantity")
        if pd.isna(qty):
            errors.append(
                {"row": row_num, "field": "quantity", "value": qty, "message": "Số lượng không được để trống"}
            )
        else:
            try:
                qty_val = float(qty)
                if qty_val <= 0:
                    errors.append(
                        {
                            "row": row_num,
                            "field": "quantity",
                            "value": qty,
                            "message": "Số lượng phải lớn hơn 0",
                        }
                    )
            except (TypeError, ValueError):
                errors.append(
                    {
                        "row": row_num,
                        "field": "quantity",
                        "value": qty,
                        "message": "Số lượng không hợp lệ (phải là số)",
                    }
                )

        if "unit_price" in normalized.columns and not pd.isna(row.get("unit_price")):
            try:
                float(row["unit_price"])
            except (TypeError, ValueError):
                errors.append(
                    {
                        "row": row_num,
                        "field": "unit_price",
                        "value": row["unit_price"],
                        "message": "Đơn giá không hợp lệ",
                    }
                )

    return normalized, errors
