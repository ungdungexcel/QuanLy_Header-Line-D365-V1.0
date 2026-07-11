"""Cấu hình cột chuẩn cho import đơn hàng và export D365."""

# Cột bắt buộc trong file Excel đầu vào (chấp nhận alias)
REQUIRED_FIELDS = {
    "po_number": [
        "Purchase order number",
        "PO#",
        "PO Number",
        "PO",
        "Số PO",
        "Ma PO",
        "Mã PO",
    ],
    "customer": [
        "Store code",
        "Customer",
        "Khách hàng",
        "Account",
        "Customer Account",
        "Mã KH",
        "Vendor code",
    ],
    "item": [
        "Item number",
        "Item",
        "Item Number",
        "Mã hàng",
        "Ma hang",
        "SKU",
        "Product",
    ],
    "quantity": [
        "Order qty",
        "Quantity",
        "Qty",
        "Số lượng",
        "So luong",
        "SL",
    ],
}

# Cột tùy chọn
OPTIONAL_FIELDS = {
    "product_name": ["Product name", "Tên hàng", "Ten hang", "Product Name"],
    "vendor_code": ["Vendor code", "Mã NCC", "Ma NCC"],
    "vendor_name": ["Vendor name", "Tên NCC", "Ten NCC"],
    "store_name": ["Store name", "Tên cửa hàng", "Ten cua hang"],
    "order_date": ["Order date", "Ngày đặt", "Ngay dat"],
    "unit_price": ["Unit Price", "Price", "Đơn giá", "Don gia", "Giá"],
    "delivery_date": ["Delivery date", "Delivery Date", "Ngày giao", "Ngay giao", "Requested Date"],
    "notes": ["Memo", "Notes", "Ghi chú", "Ghi chu", "Note", "Remark"],
    "site": ["Site", "Site ID", "Site code", "Site Code", "Ma Site", "Mã Site"],
    "warehouse": [
        "Warehouse",
        "Kho",
        "WH",
        "W/H",
        "Warehouse ID",
        "Warehouse Code",
        "Ma kho",
        "Mã kho",
        "Ma Kho",
    ],
    "uom": ["Unit", "UOM", "ĐVT", "DVT"],
    "line_number": ["Line", "Line#", "STT", "Dòng"],
}

# Cot export Header theo mau Header_Template.csv (dung de tai len lai)
D365_HEADER_COLUMNS = [
    "Số SO#",
    "Sales order",
    "Status",
    "Requested ship date",
    "Requested receipt date",
    "VAT Invoice date",
    "Customer account",
    "Invoice account",
    "Delivery address name",
    "Pool",
    "VAT purchaser name",
    "Site",
    "Warehouse",
    "Customer requisition",
    "Customer reference",
    "Mode of delivery",
]

# Cột export D365 Line theo mau Line Template (+ Thuế VAT cột O)
D365_LINE_COLUMNS = [
    "Company",
    "Line status",
    "Site",
    "Warehouse",
    "Sales order",
    "External item number",
    "Item number",
    "Delivery address name",
    "Text",
    "Unit",
    "Quantity",
    "Unit price",
    "Net amount",
    "Requested receipt date",
    "Thuế VAT",
]

# Cot bat buoc tren Line Template (nen vang, chu do)
D365_LINE_REQUIRED_COLUMNS = {
    "Site",
    "Warehouse",
    "Sales order",
    "External item number",
    "Item number",
    "Quantity",
    "Unit price",
    "Requested receipt date",
}

DEFAULT_CURRENCY = "VND"

# Dinh dang ngay khi xuat Excel: dd/mm/yyyy
EXPORT_DATETIME_FORMAT = "%d/%m/%Y"
