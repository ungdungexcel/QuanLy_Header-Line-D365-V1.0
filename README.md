# Tách đơn hàng Excel → D365

Ứng dụng **Streamlit** hỗ trợ tách đơn hàng từ file Excel/CSV, kiểm tra dữ liệu, tạo Sales Order tạm, xuất Header/Line theo mẫu Dynamics 365 và ghép lại số Sales order sau khi import từ D365.

## Quy trình sử dụng

```
Upload PO → Kiểm tra → Tách SO → Xuất Header → Import Header D365 → Ghép Line → Tải file tổng hợp
```

1. **Chuẩn bị** — tải mẫu PO và kiểm tra bảng thuế VAT (sidebar).
2. **Upload PO** — kéo thả file Excel/CSV vào ứng dụng.
3. **Kiểm tra** — xem lỗi thiếu cột, thiếu VAT, dữ liệu không hợp lệ.
4. **Tách SO** — hệ thống tách theo PO + Mã khách hàng + VAT.
5. **Xuất D365** — xuất Header → import vào D365 lấy Sales order → upload lại Header → tải file đã ghép (Header + Line + Báo cáo VAT).

## Tính năng chính

- Đọc file PO (Excel/CSV), nhận diện cột linh hoạt (alias tiếng Việt / tiếng Anh)
- Kiểm tra dữ liệu bắt buộc và bảng thuế VAT
- Tách Sales Order theo PO + Khách hàng + VAT
- Xuất Header / Line theo mẫu D365
- Upload Header từ D365 để map Sales order và ghép Line
- Xuất file Excel tổng hợp: Header + Line + Báo cáo VAT
- Giao diện tab: Dữ liệu · Kiểm tra · Tách SO · Xuất D365

## Yêu cầu hệ thống

- Python 3.10+ (khuyến nghị)
- Các thư viện trong `requirements.txt`

## Cài đặt và chạy

```bash
# 1. Vào thư mục dự án
cd "QuanLy_TaoDonHang (1)"

# 2. Cài thư viện
pip install -r requirements.txt

# 3. Chạy ứng dụng
streamlit run app.py
```

Mở trình duyệt tại: [http://localhost:8501](http://localhost:8501)

## Cột bắt buộc trong file PO

| Trường | Ví dụ tên cột chấp nhận |
|--------|-------------------------|
| Số PO | `PO#`, `Purchase order number`, `Số PO` |
| Mã khách hàng | `Store code`, `Customer`, `Mã KH` |
| Mã hàng | `Item number`, `Item`, `Mã hàng`, `SKU` |
| Số lượng | `Order qty`, `Quantity`, `Qty`, `Số lượng` |

Các cột tùy chọn: tên hàng, NCC, ngày đặt/giao, đơn giá, đơn vị, Site, Warehouse, ghi chú…

## Cấu trúc thư mục

```
QuanLy_TaoDonHang (1)/
├── app.py                 # Giao diện Streamlit
├── config.py              # Cấu hình cột chuẩn
├── requirements.txt       # Thư viện Python
├── README.md
├── templates/             # File mẫu
│   ├── PO Import_Template.csv
│   ├── Header_Template.csv
│   ├── Line_Template.csv
│   └── VAT_Template.csv
└── utils/
    ├── loader.py          # Đọc file PO / mẫu
    ├── validator.py       # Kiểm tra & chuẩn hóa cột
    ├── processor.py       # Tách SO, báo cáo VAT
    ├── tax.py             # Xử lý thuế VAT
    ├── exporter.py        # Xuất Header / Line D365
    └── d365_merge.py      # Ghép Sales order từ Header D365
```

## Lưu ý

- Ứng dụng xử lý **offline** trên máy local — **không gọi API** Dynamics 365.
- Không chứa API key / mật khẩu trong mã nguồn.
- File PO và Header D365 chỉ được xử lý trên máy chạy Streamlit, không gửi lên server bên ngoài.

## Giấy phép

Dùng nội bộ / theo quy định của đơn vị triển khai.
