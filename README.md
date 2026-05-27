# The Supply Chain Race 2026 - CKD Delivery Planning Optimizer

Bản đã sửa lỗi `Worksheet named 'đơn đặt hàng' not found`.

## Điểm sửa trong bản này

- App tự đọc danh sách sheet trong file upload.
- Có phần mapping thủ công trong sidebar:
  - Sheet đơn đặt hàng
  - Sheet kế hoạch sản xuất
  - Sheet bảng giá vận tải
- Core optimizer có fuzzy matching sheet name: không phụ thuộc 100% vào dấu tiếng Việt/chữ hoa-thường/khoảng trắng.
- Nếu chọn sai file hoặc thiếu sheet, app báo lỗi dễ hiểu và hiện danh sách sheet đang có, không chỉ ném traceback pandas.

## Cách chạy local

```bash
pip install -r requirements.txt
streamlit run app.py
```

## File cần upload

1. Worksheet chính `.xlsx`, bắt buộc có 3 bảng/sheet tương ứng:
   - đơn đặt hàng
   - kế hoạch sản xuất
   - bảng giá vận tải
2. File khung giờ đại lý `.xlsx` là optional.

Nếu tên sheet khác, dùng dropdown trong sidebar để map đúng sheet.

## Logic tối ưu

- Dùng toàn bộ nguồn xe trong kế hoạch sản xuất/input, bao gồm tồn tháng trước nếu có.
- FIFO theo model/SKU.
- 4 giờ sau xuất xưởng được tính theo giờ làm việc thực tế.
- Không xuất bãi Chủ nhật nếu bật constraint.
- Chi phí theo chuyến/full slot, kể cả slot rỗng.
- Không ghép đa đại lý mặc định; phí 300,000 VND/drop không phát sinh.


## Update notes
- Export workbook no longer hard-codes the exact sheet name `kế hoạch sản xuất`.
- The selected production-sheet mapping from the sidebar is now used for both reading and writing.
- The app also resolves capitalization, accents, and trailing spaces robustly.


## UI update
- Title changed to: The Supply Chain Race 2026 - CKD Delivery Planning Optimizer.
- Dashboard KPI now includes FIFO score (/5).
