# Reference Engine

Bản tái hiện độc lập bằng Python của pipeline tính Qdd/Qdư, viết trực tiếp từ công thức Excel thật đã trích xuất trong [docs/04_Algorithm_Specification.md](../../docs/04_Algorithm_Specification.md) (không phải từ mô tả nghiệp vụ). Dùng để:

- Kiểm tra tài liệu Algorithm Specification có đúng với công thức gốc không.
- Đối chiếu với dữ liệu vận hành thật + kết quả tính tay, xem [docs/15_Accuracy_Validation_2026-07.md](../../docs/15_Accuracy_Validation_2026-07.md).
- Làm baseline tham khảo khi viết `QDD-Core-Library` (Apps Script) ở Giai đoạn 2/3 — xem [docs/05_System_Architecture.md](../../docs/05_System_Architecture.md).

## File

- `load_xlsx.py` — vá lỗi đọc `.xlsx` do giá trị `family` font vượt chuẩn (gặp ở một số file xuất từ hệ thống điều độ), rồi mới load bằng `openpyxl`.
- `qdd_engine.py` — cài đặt Ramp Engine, dựng đoạn công suất, tích phân hình thang, chọn công suất hiệu lực (R01-R03).
- `run_validation.py` — chạy đối chiếu trên dữ liệu trong `test-data/` (không có sẵn trong repo public, xem `test-data/README.md`) với `test-data/Kiểm tra Qdu ngày 200726.xlsx`.

## Chạy

Cần dữ liệu thật trong `test-data/` (chỉ có trên máy đã đặt dữ liệu vào, xem `test-data/README.md`):

```bash
python3 run_validation.py
```

## Giới hạn hiện tại

- P0 (công suất đầu ngày) được suy ra từ giá trị tham chiếu chu kỳ 1, chỉ an toàn khi lệnh đầu tiên trong ngày bắt đầu sau 00:30. Chưa xử lý carry-over qua nửa đêm (R07) — cần bổ sung khi có dữ liệu phù hợp (UAT-04).
- Chỉ đọc CSV chuẩn (1 dòng/bản ghi) — đủ dùng cho mục đích đối chiếu.
