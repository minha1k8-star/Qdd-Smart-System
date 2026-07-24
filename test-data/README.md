# Test Data

Nơi chứa bộ dữ liệu thực tế (danh sách lệnh điều độ + CSV công tơ 6001/6303) dùng để:

1. **Kiểm tra độ chính xác** của bản Excel/VBA v1.3.1 hiện tại trước khi tách bản chính thức khỏi công cụ kiểm thử.
2. Làm **baseline đối chiếu 1:1** khi sau này chuyển thuật toán sang `QDD-Core-Library` (Apps Script) — xem [ROADMAP.md](../ROADMAP.md) Giai đoạn 3.
3. Chạy lại 31 test case UAT ở [docs/09_Test_Cases.md](../docs/09_Test_Cases.md) với dữ liệu thật thay vì dữ liệu mẫu.

## Cấu trúc gợi ý

```
test-data/
└── YYYY-MM-DD_<mô tả ngắn>/
    ├── danh_sach_lenh.xlsx (hoặc .csv)
    ├── CSV_6001.csv
    ├── CSV_6303.csv
    ├── ket_qua_ky_vong.xlsx   # nếu có kết quả đã tính tay/đã xác nhận để đối chiếu
    └── ghi_chu.md             # tình huống đặc biệt cần lưu ý (ví dụ: ramp qua 00:00, lệnh bị dừng sớm...)
```

Mỗi bộ dữ liệu nên ghi rõ: ngày tính, tổ máy (S1/S2), kết quả Qdd/Qdư đã được xác nhận đúng theo cách tính thủ công/kết quả gốc (nếu có), và tình huống nghiệp vụ đặc biệt mà bộ dữ liệu đó nhằm kiểm tra (map với mã UAT tương ứng nếu phù hợp).

## Lưu ý về quyền riêng tư

Repo này là **public**. Dữ liệu vận hành thực tế (lệnh điều độ, CSV công tơ) có thể nhạy cảm về hạ tầng điện, nên **toàn bộ nội dung trong `test-data/` (trừ file README.md này) đã được thêm vào `.gitignore`** — chỉ lưu cục bộ trên máy, không bao giờ được đưa lên GitHub. Không xoá dòng `test-data/*` trong `.gitignore` ở gốc repo nếu chưa có quyết định khác.
