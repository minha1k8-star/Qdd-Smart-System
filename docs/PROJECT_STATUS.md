# Project Status

Cập nhật lần cuối: 2026-07-26.

> **Trạng thái tổng quát**: hệ thống đã **được đơn vị đưa vào sử dụng thật** trên Sheet của nhà máy Duyên Hải 1. Thư viện `QDD-Core-Library` ở **version 4**, 54/54 test cục bộ pass.

## Đã xong

| Hạng mục | Trạng thái |
|---|---|
| `QDD-Core-Library` — toàn bộ thuật toán R01-R14 + R07 | ✅ version 4, 54/54 test pass |
| Sheet mẫu nhà máy đầu tiên, gắn Library | ✅ đang dùng thật |
| Bảng điều khiển 6 mục | ✅ |
| Nhập danh sách lệnh từ file Excel | ✅ |
| Tải nhiều CSV cùng lúc, tự nhận diện | ✅ |
| Cảnh báo tự động (lệnh 0-0, ngày không lệnh, nối tiếp ramp) | ✅ |
| Xuất báo cáo Excel/PDF + báo cáo tháng | ✅ |
| Kiểm chứng độ chính xác tổ **S1** bằng dữ liệu thật | ✅ 2 đợt, 16 tổ hợp — xem [15_Accuracy_Validation_2026-07.md](15_Accuracy_Validation_2026-07.md) |
| Kiểm chứng **R07** (ramp qua nửa đêm) bằng dữ liệu thật | ✅ ngày 23→24/07/2026 |
| Tài liệu: quy tắc nghiệp vụ, thuật toán, schema, triển khai, hướng dẫn trong file | ✅ |
| Hỗ trợ số tổ máy bất kỳ (thêm tổ = thêm dòng cấu hình) | ✅ |
| Hồ sơ sáng kiến kỹ thuật | ✅ bản thảo — còn các mục [CẦN ĐIỀN] là số liệu của đơn vị |

## Đang nợ

| Hạng mục | Ghi chú |
|---|---|
| **Kiểm chứng tổ S2** (TC-29) | Ưu tiên cao nhất. Toàn bộ đối chiếu đến nay mới chạy S1; mã công tơ, P0 và chuỗi ngày của S2 độc lập nên không suy ra được từ S1 |
| Kiểm thử giữ P0 nhập tay (TC-30) | Code đã xử lý, chưa chạy thử thật trên Sheet |
| Khởi động lại sau sự cố (TC-31) | Chờ dữ liệu vận hành đúng mẫu "CS ra lệnh = CS hoàn thành = tải thật" |

| Triển khai nhà máy thứ hai | Đồng thời là phép thử cho tài liệu [16_Huong_Dan_Trien_Khai.md](16_Huong_Dan_Trien_Khai.md) |

## Rủi ro / việc cần theo dõi

- **Chỉ mới kiểm chứng một tổ máy.** Kết quả S1 khớp không bảo đảm S2 khớp — cấu hình mã công tơ khác, chuỗi P0 khác. Cần đối chiếu vài ngày S2 trước khi tin số liệu S2.
- **Quy trình nhập liệu vận hành** cần ghi đủ CS hoàn thành khi khởi động lại tổ máy sau sự cố. Thiếu trường này thì hệ thống không tự suy ra được tổ máy đã ngừng (xem [15_Accuracy_Validation_2026-07.md](15_Accuracy_Validation_2026-07.md), UAT lệnh 0-0).
- **Bảng tính tay đối chiếu có 3 ô sai** đã phát hiện (ngày 17/07 chu kỳ 40, ngày 19/07 chu kỳ 42). Nên báo lại người phụ trách bảng tính tay để sửa bản gốc, tránh lần sau đối chiếu lại nghi ngờ nhầm hệ thống.
- **Ngày khởi động lại tổ máy sau khi ngừng** phải xử lý tay: hệ thống vẫn tính giai đoạn tăng tải từ 0 theo ramp thường (cao hơn thực tế). Đã quyết định không tự động hoá — xem R15 ở [03_Business_Rules.md](03_Business_Rules.md#r15--khởi-động-lại-tổ-máy-sau-khi-ngừng).
- **Cập nhật thư viện không tự động.** Khi có nhiều nhà máy dùng chung Library, mỗi nhà máy phải tự đổi sang version mới — đúng theo kiến trúc đã chọn, nhưng cần quy trình thông báo.
