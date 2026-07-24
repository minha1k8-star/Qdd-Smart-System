# Project Status

Cập nhật lần cuối: 2026-07-24.

## Sprint 0 — Documentation Foundation

| Hạng mục | Trạng thái |
|---|---|
| Repository & cấu trúc thư mục | ✅ |
| `CLAUDE.md` / `AGENTS.md` / `docs/AI_CONTEXT.md` | ✅ |
| `README.md` | ✅ |
| `docs/03_Business_Rules.md` | ✅ |
| `docs/06_Database_Design.md` | ✅ |
| `docs/09_Test_Cases.md` | ✅ (nội dung có, nhưng toàn bộ 31 case đang "Chưa chạy") |
| `legacy/` — bản Excel/VBA v1.3.1 | ✅ |
| `docs/00_Project_Overview.md` | ✅ |
| `docs/05_System_Architecture.md` (quyết định kiến trúc Giai đoạn 2) | ✅ |
| `docs/14_Knowledge_Transfer.md` | ✅ |
| `docs/04_Algorithm_Specification.md` | ✅ (trích xuất từ công thức Excel thật, không phải suy diễn) |
| PRD, SRS đầy đủ (dạng chi tiết 30-40 trang như đề xuất ban đầu) | ⬜ chưa bắt đầu — có thể không cần thiết nếu các doc hiện tại đã đủ, cần xác nhận |
| Thực thi 31 UAT trên bản v1.3.1 | ⬜ chưa bắt đầu — ưu tiên UAT-11 (nghi vấn hành vi CSV sai ngày, xem 09_Test_Cases.md) |
| Kiểm tra độ chính xác bằng dữ liệu thực (10 ngày, 07/2026) | ✅ — 18/19 tổ hợp khớp; 1 trường hợp lệch (07/07 S2, sự cố) đã điều tra và xác nhận là hành vi đúng (UAT-32), không phải lỗi. Xem [15_Accuracy_Validation_2026-07.md](15_Accuracy_Validation_2026-07.md) |

## Sprint 1 — QDD-Core-Library + Google Sheets Template

🟢 Đã triển khai thật. `QDD-Core-Library` đã push lên Apps Script (đã login, đã publish version 1). Sheet mẫu đầu tiên (Duyên Hải 1) đã tạo, gắn Library, có menu 6 chức năng (thiết lập, lưu CSV, tính 1 ngày, tính hàng loạt, báo cáo tháng). 29/29 test cục bộ pass. Chưa xong: carry-over R07 (phải nhập tay P0), cảnh báo UAT-34, xuất file Excel/PDF riêng, đủ 15 chức năng như VBA.

## Sprint 2 — Algorithm (migrate)

⬜ Chưa bắt đầu.

## Sprint 3 — Report

⬜ Chưa bắt đầu.

## Sprint 4 — Dashboard

⬜ Chưa bắt đầu.

## Rủi ro / việc cần quyết định

- ~~UAT-32: nghi vấn lệnh dừng máy CS=0 bị loại sai~~ — **đã đóng, xác nhận hành vi gốc đúng nghiệp vụ** (lệnh 0-0 chủ đích không tính; khởi động lại cần CS ra lệnh=CS hoàn thành cùng giá trị). Việc còn lại: cập nhật quy trình nhập liệu vận hành để ghi đủ CS hoàn thành khi khởi động lại. Xem [15_Accuracy_Validation_2026-07.md](15_Accuracy_Validation_2026-07.md).
- Bộ UAT có khoảng trống mã số UAT-20 đến UAT-23 (xem [09_Test_Cases.md](09_Test_Cases.md)) — cần xác nhận với người phụ trách nghiệp vụ.
- ~~Chưa có quyết định chính thức về việc có tiếp tục migrate sang Apps Script hay không~~ — **đã xác nhận 2026-07-24**: hướng Google Sheets + Apps Script Library (không Web App trung tâm), xem [05_System_Architecture.md](05_System_Architecture.md). Việc còn lại: bắt đầu Giai đoạn 2 sau khi hoàn tất `04_Algorithm_Specification.md` và chạy xong 31 UAT baseline.
- Lịch sử phiên bản trước v1.2.6 chưa đầy đủ (xem [CHANGELOG.md](../CHANGELOG.md)).
