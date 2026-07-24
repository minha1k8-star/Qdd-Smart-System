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

## Sprint 1 — QDD-Core-Library + Google Sheets Template

⬜ Chưa bắt đầu — chờ hoàn tất Sprint 0 (`04_Algorithm_Specification.md`, chạy UAT baseline trên v1.3.1). Hướng đã xác nhận: Apps Script Library dùng chung + Google Sheets template/nhà máy, không làm Web App (xem [05_System_Architecture.md](05_System_Architecture.md)).

## Sprint 2 — Algorithm (migrate)

⬜ Chưa bắt đầu.

## Sprint 3 — Report

⬜ Chưa bắt đầu.

## Sprint 4 — Dashboard

⬜ Chưa bắt đầu.

## Rủi ro / việc cần quyết định

- Bộ UAT có khoảng trống mã số UAT-20 đến UAT-23 (xem [09_Test_Cases.md](09_Test_Cases.md)) — cần xác nhận với người phụ trách nghiệp vụ.
- ~~Chưa có quyết định chính thức về việc có tiếp tục migrate sang Apps Script hay không~~ — **đã xác nhận 2026-07-24**: hướng Google Sheets + Apps Script Library (không Web App trung tâm), xem [05_System_Architecture.md](05_System_Architecture.md). Việc còn lại: bắt đầu Giai đoạn 2 sau khi hoàn tất `04_Algorithm_Specification.md` và chạy xong 31 UAT baseline.
- Lịch sử phiên bản trước v1.2.6 chưa đầy đủ (xem [CHANGELOG.md](../CHANGELOG.md)).
