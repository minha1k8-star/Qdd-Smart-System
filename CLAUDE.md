# CLAUDE.md

Hướng dẫn dành cho Claude Code khi làm việc trong repo này.

## Đọc trước khi làm bất kỳ việc gì

1. [AGENTS.md](AGENTS.md) — quy tắc làm việc bắt buộc (khi nào phải cập nhật Test/Changelog/Docs).
2. [docs/AI_CONTEXT.md](docs/AI_CONTEXT.md) — bối cảnh đầy đủ: dự án là gì, vì sao tồn tại, lịch sử quyết định thiết kế.
3. [docs/03_Business_Rules.md](docs/03_Business_Rules.md) — quy tắc nghiệp vụ cốt lõi (R01-R14). **Không được thay đổi nếu chưa có Test Case tương ứng.**

## Tóm tắt dự án

QDD Smart System tính công suất điều độ (**Qdd**) và công suất dư (**Qdu**) cho tổ máy nhiệt điện, từ danh sách lệnh điều độ và 2 file CSV công tơ (6001, 6303).

**Hiện trạng**: bản **Google Sheets + Apps Script** (`src/QDD-Core-Library` + `src/NhaMay-Mau-Template`) **đã đưa vào sử dụng thật**. Kiến trúc: thư viện code dùng chung, không làm Web App trung tâm — xem [docs/05_System_Architecture.md](docs/05_System_Architecture.md).

Độ chính xác được chứng minh bằng đối chiếu với **bảng tính tay** trên dữ liệu thật — xem [docs/15_Accuracy_Validation_2026-07.md](docs/15_Accuracy_Validation_2026-07.md). Đây không phải phần mềm ERP — phạm vi hẹp, độ chính xác số liệu quan trọng hơn tính năng.

## Gợi ý Skills khi làm việc với repo này

- **xlsx** — dùng khi cần đọc file Excel của người dùng để đối chiếu (danh sách lệnh gốc, bảng tính tay `Kiểm tra Qdu`). Không dùng để sửa code.
- **docx / pdf** — chỉ dùng nếu được yêu cầu xuất tài liệu bàn giao dạng Word/PDF cho khách hàng; mặc định tài liệu kỹ thuật nằm ở Markdown trong `docs/`.
- Mã nguồn hệ thống là Apps Script (`.js` + `.html`) trong `src/` — sửa bằng Read/Edit như code bình thường, rồi `clasp push`.

## Quy ước ngôn ngữ

Toàn bộ thuật ngữ nghiệp vụ giữ nguyên tiếng Việt và viết tắt gốc trong code/sheet: `Qdd`, `Qdu` (Qdư), `Qdc`, `Qmp`, `Qdd_V`, `BĐTH`, `SO`, `MO`, `AUTO_CARRY`, `S1`/`S2`. Không dịch sang tiếng Anh trong tài liệu kỹ thuật vì phải khớp với tên cột/sheet thật trong Sheet nghiệp vụ.
