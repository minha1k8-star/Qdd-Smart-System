# CLAUDE.md

Hướng dẫn dành cho Claude Code khi làm việc trong repo này.

## Đọc trước khi làm bất kỳ việc gì

1. [AGENTS.md](AGENTS.md) — quy tắc làm việc bắt buộc (khi nào phải cập nhật Test/Changelog/Docs).
2. [docs/AI_CONTEXT.md](docs/AI_CONTEXT.md) — bối cảnh đầy đủ: dự án là gì, vì sao tồn tại, lịch sử quyết định thiết kế.
3. [docs/03_Business_Rules.md](docs/03_Business_Rules.md) — quy tắc nghiệp vụ cốt lõi (R01-R14). **Không được thay đổi nếu chưa có Test Case tương ứng.**

## Tóm tắt dự án

QDD Smart System tính công suất điều độ (**Qdd**) và công suất dư (**Qdu**) cho tổ máy nhiệt điện, từ danh sách lệnh điều độ và 2 file CSV công tơ (6001, 6303). Hiện trạng: công cụ **Excel + VBA** (bản chính thức v1.3.1, single-module, xem `legacy/`). Định hướng đã xác nhận (chưa triển khai code): **Google Sheets + Apps Script gắn liền từng file + Thư viện code dùng chung** (không làm Web App trung tâm) — dự án hướng tới nhân rộng cho nhiều nhà máy như một sáng kiến kỹ thuật. Xem [docs/05_System_Architecture.md](docs/05_System_Architecture.md) và [ROADMAP.md](ROADMAP.md).

Thuật toán đã được kiểm chứng bằng dữ liệu thực tế vận hành. Đây không phải phần mềm ERP — phạm vi hẹp, độ chính xác số liệu quan trọng hơn tính năng.

## Gợi ý Skills khi làm việc với repo này

- **xlsx** — dùng khi đọc/sửa `legacy/*.xlsm`, `legacy/*.xlsx` (sheet `QUY_TAC_NGHIEP_VU`, `KIEM_THU_UAT`, `CAI_DAT`, `TINH_TOAN`... là nguồn sự thật cho business rule, không phải chỉ có trong docs Markdown). Không dùng skill này để sửa code VBA — file `.bas` là text thuần, sửa bằng Read/Edit như code bình thường.
- **docx / pdf** — chỉ dùng nếu được yêu cầu xuất tài liệu bàn giao dạng Word/PDF cho khách hàng; mặc định tài liệu kỹ thuật nằm ở Markdown trong `docs/`.
- Không có skill VBA chuyên biệt — coi `.bas` như mã nguồn text, đối chiếu với sheet `KIEM_THU_UAT` (danh sách UAT) trước khi sửa logic tính toán.

## Quy ước ngôn ngữ

Toàn bộ thuật ngữ nghiệp vụ giữ nguyên tiếng Việt và viết tắt gốc trong code/sheet: `Qdd`, `Qdu` (Qdư), `Qdc`, `Qmp`, `Qdd_V`, `BĐTH`, `SO`, `MO`, `AUTO_CARRY`, `S1`/`S2`. Không dịch sang tiếng Anh trong tài liệu kỹ thuật vì phải khớp với tên cột/sheet thật trong workbook.
