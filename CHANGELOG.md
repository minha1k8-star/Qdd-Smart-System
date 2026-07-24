# Changelog

Định dạng: mỗi mục ghi ngày (nếu biết), thay đổi, và lý do khi có thể xác định từ tài liệu gốc.

## [Unreleased] — Documentation Foundation

- Khởi tạo repo GitHub `Qdd-Smart-System`.
- Thêm `CLAUDE.md`, `AGENTS.md`, `docs/AI_CONTEXT.md` — hướng dẫn làm việc cho AI agent.
- Thêm `README.md`, `docs/00_Project_Overview.md`, `docs/03_Business_Rules.md`, `docs/06_Database_Design.md`, `docs/09_Test_Cases.md`, `docs/14_Knowledge_Transfer.md`, `docs/PROJECT_STATUS.md`, `ROADMAP.md` — dựa trên nội dung thật của workbook v1.3.0/v1.3.1 và lịch sử trao đổi phát triển.
- Đưa bản Excel/VBA v1.3.1 (bản chính thức đang chạy) vào `legacy/`.
- Xác nhận hướng kiến trúc Giai đoạn 2: **Google Sheets + Apps Script Library dùng chung**, không làm Web App trung tâm — ghi lại quyết định và lý do ở `docs/05_System_Architecture.md`.
- Thêm `docs/04_Algorithm_Specification.md` — trích xuất trực tiếp từ công thức Excel thật (`XU_LY_LENH`, `DOAN_CONG_SUAT`, `DIEN_TICH`, `TINH_TOAN`) và code VBA, phát hiện quan trọng: toàn bộ Ramp Engine/nội suy/tích phân diện tích nằm ở công thức Excel, không nằm trong VBA; đồng thời phát hiện nghi vấn UAT-11 (CSV sai ngày có thể không còn bị từ chối trong v1.3.1).

## v1.3.1 — Sửa lỗi compile All-in-One

- Nguyên nhân: 13 hằng số `Private Const M_...` của phần báo cáo tháng bị khai báo sau `End Sub`. VBA yêu cầu mọi khai báo cấp module phải nằm trước thủ tục đầu tiên.
- Đã chuyển toàn bộ các hằng số này lên đầu module, trước `Auto_Open`.
- Kết quả kiểm tra cấu trúc module: Sub/End Sub 42/42, Function/End Function 74/74, khai báo cấp module sau thủ tục: 0, tên thủ tục trùng: 0, `Application.Run` thực thi: 0.

## v1.3.0 — Hợp nhất module

- Core tính toán và báo cáo tháng được hợp nhất vào một module VBA duy nhất (`modQdu_v1_3_0_AllInOne`).
- Loại bỏ hoàn toàn việc gọi macro báo cáo tháng bằng `Application.Run` (nguồn gây lỗi "Cannot run the macro" trên Windows ở các bản trước).
- "Chạy chính thức" (nút 7) không còn yêu cầu nhập Qdd gốc; không tự tạo backup tự động (dùng nút 11 để sao lưu thủ công).

## v1.2.7 / v1.2.6 và trước đó

- v1.2.6: sửa lỗi xuất báo cáo tháng trên macOS khi file mẫu chứa ô gộp (merged cell), gây lỗi khi ghi dữ liệu vào vùng bị gộp.
- Trước v1.2.6: từng gặp các lỗi được ghi nhận trong quá trình phát triển — xem [docs/14_Knowledge_Transfer.md](docs/14_Knowledge_Transfer.md) để biết chi tiết bối cảnh và cách khắc phục từng lỗi.

## Ghi chú

Lịch sử phiên bản trước v1.2.6 chưa được ghi lại đầy đủ dưới dạng changelog có ngày tháng cụ thể — chỉ có trong lịch sử trao đổi phát triển (ChatGPT). Nếu tìm thấy các bản `.bas`/`.xlsm` cũ hơn, nên bổ sung lại các mốc còn thiếu vào changelog này.
