# AI CONTEXT

Đây là tài liệu đầu tiên mà một AI agent (Claude Code, Codex, ChatGPT...) nên đọc khi bắt đầu làm việc trên dự án **QDD Smart System (QSS)**.

## Đây là gì

Một hệ thống tính toán **Qdd** (công suất điều độ — dispatch power) và **Qdu/Qdư** (công suất dư — surplus power) cho tổ máy nhiệt điện, dựa trên:

- Danh sách lệnh điều độ (thời điểm ra lệnh, công suất ra lệnh, công suất hoàn thành...).
- Hai file CSV từ công tơ đo đếm: **6001** (cho Qdc) và **6303** (cho Qmp).

Kết quả là bảng 48 chu kỳ/ngày (mỗi chu kỳ 30 phút) so sánh công suất theo lệnh điều độ với công suất thực đo, dùng để xác định phần công suất dư ngoài dải dung sai ±3%.

**Đây KHÔNG phải phần mềm ERP.** Phạm vi rất hẹp và có chủ đích: một công cụ tính toán chuyên biệt cho một quy trình nghiệp vụ nhà máy điện cụ thể. Độ chính xác con số quan trọng hơn tính năng hay giao diện.

## Hiện trạng thật (không phải kế hoạch)

- Hệ thống đang **chạy chính thức** là một file Excel (`.xlsm`) với một module VBA duy nhất (~4000 dòng), phiên bản **v1.3.1**. Xem `legacy/`.
- Kiến trúc: **công thức Excel kiểm toán được** (người dùng nhìn thấy và verify từng phép tính) + **VBA chỉ điều phối nhập liệu, tạo nút bấm, và kiểm tra cấu trúc** — VBA không "giấu" logic tính toán quan trọng trong code, phần lớn nằm ở sheet `TINH_TOAN`, `DOAN_CONG_SUAT`, `DIEN_TICH`.
- Bản thân workbook đã tự chứa tài liệu nghiệp vụ ở các sheet `QUY_TAC_NGHIEP_VU` (14 quy tắc R01-R14), `KIEM_THU_UAT` (31 test case UAT), `THONG_TIN_HE_THONG`, `HUONG_DAN`. Các sheet này là **nguồn sự thật gốc** — tài liệu Markdown trong `docs/` được viết dựa trên các sheet này, không phải ngược lại.
- Repo GitHub này (`Qdd-Smart-System`) được tạo sau, với mục tiêu đưa toàn bộ tri thức đã tích luỹ (qua nhiều tháng làm việc với ChatGPT) thành tài liệu chuẩn, độc lập với "trí nhớ của AI" hay lịch sử chat.

## Vì sao dự án tồn tại (bối cảnh lịch sử)

Người dùng ban đầu tính Qdd thủ công trên Excel, sau đó thử vài hướng trước khi chốt phương án hiện tại:

1. **Ý tưởng đầu**: upload CSV lên website để tự tính Qdd — bị huỷ vì sếp đổi hướng.
2. **Hướng 2**: một file Excel tổng hợp nhiều sheet (CSV 6001, CSV 6303, danh sách lệnh) tự tính ra kết quả — đây là hướng được chọn và phát triển dần qua nhiều phiên bản (đến v1.3.1).
3. **Định hướng đã xác nhận, chưa triển khai code**: chuyển sang **Google Sheets + Apps Script gắn liền từng file + Thư viện code dùng chung (Apps Script Library)**, **không làm Web App trung tâm**. Quyết định này xác nhận ngày 2026-07-24, xuất phát từ mục tiêu dự án được đề xuất công nhận **sáng kiến kỹ thuật** và nhân rộng cho nhiều nhà máy khác — cần vừa nhân rộng dễ (sửa 1 chỗ, dùng nhiều nơi) vừa giữ được khả năng audit công thức bằng mắt (lý do dự án từng từ chối hướng website thuần tuý ở vòng trao đổi đầu tiên). Xem lý do đầy đủ và các phương án đã cân nhắc ở `docs/05_System_Architecture.md`.

## Các lỗi/quyết định kỹ thuật quan trọng đã gặp (rút gọn — chi tiết ở docs/14_Knowledge_Transfer.md)

- Đã từng có bug lớn liên quan đến CSV: Excel trên **Mac** dán CSV vào một cột duy nhất thay vì tách sẵn 50 cột như trên Windows → công cụ phải tự phát hiện và tách.
- Lỗi lệch lớn giờ 19h30-20h00 ngày 17 khi kiểm thử — nguyên nhân liên quan đến cách chọn cột thời điểm bắt đầu ramp (phải dùng cột BĐTH, không dùng thời điểm hoàn thành).
- Lỗi báo cáo tháng hiểu sai định dạng tháng khi sang ngày mới → gây nhảy sai tháng trong `LICH_SU_THANG`.
- Lỗi compile v1.3.0 → v1.3.1: 13 hằng số `Private Const M_...` của phần báo cáo tháng bị khai báo sau `End Sub` — VBA bắt buộc mọi khai báo cấp module phải đứng trước thủ tục đầu tiên. Đã sửa bằng cách chuyển toàn bộ lên đầu module, trước `Auto_Open`.
- Từng có lỗi xuất báo cáo tháng trên Mac do ô gộp (merged cell) — đã vá ở v1.2.6.

## Quy tắc khi làm việc (bắt buộc đọc thêm)

Xem [AGENTS.md](../AGENTS.md) ở gốc repo — đặc biệt: **không thay đổi Business Rule hoặc Algorithm nếu chưa có Test Case tương ứng**.
