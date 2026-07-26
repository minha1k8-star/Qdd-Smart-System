# AI CONTEXT

Đây là tài liệu đầu tiên mà một AI agent (Claude Code, Codex, ChatGPT...) nên đọc khi bắt đầu làm việc trên dự án **QDD Smart System (QSS)**.

## Đây là gì

Một hệ thống tính toán **Qdd** (công suất điều độ — dispatch power) và **Qdu/Qdư** (công suất dư — surplus power) cho tổ máy nhiệt điện, dựa trên:

- Danh sách lệnh điều độ (thời điểm ra lệnh, công suất ra lệnh, công suất hoàn thành...).
- Hai file CSV từ công tơ đo đếm: **6001** (cho Qdc) và **6303** (cho Qmp).

Kết quả là bảng 48 chu kỳ/ngày (mỗi chu kỳ 30 phút) so sánh công suất theo lệnh điều độ với công suất thực đo, dùng để xác định phần công suất dư ngoài dải dung sai ±3%.

**Đây KHÔNG phải phần mềm ERP.** Phạm vi rất hẹp và có chủ đích: một công cụ tính toán chuyên biệt cho một quy trình nghiệp vụ nhà máy điện cụ thể. Độ chính xác con số quan trọng hơn tính năng hay giao diện.

## Hiện trạng thật (không phải kế hoạch)

- Hệ thống đang được sử dụng là bản **Google Sheets + Apps Script** (`src/QDD-Core-Library` + `src/NhaMay-Mau-Template`), đã đối chiếu với dữ liệu vận hành thật — xem `docs/15_Accuracy_Validation_2026-07.md`.
- **Công thức và phương pháp tính** được đặc tả đầy đủ ở `docs/04_Algorithm_Specification.md` (thuật toán) và `docs/03_Business_Rules.md` (14 quy tắc R01-R14). Đây là hợp đồng kỹ thuật: code phải khớp tài liệu, sửa cái này phải sửa cái kia.
- Bằng chứng độ chính xác **duy nhất** là các đợt đối chiếu với **bảng tính tay độc lập** ghi ở `docs/15_Accuracy_Validation_2026-07.md` (2 đợt, 16 tổ hợp ngày+tổ máy, tổ S1). Tổ S2 **chưa được kiểm chứng**.
- Repo GitHub này (`Qdd-Smart-System`) được tạo sau, với mục tiêu đưa toàn bộ tri thức đã tích luỹ (qua nhiều tháng làm việc với ChatGPT) thành tài liệu chuẩn, độc lập với "trí nhớ của AI" hay lịch sử chat.

## Vì sao dự án tồn tại (bối cảnh lịch sử)

Người dùng ban đầu tính Qdd thủ công trên Excel, sau đó thử vài hướng trước khi chốt phương án hiện tại:

1. **Ý tưởng đầu**: upload CSV lên website để tự tính Qdd — bị huỷ vì sếp đổi hướng.
2. **Hướng đang dùng**: **Google Sheets + Apps Script gắn liền từng file + Thư viện code dùng chung (Apps Script Library)**, **không làm Web App trung tâm**. Quyết định xác nhận ngày 2026-07-24, xuất phát từ mục tiêu dự án được đề xuất công nhận **sáng kiến kỹ thuật** và nhân rộng cho nhiều nhà máy khác — cần vừa nhân rộng dễ (sửa 1 chỗ, dùng nhiều nơi) vừa giữ được khả năng audit công thức bằng mắt (lý do dự án từng từ chối hướng website thuần tuý ở vòng trao đổi đầu tiên). Xem lý do đầy đủ và các phương án đã cân nhắc ở `docs/05_System_Architecture.md`. Đã triển khai và đưa vào sử dụng thật từ 07/2026.

## Các lỗi đã gặp thật (rút gọn — chi tiết ở docs/14_Knowledge_Transfer.md)

Ba lỗi dưới đây thuộc nhóm nguy hiểm nhất: **kết quả sai nhưng không báo lỗi gì**.

- **`instanceof` qua ranh giới Apps Script Library**: Date tạo ở script gọi không thoả `instanceof Date` bên trong thư viện → toàn bộ lệnh bị loại âm thầm, Qdd phẳng bằng P0 cả ngày.
- **Múi giờ khi chuyển file Excel qua Drive**: bản Sheets tạm lấy múi giờ mặc định của tài khoản Google → mọi lệnh lệch **+14 giờ**, nhảy sang ngày khác.
- **P0 lấy nhầm Qdd chu kỳ 48**: P0 phải là công suất **tại đúng 24:00**, không phải giá trị trung bình 23:30-24:00.

Ngoài ra: mốc bắt đầu ramp phải là **BĐTH** (không phải thời điểm hoàn thành); ramp bị cắt giữa chừng phải lấy **giá trị nội suy tại điểm cắt**, không phải công suất mục tiêu.

## Quy tắc khi làm việc (bắt buộc đọc thêm)

Xem [AGENTS.md](../AGENTS.md) ở gốc repo — đặc biệt: **không thay đổi Business Rule hoặc Algorithm nếu chưa có Test Case tương ứng**.
