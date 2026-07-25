# AGENTS.md

Quy tắc làm việc bắt buộc cho mọi AI agent (Claude Code, Codex, ChatGPT, hoặc con người) khi đóng góp vào repo QDD Smart System.

## Cạm bẫy kỹ thuật đã gặp thật (đọc trước khi sửa code Apps Script)

- **KHÔNG dùng `instanceof` cho dữ liệu truyền qua ranh giới Apps Script Library** (`Date`, `Array`, `Error`...). Mỗi scope có constructor riêng, nên `x instanceof Date` luôn sai khi `x` được tạo ở script gọi. Lỗi này từng khiến toàn bộ lệnh điều độ bị loại âm thầm, Qdd phẳng bằng P0 cả ngày mà không báo lỗi gì. Dùng duck typing (`typeof v.getTime === 'function'`).
- **Không đọc dữ liệu sheet theo VỊ TRÍ cột** — luôn dò theo TÊN tiêu đề. Người dùng có thể dán file gốc với thứ tự cột khác, hoặc sheet còn cấu trúc của bản cũ.
- Sau khi sửa `src/QDD-Core-Library/`, phải **tạo version mới** (`npx clasp version "..."`) và cập nhật số version trong `src/NhaMay-Mau-Template/appsscript.json` — nếu không, Sheet vẫn chạy code cũ.

## Nguyên tắc cốt lõi

1. **Không thay đổi Business Rule** (`docs/03_Business_Rules.md`, R01-R14 và các quy tắc bổ sung) nếu chưa có Test Case tương ứng trong `docs/09_Test_Cases.md` / sheet `KIEM_THU_UAT`.
2. **Không thay đổi thuật toán** (`docs/04_Algorithm_Specification.md`) mà không đối chiếu với dữ liệu thực tế đã xác nhận (xem `legacy/` và `docs/14_Knowledge_Transfer.md`).
3. **Không đổi tên Sheet Excel** hoặc cột dữ liệu nếu chưa cập nhật `docs/06_Database_Design.md` và phần Export liên quan.
4. **Không xoá `LICH_SU_THANG` hoặc các sheet `LS_...`** trong bất kỳ workbook nào — đây là dữ liệu lịch sử tháng, không thể tái tạo.

## Quy trình khi sửa đổi

Mọi thay đổi liên quan đến thuật toán hoặc quy tắc nghiệp vụ phải theo đúng thứ tự:

```
Sửa đổi
  ↓
Cập nhật Test Case (docs/09_Test_Cases.md hoặc sheet KIEM_THU_UAT)
  ↓
Cập nhật docs/03_Business_Rules.md và/hoặc docs/04_Algorithm_Specification.md
  ↓
Cập nhật CHANGELOG.md
  ↓
Cập nhật docs/PROJECT_STATUS.md nếu thay đổi tiến độ/sprint
  ↓
Commit (mô tả rõ lý do thay đổi, không chỉ mô tả thay đổi gì)
```

## Trước khi bắt đầu bất kỳ phiên làm việc nào

Đọc theo thứ tự: `docs/AI_CONTEXT.md` → `AGENTS.md` (file này) → tài liệu cụ thể liên quan đến phần việc (Business Rules / Algorithm / Database / Test Plan).

## Phạm vi hiện tại

Repo hiện đang ở **Giai đoạn 1 — Documentation Foundation**: hệ thống thật đang chạy là file Excel/VBA trong `legacy/`, chưa có mã nguồn ứng dụng mới. Không tự ý bắt đầu viết code Apps Script/Web App khi chưa được yêu cầu — xem `docs/ROADMAP.md` để biết giai đoạn tiếp theo và điều kiện chuyển giai đoạn.

## Ngôn ngữ

Tài liệu và thuật ngữ nghiệp vụ viết bằng tiếng Việt, giữ nguyên các từ viết tắt gốc (Qdd, Qdu, Qdc, Qmp, BĐTH, SO, MO...) vì phải khớp với tên cột/sheet thật trong workbook nghiệp vụ.
