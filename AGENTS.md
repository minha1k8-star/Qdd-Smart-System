# AGENTS.md

Quy tắc làm việc bắt buộc cho mọi AI agent (Claude Code, Codex, ChatGPT, hoặc con người) khi đóng góp vào repo QDD Smart System.

## Cạm bẫy kỹ thuật đã gặp thật (đọc trước khi sửa code Apps Script)

- **KHÔNG dùng `instanceof` cho dữ liệu truyền qua ranh giới Apps Script Library** (`Date`, `Array`, `Error`...). Mỗi scope có constructor riêng, nên `x instanceof Date` luôn sai khi `x` được tạo ở script gọi. Lỗi này từng khiến toàn bộ lệnh điều độ bị loại âm thầm, Qdd phẳng bằng P0 cả ngày mà không báo lỗi gì. Dùng duck typing (`typeof v.getTime === 'function'`).
- **File Excel chuyển đổi qua Drive phải đặt lại MÚI GIỜ trước khi đọc.** Ngày-giờ trong `.xlsx` là giá trị "trần", không kèm múi giờ. Bản Google Sheets tạm do Drive tạo ra lấy múi giờ mặc định của tài khoản Google (thường `America/Los_Angeles`), nên Apps Script đọc ra lệch **+14 giờ** so với Sheet đích `Asia/Ho_Chi_Minh` — lệnh nhảy sang ngày khác, kết quả sai dây chuyền mà **không có lỗi nào được báo**. Luôn gọi `setSpreadsheetTimeZone` cho file tạm trước khi `getValues()` (xem `alignTimeZoneWithTargetSheet_`).
- **Không đọc dữ liệu sheet theo VỊ TRÍ cột** — luôn dò theo TÊN tiêu đề. Người dùng có thể dán file gốc với thứ tự cột khác, hoặc sheet còn cấu trúc của bản cũ.
- **Sheet chỉ gọi được hàm khai báo ở `Public.js`.** Export trong module nội bộ (`QDD.CommandFilter`...) là chưa đủ — Sheet gọi `QDDCoreLibrary.<tên>` sẽ báo `is not a function`. Test cục bộ KHÔNG tự bắt được vì nó gọi thẳng module nội bộ; đã có test riêng quét mọi lời gọi `QDDCoreLibrary.*` trong Sheet mẫu và đối chiếu với `Public.js`.
- Sau khi sửa `src/QDD-Core-Library/`, phải **tạo version mới** (`npx clasp version "..."`) và cập nhật số version trong `src/NhaMay-Mau-Template/appsscript.json` — nếu không, Sheet vẫn chạy code cũ.

## Nguyên tắc cốt lõi

1. **Không thay đổi Business Rule** (`docs/03_Business_Rules.md`, R01-R14 và các quy tắc bổ sung) nếu chưa có Test Case tương ứng trong `docs/09_Test_Cases.md`.
2. **Không thay đổi thuật toán** (`docs/04_Algorithm_Specification.md`) mà không đối chiếu lại với dữ liệu thực tế đã xác nhận — xem `docs/15_Accuracy_Validation_2026-07.md` và `docs/14_Knowledge_Transfer.md`.
3. **Không đổi tên sheet hoặc cột dữ liệu** nếu chưa cập nhật `docs/06_Database_Design.md` và phần xuất báo cáo liên quan.
4. **Không xoá dữ liệu lịch sử** trong Sheet của nhà máy (`KET_QUA`, `P0_NGAY`) khi chưa chắc chắn — mất chuỗi P0 thì phải nhập tay lại từ đầu.

## Quy trình khi sửa đổi

Mọi thay đổi liên quan đến thuật toán hoặc quy tắc nghiệp vụ phải theo đúng thứ tự:

```
Sửa đổi
  ↓
Cập nhật Test Case (docs/09_Test_Cases.md)
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

Hệ thống đang dùng thật là bản **Google Sheets + Apps Script** trong `src/`. Mã nguồn sửa ở đây phải `clasp push` mới có hiệu lực trên Sheet của nhà máy.

Bằng chứng độ chính xác của hệ thống nằm ở `docs/15_Accuracy_Validation_2026-07.md` — đối chiếu từng chu kỳ với bảng tính tay độc lập trên dữ liệu vận hành thật. Mọi khẳng định về độ tin cậy phải dẫn về đó, không dẫn về cảm tính.

## Ngôn ngữ

Tài liệu và thuật ngữ nghiệp vụ viết bằng tiếng Việt, giữ nguyên các từ viết tắt gốc (Qdd, Qdu, Qdc, Qmp, BĐTH, SO, MO...) vì phải khớp với tên cột/sheet thật trong file nghiệp vụ.
