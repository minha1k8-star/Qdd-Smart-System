# System Architecture — Định hướng Giai đoạn 2

## Quyết định kiến trúc (xác nhận ngày 2026-07-24)

**Chọn: Google Sheets + Apps Script gắn liền từng file (container-bound) + Thư viện code dùng chung (Apps Script Library).**

**Không chọn: Web App trung tâm** (một địa chỉ URL phục vụ nhiều nhà máy cùng lúc).

### Bối cảnh quyết định

Dự án đang được xây dựng để **đề xuất công nhận sáng kiến kỹ thuật**, với kỳ vọng nếu được công nhận sẽ **áp dụng cho nhiều nhà máy khác** ngoài nhà máy hiện tại. Điều này đặt ra 2 yêu cầu tưởng như mâu thuẫn:

1. Phải **nhân rộng được** dễ dàng, không lặp lại vấn đề cũ (mỗi nơi một bản VBA riêng, sửa lỗi ở đâu chỉ có tác dụng ở đó).
2. Phải **giữ được khả năng kiểm tra công thức bằng mắt** (xem [00_Project_Overview.md](00_Project_Overview.md), [14_Knowledge_Transfer.md](14_Knowledge_Transfer.md)) — đây là lý do dự án từng từ chối hướng "website tự tính" ngay từ vòng trao đổi đầu tiên.

### Các phương án đã cân nhắc

| Phương án | Nhân rộng | Audit-ability | Rủi ro vận hành |
|---|---|---|---|
| A. Web App trung tâm (1 URL, nhiều nhà máy) | Rất tốt — 1 điểm truy cập chung | Kém hơn — chỉ thấy phần chủ động hiển thị trên giao diện web | Cao — một lỗi ảnh hưởng mọi nhà máy cùng lúc; cần người quản trị xuyên suốt; phụ thuộc hạn mức (quota) Apps Script của tài khoản sở hữu |
| B. Web App riêng từng nhà máy (deploy N lần) | Trung bình — vẫn phải deploy riêng từng nơi | Kém hơn, tương tự A | Trung bình — tách rủi ro theo từng nhà máy nhưng vẫn phải tự deploy N lần |
| **C. Google Sheets + Apps Script Library (đã chọn)** | Tốt — phát bản mẫu (template) cho từng nhà máy, logic cốt lõi dùng chung qua Library | Giữ nguyên/tốt nhất — mọi sheet trung gian vẫn hiển thị như Excel hiện tại | Thấp hơn — không có "một điểm chịu trách nhiệm duy nhất"; mỗi nhà máy tự chủ vận hành bản của mình |

### Lý do chọn phương án C

- Giữ trọn triết lý đã được xác nhận qua thực tế nhiều tháng: **công thức nhìn thấy được, không hộp đen**.
- Việc "sửa một chỗ, áp dụng nhiều nơi" — yếu tố quan trọng cho hồ sơ sáng kiến — đạt được qua **Apps Script Library** (thư viện code dùng chung, có version), không bắt buộc phải qua kiến trúc Web App tập trung.
- Không cần vai trò "quản trị viên trung tâm" theo dõi một hệ thống chạy cho mọi nhà máy — giảm gánh nặng vận hành dài hạn cho người phụ trách.
- Loại bỏ hoàn toàn nhóm lỗi tốn công sức nhất trong lịch sử dự án (khác biệt hành vi VBA giữa Mac/Windows), vì Apps Script chạy trên nền tảng thống nhất (trình duyệt + cloud của Google), không phụ thuộc hệ điều hành máy người dùng.

### Đánh đổi cần lưu ý

- Không có sẵn màn hình tổng hợp/giám sát nhiều nhà máy cùng lúc (nếu sau này cần, có thể bổ sung một Sheet/Web App tổng hợp riêng, đọc dữ liệu từ các Sheet nhà máy qua `IMPORTRANGE` hoặc Apps Script — không ảnh hưởng kiến trúc cốt lõi).
- Khi thư viện code cập nhật (sửa lỗi thuật toán), **từng nhà máy phải tự cập nhật phiên bản thư viện** trong bản của mình — không tự động đồng loạt như Web App trung tâm. Cần quy trình thông báo/hướng dẫn cập nhật rõ ràng (xem mục "Quy trình phân phối" bên dưới).

## Kiến trúc mục tiêu

```
┌─────────────────────────────────────────────────────────┐
│  QDD-Core-Library (Apps Script Library, 1 nguồn duy nhất) │
│  - Ramp Engine (dựng đường cong công suất)                │
│  - Nội suy khi ngắt ramp                                   │
│  - Chuyển tiếp trạng thái qua 00:00 (carry-over)           │
│  - Tích phân diện tích theo 48 chu kỳ → Qdd                │
│  - Quy đổi Qdd_V, Qdc, Qmp, Qdư                             │
│  - Parser CSV (6001/6303)                                  │
│  Quản lý bằng clasp + Git, có version, có test tự động      │
│  (dựa trên docs/09_Test_Cases.md, viết lại thành test code)│
└─────────────────────────────────────────────────────────┘
                    │ (gọi qua Library, theo version)
     ┌──────────────┼──────────────┬──────────────┐
     ▼                              ▼               ▼
┌──────────┐                 ┌──────────┐    ┌──────────┐
│ Nhà máy A │                 │ Nhà máy B │    │ Nhà máy C │
│ Google    │                 │ Google    │    │ Google    │
│ Sheets +  │                 │ Sheets +  │    │ Sheets +  │
│ Apps      │                 │ Apps      │    │ Apps      │
│ Script    │                 │ Script    │    │ Script    │
│ gắn liền  │                 │ gắn liền  │    │ gắn liền  │
│ (menu,    │                 │ (menu,    │    │ (menu,    │
│ sidebar   │                 │ sidebar   │    │ sidebar   │
│ nhập CSV) │                 │ nhập CSV) │    │ nhập CSV) │
│ Cấu hình  │                 │ Cấu hình  │    │ Cấu hình  │
│ riêng     │                 │ riêng     │    │ riêng     │
│ (CAI_DAT) │                 │ (CAI_DAT) │    │ (CAI_DAT) │
└──────────┘                 └──────────┘    └──────────┘
```

Mỗi nhà máy: dữ liệu, cấu hình (`CAI_DAT` — tốc độ ramp, hệ số Qdd_V, dung sai, nguồn công tơ có thể khác nhau) và quyền truy cập **hoàn toàn độc lập** với nhà máy khác — không có kho dữ liệu chung nào chứa dữ liệu của tất cả nhà máy.

## Quy trình phân phối (khi nhân rộng cho nhà máy mới)

1. Tạo bản sao ("Make a copy") từ Google Sheets mẫu chính thức.
2. Sheet mẫu đã gắn sẵn tham chiếu tới `QDD-Core-Library` phiên bản ổn định mới nhất tại thời điểm phát hành.
3. Người phụ trách nhà máy mới điền `CAI_DAT` theo thông số thực tế của nhà máy đó.
4. Khi thư viện có bản vá lỗi, thông báo cho các nhà máy đang dùng để họ tự cập nhật số phiên bản Library trong Apps Script Editor (không cần thay đổi gì khác).

## Việc cần làm tiếp (Giai đoạn 2)

- [ ] Viết `docs/04_Algorithm_Specification.md` làm đặc tả kỹ thuật cho `QDD-Core-Library` (đang chờ phân tích chi tiết VBA hiện tại).
- [ ] Thiết lập repo/thư mục cho `QDD-Core-Library` (quản lý bằng `clasp`), migrate logic từ VBA sang Apps Script theo đặc tả.
- [ ] Viết test tự động dựa trên `docs/09_Test_Cases.md` (31 UAT) — chạy được cả trên bản VBA cũ (baseline) và bản Apps Script mới để đối chiếu 1:1.
- [ ] Tạo Google Sheets mẫu (template) cho nhà máy đầu tiên, gắn Library.
