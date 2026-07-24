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

## Nhiều ngày, nhiều tổ máy cùng lúc (khác biệt kiến trúc lớn so với VBA)

Bản Excel/VBA chỉ tính được **1 ngày tại một thời điểm** — không phải giới hạn thuật toán, mà do `CAI_DAT!B4` là một ô cấu hình toàn cục duy nhất mà mọi sheet downstream (`LENH_GOC`, `LENH_DIEU_DO`, `XU_LY_LENH`...) đều gắn theo, nên tính ngày khác phải ghi đè ngày cũ. Đây là lý do VBA cần cơ chế "snapshot" (`LICH_SU_THANG`, sheet ẩn `LS_...`) để không mất dữ liệu khi ghi đè.

`QDD-Core-Library` không có giới hạn này: hàm tính toán nhận **ngày + tổ máy làm tham số** (`QDD.QddCalculator.calculateDay`, `QDD.BatchCalculator.calculateMultiple`), không có "ô cấu hình toàn cục" nào bị ghi đè giữa các lần gọi — có thể tính hàng chục ngày × nhiều tổ máy trong cùng một lần chạy.

**Luồng nhập dữ liệu** (xác nhận với người dùng 2026-07-24): hỗ trợ **cả hai**:
1. Nhập từng ngày như thói quen cũ — nhưng thay vì ghi đè, dữ liệu được **tích luỹ thêm** vào một sheet lệnh dùng chung có cột ngày (tương tự cấu trúc `LENH_GOC` nhưng trải nhiều ngày), và một sheet CSV tích luỹ tương tự theo (ngày, mã công tơ).
2. Upload hàng loạt nhiều ngày cùng lúc (nhiều file CSV + danh sách lệnh) — hệ thống tự nhận diện ngày của từng file rồi tính hàng loạt bằng `QDD.BatchCalculator`.

**Báo cáo tháng**: vì mỗi ngày độc lập ngay từ khi tính (không sheet nào bị ghi đè), báo cáo tháng (`QDD.MonthlyReport.aggregate`) tính **trực tiếp từ dữ liệu gốc** bất cứ lúc nào, không cần bước "chốt" dữ liệu như VBA.

**Xuất báo cáo từng ngày**: mỗi phần tử kết quả của `calculateDay`/`calculateMultiple` đã là báo cáo ngày đầy đủ (48 chu kỳ), có thể xuất độc lập bất cứ lúc nào — không phải tính năng riêng cần xây thêm ở tầng engine.

Chi tiết cài đặt: [`src/QDD-Core-Library/README.md`](../src/QDD-Core-Library/README.md).

## Việc cần làm tiếp (Giai đoạn 2)

- [x] Viết `docs/04_Algorithm_Specification.md` làm đặc tả kỹ thuật cho `QDD-Core-Library`.
- [x] Thiết lập `src/QDD-Core-Library/` (quản lý bằng `clasp`), port CommandFilter/RampEngine/Segments/AreaIntegration/QddCalculator/CsvParser/BatchCalculator/MonthlyReport từ Python đã kiểm chứng, 29/29 test cục bộ pass.
- [ ] Viết test tự động dựa trên `docs/09_Test_Cases.md` (31 UAT) — hiện chỉ có test smoke bằng dữ liệu giả lập, chưa map đủ 31 case UAT.
- [ ] Port carry-over qua nửa đêm (R07), báo cáo tháng phía Sheet (khác `MonthlyReport.js` — cần lớp xuất/định dạng), cảnh báo lệnh 0-0 (UAT-34).
- [ ] Thiết kế sheet lưu trữ nhiều ngày (lệnh tích luỹ có cột ngày, CSV tích luỹ theo ngày+công tơ) — xem mục "Nhiều ngày, nhiều tổ máy cùng lúc" ở trên.
- [ ] Triển khai thật lên Apps Script (`clasp login`/`create`/`push`) — cần người dùng tự đăng nhập Google.
- [ ] Tạo Google Sheets mẫu (template) cho nhà máy đầu tiên, gắn Library.
