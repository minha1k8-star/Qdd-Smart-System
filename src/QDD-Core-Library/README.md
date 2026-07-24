# QDD-Core-Library

Apps Script Library dùng chung cho QDD Smart System — logic Ramp Engine, chọn công suất hiệu lực (R01-R03), tích phân diện tích 48 chu kỳ (R08), quy đổi Qdd/Qdd_V/Qdc/Qmp/Qdư (R09-R14). Đây là bước đầu của **Giai đoạn 2/3** trong [ROADMAP.md](../../ROADMAP.md) — thay thế phần "engine" của VBA, giữ nguyên triết lý mỗi nhà máy có Sheet + cấu hình riêng (xem [docs/05_System_Architecture.md](../../docs/05_System_Architecture.md)).

**Nguồn logic**: port trực tiếp (gần như dòng-đối-dòng) từ [`tools/reference_engine/qdd_engine.py`](../../tools/reference_engine/qdd_engine.py) — bản Python đã được đối chiếu và khớp với 18/19 tổ hợp ngày+tổ máy dữ liệu vận hành thật (xem [docs/15_Accuracy_Validation_2026-07.md](../../docs/15_Accuracy_Validation_2026-07.md)). Quy tắc UAT-32 (lệnh 0-0 không được tính) đã được mã hoá thành test tự động, không phụ thuộc trí nhớ.

## Cấu trúc

| File | Vai trò | Tương ứng docs/04 |
|---|---|---|
| `Config.js` | Hằng số nghiệp vụ (tốc độ ramp, hệ số Qdd_V, dung sai...) | Mục 9 |
| `CommandFilter.js` | R01-R03: chọn lệnh hiệu lực + công suất hiệu lực | Mục 1 |
| `RampEngine.js` | R06: nội suy ngắt ramp giữa các lệnh | Mục 2 |
| `Segments.js` | Dựng đoạn công suất phủ 24h | Mục 3 |
| `AreaIntegration.js` | R08: tích phân hình thang 48 chu kỳ → Qdd | Mục 4 |
| `CsvParser.js` | R10/R12: đọc dòng KwhGiao từ CSV | Mục 7 (một phần) |
| `QddCalculator.js` | R08-R14: tổng hợp kết quả 48 chu kỳ | Mục 5 |
| `BatchCalculator.js` | Tính nhiều ngày/nhiều tổ máy cùng lúc (khác VBA — không giới hạn 1 ngày) | — |
| `MonthlyReport.js` | Tổng hợp báo cáo tháng trực tiếp từ nhiều ngày, không cần snapshot | Mục 8 (thay thế cách làm VBA) |
| `Public.js` | Hàm công khai khi dùng làm Library | — |

## Nhiều ngày, nhiều tổ máy cùng lúc (khác biệt lớn nhất so với VBA)

Bản Excel/VBA chỉ tính được **1 ngày tại một thời điểm** vì `CAI_DAT!B4` là một ô cấu hình toàn cục duy nhất — mọi sheet downstream đều gắn với đúng ngày đó. Thư viện này không có giới hạn đó: mỗi lệnh gọi `calculateDay`/`QDD.BatchCalculator.calculateMultiple` nhận **ngày + tổ máy làm tham số**, độc lập hoàn toàn với các lệnh gọi khác. Có thể tính 30 ngày × 2 tổ máy trong cùng một lần chạy.

Hệ quả: **báo cáo tháng không cần cơ chế "snapshot đóng băng công thức"** như VBA (`LICH_SU_THANG`, sheet ẩn `LS_...`) — chỉ cần dữ liệu đầu vào (danh sách lệnh, CSV) được lưu theo ngày (không ghi đè), `QDD.MonthlyReport.aggregate` có thể tính lại báo cáo tháng bất cứ lúc nào trực tiếp từ dữ liệu gốc.

**Luồng nhập dữ liệu dự kiến** (xem docs/05_System_Architecture.md): vừa hỗ trợ nhập từng ngày (tích luỹ dần vào một sheet lệnh chung có cột ngày, không ghi đè ngày cũ) vừa hỗ trợ upload hàng loạt nhiều ngày một lúc (nhiều file CSV + danh sách lệnh cùng lúc, hệ thống tự nhận diện ngày của từng file).

**Xuất báo cáo từng ngày đã nhập**: không cần logic riêng — mỗi phần tử trong kết quả của `calculateMultiple`/`calculateDay` đã là một báo cáo ngày đầy đủ (48 chu kỳ Qdd/Qdd_V/Qdc/Qmp/Qdu), độc lập với các ngày khác, xuất bất cứ lúc nào.

## Chưa cài đặt (còn lại của Giai đoạn 3)

- **R07 — chuyển tiếp qua nửa đêm** (carry-over, `AUTO_CARRY`): chưa port. Trong VBA đây là `CalculateEndOfDayState`/`EnsureCarryOverCommand` (docs/04 mục 6) — cơ chế chèn một lệnh tổng hợp vào đầu ngày sau, không phải một engine tính toán riêng. Cần dùng nguồn dữ liệu tương tự `TRANG_THAI_CONG_SUAT`.
- **Phát hiện CSV bị Excel-Mac dồn cột**: không áp dụng cho Apps Script vì không đi qua Excel để mở CSV — nhưng vẫn cần xử lý CSV do người dùng dán tay/upload có thể có định dạng khác nhau (dấu phân cách, BOM...).
- **Báo cáo tháng / snapshot** (`LICH_SU_THANG`, `LS_...`): chưa bắt đầu.
- **UAT-34 (cảnh báo lệnh 0-0)**: xem [ROADMAP.md](../../ROADMAP.md) — engine hiện tại loại lệnh 0-0 đúng nghiệp vụ nhưng KHÔNG cảnh báo cho người vận hành biết, cần bổ sung.

## Kiểm tra cục bộ (không cần Google, không cần dữ liệu thật)

```bash
cd src/QDD-Core-Library
node tests/run_tests.js
```

Test dùng dữ liệu giả lập, tính tay đối chiếu — không đọc `test-data/`. Trước khi đổi bất kỳ công thức nào trong các file `.js`, chạy lại lệnh trên để đảm bảo không phá vỡ hành vi đã xác nhận (đặc biệt là quy tắc UAT-32 — lệnh 0-0 phải luôn bị loại, không tự ý đổi lại `>0` thành `>=0`, xem [docs/09_Test_Cases.md](../../docs/09_Test_Cases.md) và [AGENTS.md](../../AGENTS.md)).

## Triển khai thật lên Google Apps Script (cần làm thủ công — cần tài khoản Google của bạn)

Các bước sau **cần bạn tự thực hiện**, vì phải đăng nhập tài khoản Google qua trình duyệt — không thể tự động hoá thay bạn được.

```bash
cd src/QDD-Core-Library
npm install              # cài clasp cục bộ trong project
npx clasp login           # mở trình duyệt, đăng nhập tài khoản Google sẽ sở hữu thư viện
npx clasp create --type standalone --title "QDD-Core-Library"   # tạo project Apps Script mới, sinh ra .clasp.json (không commit)
npx clasp push             # đẩy toàn bộ code .js lên Apps Script
npx clasp open              # mở project trên trình duyệt để đặt tên phiên bản, publish làm Library
```

Sau khi `clasp create`, một file `.clasp.json` sẽ xuất hiện chứa `scriptId` — đây là ID riêng của bạn, đã được gitignore, không lên GitHub.

### Publish làm Library và lấy Script ID

1. Trong Apps Script Editor (`npx clasp open`): **Deploy → New deployment → Library**.
2. Ghi lại **Script ID** (Project Settings) — đây là ID mà các Sheet của từng nhà máy sẽ dùng để "Add a library".
3. Mỗi khi sửa code và muốn các nhà máy dùng bản mới: `npx clasp push` rồi tạo **version mới** trong Deploy — các nhà máy tự chọn cập nhật phiên bản khi sẵn sàng (không tự động, đúng theo kiến trúc đã chọn ở [docs/05](../../docs/05_System_Architecture.md)).

### Quan trọng: tài khoản sở hữu

Nên tạo project này dưới **tài khoản Google Workspace của đơn vị** (nếu có), không phải Gmail cá nhân — để đảm bảo tính liên tục lâu dài khi sáng kiến được nhân rộng (xem lưu ý ở [docs/05_System_Architecture.md](../../docs/05_System_Architecture.md)).
