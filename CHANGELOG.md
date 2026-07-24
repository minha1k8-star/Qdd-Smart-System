# Changelog

Định dạng: mỗi mục ghi ngày (nếu biết), thay đổi, và lý do khi có thể xác định từ tài liệu gốc.

## [Unreleased] — Xuất báo cáo ngày/khoảng ngày ra file riêng

- Thêm `ExportReport.js` + mục "6. Xuất báo cáo" trong sidebar: chọn khoảng ngày + tổ máy + định dạng (Excel/PDF), xuất từ dữ liệu đã có trong `KET_QUA` (không tính lại) thành 1 file lưu vào Google Drive, sidebar hiện link tải.
- Layout xuất: **mỗi ngày 1 tab**, các tổ máy đã chọn nằm **cạnh nhau trong cùng tab** — khớp layout file báo cáo gốc (S1/S2 cạnh nhau), không phải mỗi tổ máy 1 tab riêng.
- Đã đẩy lên Sheet thật.

## Tải CSV hàng loạt

- **Tải CSV hàng loạt** (mục mới trong sidebar): chọn nhiều file CSV cùng lúc, tự nhận diện tổ máy/loại dữ liệu bằng cách so tên file với mã công tơ đã cấu hình (tận dụng quy ước đặt tên thật `<ngày><tháng><mã công tơ>.CSV`), tự đọc ngày từng file — phục vụ tính hàng loạt nhiều ngày mà không phải upload từng file một. File không tự nhận diện được báo riêng theo tên, xử lý tay bằng mục "Lưu CSV" (1 file).
- Quyết định: **danh sách lệnh vẫn nhập/dán tay vào sheet LENH**, không làm import file — nguồn là .xlsx nhị phân, không đọc trực tiếp bằng JS phía trình duyệt như CSV (text thuần) được; người dùng xác nhận copy tay là đủ đơn giản, không cần đầu tư thêm.

## Sidebar UI thay cho hộp thoại prompt

Phản hồi người dùng: thao tác qua `prompt()` gõ tay ngày tháng + phải tự điền P0 + CSV qua 2 bước cảm giác rời rạc, khó dùng hơn "1 nút bấm" của VBA cũ. Đã làm lại:

- **`Sidebar.html` + `Controller.js`** thay toàn bộ menu prompt-based: 1 bảng điều khiển duy nhất, có lịch chọn ngày (`<input type=date>`), upload file CSV trực tiếp (đọc bằng FileReader phía trình duyệt, gửi thẳng lên server - bỏ hẳn bước Import vào CSV_STAGING).
- **Tự động suy ra P0** từ chu kỳ cuối ngày liền trước đã tính (`readOrInferP0_`) — chỉ còn cần nhập tay P0_NGAY cho đúng 1 lần (ngày đầu tiên dùng hệ thống). Đây là xấp xỉ, chưa phải carry-over R07 đầy đủ.
- Bỏ sheet `CSV_STAGING`, menu rút gọn còn 2 mục (Bảng điều khiển, Thiết lập sheet).
- Đã đẩy lên Sheet thật (Duyên Hải 1).

## Giai đoạn 2: triển khai thật + Sheet mẫu

- **Triển khai thật QDD-Core-Library lên Google Apps Script** (đăng nhập, tạo project, push code). Khắc phục lỗi "Premature close" của `clasp login` (do Node.js v24 quá mới) bằng cách cài thêm Node 20 LTS qua Homebrew chỉ để chạy đăng nhập. Publish version 1 làm Library, Script ID `10_vjTSgVjZodA7xTkJ_qJaGom3JDx_tnYE0YgWA_cphh1Q7g_lTKMLUO`.
- Thêm `.claspignore` (loại file test Node ra khỏi bản đẩy Apps Script - nếu không sẽ lỗi toàn bộ thư viện vì `require`/`fs` không tồn tại trong môi trường Apps Script).
- **Tạo Sheet mẫu nhà máy đầu tiên** (`src/NhaMay-Mau-Template/`, Duyên Hải 1) — Google Sheets thật, gắn `QDD-Core-Library`, có menu 6 chức năng: thiết lập sheet, lưu CSV (tận dụng File > Import có sẵn của Sheets), tính 1 ngày, tính hàng loạt nhiều ngày/nhiều tổ máy, tổng hợp báo cáo tháng trực tiếp từ kết quả đã có. Giới hạn: chưa có carry-over R07 (P0 phải nhập tay vào sheet `P0_NGAY`).

## Giai đoạn 2: khởi tạo QDD-Core-Library

- Khởi tạo `src/QDD-Core-Library/` (Apps Script Library, quản lý bằng `clasp`). Port CommandFilter (R01-R03, giữ đúng quy tắc UAT-32), RampEngine (R06, nội suy ngắt ramp), Segments, AreaIntegration (R08, tích phân hình thang), QddCalculator (R09-R14), CsvParser (R10/R12) — trực tiếp từ `tools/reference_engine/qdd_engine.py` đã kiểm chứng bằng dữ liệu thật.
- Thêm bộ test cục bộ chạy bằng Node (`tests/run_tests.js`, 17/17 pass), không cần deploy Apps Script, không cần dữ liệu thật — mã hoá quy tắc UAT-32 thành test tự động để tránh sửa nhầm lại.
- Thêm `BatchCalculator.js` (tính nhiều ngày/nhiều tổ máy cùng lúc — khác biệt kiến trúc lớn so với VBA vốn chỉ tính được 1 ngày) và `MonthlyReport.js` (tổng hợp báo cáo tháng trực tiếp từ dữ liệu gốc, không cần snapshot như VBA). Cập nhật `docs/05_System_Architecture.md` với thiết kế luồng nhập dữ liệu (vừa từng ngày vừa upload hàng loạt). Tổng 29/29 test pass.
- Chưa port: carry-over qua nửa đêm (R07), báo cáo tháng/snapshot phía Sheet (định dạng/xuất), cảnh báo lệnh 0-0 (UAT-34), thiết kế sheet lưu trữ nhiều ngày. Chưa triển khai thật lên Apps Script (cần người dùng tự `clasp login`).

## Documentation Foundation

- Khởi tạo repo GitHub `Qdd-Smart-System`.
- Thêm `CLAUDE.md`, `AGENTS.md`, `docs/AI_CONTEXT.md` — hướng dẫn làm việc cho AI agent.
- Thêm `README.md`, `docs/00_Project_Overview.md`, `docs/03_Business_Rules.md`, `docs/06_Database_Design.md`, `docs/09_Test_Cases.md`, `docs/14_Knowledge_Transfer.md`, `docs/PROJECT_STATUS.md`, `ROADMAP.md` — dựa trên nội dung thật của workbook v1.3.0/v1.3.1 và lịch sử trao đổi phát triển.
- Đưa bản Excel/VBA v1.3.1 (bản chính thức đang chạy) vào `legacy/`.
- Xác nhận hướng kiến trúc Giai đoạn 2: **Google Sheets + Apps Script Library dùng chung**, không làm Web App trung tâm — ghi lại quyết định và lý do ở `docs/05_System_Architecture.md`.
- Thêm `docs/04_Algorithm_Specification.md` — trích xuất trực tiếp từ công thức Excel thật (`XU_LY_LENH`, `DOAN_CONG_SUAT`, `DIEN_TICH`, `TINH_TOAN`) và code VBA, phát hiện quan trọng: toàn bộ Ramp Engine/nội suy/tích phân diện tích nằm ở công thức Excel, không nằm trong VBA; đồng thời phát hiện nghi vấn UAT-11 (CSV sai ngày có thể không còn bị từ chối trong v1.3.1).
- Kiểm tra độ chính xác bằng dữ liệu vận hành thật (10 ngày, 07/2026, nhà máy Duyên Hải 1): 18/19 tổ hợp ngày+tổ máy khớp gần tuyệt đối với kết quả tính tay. 1 trường hợp lệch (07/07, tổ S2, sự cố/trip) — điều tra ban đầu nghi là lỗ hổng công thức (thử vá `O>0`→`O>=0`), sau khi xác nhận trực tiếp với người phụ trách nghiệp vụ thì **kết luận hành vi gốc đúng theo thiết kế** (lệnh 0-0 chủ đích không tính; khởi động lại cần CS ra lệnh=CS hoàn thành cùng giá trị) — đã hoàn tác bản vá thử nghiệm, không sửa `legacy/`. Thêm UAT-32 (đã đóng), UAT-33, và UAT-34 (đề xuất cảnh báo khi phát hiện lệnh 0-0, kế hoạch Giai đoạn 2/3 — không tự sửa số). Xem `docs/15_Accuracy_Validation_2026-07.md`, `docs/14_Knowledge_Transfer.md`, `ROADMAP.md`.

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
