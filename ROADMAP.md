# Roadmap

## Giai đoạn 1 — Documentation Foundation (đang thực hiện)

Mục tiêu: hệ thống hoá toàn bộ tri thức nghiệp vụ/kỹ thuật đã tích luỹ qua nhiều tháng phát triển thành tài liệu chuẩn trong repo, độc lập với lịch sử chat hay trí nhớ của AI, **trước khi viết thêm bất kỳ dòng code mới nào**.

- [x] Khởi tạo repository GitHub.
- [x] Đưa bản Excel/VBA v1.3.1 chính thức vào `legacy/`.
- [x] Tài liệu hướng dẫn AI: `CLAUDE.md`, `AGENTS.md`, `docs/AI_CONTEXT.md`.
- [x] `docs/03_Business_Rules.md` — 14 quy tắc nghiệp vụ (R01-R14).
- [x] `docs/06_Database_Design.md` — schema các sheet.
- [x] `docs/09_Test_Cases.md` — 31 test case UAT (chưa thực thi).
- [ ] `docs/04_Algorithm_Specification.md` — chi tiết thuật toán Ramp Engine, nội suy, tính diện tích.
- [ ] `docs/14_Knowledge_Transfer.md` — lịch sử quyết định kỹ thuật, lỗi đã gặp.
- [ ] Thực thi 31 test case UAT trên bản v1.3.1 và ghi nhận kết quả thật (hiện toàn bộ đang "Chưa chạy").

**Điều kiện để chuyển sang Giai đoạn 2**: bộ tài liệu ở trên đầy đủ và được người phụ trách nghiệp vụ xác nhận đúng với hệ thống thực tế; ít nhất các UAT liên quan đến Ramp Engine (UAT-04, 05-08) đã chạy và Đạt trên bản v1.3.1 để có baseline đối chiếu.

## Giai đoạn 2 — Khởi tạo ứng dụng mới ✅ hoàn thành

**Xác nhận ngày 2026-07-24**: chuyển sang **Google Sheets + Apps Script gắn liền từng file + Thư viện code dùng chung (Apps Script Library)**. **Không làm Web App trung tâm** — lý do và phương án đã cân nhắc xem [05_System_Architecture.md](docs/05_System_Architecture.md). Quyết định này xuất phát từ mục tiêu dự án: được xây dựng để đề xuất công nhận **sáng kiến kỹ thuật**, dự kiến nhân rộng cho nhiều nhà máy khác — cần vừa nhân rộng dễ dàng vừa giữ được khả năng audit công thức bằng mắt.

- [x] Thiết lập khung `QDD-Core-Library` (`src/QDD-Core-Library/`, quản lý bằng `clasp` + Git) — đã port CommandFilter (R01-R03), RampEngine (R06), Segments, AreaIntegration (R08), QddCalculator (R09-R14), CsvParser (R10/R12), có test cục bộ (17/17 pass, không cần dữ liệu thật). Xem `src/QDD-Core-Library/README.md`.
- [x] Port nốt: carry-over qua nửa đêm (R07), báo cáo tháng, cảnh báo lệnh 0-0 (UAT-34). Thư viện hiện ở **version 4**, 45/45 test cục bộ pass.
- [x] Triển khai thật lên Apps Script (`clasp login`/`create`/`push`) — đã đăng nhập, đã push (khắc phục lỗi "Premature close" do Node quá mới bằng Node 20 qua Homebrew).
- [x] Tạo Google Sheets mẫu (template) cho nhà máy đầu tiên (Duyên Hải 1), gắn Library — xem `src/NhaMay-Mau-Template/README.md`.
- [x] Bảng điều khiển (sidebar) 6 mục: Tải CSV · Nhập danh sách lệnh · Tính · Báo cáo tháng · Xuất báo cáo · Dọn dữ liệu cũ. Không làm backup/kiểm tra hệ thống như bản VBA — **đã thống nhất bỏ**: Google Sheets có sẵn lịch sử phiên bản, còn "kiểm tra hệ thống" sinh ra chỉ để dò lỗi riêng của Excel/VBA.
- [x] Nhập danh sách lệnh **thẳng từ file Excel** tải lên (không copy tay, không sheet trung gian).
- [x] Tải CSV nhiều file cùng lúc, tự nhận diện tổ máy/loại dữ liệu theo tên file và tự đọc ngày trong file.
- [x] Carry-over qua nửa đêm (R07) — P0 tự ghi sau mỗi lần tính; ramp dở dang lúc 24:00 tự chạy tiếp sang ngày sau. **Đã kiểm chứng bằng dữ liệu thật** (23→24/07/2026), xem [15_Accuracy_Validation_2026-07.md](docs/15_Accuracy_Validation_2026-07.md).
- [x] Cảnh báo lệnh 0-0 (UAT-34) và cảnh báo ngày không có lệnh nào.
- [x] Xuất báo cáo ra file Excel/PDF riêng (theo ngày, khoảng ngày, hoặc cả tháng), tên file mang đúng kỳ báo cáo.

## Giai đoạn 3 — Chuyển thuật toán ✅ hoàn thành phần tính toán

- [x] Chuyển toàn bộ logic Qdd/Qdư từ VBA sang `QDD-Core-Library` (Apps Script), bám sát `docs/04_Algorithm_Specification.md`.
- [x] **Cảnh báo lệnh "0-0" (UAT-34)** — chỉ cảnh báo, không tự ý sửa số. Xem [14_Knowledge_Transfer.md](docs/14_Knowledge_Transfer.md#giới-hạn-vận-hành-lệnh-0-0-không-được-tự-động-phát-hiện-072026-uat-32).
- [x] Kiểm thử với dữ liệu vận hành thật, đối chiếu 1:1 với bảng tính tay — **2 đợt, 16 tổ hợp (ngày, tổ máy)**, xem [15_Accuracy_Validation_2026-07.md](docs/15_Accuracy_Validation_2026-07.md). Đợt 2 chạy trực tiếp trên bản Google Sheets, gồm cả tình huống ramp vắt qua nửa đêm.
- [ ] Viết test tự động dựa trên 31 UAT case ở `docs/09_Test_Cases.md` — hiện có 45 test cục bộ nhưng **chưa ánh xạ 1-1 với mã UAT**; toàn bộ 31 case vẫn ở trạng thái "Chưa chạy" trong tài liệu.
- [ ] Kiểm chứng **tổ S2** trên bản Google Sheets (mới chỉ chạy S1).

## Giai đoạn 4 — Hoàn thiện (đang dở)

- [x] Báo cáo tháng (tổng hợp vào `BAO_CAO_THANG` + xuất file gộp cả tháng).
- [x] Tài liệu người dùng: sheet `HUONG_DAN` ngay trong file + [16_Huong_Dan_Trien_Khai.md](docs/16_Huong_Dan_Trien_Khai.md).
- [x] Hồ sơ sáng kiến: [17_Thuyet_Minh_Sang_Kien.md](docs/17_Thuyet_Minh_Sang_Kien.md) (còn các mục **[CẦN ĐIỀN]** là số liệu vận hành/kinh tế chỉ đơn vị mới có).
- [ ] Dashboard.
- [ ] Nhật ký thao tác.
- [ ] Đóng gói phát hành cho nhà máy thứ hai (kiểm chứng lại quy trình trong `16_Huong_Dan_Trien_Khai.md` bằng một người chưa từng dựng hệ thống).

## Vì sao tách thành 2 giai đoạn tài liệu và code

Xem [docs/AI_CONTEXT.md](docs/AI_CONTEXT.md): hệ thống Excel/VBA hiện tại đã được kiểm chứng bằng dữ liệu vận hành thực tế qua nhiều tháng, chứa nhiều quy tắc nghiệp vụ tinh vi (ví dụ: chọn công suất hiệu lực theo loại lệnh, ngắt ramp, chuyển tiếp qua nửa đêm) mà nếu viết lại vội vàng sang nền tảng khác sẽ dễ mất chính xác. Tài liệu hoá trước giúp bất kỳ ai (người hoặc AI) viết lại hệ thống sau này có một "hợp đồng" rõ ràng để đối chiếu, thay vì phải suy đoán lại từ đầu.
