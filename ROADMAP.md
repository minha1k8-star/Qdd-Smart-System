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

## Giai đoạn 2 — Khởi tạo ứng dụng mới (đã xác nhận hướng, chưa bắt đầu code)

**Xác nhận ngày 2026-07-24**: chuyển sang **Google Sheets + Apps Script gắn liền từng file + Thư viện code dùng chung (Apps Script Library)**. **Không làm Web App trung tâm** — lý do và phương án đã cân nhắc xem [05_System_Architecture.md](docs/05_System_Architecture.md). Quyết định này xuất phát từ mục tiêu dự án: được xây dựng để đề xuất công nhận **sáng kiến kỹ thuật**, dự kiến nhân rộng cho nhiều nhà máy khác — cần vừa nhân rộng dễ dàng vừa giữ được khả năng audit công thức bằng mắt.

- [x] Thiết lập khung `QDD-Core-Library` (`src/QDD-Core-Library/`, quản lý bằng `clasp` + Git) — đã port CommandFilter (R01-R03), RampEngine (R06), Segments, AreaIntegration (R08), QddCalculator (R09-R14), CsvParser (R10/R12), có test cục bộ (17/17 pass, không cần dữ liệu thật). Xem `src/QDD-Core-Library/README.md`.
- [ ] Chưa port: carry-over qua nửa đêm (R07), báo cáo tháng/snapshot, cảnh báo lệnh 0-0 (UAT-34).
- [x] Triển khai thật lên Apps Script (`clasp login`/`create`/`push`) — đã đăng nhập, đã push (khắc phục lỗi "Premature close" do Node quá mới bằng Node 20 qua Homebrew).
- [x] Tạo Google Sheets mẫu (template) cho nhà máy đầu tiên (Duyên Hải 1), gắn Library — xem `src/NhaMay-Mau-Template/README.md`.
- [x] Menu Apps Script cơ bản (6 mục: thiết lập, lưu CSV, tính 1 ngày, tính hàng loạt, báo cáo tháng) — chưa đủ 15 mục như VBA (thiếu backup, kiểm tra hệ thống, xuất file riêng).
- [x] Carry-over qua nửa đêm (R07) — P0 tự ghi sau mỗi lần tính; ramp dở dang lúc 24:00 tự chạy tiếp sang ngày sau.
- [x] Cảnh báo lệnh 0-0 (UAT-34) và cảnh báo ngày không có lệnh nào.
- [x] Xuất báo cáo ra file Excel/PDF riêng (theo ngày, khoảng ngày, hoặc cả tháng).

## Giai đoạn 3 — Chuyển thuật toán (chưa bắt đầu)

- [ ] Chuyển toàn bộ logic Qdd/Qdư từ VBA sang `QDD-Core-Library` (Apps Script), bám sát `docs/04_Algorithm_Specification.md`.
- [ ] Viết test tự động dựa trên 31 UAT case ở `docs/09_Test_Cases.md`.
- [ ] Kiểm thử với bộ dữ liệu đã xác nhận từ hệ thống Excel/VBA (đối chiếu số liệu 1:1, không chỉ chạy UAT mới).
- [ ] **Cảnh báo lệnh "0-0" (UAT-34)**: khi phát hiện một lệnh có CS ra lệnh = CS hoàn thành = 0 trong ngày đang tính, hiển thị cảnh báo rõ ràng cho người vận hành thay vì âm thầm giữ nguyên công suất trước đó — xem [14_Knowledge_Transfer.md](docs/14_Knowledge_Transfer.md#giới-hạn-vận-hành-lệnh-0-0-không-được-tự-động-phát-hiện-072026-uat-32). Chỉ cảnh báo, không tự ý sửa số.

## Giai đoạn 4 — Hoàn thiện (chưa bắt đầu)

- [ ] Báo cáo tháng.
- [ ] Dashboard.
- [ ] Nhật ký.
- [ ] Tài liệu người dùng/quản trị.
- [ ] Đóng gói phát hành.

## Vì sao tách thành 2 giai đoạn tài liệu và code

Xem [docs/AI_CONTEXT.md](docs/AI_CONTEXT.md): hệ thống Excel/VBA hiện tại đã được kiểm chứng bằng dữ liệu vận hành thực tế qua nhiều tháng, chứa nhiều quy tắc nghiệp vụ tinh vi (ví dụ: chọn công suất hiệu lực theo loại lệnh, ngắt ramp, chuyển tiếp qua nửa đêm) mà nếu viết lại vội vàng sang nền tảng khác sẽ dễ mất chính xác. Tài liệu hoá trước giúp bất kỳ ai (người hoặc AI) viết lại hệ thống sau này có một "hợp đồng" rõ ràng để đối chiếu, thay vì phải suy đoán lại từ đầu.
