# Roadmap

## Giai đoạn 1 — Xây dựng hệ thống ✅ hoàn thành (07/2026)

Nền tảng: **Google Sheets + Apps Script gắn liền từng file + thư viện code dùng chung** (`QDD-Core-Library`). Không làm Web App trung tâm — lý do và các phương án đã cân nhắc xem [05_System_Architecture.md](docs/05_System_Architecture.md).

- [x] `QDD-Core-Library` — toàn bộ thuật toán: chọn công suất hiệu lực (R01-R03), Ramp Engine (R05-R06), dựng đoạn công suất, tích phân diện tích (R08), quy đổi Qdd/Qdd_V/Qdc/Qmp/Qdư (R09-R14), carry-over qua nửa đêm (R07), đọc CSV (R10/R12), báo cáo tháng. Quản lý bằng `clasp` + Git, **54/54 test cục bộ pass**.
- [x] Sheet mẫu cho nhà máy đầu tiên (Duyên Hải 1), gắn Library — xem [src/NhaMay-Mau-Template/README.md](src/NhaMay-Mau-Template/README.md).
- [x] Bảng điều khiển 6 mục: Tải CSV · Nhập danh sách lệnh · Tính · Báo cáo tháng · Xuất báo cáo · Dọn dữ liệu cũ.
- [x] Nhập danh sách lệnh **thẳng từ file Excel** tải lên — không copy tay, tự dò sheet và dòng tiêu đề, gộp theo ID Lệnh.
- [x] Tải nhiều CSV cùng lúc, tự nhận diện tổ máy/loại dữ liệu theo tên file và tự đọc ngày trong file.
- [x] Cảnh báo tự động: lệnh 0-0, ngày không có lệnh nào, ngày được nối tiếp ramp từ hôm trước.
- [x] Xuất báo cáo Excel/PDF theo ngày, khoảng ngày hoặc cả tháng; layout bám mẫu `Kiểm tra Qdu`, tên file mang đúng kỳ báo cáo.
- [x] Sheet `HUONG_DAN` ngay trong file + [hướng dẫn triển khai cho nhà máy mới](docs/16_Huong_Dan_Trien_Khai.md).

## Giai đoạn 2 — Kiểm chứng độ chính xác ✅ hoàn thành cho tổ S1

- [x] Đối chiếu 1:1 với bảng tính tay độc lập trên dữ liệu vận hành thật — **2 đợt, 16 tổ hợp (ngày, tổ máy)**, xem [15_Accuracy_Validation_2026-07.md](docs/15_Accuracy_Validation_2026-07.md).
- [x] Kiểm chứng tình huống khó nhất: **ramp vắt qua nửa đêm (R07)** bằng dữ liệu thật ngày 23→24/07/2026.
- [x] Truy nguyên 3 ô lệch ~1 MW — xác định là **bảng tính tay sai**, hệ thống đúng (có tính tay độc lập từng chu kỳ để chứng minh).
- [ ] **Kiểm chứng tổ S2** — mới chỉ chạy S1 (TC-29 trong [09_Test_Cases.md](docs/09_Test_Cases.md)).
- [ ] Kiểm thử việc giữ nguyên P0 do người dùng nhập tay (TC-30).
- [ ] Kiểm chứng tình huống khởi động lại sau sự cố — chờ dữ liệu đúng mẫu (TC-31).

## Giai đoạn 3 — Nhân rộng (đang làm)

- [x] Tài liệu triển khai từ đầu trên tài khoản Google khác, gồm cả xử lý sự cố thường gặp và cấu hình múi giờ.
- [x] Hồ sơ sáng kiến kỹ thuật: [17_Thuyet_Minh_Sang_Kien.md](docs/17_Thuyet_Minh_Sang_Kien.md) — còn các mục **[CẦN ĐIỀN]** là số liệu vận hành/kinh tế chỉ đơn vị mới có.
- [ ] Triển khai cho nhà máy thứ hai và **kiểm chứng lại quy trình triển khai** bằng một người chưa từng dựng hệ thống.
- [ ] Quy trình cập nhật phiên bản thư viện khi có nhiều nhà máy cùng dùng.

## Giai đoạn 4 — Mở rộng

- [x] Hỗ trợ nhà máy có **số tổ máy bất kỳ** — thêm tổ chỉ cần thêm dòng cấu hình, không sửa code.
- [ ] Cập nhật quy trình nhập liệu vận hành: ghi đủ CS hoàn thành khi khởi động lại tổ máy, để dữ liệu đủ điều kiện tính tự động (xem [15_Accuracy_Validation_2026-07.md](docs/15_Accuracy_Validation_2026-07.md)).

## Đã cân nhắc và quyết định KHÔNG làm

- **Dashboard tổng quan** và **nhật ký thao tác**: đã dựng thử nhật ký rồi gỡ bỏ theo quyết định nghiệp vụ — hệ thống giữ phạm vi hẹp đúng một việc là tính và xuất báo cáo Qdd/Qdư. Google Sheets đã có sẵn `Tệp → Lịch sử phiên bản` cho nhu cầu truy vết.

## Nguyên tắc xuyên suốt

**Độ chính xác số liệu quan trọng hơn tính năng.** Đây là công cụ tính toán chuyên biệt cho một quy trình nghiệp vụ hẹp, không phải phần mềm quản lý. Mọi tính năng mới đều phải trả lời được câu hỏi: nó có làm tăng rủi ro sai số liệu không?

Hệ quả cụ thể trong thiết kế:
- Mỗi việc chỉ có **một luồng code duy nhất** cần kiểm chứng (không có "tính 1 ngày" song song với "tính hàng loạt").
- Mọi thao tác không hoàn tác được đều **mặc định tắt**.
- Kết quả làm tròn ở **hiển thị**, không làm tròn giá trị thật.
- Tình huống mà máy không phân biệt được đúng/sai thì **cảnh báo cho người vận hành**, không tự ý sửa số.
