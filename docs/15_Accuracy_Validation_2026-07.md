# Kiểm tra độ chính xác — dữ liệu thực tế tháng 7/2026

Đợt kiểm tra đầu tiên dùng dữ liệu vận hành thật (nhà máy Duyên Hải 1), đối chiếu thuật toán đã tài liệu hoá ở [04_Algorithm_Specification.md](04_Algorithm_Specification.md) với **kết quả tính tay độc lập** (không phải xuất từ công cụ VBA) trong file `test-data/Kiểm tra Qdu ngày 200726.xlsx`.

## Phương pháp

1. Cài đặt lại chính xác pipeline (chọn công suất hiệu lực R01-R03 → Ramp Engine nội suy R06 → dựng đoạn công suất → tích phân hình thang 48 chu kỳ) bằng Python, bám sát công thức Excel gốc đã trích xuất — không phải viết lại theo mô tả nghiệp vụ.
2. Chạy trên **10 ngày có đủ dữ liệu** (danh sách lệnh + CSV công tơ + kết quả tính tay đối chiếu): 01, 02, 03, 06, 07, 10, 16, 17, 18, 19/07/2026.
3. P0 (công suất đầu ngày) được suy ra an toàn từ chính giá trị Qdd chu kỳ 1 của bảng tính tay, chỉ khi lệnh đầu tiên trong ngày bắt đầu sau 00:30 (khi đó chu kỳ 1 chắc chắn là đoạn giữ nguyên công suất, không bị ramp cắt ngang) — áp dụng được cho cả 19/19 tổ hợp ngày+tổ máy kiểm tra.
4. So sánh Qdd tính được với Qdd tính tay, từng chu kỳ trong 48 chu kỳ.

Công cụ: `test-data/tools/qdd_engine.py` + `run_validation.py` (chạy cục bộ, không commit dữ liệu vận hành thật lên GitHub — xem `test-data/README.md`).

## Kết quả tổng quan

**18/19 tổ hợp (ngày, tổ máy) khớp gần như tuyệt đối** — sai lệch tối đa 0.000–0.015 MW, nằm trong sai số làm tròn của bảng tính tay, không phải sai lệch thuật toán.

| Ngày | S1 | S2 |
|---|---|---|
| 01/07 | ✅ khớp | ✅ khớp |
| 02/07 | ✅ khớp | ✅ khớp (lệch 0.015 MW, làm tròn) |
| 03/07 | ✅ khớp | ✅ khớp |
| 06/07 | ✅ khớp | ✅ khớp (lệch 0.011 MW, làm tròn) |
| 07/07 | ✅ khớp | ⚠️ **lệch ở giai đoạn sự cố — đã xác nhận không phải lỗi công thức, xem bên dưới** |
| 10/07 | ✅ khớp | (không có lệnh S2 hôm đó) |
| 16/07 | ✅ khớp | (không có lệnh S2 hôm đó) |
| 17/07 | ✅ khớp | (không có lệnh S2 hôm đó) |
| 18/07 | ✅ khớp | (không có lệnh S2 hôm đó) |
| 19/07 | ✅ khớp | (không có lệnh S2 hôm đó) |

**Kết luận chính**: thuật toán Ramp Engine + tích phân diện tích đã tài liệu hoá ở `04_Algorithm_Specification.md` là **chính xác**, khớp hoàn toàn với vận hành thực tế trong mọi trường hợp có lệnh thay đổi/duy trì công suất bình thường (SO, MO, MO dừng sớm, nhiều lệnh chồng thời gian, ngày không có lệnh mới).

## Điều tra ban đầu (đã kết luận: KHÔNG phải lỗi công thức)

Ngày 07/07, tổ S2: có một lệnh dừng máy (nội dung "Ngừng tổ máy") lúc khoảng 10h34, với **CS ra lệnh = 0 MW** và **CS hoàn thành = 0 MW**, bị công thức hiện tại loại khỏi `LENH_DIEU_DO` (điều kiện `AND(M="MO", O>0)` — xem [03_Business_Rules.md](03_Business_Rules.md#r01r03--chọn-công-suất-hiệu-lực-theo-loại-lệnh)), khiến công cụ giữ nguyên công suất trước đó (435,7 MW) thay vì về 0 như bảng tính tay.

Ban đầu mình nghi ngờ đây là lỗ hổng thật và đã thử vá công thức (`O>0`→`O>=0`). **Sau khi trao đổi trực tiếp với người phụ trách nghiệp vụ, xác nhận đây KHÔNG phải lỗi — mà là quy tắc nghiệp vụ đã có chủ đích:**

> "Trip tổ máy vẫn tính bình thường. Khi nào có lệnh 0-0 (cả CS ra lệnh và CS hoàn thành đều = 0) thì không tính. Còn khi khởi động tổ máy thì khi CS ra lệnh và CS hoàn thành đều hoàn thành ở cùng một tải (ví dụ 435,7-435,7) thì mới tính là đã hoàn thành và bắt đầu tính Qdu."

Nói cách khác: **công thức gốc `>0` (loại bỏ lệnh 0-0) là đúng** — bản vá đã được **hoàn tác**, không đưa vào `legacy/`.

### Vì sao bảng tính tay vẫn cho ra Qdd=0 trong giai đoạn sự cố

Đối chiếu với quy tắc chính thức ở trên: lệnh "khởi động lại" đúng chuẩn phải có **CS ra lệnh = CS hoàn thành = 435,7** (cả hai cùng giá trị) thì mới được coi là hoàn thành và bắt đầu tính lại. Nhưng lệnh "Hoà lưới" thực tế trong dữ liệu ngày 07/07 chỉ có `CS ra lệnh=435,7`, còn `CS hoàn thành=0` — **không khớp mẫu 435,7-435,7** mà quy tắc yêu cầu. Đây là một khoảng trống trong **cách ghi dữ liệu lúc xảy ra sự cố** (dữ liệu lịch sử cụ thể của ngày này), không phải trong bản thân quy tắc hay công thức.

→ Kết luận: người tính tay cho ngày 07/07 đã dùng thêm hiểu biết thực tế về sự cố (biết chắc tổ máy đã ngừng) để điền Qdd=0 thủ công, thay vì thuần tuý chạy công thức trên đúng dữ liệu danh sách lệnh đã xuất — không phải điều công thức tự động có thể suy ra được từ dữ liệu này. Không có hành động sửa nào cần làm ở đây.

## Việc cần làm tiếp

- [x] Điều tra và xác nhận với người phụ trách nghiệp vụ về quy tắc lệnh 0-0 và khởi động lại — **kết luận: công thức gốc đúng, không cần sửa** (UAT-32 đóng, xem cập nhật ở [09_Test_Cases.md](09_Test_Cases.md)).
- [ ] Ghi rõ quy tắc "khởi động lại cần CS ra lệnh = CS hoàn thành cùng giá trị" vào quy trình nhập liệu vận hành, để dữ liệu tương lai đủ điều kiện tính tự động (không cần can thiệp tay như ngày 07/07).
- [ ] Mở rộng kiểm tra sang các ngày còn thiếu danh sách lệnh (04, 05, 08, 09, 20/07) khi có dữ liệu.
- [ ] Kiểm tra thêm tình huống ramp qua 00:00 (UAT-04) — chưa có ngày nào trong đợt này rơi đúng tình huống ramp chưa hoàn tất lúc nửa đêm.
