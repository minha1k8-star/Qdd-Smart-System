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
| 07/07 | ✅ khớp | ❌ **lệch nghiêm trọng — xem bên dưới** |
| 10/07 | ✅ khớp | (không có lệnh S2 hôm đó) |
| 16/07 | ✅ khớp | (không có lệnh S2 hôm đó) |
| 17/07 | ✅ khớp | (không có lệnh S2 hôm đó) |
| 18/07 | ✅ khớp | (không có lệnh S2 hôm đó) |
| 19/07 | ✅ khớp | (không có lệnh S2 hôm đó) |

**Kết luận chính**: thuật toán Ramp Engine + tích phân diện tích đã tài liệu hoá ở `04_Algorithm_Specification.md` là **chính xác**, khớp hoàn toàn với vận hành thực tế trong mọi trường hợp có lệnh thay đổi/duy trì công suất bình thường (SO, MO, MO dừng sớm, nhiều lệnh chồng thời gian, ngày không có lệnh mới).

## Phát hiện quan trọng: lệnh dừng máy với CS ra lệnh = 0 bị loại khỏi tính toán

Ngày 07/07, tổ S2: có một lệnh dừng máy (nội dung "Ngừng tổ máy") lúc khoảng 10h34, với **CS ra lệnh = 0 MW** và **CS hoàn thành = 0 MW**. Theo đúng công thức thật đang dùng trong workbook (cột T `LENH_DIEU_DO`, xem [03_Business_Rules.md](03_Business_Rules.md#r01r03--chọn-công-suất-hiệu-lực-theo-loại-lệnh)):

- Lệnh loại "MO" chỉ hợp lệ khi **CS ra lệnh > 0** (điều kiện `AND(M="MO", O>0)`).
- Lệnh loại "SO" chỉ hợp lệ khi **CS hoàn thành > 0** (điều kiện `AND(M="SO", P>0)`).

Lệnh dừng máy nói trên có cả hai giá trị công suất bằng 0 → **không thoả điều kiện nào**, bị loại hoàn toàn khỏi `LENH_DIEU_DO` (cột T = 0), dù đây là một lệnh dừng máy có thật, đã hoàn thành (`Hoàn thành = 1`).

**Hậu quả**: từ thời điểm dừng máy (~10h30) đến hết ngày, thuật toán tiếp tục giữ nguyên công suất trước đó (435,7 MW) thay vì về 0. Bảng tính tay đối chiếu tính đúng (Qdd = 0 từ chu kỳ 22 trở đi, có cả một chu kỳ ramp-xuống thủ công ở chu kỳ 21). Sai lệch: **+435,7 MW ở 27/48 chu kỳ** (từ khoảng 10h30 đến 24h00).

**Đây là một khoảng trống thật trong quy tắc nghiệp vụ hiện tại (R01-R03), không phải lỗi nhập liệu** — bất kỳ lệnh dừng máy nào có công suất mục tiêu bằng 0 (rất tự nhiên với một lệnh "Ngừng tổ máy") sẽ luôn bị điều kiện `>0` loại bỏ.

### Đề xuất hướng xử lý (cần người phụ trách nghiệp vụ xác nhận)

Một trong các hướng sau (hoặc kết hợp):

1. Nới điều kiện từ `>0` thành `>=0` cho trường hợp MO, khi đi kèm dấu hiệu là lệnh dừng máy (ví dụ dựa vào "Nội dung lệnh" chứa từ khoá "Ngừng"/"Dừng máy").
2. Thêm một quy tắc mới (R15?) riêng cho "Lệnh dừng máy về 0": nếu `Nội dung lệnh` là lệnh ngừng tổ máy và `Hoàn thành=1`, dùng P hiệu lực = 0 bất kể giá trị CS ra lệnh/CS hoàn thành.
3. Yêu cầu quy trình vận hành luôn ghi CS hoàn thành > 0 cho lệnh dừng máy (ví dụ một giá trị rất nhỏ khác 0) — rủi ro: phụ thuộc vào kỷ luật nhập liệu, không phải giải pháp kỹ thuật bền vững.

Đây là phát hiện **cần ưu tiên sửa trước khi migrate sang Apps Script** (xem [ROADMAP.md](../ROADMAP.md) Giai đoạn 3) — nếu không, lỗi này sẽ được migrate nguyên vẹn sang hệ thống mới.

## Việc cần làm tiếp

- [ ] Xác nhận với người phụ trách nghiệp vụ hướng xử lý lệnh dừng máy CS=0 (mục trên).
- [ ] Thêm test case cho tình huống này vào [09_Test_Cases.md](09_Test_Cases.md) (đã thêm UAT-32, xem bên dưới).
- [ ] Mở rộng kiểm tra sang các ngày còn thiếu danh sách lệnh (04, 05, 08, 09, 20/07) khi có dữ liệu.
- [ ] Kiểm tra thêm tình huống ramp qua 00:00 (UAT-04) — chưa có ngày nào trong đợt này rơi đúng tình huống ramp chưa hoàn tất lúc nửa đêm.
