# Kiểm tra độ chính xác — dữ liệu thực tế tháng 7/2026

Đợt kiểm tra đầu tiên dùng dữ liệu vận hành thật (nhà máy Duyên Hải 1), đối chiếu thuật toán đã tài liệu hoá ở [04_Algorithm_Specification.md](04_Algorithm_Specification.md) với **kết quả tính tay độc lập** trong file `test-data/Kiểm tra Qdu ngày 200726.xlsx`.

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

Nói cách khác: **quy tắc `>0` (loại bỏ lệnh 0-0) là đúng** — bản vá thử nghiệm đã được **hoàn tác**.

### Vì sao bảng tính tay vẫn cho ra Qdd=0 trong giai đoạn sự cố

Đối chiếu với quy tắc chính thức ở trên: lệnh "khởi động lại" đúng chuẩn phải có **CS ra lệnh = CS hoàn thành = 435,7** (cả hai cùng giá trị) thì mới được coi là hoàn thành và bắt đầu tính lại. Nhưng lệnh "Hoà lưới" thực tế trong dữ liệu ngày 07/07 chỉ có `CS ra lệnh=435,7`, còn `CS hoàn thành=0` — **không khớp mẫu 435,7-435,7** mà quy tắc yêu cầu. Đây là một khoảng trống trong **cách ghi dữ liệu lúc xảy ra sự cố** (dữ liệu lịch sử cụ thể của ngày này), không phải trong bản thân quy tắc hay công thức.

→ Kết luận: người tính tay cho ngày 07/07 đã dùng thêm hiểu biết thực tế về sự cố (biết chắc tổ máy đã ngừng) để điền Qdd=0 thủ công, thay vì thuần tuý chạy công thức trên đúng dữ liệu danh sách lệnh đã xuất — không phải điều công thức tự động có thể suy ra được từ dữ liệu này. Không có hành động sửa nào cần làm ở đây.

---

# Đợt 2 — kiểm chứng trên chính bản Google Sheets (26/07/2026)

Khác đợt 1 (chạy bằng bản tái hiện Python), đợt này đối chiếu **kết quả do chính Sheet + `QDD-Core-Library` tính ra** với bảng tính tay, trên 6 ngày: 17, 18, 19 và 23, 24, 25/07/2026 (tổ S1).

Nguồn đối chiếu: `test-data/Kiểm tra Qdu ngày 200726.xlsx` (ngày 17-19) và `Kiểm tra Qdu ngày 260726.xlsx` (ngày 23-25).

## Kết quả

| Ngày | Lệch Qdd tối đa | Số chu kỳ lệch > 0,001 MW | Ghi chú |
|---|---|---|---|
| 17/07 | 1,2215 MW | 6/48 | 1 ô bảng tính tay sai, xem bên dưới |
| 18/07 | 0,0538 MW | 5/48 | làm tròn |
| 19/07 | 1,0114 MW | 8/48 | 1 ô bảng tính tay sai, xem bên dưới |
| 23/07 | 0,0359 MW | 5/48 | làm tròn |
| 24/07 | 0,1171 MW | 12/48 | làm tròn (ngày nhận carry-over) |
| 25/07 | 0,0000 MW | 0/48 | ngày không có lệnh nào |

**Qdư**: bằng 0 ở toàn bộ 48 chu kỳ của cả 6 ngày, ở cả hệ thống lẫn bảng tính tay — khớp tuyệt đối.

## R07 (ramp vắt qua nửa đêm) — lần đầu kiểm chứng bằng dữ liệu thật

Trước đợt này, R07 mới chỉ được chứng minh bằng test tổng hợp vì chưa có ngày vận hành nào rơi đúng tình huống. Ngày **23→24/07/2026** là ngày đầu tiên có:

- Lệnh `G14001.2026.3037` lúc **23:50:46** ngày 23 (MO, bị dừng → công suất hiệu lực = CS hoàn thành = 534 MW), trong khi tổ máy đang ở 622,5 MW.
- Lúc 24:00 ramp **chưa xong**: còn **590,1833 MW**, cần thêm **963 giây** nữa mới xuống 534 MW.
- Ngày 24 nhận `carryTarget = 534` và chạy tiếp phần ramp còn lại trước khi nhận lệnh mới lúc 00:17:55.

Kiểm chứng bằng cách chạy lại ngày 24 **có** và **không có** carry:

| Chu kỳ | Có carry (R07) | Không carry | Bảng tính tay |
|---|---|---|---|
| 01 | **557,55** | 598,23 | **557,62** |
| 02 | **599,32** | 609,49 | **599,21** |

Không có R07 thì chu kỳ 1 sai khoảng **40 MW**. → **R07 đã được xác nhận bằng dữ liệu vận hành thật**, không còn là tính năng chỉ có test bao phủ.

## Ba ô sai trong bảng tính tay (hệ thống đúng)

Ba chu kỳ lệch ~1 MW đều nằm **ngay sau một lệnh bị Dừng**:

| Ngày | Chu kỳ | Tính tay | Hệ thống | Tính tay độc lập bằng giấy bút |
|---|---|---|---|---|
| 17/07 | 40 (19:30-20:00) | 479,036 | **477,814** | **477,815** |
| 19/07 | 42 (20:30-21:00) | 436,568 | **437,579** | **437,58** |

Cách kiểm chứng chu kỳ 40 ngày 17/07: giữ 480,4 MW đến 19:53:17 (1397 s), giảm về 460 MW mất 349,7 s ở tốc độ 3,5 MW/phút, còn lại 53,3 s ở 460 MW → trung bình **477,815 MW**. Hệ thống ra 477,8144 — lệch 0,0006 MW.

→ Kết luận: **hệ thống đúng theo đúng thuật toán đã tài liệu hoá; ba ô đó trong bảng tính tay sai.** Người tính tay có vẻ áp thời điểm dừng muộn hơn thực tế.

> Điều này cũng đính chính kết luận "17/07 và 18/07 khớp tuyệt đối 0,0000 MW" ghi ở phần trên của tài liệu này: mức khớp thực tế của ngày 17/07 là **0,0075 MW ở các chu kỳ bình thường và 1,22 MW ở một chu kỳ mà bảng tính tay sai**.

## Quy tắc nghiệp vụ được xác nhận lại trong đợt này

Người phụ trách nghiệp vụ xác nhận: *"Ở cột lý do dừng/hủy có ghi chú thì cột công suất ra lệnh bằng CS hoàn thành."* Đây chính là **R03** (và R01 cho lệnh SO) đã cài sẵn trong `CommandFilter.selectEffective` — kiểm tra trên 4 lệnh bị dừng của các ngày 17-19/07 đều áp đúng, không cần sửa gì.

## Việc cần làm tiếp

- [x] Điều tra và xác nhận với người phụ trách nghiệp vụ về quy tắc lệnh 0-0 và khởi động lại — **kết luận: công thức gốc đúng, không cần sửa** (UAT-32 đóng, xem cập nhật ở [09_Test_Cases.md](09_Test_Cases.md)).
- [x] Kiểm tra tình huống ramp qua 00:00 (UAT-04) — **đã kiểm chứng bằng dữ liệu thật ngày 23→24/07**, xem Đợt 2 ở trên.
- [ ] Ghi rõ quy tắc "khởi động lại cần CS ra lệnh = CS hoàn thành cùng giá trị" vào quy trình nhập liệu vận hành, để dữ liệu tương lai đủ điều kiện tính tự động (không cần can thiệp tay như ngày 07/07).
- [ ] Mở rộng kiểm tra sang các ngày còn thiếu danh sách lệnh (04, 05, 08, 09/07) khi có dữ liệu.
- [ ] Kiểm tra tổ **S2** trên bản Google Sheets — toàn bộ đợt 2 mới chỉ chạy S1.
- [ ] Báo lại cho người phụ trách bảng tính tay về 3 ô sai ở trên, để bản đối chiếu gốc được sửa.
