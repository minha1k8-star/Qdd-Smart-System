# Test Cases — QDD Smart System

Hai lớp kiểm thử:

1. **Test tự động** (`src/QDD-Core-Library/tests/run_tests.js`) — chạy bằng Node, không cần Google, không cần dữ liệu thật. Khoá lại từng lỗi đã gặp. Chạy: `node tests/run_tests.js`. Hiện **54/54 pass**.
2. **Đối chiếu dữ liệu thật** — so kết quả hệ thống với bảng tính tay độc lập trên dữ liệu vận hành. Kết quả ghi ở [15_Accuracy_Validation_2026-07.md](15_Accuracy_Validation_2026-07.md).

> Theo [AGENTS.md](../AGENTS.md): thay đổi ở [03_Business_Rules.md](03_Business_Rules.md) hoặc [04_Algorithm_Specification.md](04_Algorithm_Specification.md) phải có test case tương ứng ở đây.

## Trạng thái

| ID | Tình huống | Cách kiểm | Trạng thái |
|---|---|---|---|
| TC-01 | Lọc đúng ngày và tổ máy từ danh sách nhiều ngày | Test tự động + dữ liệu thật (file 63 lệnh nhiều ngày) | ✅ Đạt |
| TC-02 | SO hợp lệ → P hiệu lực = CS hoàn thành (R01) | Test tự động + lệnh SO ngày 19/07 | ✅ Đạt |
| TC-03 | MO bình thường → P hiệu lực = CS ra lệnh (R02) | Test tự động + dữ liệu thật | ✅ Đạt |
| TC-04 | MO bị dừng → P hiệu lực = CS hoàn thành (R03) | Test tự động + lệnh 2983, 3035, 3037 | ✅ Đạt |
| TC-05 | Lệnh 0-0 bị loại khỏi tính toán | Test tự động | ✅ Đạt |
| TC-06 | Cảnh báo khi có lệnh 0-0 trong ngày đang tính | Kiểm tay trên Sheet | ✅ Đạt |
| TC-07 | Lệnh mới ngắt ramp đang chạy → nội suy tại điểm cắt (R06) | Test tự động + dữ liệu thật | ✅ Đạt |
| TC-08 | Ramp bị cắt lúc 24:00 → công suất cuối là giá trị nội suy | Test tự động | ✅ Đạt |
| TC-09 | Tích phân hình thang 48 chu kỳ (R08) | Test tự động + đối chiếu bảng tính tay từng chu kỳ | ✅ Đạt |
| TC-10 | Qdd_V, Qdc, P_Qdc, Qmp, ngưỡng, Qdư (R09–R14) | Test tự động + dữ liệu thật | ✅ Đạt |
| TC-11 | Ngày không có lệnh nào → Qdd = P0 suốt 48 chu kỳ | Test tự động + ngày 25/07 | ✅ Đạt |
| TC-12 | Cảnh báo khi ngày tính không có lệnh hiệu lực nào | Kiểm tay trên Sheet | ✅ Đạt |
| TC-13 | P0 tự ghi cho ngày kế tiếp = công suất tại 24:00 | Test tự động + chuỗi 17→18→19/07 | ✅ Đạt |
| TC-14 | **Ramp vắt qua nửa đêm (R07)** — ngày sau chạy tiếp tới mục tiêu | Test tự động (9 test) + dữ liệu thật 23→24/07 | ✅ Đạt |
| TC-15 | Date truyền qua ranh giới Library vẫn nhận đúng | Test hồi quy tự động | ✅ Đạt |
| TC-16 | Đọc CSV: lấy đúng dòng KwhGiao, đủ 48 giá trị số | Test tự động | ✅ Đạt |
| TC-17 | Nhận diện tổ máy/loại dữ liệu theo tên file CSV | Kiểm tay + dữ liệu thật | ✅ Đạt |
| TC-18 | Đọc ngày từ nội dung file CSV | Test tự động | ✅ Đạt |
| TC-19 | Nhập lệnh từ file Excel: dò đúng sheet và dòng tiêu đề | Kiểm tay với file gốc (tiêu đề ở dòng 3, 3 sheet) | ✅ Đạt |
| TC-20 | Nhập lệnh từ file Excel: **không lệch múi giờ** | Đối chiếu khoảng BĐTH với file gốc | ✅ Đạt |
| TC-21 | Nhập lại cùng file → cập nhật theo ID, không nhân đôi | Kiểm tay (63 lệnh → 0 mới, 63 cập nhật) | ✅ Đạt |
| TC-22 | Đọc dữ liệu theo TÊN cột, không theo vị trí | Test tự động | ✅ Đạt |
| TC-23 | Tính nhiều ngày liên tiếp; ngày thiếu dữ liệu không chặn ngày khác | Kiểm tay + dữ liệu thật | ✅ Đạt |
| TC-24 | Nhiều lệnh trong ngày (31/45/60/80/120) | Test tay bằng Node | ✅ Đạt — đủ 48 chu kỳ hợp lệ, < 5 ms |
| TC-25 | Báo cáo tháng tổng hợp đúng theo (ngày, tổ máy) | Test tự động | ✅ Đạt |
| TC-26 | File xuất: mỗi ngày 1 tab, S1/S2 cạnh nhau, tổ trống vẫn giữ bảng | Kiểm tay, đối chiếu mẫu gốc | ✅ Đạt |
| TC-27 | Tên file xuất mang đúng kỳ báo cáo | Test tay hàm đặt tên (4 trường hợp) | ✅ Đạt |
| TC-28 | Dọn dữ liệu cũ giữ lại CAI_DAT và P0 ngày kế tiếp | Kiểm tay | ✅ Đạt |
| **TC-29** | **Tổ máy S2**: lọc đúng S2, P0 và lịch sử độc lập với S1 | Cần dữ liệu thật S2 + bảng tính tay | 🟡 **Chưa chạy** |
| **TC-30** | **Giữ P0 người dùng nhập tay**, không bị ghi đè bởi P0 tự động | Cần kiểm tay trên Sheet | 🟡 **Chưa chạy** |
| **TC-31** | **Khởi động lại sau sự cố** (CS ra lệnh = CS hoàn thành = tải thật) | Cần dữ liệu thật đúng mẫu | 🟡 **Chưa chạy** |
| **TC-32** | **R15 — Qdd chỉ tính từ lúc hoàn thành lệnh khởi động**, giai đoạn tăng tải từ 0 không tính | — | ⚪ **Quyết định KHÔNG triển khai** (26/07/2026), xử lý ở khâu vận hành |

## Các case còn lại

**TC-29 — tổ máy S2** là ưu tiên cao nhất. Toàn bộ đối chiếu dữ liệu thật đến nay mới chạy **S1**. Mã công tơ, P0 và chuỗi ngày của S2 độc lập hoàn toàn với S1, nên không suy ra được từ kết quả S1. Cần một vài ngày S2 có bảng tính tay để đối chiếu.

**TC-30 — giữ P0 nhập tay**: code đã xử lý (`saveNextDayP0_` bỏ qua dòng có ghi chú không bắt đầu bằng "Tự động"), nhưng chưa chạy thử thật trên Sheet.

**TC-31 — khởi động lại sau sự cố**: chờ có dữ liệu vận hành đúng mẫu "CS ra lệnh = CS hoàn thành = tải thật". Dữ liệu 07/07/2026 không dùng được vì chỉ có CS ra lệnh, thiếu CS hoàn thành.

**TC-32 — R15**: **quyết định KHÔNG triển khai bằng code** (26/07/2026) — trưởng ca biết thời điểm nào bắt đầu tính Qdd nên tự quyết định khi nào nhập dữ liệu và bấm tính. Lưu ý vận hành: riêng ngày khởi động lại, nếu bấm tính cả ngày thì giai đoạn tăng tải từ 0 vẫn được tính theo ramp thường (cao hơn thực tế) — ngày đó xử lý tay hoặc bắt đầu tính từ ngày kế tiếp. Xem [03_Business_Rules.md](03_Business_Rules.md#r15--khởi-động-lại-tổ-máy-sau-khi-ngừng).

## Cách chạy test tự động

```bash
cd src/QDD-Core-Library
node tests/run_tests.js
```

Bộ test không cần Google, không cần dữ liệu vận hành thật, chạy dưới 1 giây. **Có test đỏ thì không được đưa vào dùng** — mỗi test khoá lại một lỗi đã từng gây sai số liệu thật (xem [14_Knowledge_Transfer.md](14_Knowledge_Transfer.md)).
