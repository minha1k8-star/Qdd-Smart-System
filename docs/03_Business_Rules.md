# Business Rules — QDD Smart System

Nguồn gốc: sheet `QUY_TAC_NGHIEP_VU` và `HUONG_DAN` trong workbook chính thức (`legacy/CongCu_Tinh_Qdd_Qdu_v1_3_0_AllInOne.xlsm`). Đây là **nguồn sự thật gốc** — nếu tài liệu này và workbook lệch nhau, workbook thắng, phải cập nhật lại tài liệu.

> Không thay đổi bất kỳ quy tắc nào dưới đây nếu chưa có Test Case tương ứng trong [09_Test_Cases.md](09_Test_Cases.md). Xem quy trình bắt buộc ở [AGENTS.md](../AGENTS.md).

## Danh mục quy tắc (R01–R14)

| Mã | Nhóm | Quy tắc | Đầu vào | Kết quả | Ghi chú |
|----|------|---------|---------|---------|---------|
| R01 | Lệnh SO | Chỉ nhận lệnh hoàn thành và CS hoàn thành dương. | `P=1, F>0, Y=SO` | P hiệu lực = F | |
| R02 | Lệnh MO | MO bình thường dùng CS ra lệnh. | `P=1, E>0, Y=MO, Q<>TRUE` | P hiệu lực = E | |
| R03 | MO dừng | MO dừng sớm dùng CS hoàn thành thực tế. | `P=1, Q=TRUE, F>0` | P hiệu lực = F | |
| R04 | Thời điểm | Ramp bắt đầu tại thời điểm BĐTH. | Cột G | Giây trong ngày | **Không dùng cột H (thời điểm hoàn thành) làm điểm bắt đầu** |
| R05 | Tốc độ | Tốc độ thay đổi công suất mặc định. | `CAI_DAT!B7` | 3,5 MW/phút | Có thể cấu hình |
| R06 | Ngắt ramp | Lệnh mới ngắt ramp cũ tại công suất nội suy. | Hai lệnh chồng thời gian | Ramp mới bắt đầu từ P hiện tại (nội suy tại thời điểm lệnh mới) | |
| R07 | Qua ngày | Lưu P tại 24:00 và phần ramp còn lại. | `TRANG_THAI_CONG_SUAT` | Tạo `AUTO_CARRY` cho ngày sau | Theo từng S1/S2 độc lập |
| R08 | Qdd | Trung bình diện tích công suất trong 1.800 giây (mỗi chu kỳ 30 phút). | Đường P(t) | MW trung bình | |
| R09 | Qdd_V | Quy đổi Qdd sang Qdd_V. | Qdd, hệ số 0,9188 | `Qdd_V = Qdd / 2 × 0,9188` | |
| R10 | Qdc | Lấy KwhGiao công tơ 6001 chia 1.000. | 48 giá trị CSV 6001 | MWh | |
| R11 | P_Qdc | Quy đổi Qdc về MW chu kỳ 30 phút. | Qdc | `P_Qdc = Qdc × 2` | |
| R12 | Qmp | Lấy KwhGiao công tơ 6303 chia 1.000. | 48 giá trị CSV 6303 | MWh | |
| R13 | Dải ±3% | So sánh P_Qdc với Qdd±3%. | Qdd, P_Qdc | Trong dải / ngoài dải | |
| R14 | Qdư | Trong dải: 0. Ngoài dải: `Qmp − Qdd_V`. | Kết quả R13, Qmp, Qdd_V | MWh | |

## Diễn giải chi tiết cho các quy tắc dễ nhầm

### R01–R03 — Chọn công suất hiệu lực theo loại lệnh

Danh sách lệnh gốc (`LENH_GOC`) có nhiều loại lệnh (`Y`), công suất ra lệnh (`E`) khác công suất hoàn thành (`F`), và có thể bị dừng sớm (`Q=TRUE`). Quy tắc chọn "công suất hiệu lực" (P hiệu lực) dùng để dựng ramp phụ thuộc vào 3 trường hợp loại trừ lẫn nhau:

- **SO hợp lệ** (lệnh đóng máy đã hoàn thành, `F>0`) → dùng `F` (công suất hoàn thành).
- **MO bình thường** (không bị dừng sớm) → dùng `E` (công suất ra lệnh).
- **MO bị dừng sớm** (`Q=TRUE`, có `F>0`) → dùng `F` (công suất hoàn thành thực tế), không dùng `E` vì lệnh chưa đạt mục tiêu ban đầu.

Lệnh nào không thoả điều kiện `P=1` (hợp lệ) bị loại khỏi `LENH_DIEU_DO` — không tham gia tính toán.

> **Khoảng trống đã phát hiện qua dữ liệu thực tế (07/2026)**: điều kiện `E>0`/`P>0` khiến một lệnh **dừng máy hợp lệ với công suất mục tiêu = 0** (rất tự nhiên với lệnh "Ngừng tổ máy") bị loại hoàn toàn khỏi tính toán, làm công suất bị giữ sai ở mức trước đó thay vì về 0. Xem chi tiết và đề xuất xử lý ở [15_Accuracy_Validation_2026-07.md](15_Accuracy_Validation_2026-07.md) và test case [UAT-32](09_Test_Cases.md).

### R04 — Vì sao không dùng thời điểm hoàn thành

Ramp phải bắt đầu tại **thời điểm bắt đầu thực hiện (BĐTH, cột G)**, không phải thời điểm hoàn thành (cột H). Đây từng là nguồn gốc lỗi lệch số liệu lớn trong kiểm thử (xem [14_Knowledge_Transfer.md](14_Knowledge_Transfer.md)) khi nhầm lẫn hai cột này.

### R06 — Ngắt ramp

Khi một lệnh mới có BĐTH rơi vào giữa một ramp đang chạy (do lệnh trước chưa đạt công suất mục tiêu), ramp cũ bị cắt tại đúng thời điểm BĐTH của lệnh mới. Công suất tại điểm cắt được tính bằng **nội suy tuyến tính** trên ramp cũ (không nhảy về công suất mục tiêu cũ hay công suất ra lệnh mới), rồi từ đó bắt đầu ramp mới hướng tới mục tiêu của lệnh mới.

### R07 — Chuyển tiếp qua nửa đêm

Nếu ramp chưa hoàn tất khi hết ngày (24:00), hệ thống lưu công suất P tại 24:00 vào sheet `TRANG_THAI_CONG_SUAT`, tách riêng theo từng tổ máy (S1/S2). Khi tính ngày kế tiếp, hệ thống tạo một đoạn `AUTO_CARRY` để ramp tiếp tục đúng theo tiến độ dở dang, thay vì reset về 0 hay bắt đầu lại.

### R08–R09 — Qdd và Qdd_V

Qdd của mỗi chu kỳ 30 phút (1.800 giây) là **trung bình diện tích** dưới đường cong công suất P(t) trong chu kỳ đó (xem `DIEN_TICH`, `DOAN_CONG_SUAT` — chi tiết thuật toán ở [04_Algorithm_Specification.md](04_Algorithm_Specification.md)). Qdd_V là giá trị quy đổi dùng hệ số hiệu suất 0,9188, đồng thời chia đôi để đưa MW trung bình 30 phút về đơn vị năng lượng nửa chu kỳ tương ứng với Qdc/Qmp.

### R10–R12 — Nguồn dữ liệu CSV

Qdc và Qmp **không đến từ tính toán**, mà đọc trực tiếp từ dòng `KwhGiao` trong 2 file CSV công tơ (6001 và 6303 tương ứng) — mỗi file có 48 giá trị số cho 48 chu kỳ. Cột ngày trong CSV bị **bỏ qua hoàn toàn** khi nhập (chỉ dùng số, không dùng ngày trong file làm căn cứ — ngày tính lấy từ `CAI_DAT`).

### R13–R14 — Xác định Qdư

Dải dung sai mặc định là **±3%** quanh Qdd (dung sai này cấu hình được ở `CAI_DAT`, mặc định 0,03). Nếu `P_Qdc` nằm trong dải `[Qdd×0,97, Qdd×1,03]` → Qdư = 0 (không có công suất dư). Nếu nằm ngoài dải → Qdư = `Qmp − Qdd_V`, có thể âm hoặc dương tuỳ chiều lệch.

## Quy tắc phạm vi áp dụng

Toàn bộ quy tắc trên áp dụng được cho **cả tổ máy S1 và S2**, với điều kiện cấu trúc danh sách lệnh, CSV 6001/6303 và quy tắc nghiệp vụ giống nhau giữa hai tổ. Lịch sử công suất (P) của mỗi tổ được lưu và xử lý độc lập (xem R07, UAT-15 ở [09_Test_Cases.md](09_Test_Cases.md)).

## Ghi chú

Sheet `HUONG_DAN` trong workbook có thêm một số mục tiêu đề (SO, MO thường, Ramp, P đầu ngày, Qdd, Qdd_V, Dải kiểm tra...) chưa được điền mô tả chi tiết trong bản v1.3.0 — nội dung tương ứng đã có đầy đủ trong bảng R01–R14 ở trên và không bị mất thông tin.
