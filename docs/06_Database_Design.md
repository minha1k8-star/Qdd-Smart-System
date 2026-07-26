# Database Design — QDD Smart System

"Cơ sở dữ liệu" của hệ thống là các sheet trong Google Sheets của từng nhà máy. Mỗi nhà máy có một file riêng; toàn bộ sheet dưới đây do menu **QDD Smart System → Thiết lập sheet** tạo ra (`src/NhaMay-Mau-Template/SheetSetup.js`).

**Nguyên tắc thiết kế**: mỗi sheet lưu **nhiều ngày trong cùng một bảng** (có cột `Ngày`), không phải mỗi ngày một sheet. Nhờ đó tính được nhiều ngày cùng lúc và báo cáo tháng đọc thẳng dữ liệu gốc, không cần "đóng băng" kết quả từng ngày.

> **Quy tắc bất di bất dịch**: mọi chỗ đọc dữ liệu đều **dò cột theo TÊN tiêu đề**, không theo vị trí cột. Người dùng có thể dán file gốc với thứ tự cột khác, hoặc sheet còn cấu trúc của bản cũ. Đọc theo vị trí đã từng làm lệch cột và loại sạch lệnh mà không báo lỗi.

---

## `HUONG_DAN`

Hướng dẫn sử dụng đầy đủ, luôn nằm ngoài cùng bên trái. **Do hệ thống ghi lại mỗi lần chạy "Thiết lập sheet"** — không sửa tay, sửa sẽ mất khi thiết lập lại. Nội dung nguồn ở `HuongDan.js`.

## `CAI_DAT` — Cấu hình nhà máy

Hai cột: nhãn (A) và giá trị (B). Đọc theo **nhãn**, không theo số dòng.

**Cấu hình chung:**

| Nhãn | Mặc định | Dùng cho |
|---|---|---|
| Tên nhà máy | | |
| Tốc độ ramp (MW/phút) | `3.5` | R05, R06 — tính thời lượng ramp |
| Hệ số Qdd_V | `0.9188` | R09 |
| Dung sai (+-) | `0.03` | R13 — dải ±3% |

**Cấu hình theo từng tổ máy** — 3 dòng cho mỗi tổ:

| Nhãn | Mặc định (Duyên Hải 1) | Dùng cho |
|---|---|---|
| Mã công tơ Qdc - S1 | `001` | R10 — nhận diện file CSV theo tên |
| Mã công tơ Qmp - S1 | `303` | R12 |
| Nhãn báo cáo - S1 | `S1DH1` | tiêu đề khối trong file xuất |
| Mã công tơ Qdc - S2 | `002` | |
| Mã công tơ Qmp - S2 | `301` | |
| Nhãn báo cáo - S2 | `S2DH1` | |

> **Danh sách tổ máy được suy ra từ chính các dòng này** (`getConfiguredUnits_` quét các nhãn dạng `Mã công tơ Qdc - <tổ máy>`). Nhà máy có **3 tổ máy trở lên** chỉ cần **thêm 3 dòng** cho mỗi tổ (`Mã công tơ Qdc - S3`, `Mã công tơ Qmp - S3`, `Nhãn báo cáo - S3`) — sidebar và báo cáo tự hiện đủ, **không phải sửa một dòng code nào**. Thứ tự trong `CAI_DAT` quyết định thứ tự hiển thị.

> **Ô mã công tơ phải ở định dạng "Văn bản thuần"** — Google Sheets tự hiểu `001` là số 1 rồi cắt mất hai số 0 đầu. "Thiết lập sheet" tự đặt định dạng này. Nếu vẫn thấy hiển thị `1`, gõ lại `001` là được (hoặc ghi `csv001` — hệ thống cũng hiểu). Việc so khớp chịu được cả trường hợp mất số 0, nên số liệu không sai, chỉ là nhìn khó hiểu.

> **MÃ CÔNG TƠ KHÔNG KÈM CHỮ SỐ NĂM.** Tên file CSV có dạng `<ngày><tháng><năm 1 chữ số><mã công tơ>` — `17076001.CSV` = ngày 17, tháng 07, **năm 2026 (số 6)**, công tơ **001**. Điền `6001` thì sang 2027 tên file thành `17077001.CSV` và **hệ thống không nhận ra file nào**. Chạy "Thiết lập sheet" sẽ tự bỏ chữ số năm khỏi mã 4 chữ số và sửa luôn các dòng `CSV_DATA` cũ theo, có báo rõ đã đổi những gì.

**Mã công tơ và nhãn báo cáo khác nhau giữa các nhà máy** — bắt buộc sửa khi triển khai cho nhà máy mới, giá trị mặc định chỉ đúng cho Duyên Hải 1.

`CAI_DAT` **không bị xoá** khi dọn dữ liệu.

## `LENH` — Danh sách lệnh điều độ

**25 cột, đúng thứ tự và tên như file gốc `DanhSachLenhKetThuc`** để có thể dán thẳng file gốc vào từ dòng 2 mà không lệch cột. Thông thường không cần dán tay — sidebar mục 2 nhập trực tiếp từ file Excel.

| # | Cột | Vai trò trong thuật toán |
|---|---|---|
| 1 | ID Lệnh | khoá gộp khi nhập lại (trùng ID → cập nhật, không nhân đôi) |
| 2 | Nhà máy | |
| 3 | Tổ máy | lọc S1/S2 |
| 4 | Nội dung lệnh | |
| 5 | CS ra lệnh (MW) | P hiệu lực khi MO bình thường (R02) |
| 6 | CS hoàn thành (MW) | P hiệu lực khi SO (R01) hoặc MO bị dừng (R03) |
| 7 | Thời điểm BĐTH | **mốc bắt đầu ramp (R04)** — không dùng "Thời điểm hoàn thành" |
| 8 | Thời điểm hoàn thành | tham khảo |
| 9–15 | Người ra lệnh, Người thực hiện, AGC, Nhiên liệu, Lý do lệnh, Ghi chú ra lệnh, Ghi chú hoàn thành | tham khảo nghiệp vụ |
| 16 | Hoàn thành | `1` = đã hoàn thành; điều kiện bắt buộc để lệnh có hiệu lực |
| 17 | Dừng lệnh | `TRUE` → dùng CS hoàn thành (R03) |
| 18–22 | Thời điểm dừng, Lý do dừng, Người dừng, Lý do hủy, Người hủy | tham khảo |
| 23–24 | Lệnh cụm, Lệnh nhập lại | tham khảo |
| 25 | Nguồn lệnh | `SO` hoặc `MO` — quyết định quy tắc chọn công suất |

Thuật toán chỉ dùng **9 trường** (3, 5, 6, 7, 16, 17, 25 và ID, Nhà máy); các cột còn lại giữ nguyên để đối chiếu nghiệp vụ.

Sheet tự sắp xếp theo **Thời điểm BĐTH** tăng dần sau mỗi lần nhập.

## `CSV_DATA` — Dữ liệu công tơ

| Cột | Nội dung |
|---|---|
| 1 | Ngày |
| 2 | Mã công tơ (vd `001`) — không kèm chữ số năm |
| 3–50 | Chu kỳ 1 … Chu kỳ 48 — giá trị `KwhGiao` gốc, **chưa chia 1000** |

Mỗi dòng = một cặp (ngày, mã công tơ). Ghi lại cùng cặp sẽ **ghi đè**, không nhân đôi. Tự sắp xếp theo Ngày rồi Mã công tơ.

Việc quy đổi `/1000` và `×2` (ra P_Qdc) thực hiện lúc tính, không lưu sẵn — xem [04_Algorithm_Specification.md](04_Algorithm_Specification.md) mục 5.

## `P0_NGAY` — Công suất đầu ngày

| Cột | Nội dung |
|---|---|
| 1 | Ngày |
| 2 | Tổ máy |
| 3 | P0 (MW) — công suất tại **đúng 00:00** của ngày đó |
| 4 | Ghi chú — `Tự động từ cuối ngày dd/MM/yyyy` nếu do hệ thống ghi |
| 5 | Ramp tiếp đến (MW) — mục tiêu ramp còn dở dang từ ngày trước (**R07**) |

Sau mỗi lần tính, hệ thống **tự ghi dòng cho ngày kế tiếp**. Dòng do người dùng nhập tay (ghi chú không bắt đầu bằng "Tự động") **không bị ghi đè** — chỉ ngày đầu tiên dùng hệ thống mới cần nhập tay.

> P0 là công suất **tại 24:00** của ngày trước, **không phải** Qdd chu kỳ 48 (là giá trị trung bình 23:30–24:00). Nhầm hai giá trị này làm sai lan ra cả ngày.

Dòng P0 của ngày kế tiếp **được giữ lại** khi dọn dữ liệu cũ — nếu xoá thì ngày tiếp theo không tính được.

## `KET_QUA` — Kết quả tính

Mỗi (ngày, tổ máy) chiếm 48 dòng. Ngày mới nhất nằm trên cùng; trong cùng một ngày thì tổ máy và chu kỳ tăng dần.

| Cột | Nội dung |
|---|---|
| 1 | Ngày |
| 2 | Tổ máy |
| 3 | Chu kỳ — dạng `01 [00:00-00:30]` |
| 4 | Qdd (MW) |
| 5 | Qdd_V (MWh) |
| 6 | Qdc (MWh) |
| 7 | P_Qdc (MW) |
| 8 | Ngưỡng dưới |
| 9 | Ngưỡng trên |
| 10 | Qmp (MWh) |
| 11 | Qdư (MWh) |
| 12 | Dấu hiệu — `âm` / `dương` / `trong ±3%` |

Cột 4–11 hiển thị **2 số thập phân**; giá trị thật giữ đầy đủ độ chính xác vì P0 ngày kế tiếp và tổng tháng đọc lại chính các ô này.

Tính lại cùng một (ngày, tổ máy) sẽ **thay thế** 48 dòng cũ, không nhân đôi.

## `BAO_CAO_THANG` — Tổng hợp tháng

| Cột | Nội dung |
|---|---|
| 1 | Ngày |
| 2 | Tổ máy |
| 3 | Tổng Qdc (MWh) |
| 4 | Tổng Qmp (MWh) |
| 5 | Tổng Qdd_V (MWh) |
| 6 | Tổng Qdư (MWh) |

Ghi đè toàn bộ mỗi lần chạy báo cáo tháng — đây là bảng dẫn xuất, không phải dữ liệu gốc.

---

## Dữ liệu nào tái tạo được, dữ liệu nào không

| Sheet | Mất thì sao |
|---|---|
| `CAI_DAT` | Phải cấu hình lại tay — **cẩn thận nhất** |
| `LENH`, `CSV_DATA` | Nhập lại từ file gốc trên máy |
| `P0_NGAY` | Tính lại từ ngày liền trước; nếu mất cả chuỗi thì phải nhập tay P0 ngày đầu |
| `KET_QUA` | Tính lại được nếu còn lệnh + CSV |
| `BAO_CAO_THANG` | Chạy lại báo cáo tháng |

Google Sheets có sẵn `Tệp → Lịch sử phiên bản` để khôi phục khi lỡ xoá.
