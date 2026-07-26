# Algorithm Specification — QDD Smart System

Đặc tả kỹ thuật của "trái tim" hệ thống: pipeline tính Qdd/Qdư từ danh sách lệnh điều độ + CSV công tơ.

Đây là **hợp đồng kỹ thuật** của dự án: mã nguồn trong `src/QDD-Core-Library/` phải khớp với tài liệu này, và mọi thay đổi thuật toán phải sửa cả hai. Nếu tài liệu này mâu thuẫn với [03_Business_Rules.md](03_Business_Rules.md), tài liệu này chi tiết hơn (công thức cụ thể), còn 03 mô tả *vì sao* quy tắc như vậy.

Toàn bộ công thức dưới đây đã được **đối chiếu với bảng tính tay trên dữ liệu vận hành thật** — xem [15_Accuracy_Validation_2026-07.md](15_Accuracy_Validation_2026-07.md).

## Pipeline tổng quan

```
LENH (danh sách lệnh điều độ thô, nhiều ngày, nhiều tổ máy)
   │  CommandFilter.selectEffective — lọc theo ngày/tổ máy + chọn P hiệu lực (R01–R03)
   ▼
Lệnh hiệu lực (sắp theo giây trong ngày tăng dần)
   │  RampEngine — nội suy liên tục giữa các lệnh (R06)
   ▼
Bảng ramp (mỗi lệnh: P bắt đầu / P mục tiêu / thời lượng ramp / thời điểm kết thúc)
   │  Segments.build — tách mỗi lệnh thành đoạn RAMP + đoạn HOLD, cắt tại 24:00
   ▼
Đoạn công suất phủ kín 24h (P(t) tuyến tính từng đoạn)
   │  AreaIntegration — tích phân hình thang, giao từng đoạn với từng chu kỳ 30 phút
   ▼
Qdd 48 chu kỳ (MW)
   │  QddCalculator — quy đổi Qdd_V, đọc Qdc/Qmp từ CSV, so dung sai
   ▼
Kết quả: Qdd, Qdd_V, Qdc, P_Qdc, Qmp, Qdư, ngưỡng ±dung sai, dấu hiệu
```

Chuyển tiếp qua nửa đêm (R07, mục 6) **không phải một engine riêng** — nó hoạt động bằng cách chèn **một lệnh nối ảo tại 00:00:00** vào đầu danh sách lệnh hiệu lực, để chính pipeline trên xử lý y hệt một lệnh thật.

---

## 1. Chọn lệnh hiệu lực và công suất hiệu lực (`CommandFilter.js`)

Một lệnh được coi là **hợp lệ** khi thoả đồng thời:

```
BĐTH là ngày-giờ hợp lệ
VÀ ngày của BĐTH = ngày đang tính
VÀ 2 ký tự đầu của Tổ máy = 2 ký tự đầu tổ máy đang tính
VÀ Hoàn thành = 1
VÀ ( (Nguồn lệnh = "SO" VÀ CS hoàn thành > 0)
     HOẶC (Nguồn lệnh = "MO" VÀ CS ra lệnh > 0) )
```

**Công suất hiệu lực** — chính là R01–R03:

```
P_hiệu_lực = CS hoàn thành   nếu Nguồn lệnh = "SO"                              (R01)
             CS hoàn thành   nếu "MO" VÀ Dừng lệnh = TRUE VÀ CS hoàn thành > 0  (R03)
             CS ra lệnh      nếu "MO" bình thường                               (R02)
```

Điều kiện `> 0` cho cả SO và MO là **chủ đích, đã xác nhận nghiệp vụ** — lệnh "0-0" (CS ra lệnh = CS hoàn thành = 0, thường là trip/ngừng sự cố) **không được tính**. Không nới thành `>= 0`. Xem [15_Accuracy_Validation_2026-07.md](15_Accuracy_Validation_2026-07.md).

Lệnh hiệu lực được sắp theo **giây trong ngày** của BĐTH:
`seconds = giờ × 3600 + phút × 60 + giây`.

> **Cạm bẫy đã gặp**: không được kiểm tra kiểu ngày bằng `instanceof Date`. Khi Sheet gọi sang Apps Script Library, mỗi scope có constructor `Date` riêng nên `instanceof` luôn sai, làm **mọi lệnh bị loại âm thầm**. Dùng duck typing (`typeof v.getTime === 'function'`).

## 2. Ramp Engine — nội suy liên tục giữa các lệnh (`RampEngine.js`)

Công thức cốt lõi nhất của hệ thống. Với lệnh hiệu lực thứ `n` (n = 1, 2, 3…):

| Đại lượng | Công thức | Ý nghĩa |
|---|---|---|
| `B_n` (bắt đầu, giây) | giây trong ngày của BĐTH | |
| `D_n` (P mục tiêu) | công suất hiệu lực của lệnh (mục 1) | |
| **`F_n` (P bắt đầu)** | `n = 1`: `P0`.<br>`n ≥ 2`: `IF(B_n >= I_{n-1}, D_{n-1}, F_{n-1} + (D_{n-1} − F_{n-1}) × (B_n − B_{n-1}) / MAX(I_{n-1} − B_{n-1}, EPS))` | **Đây là công thức ngắt ramp (R06)** |
| `H_n` (thời lượng, giây) | `0` nếu `D_n = F_n`, ngược lại `ABS(D_n − F_n) / tốc_độ_ramp × 60` | Quãng đường công suất ÷ tốc độ (MW/phút) × 60 (R05) |
| `I_n` (kết thúc, giây) | `B_n + H_n` | Thời điểm ramp hoàn tất |

**Giải thích `F_n`** — điểm mấu chốt của cả thuật toán:

- Nếu lệnh mới bắt đầu **sau khi** ramp trước đã hoàn tất (`B_n ≥ I_{n-1}`): điểm xuất phát = mục tiêu cũ `D_{n-1}`.
- Nếu lệnh mới đến **giữa lúc ramp trước đang chạy**: nội suy tuyến tính công suất thực tế tại thời điểm `B_n` trên đường ramp trước — **không** nhảy về mục tiêu cũ, cũng **không** lấy công suất ra lệnh mới.

**Ramp không bị giới hạn trong phạm vi 1 ngày**: `I_n` có thể vượt quá 86400 giây (ramp chưa xong lúc nửa đêm). Việc cắt theo ranh giới ngày là trách nhiệm của tầng dựng đoạn (mục 3) và carry-over (mục 6), **không phải** của Ramp Engine. Giữ đúng ranh giới trách nhiệm này khi sửa code.

## 3. Dựng đoạn công suất phủ kín 24h (`Segments.js`)

Mỗi lệnh sinh ra **2 đoạn**:

- **Đoạn RAMP**: từ `B_n` đến `min(I_n, B_{n+1}, 86400)`, công suất đi từ `F_n` hướng tới `D_n`.
- **Đoạn HOLD**: từ `I_n` đến `B_{n+1}` (hoặc đến `86400` nếu là lệnh cuối ngày), giữ nguyên `D_n`.

Thêm **đoạn đầu ngày**: HOLD từ giây `0` đến `B_1` (hoặc đến `86400` nếu ngày không có lệnh nào), giữ nguyên `P0`.

> **Lỗi đã gặp, phải giữ đúng**: khi một đoạn RAMP **bị cắt giữa chừng** (lệnh mới đến, hoặc hết ngày), công suất cuối đoạn phải là **giá trị nội suy tại điểm cắt**, KHÔNG phải mục tiêu `D_n`. Ghi sai chỗ này làm sai độ dốc đoạn → sai diện tích → sai Qdd.

Không có giới hạn số lệnh/số đoạn. Đã chạy thử 120 lệnh trong một ngày, ra đủ 48 chu kỳ hợp lệ.

## 4. Tích phân diện tích theo 48 chu kỳ (`AreaIntegration.js`)

Mỗi chu kỳ `i` là khoảng `[1800×(i−1), 1800×i]` giây. Với mỗi đoạn công suất giao với chu kỳ đó:

```
overlap_start = MAX(chu_kỳ_bắt_đầu, đoạn_bắt_đầu)
overlap_end   = MIN(chu_kỳ_kết_thúc, đoạn_kết_thúc)

nếu overlap_end <= overlap_start  →  diện tích = 0

ngược lại:
  P_tại(overlap_start) = nội suy tuyến tính trên đoạn (P_đầu → P_cuối) tại overlap_start
  P_tại(overlap_end)   = nội suy tuyến tính tương tự tại overlap_end
  diện tích = (P_tại(overlap_start) + P_tại(overlap_end)) / 2 × (overlap_end − overlap_start)
```

Đây là **tích phân hình thang** (trapezoidal rule) — chính xác tuyệt đối với đoạn tuyến tính (ramp) hoặc hằng số (hold), không cần chia nhỏ hơn.

```
Qdd(chu kỳ i) = Tổng diện tích (MW·s) của chu kỳ i / 1800        (R08)
```

## 5. Kết quả cuối cho từng chu kỳ (`QddCalculator.js`)

| Đại lượng | Công thức | Quy tắc |
|---|---|---|
| Qdd (MW) | kết quả mục 4 | R08 |
| Qdd_V (MWh) | `Qdd / 2 × hệ_số_Qdd_V` (mặc định `0,9188`) | R09 |
| Qdc (MWh) | giá trị `KwhGiao` của chu kỳ trong CSV công tơ Qdc `/ 1000` | R10 |
| P_Qdc (MW) | `Qdc × 2` | R11 |
| Qmp (MWh) | giá trị `KwhGiao` của chu kỳ trong CSV công tơ Qmp `/ 1000` | R12 |
| Ngưỡng dưới / trên | `Qdd × (1 − dung_sai)` / `Qdd × (1 + dung_sai)` (mặc định `0,03`) | R13 |
| **Qdư (MWh)** | `IF(P_Qdc < ngưỡng_dưới HOẶC P_Qdc > ngưỡng_trên, Qmp − Qdd_V, 0)` | R13 + R14 |
| Dấu hiệu | `"âm"` nếu `P_Qdc < ngưỡng dưới`; `"dương"` nếu `P_Qdc > ngưỡng trên`; `"trong ±3%"` nếu nằm trong dải | |

## 6. Chuyển tiếp qua nửa đêm — R07 (`Segments.endPowerOfDay`, `carryOverOf`)

**Cuối ngày N**, sau khi tính xong:

```
nếu I_cuối (giây kết thúc ramp của lệnh cuối) <= 86400 + EPS:
    endPower = công suất tại 24:00 theo đoạn cuối cùng   (ramp đã xong, không carry)
ngược lại (ramp chưa xong lúc 24:00):
    endPower  = F_cuối + hướng × tốc_độ_ramp × (86400 − B_cuối)/60   [kẹp không vượt quá D_cuối]
    carry     = { target: D_cuối, remainingSeconds: I_cuối − 86400 }
```

`endPower` được ghi làm **P0 của ngày N+1**, `carry.target` ghi vào cột `Ramp tiếp đến (MW)`. Lưu riêng theo từng cặp (ngày, tổ máy) — S1 và S2 độc lập.

> **Lỗi đã gặp, tuyệt đối không lặp lại**: P0 của ngày sau là **công suất TẠI đúng 24:00**, KHÔNG phải Qdd của chu kỳ 48. Qdd chu kỳ 48 là công suất *trung bình* khoảng 23:30–24:00; nếu lúc đó đang ramp thì hai giá trị khác nhau, và sai ở điểm khởi đầu sẽ lan ra cả ngày hôm sau.

**Đầu ngày N+1**: nếu có `carryTarget`, chèn một **lệnh nối ảo** vào đầu danh sách lệnh hiệu lực:

```
{ id: "AUTO_CARRY", seconds: 0, p: carryTarget }
```

Không cần truyền thời gian còn lại — Ramp Engine tự tính đúng phần còn lại từ `(P0, mục tiêu, tốc độ ramp)`. Sau đó pipeline mục 1–5 xử lý y hệt một lệnh thật.

Đã kiểm chứng bằng dữ liệu vận hành thật (23→24/07/2026): thiếu bước này thì chu kỳ 1 của ngày sau sai khoảng **40 MW**.

## 7. Đọc CSV công tơ (`CsvParser.js`)

- Dòng dữ liệu cần lấy là dòng có nhãn `KwhGiao` (so khớp không phân biệt hoa/thường, đã trim).
- 48 giá trị của dòng đó là 48 chu kỳ, phải là **số** — nếu không, báo lỗi trước khi ghi vào sheet.
- **Ngày được đọc từ chính nội dung file** (cột đầu dòng dữ liệu), không bắt người dùng chọn tay. Hỗ trợ `dd-mm-yy(yy)`, `dd/mm/yy(yy)`, `yyyy-mm-dd`.
- **Tổ máy và loại dữ liệu (Qdc/Qmp) nhận diện từ TÊN file** bằng cách so với mã công tơ cấu hình trong `CAI_DAT` — tên file thật có dạng `<ngày><tháng><mã công tơ>.CSV`, ví dụ `17076001.CSV`.
- Apps Script đọc text không phụ thuộc Regional Settings của máy người dùng, nên không có lớp xử lý riêng cho từng hệ điều hành.

## 8. Báo cáo tháng (`MonthlyReport.js`)

Báo cáo tháng **đọc trực tiếp kết quả đã tính trong `KET_QUA`**, không cần cơ chế "đóng băng"/snapshot từng ngày: mỗi lần tính là một lời gọi độc lập nhận tham số (ngày, tổ máy), nên dữ liệu của mọi ngày trong tháng cùng tồn tại song song.

Tổng hợp theo từng cặp (ngày, tổ máy): tổng Qdc, tổng Qmp, tổng Qdd_V, tổng Qdư.

## 9. Hằng số và cấu hình

| Tên | Mặc định | Nguồn | Ghi chú |
|---|---|---|---|
| Tốc độ ramp | 3,5 MW/phút | `CAI_DAT` | Dùng ở mục 2 và mục 6 |
| Hệ số Qdd_V | 0,9188 | `CAI_DAT` | Mục 5 |
| Dung sai | 0,03 (±3%) | `CAI_DAT` | Mục 5 — **cấu hình được**, không hard-code |
| Số chu kỳ/ngày | 48 | `Config.PERIOD_COUNT` | |
| Giây/chu kỳ | 1800 | `Config.CYCLE_SECONDS` | |
| Giây/ngày | 86400 | `Config.SECONDS_PER_DAY` | |
| Epsilon so sánh số thực | 0,000001 | `Config.EPS` | Nội suy (mục 2) và carry-over (mục 6) |

Tốc độ ramp, hệ số và dung sai **khác nhau giữa các nhà máy** — luôn đọc từ `CAI_DAT` của từng Sheet, không dùng hằng số cứng trong thư viện.

## 10. Lưu ý khi sửa thuật toán

- Ranh giới trách nhiệm: Ramp Engine tính "nếu không bị gì cản thì bao giờ ramp xong"; việc cắt theo ngày thuộc tầng dựng đoạn và carry-over. Trộn lẫn hai việc này là nguồn của nhiều lỗi tinh vi.
- Mọi thay đổi ở đây phải có test tương ứng trong `src/QDD-Core-Library/tests/run_tests.js` và cập nhật [09_Test_Cases.md](09_Test_Cases.md) — xem [AGENTS.md](../AGENTS.md).
- Sau khi sửa thư viện phải **tạo version mới** (`npx clasp version`) và cập nhật số version trong `src/NhaMay-Mau-Template/appsscript.json`, nếu không Sheet vẫn chạy code cũ.
