# Algorithm Specification — QDD Smart System

Đây là đặc tả kỹ thuật của "trái tim" hệ thống: pipeline tính Qdd/Qdư từ danh sách lệnh hiệu lực + CSV công tơ.

**Nguồn của tài liệu này**: trích xuất trực tiếp từ **công thức Excel thật** trong `legacy/CongCu_Tinh_Qdd_Qdu_v1_3_0_AllInOne.xlsm` (sheet `XU_LY_LENH`, `DOAN_CONG_SUAT`, `DIEN_TICH`, `TINH_TOAN`) và từ phân tích code VBA (`legacy/modQdu_v1_3_1_AllInOne_FixCompile.bas`, phần chuyển tiếp qua nửa đêm và nhập CSV). Đây là tài liệu duy nhất trong repo dựa trên **công thức gốc**, không phải suy diễn từ mô tả nghiệp vụ — nếu có mâu thuẫn với [03_Business_Rules.md](03_Business_Rules.md), tài liệu này đúng hơn vì bám sát cell formula thật.

> **Phát hiện quan trọng**: toàn bộ thuật toán Ramp Engine / nội suy / tích phân diện tích **nằm ở công thức Excel**, không nằm trong code VBA. VBA chỉ điều phối (nhập liệu, kiểm tra, xuất báo cáo) và xử lý riêng phần chuyển tiếp qua nửa đêm. Đây là lý do người dùng nghiệp vụ có thể tự audit từng phép tính bằng cách click vào ô — đúng triết lý "không hộp đen" (xem [00_Project_Overview.md](00_Project_Overview.md)). Khi migrate sang Apps Script Library ([05_System_Architecture.md](05_System_Architecture.md)), toàn bộ logic mô tả dưới đây phải được viết lại thành code, vì Google Sheets formula tuy tương thích nhưng ta chủ động chuyển engine vào code có test.

## Pipeline tổng quan

```
LENH_GOC (lệnh gốc, có thể có AUTO_CARRY_ do VBA chèn từ hôm trước)
   │  công thức lọc + chọn P hiệu lực theo R01–R03
   ▼
LENH_DIEU_DO (lệnh hiệu lực, tối đa 60, sắp theo thời gian tăng dần)
   │  công thức Ramp Engine — nội suy liên tục giữa các lệnh (R06)
   ▼
XU_LY_LENH (mỗi dòng = 1 lệnh, có P bắt đầu/P mục tiêu/thời lượng ramp)
   │  công thức tách mỗi lệnh thành 2 đoạn: RAMP + HOLD
   ▼
DOAN_CONG_SUAT (121 đoạn công suất phủ kín 24h, P(t) từng đoạn)
   │  công thức tích phân hình thang, giao đoạn với từng chu kỳ 30 phút
   ▼
DIEN_TICH (diện tích MW·s mỗi đoạn × mỗi chu kỳ, tổng theo hàng)
   │  Qdd = Tổng MW.s / 1800
   ▼
TINH_TOAN (Qdd, Qdd_V, Qdc, P_Qdc, Qmp, Qdư, ngưỡng ±dung sai) → BAO_CAO_QDU
```

Chuyển tiếp qua nửa đêm (VBA, xem mục 6) không phải một engine riêng — nó hoạt động bằng cách **chèn một lệnh tổng hợp (`AUTO_CARRY_...`) vào đầu `LENH_GOC`**, để pipeline công thức ở trên xử lý y hệt một lệnh thật.

---

## 1. Chọn công suất hiệu lực (LENH_GOC → LENH_DIEU_DO, cột T/U)

Mỗi dòng `LENH_GOC` (qua vùng làm việc `LENH_DIEU_DO!J:X`) được đánh giá **hợp lệ** (cột T = 1) khi:

```
T = 1  nếu và chỉ nếu:
  N (BĐTH) là số hợp lệ
  VÀ INT(N) = CAI_DAT!B4 (đúng ngày đang tính)
  VÀ 2 ký tự đầu của Tổ máy (L) = 2 ký tự đầu CAI_DAT!B5 (đúng tổ máy)
  VÀ Q (Hoàn thành) = 1
  VÀ ( (M="SO" VÀ P>0) HOẶC (M="MO" VÀ O>0) )
```

Công suất hiệu lực (cột U) — chính là R01–R03:

```
U =  P (CS hoàn thành)   nếu M = "SO"
     P (CS hoàn thành)   nếu M = "MO" VÀ Dừng(R) = TRUE VÀ P > 0     (MO dừng sớm)
     O (CS ra lệnh)      nếu M = "MO" (bình thường)
```

Thứ tự lệnh (`W`) được đánh số tăng dần theo giây trong ngày (`V = MOD(N,1)*86400`) bằng `COUNTIFS`, xử lý cả trường hợp trùng giây (đếm thêm số dòng cùng giá trị đã xuất hiện trước nó). `LENH_DIEU_DO!A:H` là bảng kết quả cuối, sắp theo `W`, giới hạn hiển thị theo `CAI_DAT!B10` (số lệnh hiệu lực tối đa, ràng buộc bởi `MAX_EFFECTIVE_COMMANDS=60`).

## 2. Ramp Engine (sheet `XU_LY_LENH`) — nội suy liên tục giữa các lệnh

Đây là công thức cốt lõi nhất của toàn hệ thống. Mỗi dòng `n` (n=1..60) tương ứng lệnh hiệu lực thứ n theo `LENH_DIEU_DO`:

| Cột | Công thức | Ý nghĩa |
|---|---|---|
| B (Bắt đầu, giây) | `MOD(LENH_DIEU_DO!C, 1) * 86400` | Giây trong ngày của thời điểm BĐTH |
| D (P mục tiêu) | `LENH_DIEU_DO!D` | = U hiệu lực của lệnh này |
| **F (P bắt đầu)** | dòng 1: `CAI_DAT!$B$6` (P0). Dòng n≥2: `IF(B_n >= I_{n-1}, D_{n-1}, F_{n-1} + (D_{n-1}-F_{n-1}) * (B_n - B_{n-1}) / MAX(I_{n-1}-B_{n-1}, 0.000001))` | **Đây là công thức ngắt ramp (R06)**: nếu lệnh mới bắt đầu sau khi ramp trước đã hoàn tất (`B_n ≥ I_{n-1}`), điểm xuất phát = mục tiêu cũ. Nếu lệnh mới đến **giữa ramp trước**, nội suy tuyến tính công suất tại thời điểm `B_n` trên đường ramp trước (từ `F_{n-1}` đến `D_{n-1}`, trong khoảng `[B_{n-1}, I_{n-1}]`) |
| G (Loại) | `RAMP_UP` nếu D>F, `RAMP_DOWN` nếu D<F, `HOLD` nếu bằng | |
| H (Thời lượng, giây) | `IF(HOLD, 0, ABS(D-F) / CAI_DAT!$B$7 * 60)` | Quãng đường công suất ÷ tốc độ (MW/phút) × 60 = số giây cần để ramp xong (R05) |
| I (Kết thúc, giây) | `B + H` | Thời điểm ramp này hoàn tất |
| K (P cuối) | `D` | Ramp luôn hoàn tất đúng mục tiêu (trong phạm vi 1 dòng — việc bị ngắt bởi lệnh sau xử lý ở dòng sau, không phải ở đây) |

**Điểm quan trọng**: công thức không giới hạn `I` (thời điểm kết thúc ramp) trong phạm vi ngày — nếu tốc độ ramp chậm/khoảng cách công suất lớn, `I` có thể vượt quá 86400 giây (qua ngày hôm sau). Đây chính là điều kiện `carryActive` mà VBA `CalculateEndOfDayState` kiểm tra ở mục 6.

## 3. Dựng 121 đoạn công suất phủ kín 24h (sheet `DOAN_CONG_SUAT`)

Mỗi lệnh (dòng n trong `XU_LY_LENH`) sinh ra **2 đoạn** trong `DOAN_CONG_SUAT`:

- **Đoạn lẻ (2n)** — đoạn RAMP: từ `B_n` đến `I_n`, công suất từ `F_n` đến `D_n` (chép trực tiếp từ `XU_LY_LENH`).
- **Đoạn chẵn (2n+1)** — đoạn HOLD: từ `I_n` đến thời điểm bắt đầu lệnh kế tiếp (`B_{n+1}`), hoặc đến `86400` nếu là lệnh cuối cùng ngày. Giữ nguyên công suất `D_n` (mục tiêu vừa đạt).
- **Đoạn 1 (đầu ngày)** là trường hợp đặc biệt: `HOLD` từ giây 0 đến `B_1` (thời điểm lệnh đầu tiên, hoặc 86400 nếu không có lệnh nào), giữ nguyên `P0` (`CAI_DAT!B6`).

Tổng 1 (đầu ngày) + 60×2 (ramp+hold mỗi lệnh) = **121 đoạn**, khớp `MAX_EFFECTIVE_COMMANDS=60` và giới hạn `DOAN_CONG_SUAT!B122` mà VBA `VerifyWorkbookStructure` kiểm tra tồn tại công thức.

Nếu một lệnh không tồn tại (dòng `XU_LY_LENH!C` rỗng), 2 đoạn tương ứng để trống (không tính diện tích ở bước sau).

## 4. Tích phân diện tích theo 48 chu kỳ (sheet `DIEN_TICH`)

Mỗi hàng = 1 chu kỳ 30 phút (`[B, C]` = `[0,1800], [1800,3600], ...`). Mỗi cột = 1 trong 121 đoạn công suất. Công thức tại giao điểm (chu kỳ i, đoạn j):

```
overlap_start = MAX(B_chu_ky, C_đoạn_bắt_đầu)
overlap_end   = MIN(C_chu_ky, D_đoạn_kết_thúc)

nếu đoạn không giao với chu kỳ (D_đoạn ≤ B_chu_ky, hoặc C_đoạn ≥ C_chu_ky, hoặc đoạn rỗng) → diện tích = 0

ngược lại:
  P_tại(overlap_start) = nội suy tuyến tính giữa (E_đoạn=P_đầu, F_đoạn=P_cuối) tại overlap_start
  P_tại(overlap_end)   = nội suy tuyến tính tương tự tại overlap_end
  diện tích = (P_tại(overlap_start) + P_tại(overlap_end)) / 2 × (overlap_end − overlap_start)
```

Đây là **tích phân hình thang** (trapezoidal rule) chuẩn — chính xác tuyệt đối với các đoạn tuyến tính (ramp) hoặc hằng số (hold), không cần chia nhỏ hơn.

Cuối mỗi hàng: `Tổng MW.s = SUM(tất cả 121 cột diện tích)`, rồi **`Qdd (MW) = Tổng MW.s / 1800`** — đây chính xác là R08 ("trung bình diện tích công suất trong 1.800 giây").

## 5. Kết quả cuối: Qdd, Qdd_V, Qdc, Qmp, Qdư (sheet `TINH_TOAN`)

Mỗi hàng = 1 chu kỳ (48 hàng), cột chính:

| Cột | Công thức | Quy tắc |
|---|---|---|
| Qdd (MW) | `= DIEN_TICH!DV` (kết quả mục 4) | R08 |
| Qdd_V (MWh) | `= Qdd / 2 × CAI_DAT!B8` (hệ số, mặc định `0,9188`) | R09 |
| Qdc (MWh) | `= INDEX/MATCH dòng "KwhGiao" trong CSV_6001, cột theo chu kỳ / 1000` | R10 |
| P_Qdc (MW) | `= Qdc × 2` | R11 |
| Ngưỡng dưới / trên | `= Qdd × (1 − CAI_DAT!B9)` / `= Qdd × (1 + CAI_DAT!B9)` (dung sai mặc định `0,03`) | R13 |
| Qmp (MWh) | `= INDEX/MATCH dòng "KwhGiao" trong CSV_6303, cột theo chu kỳ / 1000` | R12 |
| **Qdư (MWh)** | `= IF(P_Qdc < Ngưỡng_dưới HOẶC P_Qdc > Ngưỡng_trên, Qmp − Qdd_V, 0)` | R13+R14 |
| Âm/Dương | `"âm"` nếu P_Qdc<Ngưỡng dưới, `"dương"` nếu P_Qdc>Ngưỡng trên, `"trong ±3%"` nếu trong dải | |
| Chênh lệch / Đánh giá | So Qdd tính được với "Qdd file gốc" nhập tay (nếu có) — dùng để đối chiếu thủ công (nút 4), **không** phải một phần thuật toán chính thức, xem lưu ý ở mục 7 | |

## 6. Chuyển tiếp qua nửa đêm (VBA — `CalculateEndOfDayState`, `EnsureCarryOverCommand`)

Đây là phần **duy nhất thuộc engine tính toán thực sự nằm trong VBA**, không phải công thức Excel:

**Cuối ngày N** (`SaveCurrentPowerState`, gọi sau khi "Chạy chính thức" thành công): đọc dòng cuối cùng của `XU_LY_LENH` (lệnh hiệu lực cuối ngày):

```
nếu I_cuối (giây kết thúc ramp) ≤ 86400 + EPS:
    endPower = D_cuối (ramp đã hoàn tất trước nửa đêm, không carry)
ngược lại (ramp chưa xong lúc 24:00):
    elapsedMinutes = (86400 − B_cuối) / 60
    direction = SGN(D_cuối − F_cuối)
    endPower = F_cuối + direction × CAI_DAT!B7(tốc độ) × elapsedMinutes   [kẹp không vượt quá D_cuối]
    carryActive = TRUE, carryTarget = D_cuối, remainingSeconds = I_cuối − 86400
```

Kết quả ghi vào `TRANG_THAI_CONG_SUAT` (theo từng cặp ngày+tổ máy, độc lập S1/S2 — R07).

**Đầu ngày N+1** (`ApplyAutomaticInitialPower`, nút 9 hoặc tự động khi bật P0 tự động): đọc lại trạng thái ngày N. Nếu `carryActive`:

```
Chèn 1 dòng LENH_GOC mới (đầu bảng, ID = "AUTO_CARRY_yyyymmdd_S#"):
  CS ra lệnh = CS hoàn thành = carryTarget
  Thời điểm BĐTH = 00:00:00 ngày N+1
  Thời điểm hoàn thành = 00:00:00 + remainingSeconds/86400
  Nguồn lệnh (Y) = carrySource ("SO" hoặc "MO")
```

Dòng lệnh tổng hợp này sau đó được **chính pipeline công thức ở mục 1-5 xử lý y hệt một lệnh thật** — không có logic tính toán ramp riêng cho ngày mới, chỉ là P0 ngày mới = Pend ngày trước và có thêm 1 lệnh "ảo" tiếp nối đúng tiến độ ramp dở dang.

## 7. CSV parsing (VBA — `ReadCsvWithExcel50` và chuỗi hàm liên quan)

- Không dùng `Workbooks.OpenText` (né lỗi 1004 đã biết trên một số bản Excel Windows) — mở CSV như một workbook thật (`Workbooks.Open`, thử lại với `Local:=True` nếu lỗi decimal/locale).
- **Phát hiện Mac vs Windows**: nếu Excel tự tách được ≥50 cột → đọc trực tiếp từng ô. Nếu <50 cột (toàn bộ dòng dồn vào 1-vài cột, điển hình khi Excel trên Mac mở CSV) → dựng lại chuỗi CSV chuẩn từ các ô đã mở (`BuildCsvLineFromOpenedRow`/`EncodeCsvToken`), rồi tự tách bằng bộ phân tích CSV viết tay (`ParseCsvLine`, hỗ trợ dấu ngoặc kép/escape).
- Số hợp lệ được parse độc lập với Regional Settings của hệ điều hành (`TryParseCsvNumber` dùng `Val()`, không dùng hàm phụ thuộc locale).
- Dòng `KwhGiao` được tìm bằng cách quét cột B tìm đúng chuỗi (không phân biệt hoa/thường, đã trim) "KWHGIAO"; 48 giá trị ở cột 3-50 của dòng đó phải là số, nếu không sẽ báo lỗi trước khi ghi.
- **Quan trọng — khác với mô tả nghiệp vụ ban đầu**: cột ngày (cột A) của CSV **không hề được đọc, so sánh hay dùng để từ chối import** trong bản v1.3.1 — chỉ lưu để tham khảo. Điều kiện bắt buộc duy nhất là `CAI_DAT!B4` (ngày đang cấu hình trong công cụ) phải là ngày hợp lệ trước khi bắt đầu import CSV. Hàm `ParseCsvDate` từng làm việc so sánh ngày đã được đánh dấu **legacy, không còn được gọi từ v1.1.2**. Lý do lịch sử: Excel từng đọc sai định dạng ngày tuỳ theo Regional Settings của từng máy (Mac/Windows) — xem [14_Knowledge_Transfer.md](14_Knowledge_Transfer.md). Xem thêm ghi chú ở [09_Test_Cases.md](09_Test_Cases.md) (UAT-11 mô tả hành vi "từ chối CSV sai ngày" — cần chạy lại UAT này để xác nhận đây có còn đúng với v1.3.1 hay không, vì code hiện tại không cho thấy cơ chế từ chối đó).
- **Khác biệt ở bản Apps Script**: Google Apps Script đọc chuỗi text không phụ thuộc Regional Settings của máy người dùng (không có giới hạn như Excel), nên `src/NhaMay-Mau-Template/Sidebar.html` **chủ động đọc và dùng ngày trong CSV** làm gợi ý tự động điền vào ô Ngày (người dùng xem lại/sửa trước khi lưu) — khôi phục lại tính năng đã bỏ ở VBA, vì lý do khiến VBA phải bỏ (giới hạn Excel) không còn áp dụng ở nền tảng mới.
- Sau khi ghi vào `CSV_6001`/`CSV_6303`, có bước đọc lại xác minh đúng 48/48 ô là kiểu số thật (không bị Excel tự động đổi định dạng) trước khi coi là thành công.

## 8. Báo cáo tháng / Snapshot (VBA — `QduMonth_*`)

- `QduMonth_SaveCurrentDay`: sau khi "Chạy chính thức" xong, **copy nguyên sheet `BAO_CAO_QDU`** thành sheet ẩn `LS_yyyymm_dd_S#`, rồi **chuyển toàn bộ công thức trong bản copy thành giá trị tĩnh** (`ConvertSnapshotFormulasToValues`). Đây là lý do báo cáo tháng đọc snapshot chứ không đọc `TINH_TOAN` trực tiếp: `TINH_TOAN` chỉ phản ánh ngày đang mở gần nhất, còn báo cáo tháng cần dữ liệu của **mọi ngày trong tháng cùng lúc** — chỉ snapshot đông cứng mới giữ được điều đó.
- Khoá tháng (`LICH_SU_THANG!N2`, dạng `yyyymm`) được xác định qua 3 tầng dự phòng: ô `N2` (chuẩn) → suy từ text `B2` (tương thích ngược với bản cũ) → suy từ ngày của dòng lịch sử đầu tiên (nếu 2 tầng trên đều thiếu/sai). Đây là điểm rủi ro gây lỗi "hiểu sai sang tháng mới" đã gặp trong lịch sử (xem [14_Knowledge_Transfer.md](14_Knowledge_Transfer.md)) — nếu `B2` bị sửa tay sai định dạng, tầng dự phòng có thể suy ra khoá tháng khác với kỳ vọng người dùng.
- Xuất báo cáo tháng (`QduMonth_ExportMonthlyReport`) dựng 1 sheet tổng hợp (`TONG_HOP`) + 1 sheet gộp mỗi ngày (S1 cột B:K, S2 cột L:U cạnh nhau), lấy dữ liệu từ các sheet `LS_...` ẩn.

## 9. Hằng số và giới hạn (tổng hợp từ VBA `Const` + `CAI_DAT`)

| Tên | Giá trị mặc định | Nguồn | Ghi chú |
|---|---|---|---|
| Tốc độ ramp | 3,5 MW/phút | `CAI_DAT!B7` | Dùng trong mục 2 (H) và mục 6 |
| Hệ số Qdd_V | 0,9188 | `CAI_DAT!B8` | Mục 5 |
| Dung sai ±3% | 0,03 | `CAI_DAT!B9` | Mục 5 — **có thể cấu hình**, không hard-code |
| Số lệnh hiệu lực tối đa | 60 | `CAI_DAT!B10`, `MAX_EFFECTIVE_COMMANDS` | Giới hạn 121 đoạn ở mục 3 |
| Số lệnh gốc tối đa | 199 | `MAX_COMMANDS` | `LENH_GOC` hàng 4-203 |
| Số chu kỳ/ngày | 48 | `PERIOD_COUNT` | |
| Epsilon so sánh số thực | 0,000001 | `EPS` | Dùng trong nội suy (mục 2) và carry-over (mục 6) — lưu ý có **epsilon thứ hai** `0,0000000001` dùng riêng khi xác định dấu Qdư trong phần hiển thị báo cáo, không đồng nhất với `EPS` chính |
| Số giây/ngày | 86400 | hằng số nội suy trong công thức, không đặt tên | |

## 10. Giới hạn đã biết / cần lưu ý khi migrate

- Toàn bộ vị trí cột (`LENH_DIEU_DO` cột B-X, `XU_LY_LENH` cột A-L, `TINH_TOAN` cột A-Q...) là **tham chiếu cứng theo vị trí**, không có tên trường tường minh trong công thức — khi viết lại thành code Apps Script, nên đặt tên biến rõ ràng thay vì giữ nguyên kiểu tham chiếu theo cột.
- Công thức Ramp Engine (mục 2) **không giới hạn ramp trong phạm vi 1 ngày** — việc "cắt" đúng lúc 24:00 chỉ xảy ra ở bước dựng `DOAN_CONG_SUAT` (đoạn cuối cùng bị giới hạn `MIN(..., 86400)`) và ở VBA khi lưu trạng thái cuối ngày. Khi viết lại, cần giữ đúng ranh giới trách nhiệm này: engine ramp tính "nếu không bị gì cản" thì bao giờ xong, còn việc cắt theo ngày là trách nhiệm của tầng dựng lịch/carry-over.
- Cột "Qdd file gốc" / "Chênh lệch" / "Đánh giá" trong `TINH_TOAN` là **công cụ đối chiếu thủ công** (nút 4 "Nhập Qdd gốc"), không phải một bước bắt buộc của thuật toán chính thức — không nên coi tolerance `0,1`/`0,15`/`0,3` MW ở đây là một phần của Business Rules chính thức (R01-R14), đây là ngưỡng cảnh báo khi đối chiếu, khác hẳn dung sai ±3% dùng để tính Qdư.
- Việc CSV không kiểm tra ngày (mục 7) là điểm khác biệt giữa hành vi thực tế và mô tả UAT-11 — **cần xác nhận lại khi chạy UAT**, có thể tài liệu UAT viết theo kỳ vọng thiết kế chứ không phải hành vi hiện tại của v1.3.1.
