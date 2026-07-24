# Database Design — QDD Smart System

"Cơ sở dữ liệu" của hệ thống hiện tại là các sheet trong workbook Excel chính thức (`legacy/CongCu_Tinh_Qdd_Qdu_v1_3_0_AllInOne.xlsm`). Tài liệu này mô tả schema từng sheet, lấy trực tiếp từ workbook — không suy diễn.

## Sheet có sẵn trong file mẫu (template)

### `CAI_DAT` — Cấu hình ngày tính

| Thông số | Ghi chú |
|---|---|
| Ngày tính | |
| Tổ máy | mặc định `S1` |
| Công suất đầu ngày (MW) | P0 |
| Tốc độ tăng/giảm tải (MW/phút) | mặc định `3.5` — dùng cho R05 |
| Hệ số Qdd_V | mặc định `0.9188` — dùng cho R09 |
| Dung sai | mặc định `0.03` (±3%) — dùng cho R13 |
| Số lệnh hợp lệ | đếm tự động |
| Quy tắc công suất SO | mô tả — "Dùng CS hoàn thành" |
| Quy tắc công suất MO | |
| Nguồn công suất đầu ngày | |
| Công suất cuối ngày (MW) | |
| Ramp qua 00:00 | |
| Dung sai đối chiếu Qdd (MW) | |
| Phiên bản công cụ | hiện `1.3.0`/`1.3.1` |
| Chế độ sử dụng | `CHÍNH THỨC` |
| Nguồn Qdc | `Công tơ 6001 - KwhGiao` |
| Nguồn Qmp | `Công tơ 6303 - KwhGiao` |
| Số lệnh hiệu lực tối đa/ngày | `60` |

### `LENH_GOC` — Danh sách lệnh điều độ gốc (raw)

Vùng dữ liệu: hàng 4–203 (`CMD_FIRST_ROW=4`, `CMD_LAST_ROW=203`, tối đa `MAX_COMMANDS=199` lệnh). Header ở hàng 3, cột A–Y:

| Cột | Tên | Vai trò trong Business Rules |
|---|---|---|
| A | ID Lệnh | |
| B | Nhà máy | |
| C | Tổ máy | lọc theo S1/S2 |
| D | Nội dung lệnh | |
| E | CS ra lệnh (MW) | dùng làm P hiệu lực khi MO bình thường (R02) |
| F | CS hoàn thành (MW) | dùng làm P hiệu lực khi SO hợp lệ (R01) hoặc MO dừng sớm (R03) |
| G | Thời điểm BĐTH | mốc bắt đầu ramp (R04) — **không dùng cột H** |
| H | Thời điểm hoàn thành | |
| I | Người ra lệnh | |
| J | Người thực hiện | |
| K | AGC | |
| L | Nhiên liệu | |
| M | Lý do lệnh | |
| N | Ghi chú ra lệnh | |
| O | Ghi chú hoàn thành | |
| P | Hoàn thành | cờ hợp lệ, `P=1` là điều kiện bắt buộc ở R01–R03 |
| Q | Dừng lệnh | `Q=TRUE` → lệnh MO bị dừng sớm (R03) |
| R | Thời điểm dừng | |
| S | Lý do dừng | |
| T | Người dừng | |
| U | Lý do hủy | |
| V | Người hủy | |
| W | Lệnh cụm | |
| X | Lệnh nhập lại | |
| Y | Nguồn lệnh | giá trị `SO`/`MO`, dùng trong điều kiện R01/R02 |

### `LENH_DIEU_DO` — Lệnh hiệu lực (tự động lọc từ LENH_GOC)

Sheet tính toán trung gian, tự sinh từ `LENH_GOC` qua VBA/công thức. Hai khối cột:

- **A–H** (kết quả đã xử lý): STT, Tổ máy, Thời điểm BĐTH, Công suất hiệu lực, Thời điểm hoàn thành, CS hoàn thành, Ghi chú xử lý, Thứ tự.
- **J–X** (vùng làm việc/staging): Dòng, ID, Tổ máy, Nguồn, BĐTH, CS ra lệnh, CS hoàn thành, Hoàn thành, Dừng, Thời điểm dừng, Hợp lệ, P hiệu lực, Giây, Thứ tự tăng dần, Kết luận.

Tối đa `MAX_EFFECTIVE_COMMANDS=60` lệnh hiệu lực/ngày (khớp UAT-19).

### `CSV_6001` / `CSV_6303` — Dữ liệu công tơ nhập từ CSV

Vùng dữ liệu: tối đa 50 hàng (`CSV_LAST_ROW=50`), 50 cột (A–AX). Không có header cố định trong template (do dữ liệu được ghi đè khi import). Theo R10–R12: chỉ dùng dòng `KwhGiao` và 48 giá trị chu kỳ số; **cột ngày (cột A) bị bỏ qua khi import**.

- `CSV_6001` → nguồn của **Qdc**.
- `CSV_6303` → nguồn của **Qmp**.

### `XU_LY_LENH` — Ramp Engine (mỗi dòng = 1 lệnh hiệu lực)

Header: `STT, Bắt đầu (giây), Thời điểm bắt đầu, P mục tiêu, Thời điểm hoàn thành gốc, P bắt đầu, Loại, Thời lượng (giây), Kết thúc (giây), Thời điểm kết thúc, P cuối, Ghi chú`. Vùng dữ liệu B2:L61 (khớp `MAX_EFFECTIVE_COMMANDS=60`). Đây là sheet chứa **công thức nội suy ngắt ramp** (R06) — cột "P bắt đầu" của lệnh sau phụ thuộc vào lệnh trước, xem công thức đầy đủ ở [04_Algorithm_Specification.md](04_Algorithm_Specification.md#2-ramp-engine-sheet-xu_ly_lenh--nội-suy-liên-tục-giữa-các-lệnh).

### `DOAN_CONG_SUAT` — Các đoạn công suất theo thời gian

Header: `Đoạn, Loại, Bắt đầu (giây), Kết thúc (giây), P đầu, P cuối, Thời lượng, Diễn giải`. Tối đa 121 đoạn/ngày = 1 đoạn đầu ngày + 60 lệnh × 2 đoạn (RAMP + HOLD) (`DIEN_TICH` có 121 cột đoạn tương ứng). Đây là kết quả dựng bằng công thức Excel (không phải VBA) từ `XU_LY_LENH` — công thức nội suy đầy đủ ở [04_Algorithm_Specification.md](04_Algorithm_Specification.md#2-ramp-engine-sheet-xu_ly_lenh--nội-suy-liên-tục-giữa-các-lệnh).

### `DIEN_TICH` — Tính diện tích công suất theo chu kỳ

Header: `Chu kỳ, Bắt đầu giây, Kết thúc giây, Đoạn 1 ... Đoạn 121, Tổng MW.s, Qdd MW`. Mỗi hàng là một trong 48 chu kỳ 30 phút (1.800 giây); mỗi cột `Đoạn N` là phần diện tích (MW·s) mà đoạn công suất thứ N đóng góp vào chu kỳ đó (0 nếu đoạn không giao với chu kỳ). Tổng theo hàng chia cho 1.800 giây ra `Qdd MW` (R08).

### `TINH_TOAN` — Bảng tính 48 chu kỳ

Header: `Chu kỳ, Bắt đầu, Kết thúc, Qdd (MW), Qdd_V (MWh), Qdc (MWh), P_Qdc (MW), Ngưỡng dưới, Ngưỡng trên, Qmp (MWh), Qdư (MWh), Âm/Dương, Trạng thái, Diễn giải, Qdd file gốc, Chênh lệch, Đánh giá, Ghi chú`. Đây là bảng tổng hợp kết quả cuối cùng của toàn bộ pipeline tính toán (R08–R14), có cột đối chiếu với "Qdd file gốc" nhập tay để kiểm tra sai lệch (UAT-13).

### `BAO_CAO_QDU` — Báo cáo ngày

Header: `Chu kỳ, Qdd, Qdd_V, Qdc, Qmp, Qdư, Qdư âm/dương, P_Qdc, Ngưỡng dưới, Ngưỡng trên, Ghi chú`, cùng dòng tiêu đề "Ngày ... | Tổ máy ...". Vùng dữ liệu 48 chu kỳ (hàng 4–51) + hàng tổng (52).

### `QUY_TAC_NGHIEP_VU`, `KIEM_THU_UAT`, `THONG_TIN_HE_THONG`, `HUONG_DAN`

Sheet tài liệu nội bộ, là nguồn cho [03_Business_Rules.md](03_Business_Rules.md) và [09_Test_Cases.md](09_Test_Cases.md). Không chứa dữ liệu vận hành.

## Sheet sinh ra khi vận hành (không có trong file mẫu, chỉ thấy trong mã VBA)

Các sheet sau được VBA tự tạo khi công cụ chạy lần đầu (`GetOrCreateStateSheet`, `GetOrCreateControlSheet`, `QduMonth_Initialize`...), **không có sẵn trong `CongCu_Tinh_Qdd_Qdu_v1_3_0_AllInOne.xlsm`** vì đây là file mẫu/template chưa vận hành:

| Sheet | Hằng số VBA | Vai trò |
|---|---|---|
| `DIEU_KHIEN_VBA` | `SH_CTRL` | Sheet điều khiển — chứa 15 nút bấm |
| `TRANG_THAI_CONG_SUAT` | `SH_STATE` | Lưu P tại 24:00 mỗi ngày, theo S1/S2 — nguồn của R07/AUTO_CARRY |
| `NHAT_KY_VBA` | `SH_LOG` | Log thao tác (`WriteLog`) |
| `LICH_SU_THANG` | `M_SH_HISTORY` | Snapshot lịch sử theo tháng — **không được xoá** (xem AGENTS.md) |
| `LS_...` (tiền tố) | `M_PREFIX` | Sheet snapshot riêng từng tháng, tiền tố `LS_` |

**Không được xoá `LICH_SU_THANG` hoặc bất kỳ sheet `LS_...` nào** — đây là dữ liệu lịch sử không thể tái tạo lại từ đầu.

## Giới hạn/hằng số quan trọng

| Hằng số | Giá trị | Ý nghĩa |
|---|---|---|
| `CMD_FIRST_ROW` – `CMD_LAST_ROW` | 4 – 203 | Vùng dữ liệu `LENH_GOC` |
| `MAX_COMMANDS` | 199 | Số lệnh gốc tối đa |
| `MAX_EFFECTIVE_COMMANDS` | 60 | Số lệnh hiệu lực tối đa/ngày (khớp UAT-19) |
| `CSV_LAST_ROW` | 50 | Số hàng tối đa mỗi CSV |
| `PERIOD_COUNT` | 48 | Số chu kỳ/ngày |
| `EPS` | 0.000001 | Sai số làm tròn khi so sánh số thực |
