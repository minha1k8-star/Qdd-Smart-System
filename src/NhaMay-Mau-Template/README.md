# Sheet mẫu nhà máy — Nha may mau (Duyen Hai 1)

Google Sheets đầu tiên gắn `QDD-Core-Library`, dùng làm **bản mẫu để copy cho từng nhà máy** khi nhân rộng (xem [docs/05_System_Architecture.md](../../docs/05_System_Architecture.md)).

- **Sheet thật**: https://docs.google.com/spreadsheets/d/1cJmYefz09ch2DMwJ8Mwgl8XnXH_9-M6kD0ApSB3mDQQ/edit
- **Apps Script đính kèm**: https://script.google.com/d/10wCqQ4Xf4QUcbBum4DqVMKuw4IL4A4M5iw3nMbdBPKLXq15CJ89uYA3Y/edit
- **Gắn thư viện**: `QDD-Core-Library` (Script ID `10_vjTSgVjZodA7xTkJ_qJaGom3JDx_tnYE0YgWA_cphh1Q7g_lTKMLUO`, version 1) — khai báo trong `appsscript.json`.

## Bắt đầu dùng

1. Mở Sheet ở link trên. Nếu chưa thấy menu **"QDD Smart System"** trên thanh menu, tải lại trang (F5).
2. **QDD Smart System → 1. Thiết lập sheet** — tạo đủ các sheet cần thiết (chỉ cần chạy 1 lần).
3. Điền `CAI_DAT` đúng thông số thật của nhà máy (tốc độ ramp, hệ số Qdd_V, dung sai) — mặc định đang để theo Duyên Hải 1.

## Nhập dữ liệu

### Danh sách lệnh

Nhập/dán trực tiếp vào sheet **`LENH`** — mỗi dòng 1 lệnh, có thể nhiều ngày cùng lúc (khác VBA — không cần xoá dữ liệu ngày cũ để nhập ngày mới). Cột `Thời điểm BĐTH` phải là kiểu ngày-giờ thật (không phải text).

### CSV công tơ (Qdc/Qmp)

1. Mở sheet **`CSV_STAGING`** → **File → Import → Upload** → chọn file CSV (6001 hoặc 6303) → **Insert new sheet** hoặc **Replace current sheet** (dùng tính năng nhập CSV có sẵn của Google Sheets, không tự viết lại logic tách cột như VBA).
2. Quay lại menu **QDD Smart System → 2. Lưu CSV → 6001** (hoặc **3. → 6303**), nhập đúng ngày của file đó khi được hỏi.
3. Lặp lại cho từng ngày, từng công tơ.

### P0 (công suất đầu ngày)

**Giới hạn quan trọng**: bản này **chưa cài đặt carry-over qua nửa đêm (R07)** — không tự tính P0 từ ngày trước như VBA. Phải điền tay vào sheet **`P0_NGAY`** (Ngày, Tổ máy, P0) trước khi tính ngày đó.

## Tính toán

- **QDD Smart System → 4. Tính 1 ngày** — nhập ngày + tổ máy, kết quả ghi vào `KET_QUA`.
- **QDD Smart System → 5. Tính nhiều ngày** — nhập khoảng ngày + tổ máy (vd `S1,S2`), tính hàng loạt. Ngày nào thiếu P0/CSV sẽ báo lỗi riêng, không chặn các ngày khác.
- **QDD Smart System → 6. Tổng hợp báo cáo tháng** — tổng hợp trực tiếp từ `KET_QUA` đã có (không cần tính lại, không cần snapshot như VBA).

## Chưa làm (so với VBA v1.3.1)

- Carry-over qua nửa đêm (R07) — phải điền P0 tay.
- Cảnh báo lệnh 0-0 (UAT-34, xem [ROADMAP.md](../../ROADMAP.md)).
- Xuất báo cáo ngày/tháng ra file Excel/PDF riêng — hiện chỉ ghi vào sheet `KET_QUA`/`BAO_CAO_THANG`.
- Kiểm tra cấu trúc/backup tự động (tương đương nút 11, 13 VBA).
- Schema `LENH` rút gọn (9 cột) so với `LENH_GOC` gốc (25 cột) — chỉ giữ đúng trường thuật toán cần dùng.

## Cập nhật code sau này

```bash
cd src/NhaMay-Mau-Template
npx clasp push
```

Khi `QDD-Core-Library` ra version mới, vào Apps Script Editor của Sheet này → **Libraries** → chọn version mới → Save, để cập nhật (không tự động, đúng theo kiến trúc đã chọn).
