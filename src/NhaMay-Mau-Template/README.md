# Sheet mẫu nhà máy — Nha may mau (Duyen Hai 1)

Google Sheets đầu tiên gắn `QDD-Core-Library`, dùng làm **bản mẫu để copy cho từng nhà máy** khi nhân rộng (xem [docs/05_System_Architecture.md](../../docs/05_System_Architecture.md)).

- **Sheet thật**: https://docs.google.com/spreadsheets/d/1cJmYefz09ch2DMwJ8Mwgl8XnXH_9-M6kD0ApSB3mDQQ/edit
- **Apps Script đính kèm**: https://script.google.com/d/10wCqQ4Xf4QUcbBum4DqVMKuw4IL4A4M5iw3nMbdBPKLXq15CJ89uYA3Y/edit
- **Gắn thư viện**: `QDD-Core-Library` (Script ID `10_vjTSgVjZodA7xTkJ_qJaGom3JDx_tnYE0YgWA_cphh1Q7g_lTKMLUO`, version 1) — khai báo trong `appsscript.json`.

## Bắt đầu dùng

1. Mở Sheet ở link trên. Nếu chưa thấy menu **"QDD Smart System"** trên thanh menu, tải lại trang (F5).
2. **QDD Smart System → Thiết lập sheet** — tạo đủ các sheet cần thiết (chỉ cần chạy 1 lần).
3. Điền `CAI_DAT` đúng thông số thật của nhà máy (tốc độ ramp, hệ số Qdd_V, dung sai) — mặc định đang để theo Duyên Hải 1.
4. **QDD Smart System → Bảng điều khiển** — mở sidebar bên phải, dùng cho mọi thao tác còn lại (không cần dùng menu nữa).

## Bảng điều khiển (sidebar)

Toàn bộ thao tác nằm gọn trong 1 sidebar, có lịch chọn ngày và upload file trực tiếp — không còn phải gõ tay ngày tháng hay qua 2 bước nhập CSV như trước.

### 1. Lưu CSV công tơ

Chọn ngày, chọn công tơ (6001 = Qdc, 6303 = Qmp), chọn file CSV từ máy, bấm **Lưu CSV**. Xong ngay trong 1 bước — trình duyệt tự đọc file, gửi thẳng lên hệ thống.

### 2. Tính 1 ngày

Chọn ngày (lịch), chọn tổ máy, bấm **Tính**.

**P0 (công suất đầu ngày)** được **tự động lấy từ chu kỳ cuối cùng của ngày liền trước đã tính** (nếu có) — không cần nhập tay như trước nữa, trừ **ngày đầu tiên sử dụng hệ thống** (chưa có ngày nào trước đó) thì vẫn cần điền 1 lần vào sheet `P0_NGAY`.

> Lưu ý: đây là P0 **xấp xỉ** (lấy đúng Qdd chu kỳ cuối ngày trước), chưa phải carry-over R07 đầy đủ (chưa mô phỏng ramp còn dở dang qua nửa đêm) — đủ dùng cho phần lớn trường hợp thực tế.

### 3. Tính hàng loạt

Chọn khoảng ngày + tổ máy (tick chọn S1/S2), bấm **Tính hàng loạt**. Ngày nào thiếu CSV sẽ báo trong kết quả, không chặn các ngày khác.

### 4. Báo cáo tháng

Chọn tháng, bấm **Tổng hợp** — gộp trực tiếp từ các ngày đã tính trong `KET_QUA` (không cần tính lại, không cần snapshot như VBA).

## Danh sách lệnh

Nhập/dán trực tiếp vào sheet **`LENH`** (không qua sidebar) — mỗi dòng 1 lệnh, có thể nhiều ngày cùng lúc (khác VBA — không cần xoá dữ liệu ngày cũ để nhập ngày mới). Cột `Thời điểm BĐTH` phải là kiểu ngày-giờ thật (không phải text).

## Chưa làm (so với VBA v1.3.1)

- Carry-over qua nửa đêm (R07) đầy đủ — hiện chỉ xấp xỉ bằng Qdd chu kỳ cuối ngày trước (xem mục "Tính 1 ngày" ở trên).
- Cảnh báo lệnh 0-0 (UAT-34, xem [ROADMAP.md](../../ROADMAP.md)).
- Xuất báo cáo ngày/tháng ra file Excel/PDF riêng — hiện chỉ ghi vào sheet `KET_QUA`/`BAO_CAO_THANG`.
- Kiểm tra cấu trúc/backup tự động (tương đương nút 11, 13 VBA).
- Nhập lệnh qua sidebar (hiện vẫn nhập trực tiếp vào sheet `LENH`).
- Schema `LENH` rút gọn (9 cột) so với `LENH_GOC` gốc (25 cột) — chỉ giữ đúng trường thuật toán cần dùng.

## Cập nhật code sau này

```bash
cd src/NhaMay-Mau-Template
npx clasp push
```

Khi `QDD-Core-Library` ra version mới, vào Apps Script Editor của Sheet này → **Libraries** → chọn version mới → Save, để cập nhật (không tự động, đúng theo kiến trúc đã chọn).
