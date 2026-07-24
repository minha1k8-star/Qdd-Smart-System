# Sheet mẫu nhà máy — Nha may mau (Duyen Hai 1)

Google Sheets đầu tiên gắn `QDD-Core-Library`, dùng làm **bản mẫu để copy cho từng nhà máy** khi nhân rộng (xem [docs/05_System_Architecture.md](../../docs/05_System_Architecture.md)).

- **Sheet thật**: https://docs.google.com/spreadsheets/d/1cJmYefz09ch2DMwJ8Mwgl8XnXH_9-M6kD0ApSB3mDQQ/edit
- **Apps Script đính kèm**: https://script.google.com/d/10wCqQ4Xf4QUcbBum4DqVMKuw4IL4A4M5iw3nMbdBPKLXq15CJ89uYA3Y/edit
- **Gắn thư viện**: `QDD-Core-Library` (Script ID `10_vjTSgVjZodA7xTkJ_qJaGom3JDx_tnYE0YgWA_cphh1Q7g_lTKMLUO`, version 1) — khai báo trong `appsscript.json`.

## Bắt đầu dùng

1. Mở Sheet ở link trên. Nếu chưa thấy menu **"QDD Smart System"** trên thanh menu, tải lại trang (F5).
2. **QDD Smart System → Thiết lập sheet** — tạo đủ các sheet cần thiết (chạy được nhiều lần, tự bù các dòng cấu hình còn thiếu nếu chạy lại sau khi nâng cấp).
3. Điền `CAI_DAT` đúng thông số thật của nhà máy:
   - Tốc độ ramp, hệ số Qdd_V, dung sai.
   - **Mã công tơ Qdc/Qmp cho từng tổ máy** (4 dòng riêng: Qdc-S1, Qmp-S1, Qdc-S2, Qmp-S2). ⚠️ **Mã công tơ khác nhau giữa S1 và S2, và khác nhau giữa các nhà máy** — "6001"/"6303" chỉ là ví dụ mặc định cho Duyên Hải 1 tổ S1, phải tự điền đúng mã thật khi copy Sheet cho nhà máy/tổ máy khác.
4. **QDD Smart System → Bảng điều khiển** — mở sidebar bên phải, dùng cho mọi thao tác còn lại (không cần dùng menu nữa).

## Bảng điều khiển (sidebar)

Toàn bộ thao tác nằm gọn trong 1 sidebar, có lịch chọn ngày và upload file trực tiếp — không còn phải gõ tay ngày tháng hay qua 2 bước nhập CSV như trước.

### 1. Lưu CSV công tơ

Chọn tổ máy, chọn loại dữ liệu (Qdc/Qmp), chọn file CSV từ máy — **hệ thống tự đọc ngày ghi trong file** (cột A) và điền sẵn vào ô Ngày, bạn chỉ cần kiểm tra lại/sửa nếu cần rồi bấm **Lưu CSV**. Tự tra đúng mã công tơ theo cấu hình trong `CAI_DAT` (không cần nhớ mã công tơ là số mấy).

> **Khác với VBA cũ**: bản VBA v1.3.1 chủ động **không đọc** ngày trong CSV (chỉ để tham khảo), vì Excel từng đọc sai định dạng ngày tuỳ theo Regional Settings của từng máy (Mac/Windows) — xem [docs/14_Knowledge_Transfer.md](../../docs/14_Knowledge_Transfer.md). Google Apps Script không có giới hạn đó (không phụ thuộc cài đặt vùng của máy người dùng khi đọc chuỗi text), nên bản này đọc và dùng ngày trong file làm gợi ý tự động, có kiểm tra lại bằng mắt trước khi lưu.

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
