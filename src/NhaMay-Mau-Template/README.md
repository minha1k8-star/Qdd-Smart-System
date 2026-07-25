# Sheet mẫu nhà máy — Nha may mau (Duyen Hai 1)

Google Sheets đầu tiên gắn `QDD-Core-Library`, dùng làm **bản mẫu để copy cho từng nhà máy** khi nhân rộng (xem [docs/05_System_Architecture.md](../../docs/05_System_Architecture.md)).

- **Sheet thật**: https://docs.google.com/spreadsheets/d/1cJmYefz09ch2DMwJ8Mwgl8XnXH_9-M6kD0ApSB3mDQQ/edit
- **Apps Script đính kèm**: https://script.google.com/d/10wCqQ4Xf4QUcbBum4DqVMKuw4IL4A4M5iw3nMbdBPKLXq15CJ89uYA3Y/edit
- **Gắn thư viện**: `QDD-Core-Library` (Script ID `10_vjTSgVjZodA7xTkJ_qJaGom3JDx_tnYE0YgWA_cphh1Q7g_lTKMLUO`, version 1) — khai báo trong `appsscript.json`.

## Bắt đầu dùng

1. Mở Sheet ở link trên. Nếu chưa thấy menu **"QDD Smart System"**, tải lại trang (F5).
2. **QDD Smart System → Thiết lập sheet** — tạo/nâng cấp đủ các sheet cần thiết. Chạy lại nhiều lần được: tự bù dòng cấu hình còn thiếu và tự dựng lại `LENH` đúng cấu trúc mà không mất dữ liệu.
3. Điền `CAI_DAT` đúng thông số thật của nhà máy:
   - Tốc độ ramp, hệ số Qdd_V, dung sai.
   - **Mã công tơ Qdc/Qmp cho từng tổ máy** (4 dòng: Qdc-S1, Qmp-S1, Qdc-S2, Qmp-S2). ⚠️ Mã công tơ khác nhau giữa S1/S2 và giữa các nhà máy — giá trị mặc định (6001/6303/6002/6301) chỉ đúng cho Duyên Hải 1.
   - **Nhãn báo cáo** cho từng tổ (mặc định `S1DH1`, `S2DH1`) — hiện trên đầu mỗi khối trong file xuất.
4. **QDD Smart System → Bảng điều khiển** — mở sidebar bên phải, dùng cho mọi thao tác còn lại.

> **Người dùng cần quyền Chỉnh sửa (Editor)** mới thấy menu — người chỉ có quyền Xem sẽ không chạy được Apps Script.

## Bảng điều khiển (sidebar)

### 1. Lưu CSV công tơ
Chọn tổ máy + loại dữ liệu (Qdc/Qmp) + file CSV — hệ thống **tự đọc ngày trong file** điền sẵn vào ô Ngày, và tự tra mã công tơ theo `CAI_DAT`. Nếu tên file ứng với mã công tơ **khác** lựa chọn của bạn, hệ thống **chặn lại** thay vì lưu nhầm.

> Bản VBA cũ chủ động bỏ qua ngày trong CSV vì Excel đọc sai định dạng theo Regional Settings từng máy (xem [docs/14_Knowledge_Transfer.md](../../docs/14_Knowledge_Transfer.md)). Apps Script không có giới hạn đó nên tính năng này được khôi phục.

### 2. Nhập danh sách lệnh
Hai cách, dùng cách nào cũng được:
- **Dán thẳng**: sheet `LENH` có cấu trúc **giống hệt file gốc (25 cột, Nhà máy ở cột B)** nên copy toàn bộ file gốc dán vào từ dòng 2 là xong.
- **Qua `LENH_STAGING`**: dán file gốc vào đó (giữ dòng tiêu đề) rồi bấm nút — hệ thống dò cột **theo tên tiêu đề** (không sợ lệch cột), gộp vào `LENH` theo ID Lệnh (trùng thì cập nhật, không nhân đôi) và sắp xếp lại theo thời gian.

Việc đọc dữ liệu luôn dò theo tên cột, nên `LENH` đổi thứ tự cột hay thừa/thiếu cột phụ đều không ảnh hưởng.

### 3. Tính 1 ngày
Chọn ngày + tổ máy → **Tính**.

**P0** tự lấy từ chu kỳ cuối ngày liền trước đã tính; chỉ **ngày đầu tiên** dùng hệ thống mới cần điền tay 1 dòng vào `P0_NGAY`.

P0 là **công suất tại đúng 24:00** của ngày trước (không phải Qdd trung bình chu kỳ 48). Nếu lúc 24:00 tổ máy **vẫn đang tăng/giảm tải dở dang**, hệ thống ghi thêm mục tiêu vào cột `Ramp tiếp đến (MW)` và ngày hôm sau **tự chạy tiếp phần còn lại** cho tới khi đạt mục tiêu — đúng quy tắc **R07**, tương đương cơ chế `AUTO_CARRY` của bản VBA cũ.

Tuỳ chọn **"Dọn lệnh + CSV của ngày này sau khi tính xong"** (mặc định bật) — xoá dữ liệu nguồn của đúng (ngày, tổ máy) vừa tính. Tắt đi khi đang dò lỗi để không phải nhập lại.

### 4. Tải CSV hàng loạt
Chọn **nhiều file CSV cùng lúc** — hệ thống tự nhận diện tổ máy/loại dữ liệu bằng cách so tên file với mã công tơ trong `CAI_DAT` (tên file thật dạng `<ngày><tháng><mã công tơ>.CSV`, vd `17076001.CSV`) và tự đọc ngày từ nội dung. File không nhận diện được sẽ báo riêng theo tên.

### 5. Tính hàng loạt
Chọn khoảng ngày + tổ máy. Ngày nào thiếu dữ liệu sẽ báo riêng, không chặn các ngày còn lại. Cũng có tuỳ chọn dọn dữ liệu nguồn như mục 3.

### 6. Báo cáo tháng
Chọn tháng + tổ máy + định dạng → làm 2 việc: cập nhật bảng tổng vào `BAO_CAO_THANG`, và **xuất 1 file gộp toàn bộ ngày đã tính trong tháng**.

### 7. Xuất báo cáo (ngày cụ thể)
Như mục 6 nhưng cho khoảng ngày tự do. Chỉ xuất được ngày **đã tính xong**.

**Layout file xuất** (chung cho mục 6 và 7, bám file `Kiểm tra Qdu` gốc): mỗi ngày 1 tab; A1 = `MWh`; cột A là chu kỳ dạng `01 [00:00-00:30]`; các tổ máy nằm **liền cột** (B:K = tổ 1, L:U = tổ 2) với thứ tự `Qdd | Qdd_V | Qdc | Qmp | Qdư | Qdư âm/dương | P_Qdc | Ngưỡng dưới | Ngưỡng trên | Ghi chú`; cuối bảng có hàng **Tổng ngày**. File lưu cùng thư mục Drive với Sheet, sidebar hiện link tải.

Hai điểm bám sát bản gốc:
- **Tổ máy chưa có dữ liệu vẫn giữ nguyên bảng** (để trống, không có hàng tổng) — không bị mất bảng như trước.
- Cột **Qdư âm/dương** chỉ ghi `âm` / `dương`; khi Qdư = 0 (trong dải dung sai) thì để dấu **`-`**, không ghi chữ "trong ±3%" (chữ đó chỉ dùng trong sheet `KET_QUA` nội bộ cho dễ đọc).

### 8. Dọn dữ liệu cũ
Dùng khi bàn giao Sheet cho người khác, hoặc khi dữ liệu tích luỹ nhiều tháng. Giữ nguyên `CAI_DAT` và kết quả của **N ngày gần nhất** (để P0 vẫn tự suy được), xoá phần còn lại. Có hộp xác nhận trước khi chạy.

## Cảnh báo tự động

**Lệnh "0-0"** (CS ra lệnh = CS hoàn thành = 0, thường là trip/ngừng sự cố): theo quy tắc nghiệp vụ đã xác nhận, loại lệnh này **không được tính** — nhưng hệ thống sẽ **cảnh báo rõ** để người vận hành biết Qdd giai đoạn đó có thể cao hơn thực tế và tự xử lý. Hệ thống chỉ cảnh báo, không tự sửa số (UAT-34, xem [docs/09_Test_Cases.md](../../docs/09_Test_Cases.md)).

## Chưa làm (so với VBA v1.3.1)

- Kiểm tra cấu trúc/backup tự động (tương đương nút 11, 13 VBA).

## Cập nhật code sau này

```bash
cd src/NhaMay-Mau-Template
npx clasp push
```

Khi `QDD-Core-Library` ra version mới, vào Apps Script Editor của Sheet này → **Libraries** → chọn version mới → Save (không tự động, đúng theo kiến trúc đã chọn).
