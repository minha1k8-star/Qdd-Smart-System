# Sheet mẫu nhà máy — Nha may mau (Duyen Hai 1)

Google Sheets đầu tiên gắn `QDD-Core-Library`, dùng làm **bản mẫu để copy cho từng nhà máy** khi nhân rộng (xem [docs/05_System_Architecture.md](../../docs/05_System_Architecture.md)).

- **Sheet thật**: https://docs.google.com/spreadsheets/d/1cJmYefz09ch2DMwJ8Mwgl8XnXH_9-M6kD0ApSB3mDQQ/edit
- **Apps Script đính kèm**: https://script.google.com/d/10wCqQ4Xf4QUcbBum4DqVMKuw4IL4A4M5iw3nMbdBPKLXq15CJ89uYA3Y/edit
- **Gắn thư viện**: `QDD-Core-Library` (Script ID `10_vjTSgVjZodA7xTkJ_qJaGom3JDx_tnYE0YgWA_cphh1Q7g_lTKMLUO`, version 4) — khai báo trong `appsscript.json`.

## Bắt đầu dùng

1. Mở Sheet ở link trên. Nếu chưa thấy menu **"QDD Smart System"**, tải lại trang (F5).
2. **QDD Smart System → Thiết lập sheet** — tạo/nâng cấp đủ các sheet cần thiết. Chạy lại nhiều lần được: tự bù dòng cấu hình còn thiếu và tự dựng lại `LENH` đúng cấu trúc mà không mất dữ liệu.
3. Điền `CAI_DAT` đúng thông số thật của nhà máy:
   - Tốc độ ramp, hệ số Qdd_V, dung sai.
   - **Mã công tơ Qdc/Qmp cho từng tổ máy** (4 dòng: Qdc-S1, Qmp-S1, Qdc-S2, Qmp-S2). ⚠️ Mã công tơ khác nhau giữa S1/S2 và giữa các nhà máy — giá trị mặc định (6001/6303/6002/6301) chỉ đúng cho Duyên Hải 1.
   - **Nhãn báo cáo** cho từng tổ (mặc định `S1DH1`, `S2DH1`) — hiện trên đầu mỗi khối trong file xuất.
4. **QDD Smart System → Bảng điều khiển** — mở sidebar bên phải, dùng cho mọi thao tác còn lại.

Sheet **`HUONG_DAN`** (ngoài cùng bên trái) chứa toàn bộ hướng dẫn sử dụng ngay trong file — người dùng mới đọc là làm được, không cần tìm tài liệu bên ngoài. Sheet này do hệ thống tự ghi lại mỗi lần chạy "Thiết lập sheet", đừng sửa tay.

> **Người dùng cần quyền Chỉnh sửa (Editor)** mới thấy menu — người chỉ có quyền Xem sẽ không chạy được Apps Script.

## Bảng điều khiển (sidebar)

### 1. Tải CSV công tơ
Chọn **một hoặc nhiều file CSV cùng lúc** — hệ thống tự nhận diện tổ máy/loại dữ liệu bằng cách so **tên file** với mã công tơ trong `CAI_DAT` (tên file thật dạng `<ngày><tháng><mã công tơ>.CSV`, vd `17076001.CSV`) và **tự đọc ngày từ nội dung** file. File không nhận diện được sẽ báo riêng theo tên — sửa lại tên file cho đúng chuẩn rồi tải lại.

Không còn đường "lưu từng file có chọn tay tổ máy/loại dữ liệu": việc chọn tay vừa mất công vừa từng gây lưu nhầm Qmp vào ô Qdc, trong khi nhận diện theo tên file không sai được.

### 2. Nhập danh sách lệnh
**Chọn thẳng file Excel** (`.xlsx`/`.xlsm`/`.xls`) danh sách lệnh rồi bấm nút — không mở file, không copy tay.

Luồng xử lý: file được tải lên Drive kèm yêu cầu chuyển sang Google Sheets (giữ nguyên kiểu ngày-giờ, không thành chữ) → hệ thống **dò khắp các sheet, 25 dòng đầu mỗi sheet** để tìm dòng tiêu đề có đủ các cột bắt buộc → gộp vào `LENH` theo **ID Lệnh** (trùng thì cập nhật, không nhân đôi) → sắp xếp lại theo BĐTH → **xoá file tạm trên Drive** (kể cả khi nhập lỗi).

Cột được dò **theo tên tiêu đề**, nên file gốc khác thứ tự cột, thừa/thiếu cột phụ, hay có vài dòng tiêu đề/logo phía trên bảng đều không ảnh hưởng. Cột BĐTH nhận cả ô ngày-giờ thật lẫn ô dạng chữ `dd/MM/yyyy HH:mm`.

> Vẫn có thể dán tay vào sheet `LENH` như trước (cấu trúc **giống hệt file gốc: 25 cột, Nhà máy ở cột B**, dán từ dòng 2) — cách này không qua bước gộp theo ID nên dán trùng sẽ bị nhân đôi.

### 3. Tính
Chọn **Từ ngày / Đến ngày** + tổ máy → **Tính**. Tính 1 ngày là trường hợp **Từ ngày = Đến ngày** (chọn Từ ngày xong hệ thống tự điền Đến ngày giống hệt) — chỉ có **một luồng tính duy nhất**, không tách riêng "1 ngày" và "hàng loạt", để chỉ phải kiểm chứng một đường code.

Ngày nào thiếu CSV hoặc thiếu P0 sẽ được báo riêng, không chặn các ngày còn lại.

**P0** tự lấy từ chu kỳ cuối ngày liền trước đã tính; chỉ **ngày đầu tiên** dùng hệ thống mới cần điền tay 1 dòng vào `P0_NGAY`.

P0 là **công suất tại đúng 24:00** của ngày trước (không phải Qdd trung bình chu kỳ 48). Nếu lúc 24:00 tổ máy **vẫn đang tăng/giảm tải dở dang**, hệ thống ghi thêm mục tiêu vào cột `Ramp tiếp đến (MW)` và ngày hôm sau **tự chạy tiếp phần còn lại** cho tới khi đạt mục tiêu — đúng quy tắc **R07**.

Tuỳ chọn **"Dọn lệnh + CSV của các ngày đã tính xong"** (**mặc định tắt**) — khi tự tick thì xoá dữ liệu nguồn của đúng (ngày, tổ máy) vừa tính. Để nguyên (không tick) nếu còn cần đối chiếu hoặc tính lại.

### 4. Báo cáo tháng
Chọn tháng + tổ máy + định dạng → làm 2 việc: cập nhật bảng tổng vào `BAO_CAO_THANG`, và **xuất 1 file gộp toàn bộ ngày đã tính trong tháng**.

### 5. Xuất báo cáo (ngày cụ thể)
Như mục 4 nhưng cho khoảng ngày tự do. Chỉ xuất được ngày **đã tính xong**.

**Layout file xuất** (chung cho mục 4 và 5, bám file `Kiểm tra Qdu` gốc): mỗi ngày 1 tab; A1 = `MWh`; cột A là chu kỳ dạng `01 [00:00-00:30]`; các tổ máy nằm **liền cột** (B:K = tổ 1, L:U = tổ 2) với thứ tự `Qdd | Qdd_V | Qdc | Qmp | Qdư | Qdư âm/dương | P_Qdc | Ngưỡng dưới | Ngưỡng trên | Ghi chú`; cuối bảng có hàng **Tổng ngày**. File lưu cùng thư mục Drive với Sheet, sidebar hiện link tải.

Hai điểm bám sát bản gốc:
- **Tổ máy chưa có dữ liệu vẫn giữ nguyên bảng** (để trống, không có hàng tổng) — không bị mất bảng như trước.
- Cột **Qdư âm/dương** chỉ ghi `âm` / `dương`; khi Qdư = 0 (trong dải dung sai) thì để dấu **`-`**, không ghi chữ "trong ±3%" (chữ đó chỉ dùng trong sheet `KET_QUA` nội bộ cho dễ đọc).

### 6. Dọn dữ liệu cũ
Dùng khi bàn giao Sheet cho người khác, hoặc khi dữ liệu tích luỹ nhiều tháng. Giữ nguyên `CAI_DAT` và kết quả của **N ngày gần nhất** (để P0 vẫn tự suy được), xoá phần còn lại. Có hộp xác nhận trước khi chạy.

## Cảnh báo tự động

**Lệnh "0-0"** (CS ra lệnh = CS hoàn thành = 0, thường là trip/ngừng sự cố): theo quy tắc nghiệp vụ đã xác nhận, loại lệnh này **không được tính** — nhưng hệ thống sẽ **cảnh báo rõ** để người vận hành biết Qdd giai đoạn đó có thể cao hơn thực tế và tự xử lý. Hệ thống chỉ cảnh báo, không tự sửa số (xem [docs/09_Test_Cases.md](../../docs/09_Test_Cases.md)).

## Cập nhật code sau này

```bash
cd src/NhaMay-Mau-Template
npx clasp push
```

Khi `QDD-Core-Library` ra version mới, vào Apps Script Editor của Sheet này → **Libraries** → chọn version mới → Save (không tự động, đúng theo kiến trúc đã chọn).
