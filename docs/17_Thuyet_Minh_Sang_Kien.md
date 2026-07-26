# Thuyết minh sáng kiến — Bản thảo kỹ thuật

Tài liệu này cung cấp **phần nội dung kỹ thuật** để người phụ trách biên soạn hồ sơ đề nghị công nhận sáng kiến.

> Phần giải thích **hệ thống tính như thế nào và vì sao kết quả chính xác** (kèm ví dụ tính bằng số thật) nằm ở [18_Phuong_An_Ky_Thuat.md](18_Phuong_An_Ky_Thuat.md) — nên đọc tài liệu đó trước.

> ⚠️ **Lưu ý về số liệu**: các mục đánh dấu **[CẦN ĐIỀN]** là số liệu vận hành/kinh tế mà chỉ đơn vị mới có (thời gian thao tác thực tế, số nhân sự, số nhà máy dự kiến áp dụng…). Tài liệu này **không tự suy đoán** những con số đó — mọi số liệu ghi sẵn bên dưới đều là **kết quả kỹ thuật đã kiểm chứng được**, có thể dẫn chứng lại từ mã nguồn và dữ liệu thật.

---

## 1. Tên sáng kiến

**Hệ thống tự động tính toán và lập báo cáo công suất điều độ (Qdd) và công suất dư (Qdư) cho tổ máy nhiệt điện trên nền Google Sheets + Apps Script, dùng chung thư viện thuật toán cho nhiều nhà máy.**

Tên gọi hệ thống: **QDD Smart System (QSS)**.

## 2. Lĩnh vực áp dụng

Công tác vận hành — đo đếm điện năng — đối soát công suất điều độ tại nhà máy nhiệt điện.

## 3. Thực trạng trước khi có sáng kiến

**3.1. Giai đoạn tính thủ công**

Việc xác định Qdd/Qdư được thực hiện thủ công trên Excel: đối chiếu danh sách lệnh điều độ với dữ liệu công tơ, dựng lại đường cong công suất, tính diện tích từng chu kỳ 30 phút (48 chu kỳ/ngày), rồi so với dải dung sai ±3%.

Hạn chế:
- Khối lượng tính lớn, dễ sai sót — đặc biệt khi trong ngày có nhiều lệnh, hoặc lệnh mới đến khi tổ máy đang tăng/giảm tải dở dang (phải nội suy công suất tại điểm cắt).
- Không thống nhất giữa những người thực hiện khác nhau.
- Khó truy vết: khi số liệu lệch, không biết sai ở khâu nhập liệu hay khâu tính.

**3.2. Vì sao chưa có công cụ tự động nào được đưa vào sử dụng**

Bài toán này khó tự động hoá bằng bảng tính thông thường vì ba lý do:

- **Phụ thuộc chuỗi ngày**: công suất đầu ngày phải lấy từ công suất tại đúng 24:00 của ngày trước, và nếu tổ máy còn đang tăng/giảm tải dở dang lúc nửa đêm thì phần dở dang phải chạy tiếp sang ngày sau. Một bảng tính gắn với "ngày đang tính" không xử lý được chuỗi liên tục này.
- **Nội suy khi lệnh mới ngắt ngang**: khi lệnh mới đến giữa lúc tổ máy đang tăng/giảm tải, phải nội suy công suất tại đúng điểm cắt rồi mới ramp tiếp — không phải phép tính viết được bằng một công thức đơn giản.
- **Khối lượng**: 48 chu kỳ × 2 tổ máy × 30 ngày cho mỗi tháng báo cáo.

Kết quả là công việc vẫn phải làm tay, với các hạn chế ở mục 3.1.

## 4. Nội dung giải pháp

### 4.1. Kiến trúc

Chuyển sang **Google Sheets + Apps Script**, với điểm mấu chốt là tách phần thuật toán thành **một thư viện dùng chung** (`QDD-Core-Library`):

```
              QDD-Core-Library  (thuật toán dùng chung, có kiểm soát phiên bản)
                      │
      ┌───────────────┼───────────────┐
      ▼               ▼               ▼
  Nhà máy A       Nhà máy B       Nhà máy C
  (Sheet riêng)   (Sheet riêng)   (Sheet riêng)
  cấu hình riêng  cấu hình riêng  cấu hình riêng
```

- **Thuật toán viết một lần, dùng chung** — sửa lỗi hoặc cập nhật quy tắc nghiệp vụ chỉ cần làm ở thư viện, các nhà máy cập nhật phiên bản khi sẵn sàng.
- **Dữ liệu và cấu hình của mỗi nhà máy hoàn toàn độc lập** (tốc độ ramp, hệ số quy đổi, dung sai, mã công tơ từng tổ máy) — không có kho dữ liệu chung, không ảnh hưởng lẫn nhau.
- Không phụ thuộc hệ điều hành hay phần mềm cài trên máy trạm; chạy trên trình duyệt.

Lý do chọn kiến trúc này thay vì một website tập trung: xem [05_System_Architecture.md](05_System_Architecture.md) — quyết định dựa trên yêu cầu **vừa nhân rộng dễ, vừa giữ được khả năng kiểm tra số liệu bằng mắt** của người làm nghiệp vụ.

### 4.2. Các chức năng chính

1. Nhập danh sách lệnh điều độ **thẳng từ file Excel gốc** — không cần mở file, không cần cắt sửa cột; hệ thống tự tìm bảng dữ liệu trong file và nhận diện cột theo tên tiêu đề.
2. Nhập dữ liệu công tơ từ file CSV — **một lần chọn được nhiều file**, hệ thống tự nhận diện tổ máy/loại dữ liệu theo tên file và tự đọc ngày trong file.
3. Tính Qdd/Qdư cho **một ngày hoặc nhiều ngày liên tiếp, nhiều tổ máy cùng lúc**.
4. Tự động chuyển tiếp công suất sang ngày kế tiếp, kể cả khi tổ máy **đang tăng/giảm tải dở dang lúc 24:00**.
5. **Cảnh báo tự động** các tình huống dễ gây sai số âm thầm (xem 4.3).
6. Xuất báo cáo theo ngày, khoảng ngày, hoặc cả tháng ra file Excel/PDF **theo đúng mẫu báo cáo hiện hành**.
7. Hướng dẫn sử dụng tích hợp ngay trong file.

### 4.3. Tính mới, tính sáng tạo

**a) Nhân rộng bằng thư viện dùng chung thay vì sao chép công cụ**
Điểm khác biệt căn bản so với cách làm cũ (mỗi nơi một bản sao). Một lỗi thuật toán được sửa một lần sẽ áp dụng cho mọi nhà máy, thay vì phải rà soát và cập nhật thủ công từng bản — vốn là nguyên nhân khiến các bản dùng lâu ngày bị lệch nhau.

**b) Bỏ được ràng buộc "mỗi lần một ngày"**
Nhờ thiết kế lại: mỗi lần tính là một lời gọi độc lập nhận tham số (ngày, tổ máy), thay vì phụ thuộc một ô cấu hình dùng chung. Nhờ đó tính được hàng loạt, và **báo cáo tháng lấy trực tiếp từ dữ liệu gốc** — không cần cơ chế "đóng băng" kết quả từng ngày như trước.

**c) Cảnh báo chủ động thay vì im lặng cho ra số sai**
Hệ thống phát hiện và cảnh báo rõ các tình huống mà công cụ cũ âm thầm cho ra kết quả trông "bình thường" nhưng sai:
- Ngày có **lệnh 0-0** (ngừng tổ máy do sự cố) — theo quy tắc nghiệp vụ lệnh này không được tính, nhưng nếu tổ máy thực sự đã ngừng thì Qdd đang cao hơn thực tế.
- Ngày **không có lệnh nào** — có thể do quên nhập, kết quả sẽ phẳng theo công suất đầu ngày.
- File CSV **không nhận diện được** tổ máy/loại dữ liệu theo tên file — báo riêng theo tên file thay vì lưu nhầm.

Nguyên tắc: hệ thống **chỉ cảnh báo, không tự sửa số liệu** — quyền quyết định thuộc về người vận hành.

**d) Có bộ kiểm thử tự động bảo vệ thuật toán**
Toàn bộ quy tắc nghiệp vụ được khoá lại bằng **45 trường hợp kiểm thử tự động**, chạy được trên máy cá nhân mà không cần dữ liệu vận hành thật. Mỗi lỗi phát hiện trong quá trình triển khai đều được bổ sung một trường hợp kiểm thử tương ứng để không tái diễn. Nhờ đó mỗi lỗi chỉ xảy ra đúng một lần.

**e) Toàn bộ tri thức được tài liệu hoá công khai**
Quy tắc nghiệp vụ, đặc tả thuật toán, lịch sử lỗi và lý do của từng quyết định thiết kế đều được ghi lại trong kho mã nguồn — để người tiếp nhận sau này (hoặc nhà máy khác) hiểu được hệ thống mà không phụ thuộc vào trí nhớ của người xây dựng.

## 5. Kết quả kiểm chứng

**5.1. Đối chiếu với số liệu tính tay độc lập**

Dùng dữ liệu vận hành thật của Nhà máy Duyên Hải 1 (tháng 7/2026), đối chiếu kết quả hệ thống với bảng tính tay:

| Ngày | Tổ máy | Sai lệch Qdd lớn nhất | Ghi chú |
|---|---|---|---|
| 17/07/2026 | S1 | 0,0075 MW | mức làm tròn |
| 18/07/2026 | S1 | 0,0538 MW | mức làm tròn |
| 19/07/2026 | S1 | 0,0293 MW | mức làm tròn |
| 23/07/2026 | S1 | 0,0359 MW | mức làm tròn |
| 24/07/2026 | S1 | 0,1171 MW | ngày nhận chuyển tiếp qua nửa đêm |
| 25/07/2026 | S1 | **0,0000 MW** | ngày không có lệnh nào |

**Qdư — con số dùng để đối soát — khớp tuyệt đối** ở toàn bộ 48 chu kỳ của cả 6 ngày.

Quá trình đối chiếu còn **phát hiện 3 ô sai trong chính bảng tính tay** (ngày 17/07 chu kỳ 40, ngày 19/07 chu kỳ 42, lệch ~1 MW). Đã tính tay độc lập lại từng chu kỳ đó để xác định kết quả hệ thống là đúng — cho thấy hệ thống không chỉ nhanh hơn mà còn phát hiện được sai sót của cách làm thủ công.

Ngoài ra, một bản tái hiện thuật toán độc lập bằng Python (dùng để kiểm tra chéo đặc tả) đã đối chiếu **10 ngày × 2 tổ máy**. Chi tiết toàn bộ: [15_Accuracy_Validation_2026-07.md](15_Accuracy_Validation_2026-07.md).

**5.2. Phát hiện được khoảng trống trong quy trình hiện hành**

Trong quá trình đối chiếu, hệ thống phát hiện trường hợp **ngày 07/07/2026 tổ S2**: tổ máy ngừng do sự cố nhưng lệnh ghi nhận dạng "0-0" nên theo quy tắc hiện hành không được tính vào — dẫn tới Qdd giai đoạn đó cao hơn thực tế. Đây là khoảng trống **đã tồn tại từ trước**, chỉ được phát hiện nhờ việc đối chiếu có hệ thống. Hệ thống mới đã bổ sung cảnh báo cho tình huống này.

**5.3. Kiểm thử tự động**

45/45 trường hợp kiểm thử đạt, bao phủ: chọn công suất hiệu lực theo loại lệnh, nội suy khi lệnh mới cắt ngang, tích phân diện tích 48 chu kỳ, quy đổi Qdd_V/Qdc/Qmp/Qdư theo dải dung sai, chuyển tiếp công suất qua nửa đêm, và các lỗi kỹ thuật đã từng gặp.

## 6. Hiệu quả

### 6.1. Hiệu quả kỹ thuật (đã kiểm chứng)

- Độ chính xác: **Qdư khớp tuyệt đối** với cách tính thủ công trên dữ liệu thật; sai lệch Qdd nằm trong mức làm tròn (≤ 0,12 MW).
- Phát hiện được sai sót của chính bảng tính tay (3 ô lệch ~1 MW) — tăng độ tin cậy của số liệu đối soát.
- Không phụ thuộc hệ điều hành hay phần mềm cài trên máy trạm.
- Tính được nhiều ngày, nhiều tổ máy trong một thao tác; báo cáo tháng lập trực tiếp từ dữ liệu gốc.
- Có cảnh báo chủ động cho các tình huống dễ sai sót.
- Có bộ kiểm thử tự động và tài liệu đầy đủ, giảm rủi ro khi bàn giao/nhân sự thay đổi.

### 6.2. Hiệu quả kinh tế **[CẦN ĐIỀN]**

Cần đơn vị bổ sung số liệu thực tế:

- Thời gian trung bình để tính và lập báo cáo Qdd/Qdư **cho 1 ngày** theo cách hiện tại: **[CẦN ĐIỀN]** phút
- Thời gian tương ứng khi dùng hệ thống mới: **[CẦN ĐIỀN]** phút
- Số ngày phải xử lý trong 1 tháng: **[CẦN ĐIỀN]**
- Số nhân sự tham gia: **[CẦN ĐIỀN]**
- Đơn giá giờ công áp dụng: **[CẦN ĐIỀN]**

→ Từ đó quy ra số giờ công và giá trị làm lợi hằng năm cho một nhà máy, rồi nhân với số nhà máy dự kiến áp dụng.

Ngoài ra có thể tính đến phần **giảm rủi ro sai số trong đối soát điện năng** — mức độ ảnh hưởng kinh tế của một sai sót Qdư cần đơn vị đánh giá: **[CẦN ĐIỀN]**.

**Chi phí đầu tư**: không phát sinh chi phí bản quyền phần mềm (dùng nền tảng Google Workspace/tài khoản Google sẵn có), không cần đầu tư máy chủ.

### 6.3. Khả năng nhân rộng

- Áp dụng cho **tổ máy khác trong cùng nhà máy** và **các nhà máy khác** có cùng cấu trúc dữ liệu lệnh điều độ và công tơ, chỉ cần thay đổi phần cấu hình (mã công tơ, tốc độ ramp, hệ số, dung sai).
- Triển khai cho một nhà máy mới mất khoảng **5–20 phút**, có hướng dẫn từng bước: [16_Huong_Dan_Trien_Khai.md](16_Huong_Dan_Trien_Khai.md).
- Số nhà máy dự kiến áp dụng: **[CẦN ĐIỀN]**

## 7. Phạm vi và giới hạn cần nêu rõ

Để hồ sơ trung thực, nên nêu cả những điểm chưa hoàn thiện:

- Quy tắc **lệnh 0-0** (ngừng do sự cố) hiện được xử lý bằng **cảnh báo**, chưa tự động đưa công suất về 0 — do đây là quy tắc nghiệp vụ cần cấp có thẩm quyền quyết định, không phải vấn đề kỹ thuật.
- Kết quả đối chiếu với bảng tính tay hiện mới thực hiện cho **tổ máy S1**; tổ S2 cần đối chiếu bổ sung khi có bảng tính tay tương ứng.
- Hệ thống phụ thuộc nền tảng Google; cần tài khoản và kết nối mạng.

## 8. Tài liệu kèm theo

| Nội dung | Tài liệu |
|---|---|
| Quy tắc nghiệp vụ (R01–R14) | [03_Business_Rules.md](03_Business_Rules.md) |
| **Phương án kỹ thuật** (phương pháp tính, ví dụ số, vì sao chính xác) | [18_Phuong_An_Ky_Thuat.md](18_Phuong_An_Ky_Thuat.md) |
| Đặc tả thuật toán (cho người lập trình) | [04_Algorithm_Specification.md](04_Algorithm_Specification.md) |
| Kiến trúc hệ thống và lý do lựa chọn | [05_System_Architecture.md](05_System_Architecture.md) |
| Kết quả đối chiếu với dữ liệu thật | [15_Accuracy_Validation_2026-07.md](15_Accuracy_Validation_2026-07.md) |
| Hướng dẫn triển khai cho nhà máy mới | [16_Huong_Dan_Trien_Khai.md](16_Huong_Dan_Trien_Khai.md) |
| Lịch sử lỗi và bài học kỹ thuật | [14_Knowledge_Transfer.md](14_Knowledge_Transfer.md) |
| Mã nguồn đầy đủ | https://github.com/minha1k8-star/Qdd-Smart-System |
