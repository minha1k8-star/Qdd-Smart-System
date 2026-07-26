# Phương án kỹ thuật — QDD Smart System

Tài liệu này giải thích **hệ thống tính như thế nào và vì sao kết quả chính xác**, viết cho người phụ trách kỹ thuật/nghiệp vụ đọc — không cần biết lập trình.

Dùng làm căn cứ để biên soạn hồ sơ đề nghị công nhận sáng kiến. Nội dung dành riêng cho hồ sơ (tên sáng kiến, hiệu quả kinh tế, khả năng nhân rộng) ở [17_Thuyet_Minh_Sang_Kien.md](17_Thuyet_Minh_Sang_Kien.md). Đặc tả dành cho người lập trình ở [04_Algorithm_Specification.md](04_Algorithm_Specification.md).

---

## 1. Bài toán

Mỗi ngày, với mỗi tổ máy, cần xác định **công suất dư (Qdư)** — phần điện năng phát ra nằm ngoài dải dung sai ±3% so với công suất điều độ yêu cầu.

**Đầu vào:**

| Nguồn | Nội dung |
|---|---|
| Danh sách lệnh điều độ | Mỗi lệnh có: thời điểm bắt đầu thực hiện (BĐTH), công suất ra lệnh, công suất hoàn thành, loại lệnh (SO/MO), trạng thái hoàn thành/dừng |
| CSV công tơ Qdc | 48 giá trị điện năng giao, mỗi giá trị ứng với một chu kỳ 30 phút |
| CSV công tơ Qmp | 48 giá trị tương tự, công tơ đo phía khác |
| Cấu hình nhà máy | Tốc độ tăng/giảm tải (MW/phút), hệ số quy đổi Qdd_V, dung sai, mã công tơ từng tổ máy |

**Đầu ra:** bảng 48 chu kỳ, mỗi chu kỳ gồm Qdd, Qdd_V, Qdc, P_Qdc, ngưỡng trên/dưới, Qmp, **Qdư** và dấu hiệu âm/dương.

**Khó khăn cốt lõi**: lệnh điều độ chỉ cho biết *"lúc 19:53:17 yêu cầu về 460 MW"*, không cho biết công suất tại từng thời điểm. Tổ máy không nhảy tức thời từ công suất này sang công suất khác — nó **tăng/giảm tải dần theo một tốc độ nhất định**. Muốn tính điện năng trung bình mỗi 30 phút thì phải **dựng lại được đường cong công suất theo thời gian**, rồi lấy diện tích dưới đường cong đó.

---

## 2. Phương pháp tính — 5 bước

### Bước 1 — Chọn lệnh có hiệu lực và công suất hiệu lực

Không phải lệnh nào trong danh sách cũng được đưa vào tính. Một lệnh chỉ có hiệu lực khi: đúng ngày đang tính, đúng tổ máy, đã đánh dấu **Hoàn thành**, và công suất tương ứng lớn hơn 0.

Với lệnh có hiệu lực, **công suất hiệu lực** được chọn theo loại lệnh:

| Trường hợp | Lấy công suất |
|---|---|
| Lệnh **SO** | CS hoàn thành |
| Lệnh **MO** bị **dừng** (cột lý do dừng có ghi chú) | CS hoàn thành |
| Lệnh **MO** bình thường | CS ra lệnh |

Đây là quy tắc nghiệp vụ đã được người phụ trách xác nhận, không phải quy ước kỹ thuật.

**Trường hợp đặc biệt — lệnh "0-0"**: lệnh có CS ra lệnh = CS hoàn thành = 0 (thường là trip/ngừng sự cố) **không được tính**. Hệ thống loại lệnh này nhưng **hiện cảnh báo** cho người vận hành biết, vì nếu tổ máy thực sự đã ngừng thì kết quả đang cao hơn thực tế và cần can thiệp tay.

### Bước 2 — Dựng đường cong công suất theo thời gian

Từ danh sách lệnh hiệu lực đã sắp theo thời gian, dựng đường cong công suất P(t) trong 24 giờ. Mỗi lệnh sinh ra hai đoạn:

- **Đoạn tăng/giảm tải (ramp)**: công suất đi tuyến tính từ giá trị hiện tại tới công suất mục tiêu.
- **Đoạn giữ tải (hold)**: giữ nguyên công suất mục tiêu cho tới khi có lệnh tiếp theo.

**Thời gian cần để tăng/giảm tải:**

```
Thời lượng (giây) = |Công suất mục tiêu − Công suất hiện tại| ÷ Tốc độ ramp (MW/phút) × 60
```

Ví dụ: từ 480,4 MW về 460 MW, chênh 20,4 MW, tốc độ 3,5 MW/phút → cần **349,7 giây** (gần 5 phút 50 giây).

**Điểm tinh vi nhất — lệnh mới đến khi đang tăng/giảm tải dở dang.** Khi đó công suất xuất phát của lệnh mới **không phải** công suất mục tiêu cũ (chưa đạt tới), cũng **không phải** công suất ra lệnh mới. Nó là **công suất thực tế tại đúng thời điểm bị cắt**, tính bằng nội suy tuyến tính:

```
P tại điểm cắt = P bắt đầu + (P mục tiêu − P bắt đầu) × (thời điểm cắt − thời điểm bắt đầu)
                                                        ─────────────────────────────────────
                                                        (thời điểm ramp xong − thời điểm bắt đầu)
```

Đây chính là chỗ dễ sai nhất khi tính tay, và là lý do chính khiến việc tính thủ công không thống nhất giữa những người thực hiện khác nhau.

### Bước 3 — Tính diện tích theo từng chu kỳ 30 phút

Một ngày chia thành 48 chu kỳ, mỗi chu kỳ 1.800 giây. Với mỗi chu kỳ, lấy phần đường cong công suất nằm trong chu kỳ đó và tính **diện tích dưới đường cong** bằng **công thức hình thang**:

```
Diện tích (MW·giây) = (P đầu đoạn + P cuối đoạn) ÷ 2 × Thời lượng đoạn
```

Cộng diện tích của mọi đoạn trong chu kỳ, rồi chia cho 1.800 giây để ra công suất trung bình:

```
Qdd (MW) = Tổng diện tích (MW·giây) ÷ 1.800
```

### Bước 4 — Quy đổi và so dung sai để ra Qdư

| Đại lượng | Công thức |
|---|---|
| Qdd_V (MWh) | `Qdd ÷ 2 × hệ số Qdd_V` (mặc định 0,9188) |
| Qdc (MWh) | Giá trị công tơ Qdc của chu kỳ `÷ 1000` |
| P_Qdc (MW) | `Qdc × 2` |
| Qmp (MWh) | Giá trị công tơ Qmp của chu kỳ `÷ 1000` |
| Ngưỡng dưới / trên | `Qdd × 0,97` và `Qdd × 1,03` (dung sai ±3%) |
| **Qdư (MWh)** | Nếu `P_Qdc` **nằm ngoài** dải ngưỡng: `Qmp − Qdd_V`. Nếu nằm trong dải: **0** |

Dấu hiệu ghi **âm** khi P_Qdc thấp hơn ngưỡng dưới, **dương** khi cao hơn ngưỡng trên.

### Bước 5 — Chuyển tiếp công suất qua nửa đêm

Công suất đầu ngày (P0) của một ngày chính là **công suất tại đúng thời điểm 24:00 của ngày liền trước**. Hệ thống tự ghi giá trị này sau mỗi lần tính, nên chỉ ngày đầu tiên sử dụng mới cần nhập tay.

Trường hợp khó: **lúc 24:00 tổ máy vẫn đang tăng/giảm tải dở dang**. Khi đó ngày hôm sau không được giữ nguyên công suất, mà phải **chạy tiếp phần ramp còn lại** cho tới khi đạt mục tiêu rồi mới giữ. Hệ thống ghi lại cả công suất tại 24:00 lẫn **mục tiêu ramp còn dở**, và tự nối tiếp sang ngày sau.

> **Lưu ý quan trọng**: P0 là công suất *tại* 24:00, **không phải** Qdd của chu kỳ 48. Qdd chu kỳ 48 là công suất *trung bình* khoảng 23:30–24:00 — hai giá trị này khác nhau khi đang ramp, và lấy nhầm sẽ làm sai toàn bộ ngày hôm sau.

---

## 3. Ví dụ tính đầy đủ bằng số thật

Lấy **chu kỳ 40 (19:30–20:00) ngày 17/07/2026, tổ S1** — trường hợp có lệnh cắt ngang, tức trường hợp khó.

**Dữ liệu:**
- Công suất đang giữ: **480,4 MW** (theo lệnh trước đó lúc 18:50:16, đã đạt mục tiêu).
- Lệnh mới `G14001.2026.2983` lúc **19:53:17**: loại MO, **bị dừng**, CS ra lệnh 435,7 — CS hoàn thành **460 MW**.
- Tốc độ ramp: 3,5 MW/phút.

**Bước 1** — lệnh MO bị dừng → lấy **CS hoàn thành = 460 MW** làm công suất mục tiêu (không lấy 435,7).

**Bước 2** — dựng đường cong trong chu kỳ 19:30–20:00 (giây 70.200 → 72.000):

| Đoạn | Từ → đến | Công suất | Thời lượng |
|---|---|---|---|
| Giữ tải | 19:30:00 → 19:53:17 | 480,4 MW | 1.397 giây |
| Giảm tải | 19:53:17 → 19:59:07 | 480,4 → 460 MW | 349,7 giây |
| Giữ tải | 19:59:07 → 20:00:00 | 460 MW | 53,3 giây |

Thời lượng giảm tải: `(480,4 − 460) ÷ 3,5 × 60 = 349,7 giây`.

**Bước 3** — tính diện tích từng đoạn:

| Đoạn | Phép tính | Diện tích (MW·giây) |
|---|---|---|
| Giữ tải | `480,4 × 1.397` | 671.118,8 |
| Giảm tải | `(480,4 + 460) ÷ 2 × 349,7` | 164.435,7 |
| Giữ tải | `460 × 53,3` | 24.511,4 |
| | **Tổng** | **860.065,9** |

```
Qdd = 860.065,9 ÷ 1.800 = 477,8144 MW
```

**Kết quả hệ thống tính ra: 477,8144 MW** — trùng khít với phép tính tay ở trên.

> Đáng chú ý: bảng tính tay đối chiếu ghi **479,036 MW** cho chu kỳ này. Chênh 1,22 MW. Sau khi tính lại độc lập bằng giấy bút như trên, xác định **hệ thống đúng, bảng tính tay sai** — người tính tay áp thời điểm dừng muộn hơn thực tế. Đây là một trong ba ô sai mà quá trình đối chiếu phát hiện được.

### Ví dụ chuyển tiếp qua nửa đêm

Ngày **23/07/2026**, lệnh cuối cùng lúc **23:50:46** yêu cầu về 534 MW trong khi tổ máy đang ở 622,5 MW:

```
Chênh công suất  = 622,5 − 534 = 88,5 MW
Thời gian cần    = 88,5 ÷ 3,5 × 60 = 1.517,1 giây
Thời điểm ramp xong = 23:50:46 + 1.517,1 giây = 00:16:03 ngày hôm sau
```

Vì thời điểm ramp xong vượt quá 24:00 nên:

```
Công suất tại 24:00 = 622,5 − 3,5 ÷ 60 × 554 giây = 590,1833 MW
Còn lại              = 963,1 giây ramp tiếp, mục tiêu 534 MW
```

Ngày 24/07 khởi động từ 590,1833 MW và **chạy tiếp** xuống 534 MW, thay vì giữ nguyên 590 MW.

**Kiểm chứng ảnh hưởng:**

| Chu kỳ ngày 24/07 | Có chuyển tiếp | Không chuyển tiếp | Bảng tính tay |
|---|---|---|---|
| 01 (00:00–00:30) | **557,55** | 598,23 | **557,62** |
| 02 (00:30–01:00) | **599,32** | 609,49 | **599,21** |

Nếu bỏ qua bước chuyển tiếp, chu kỳ đầu ngày sai khoảng **40 MW**.

---

## 4. Vì sao phương pháp này cho kết quả chính xác

**a) Công thức hình thang là chính xác tuyệt đối với bài toán này.**
Công thức hình thang thường chỉ là *xấp xỉ* khi tính diện tích dưới đường cong bất kỳ. Nhưng ở đây đường cong công suất chỉ gồm các đoạn **tuyến tính** (đang ramp) và **hằng số** (đang giữ tải) — với hai dạng này, công thức hình thang cho kết quả **đúng tuyệt đối về mặt toán học**, không phải xấp xỉ. Không có sai số rời rạc hoá, và không cần chia nhỏ thời gian hơn nữa.

**b) Không rời rạc hoá thời gian.**
Hệ thống không chia ngày thành các bước 1 phút hay 1 giây rồi cộng lại — cách đó luôn có sai số tích luỹ. Thay vào đó tính **chính xác từng đoạn theo giây thực tế của lệnh** (19:53:**17**, không làm tròn thành 19:53 hay 19:55).

**c) Cắt đoạn theo đúng ranh giới chu kỳ.**
Khi một đoạn ramp vắt qua ranh giới hai chu kỳ, hệ thống nội suy công suất tại đúng điểm ranh giới và chia diện tích cho hai chu kỳ theo đúng tỷ lệ — không dồn cả đoạn về một chu kỳ.

**d) Giữ nguyên độ chính xác qua chuỗi ngày.**
Kết quả hiển thị 2 số thập phân, nhưng **giá trị lưu bên trong giữ đầy đủ độ chính xác**. Vì công suất đầu ngày sau lấy từ cuối ngày trước, nếu làm tròn thật thì sai số sẽ cộng dồn qua chuỗi ngày liên tiếp.

**e) Sai số duy nhất còn lại là làm tròn của bảng đối chiếu.**
Trong 6 ngày đối chiếu gần nhất, chênh lệch lớn nhất giữa hệ thống và bảng tính tay là **0,12 MW** ở các chu kỳ đang ramp — bằng đúng mức làm tròn mà người tính tay sử dụng. Riêng **Qdư khớp tuyệt đối** ở toàn bộ 288 chu kỳ.

---

## 5. Kiến trúc hệ thống

```
                QDD-Core-Library
        (toàn bộ thuật toán, có kiểm soát phiên bản)
                        │
        ┌───────────────┼───────────────┐
        ▼               ▼               ▼
    Nhà máy A       Nhà máy B       Nhà máy C
   Sheet riêng     Sheet riêng     Sheet riêng
   cấu hình riêng  cấu hình riêng  cấu hình riêng
```

**Nền tảng**: Google Sheets + Google Apps Script. Không cần cài phần mềm, không cần máy chủ, không phát sinh chi phí bản quyền.

**Điểm mấu chốt — tách thuật toán thành thư viện dùng chung:**

| | Nếu mỗi nhà máy một bản sao | Với thư viện dùng chung |
|---|---|---|
| Sửa lỗi thuật toán | Phải cập nhật thủ công từng bản | Sửa một lần |
| Nguy cơ các bản lệch nhau | Cao, tăng dần theo thời gian | Không |
| Cấu hình riêng từng nhà máy | Có | Có — nằm ở Sheet riêng, không đụng thư viện |
| Dữ liệu vận hành | Rời rạc từng máy trạm | Trong Sheet của chính nhà máy đó |

**Phân chia trách nhiệm rõ ràng:**
- **Thư viện** chỉ tính, không biết gì về giao diện hay cách hiển thị.
- **Sheet của nhà máy** lo nhập liệu, cấu hình, hiển thị, xuất báo cáo.

Nhờ tách bạch như vậy, thuật toán kiểm thử được độc lập — chạy trên máy cá nhân, không cần tài khoản Google và không cần dữ liệu vận hành thật.

### Nhà máy có bao nhiêu tổ máy cũng dùng được

Danh sách tổ máy **không cố định trong code** mà suy ra từ chính bảng cấu hình của từng nhà máy. Nhà máy 3 tổ máy chỉ cần thêm 3 dòng cấu hình cho tổ thứ ba (mã công tơ Qdc, mã công tơ Qmp, nhãn báo cáo) — bảng điều khiển tự hiện thêm ô chọn và báo cáo tự thêm khối cho tổ đó, **không phải sửa phần mềm**.

Tương tự, mã công tơ, tốc độ tăng/giảm tải, hệ số quy đổi và dung sai đều là cấu hình riêng của từng nhà máy.

> **Lưu ý khi cấu hình mã công tơ**: tên file CSV có dạng `<ngày><tháng><năm 1 chữ số><mã công tơ>` — ví dụ `17076001.CSV` là ngày 17, tháng 07, **năm 2026 (chữ số 6)**, công tơ **001**. Cấu hình chỉ ghi phần mã công tơ (`001`), không ghi kèm chữ số năm, để sang năm sau hệ thống vẫn nhận đúng file.

### Vì sao không làm một website tập trung

Đã cân nhắc và **chủ động không chọn**, vì hai lý do:

1. **Yêu cầu kiểm tra được bằng mắt**: người làm nghiệp vụ cần nhìn thấy dữ liệu đầu vào, các bước trung gian và kết quả ngay trên bảng tính quen thuộc. Một website trả về con số cuối cùng là một "hộp đen" — khi số liệu có tranh cãi thì không truy vết được.
2. **Trách nhiệm dữ liệu**: mỗi nhà máy tự giữ dữ liệu vận hành của mình, không tập trung về một kho chung.

Chi tiết các phương án đã cân nhắc: [05_System_Architecture.md](05_System_Architecture.md).

---

## 6. Các cơ chế bảo vệ độ chính xác

Đây là phần quyết định việc hệ thống có đáng tin để dùng cho đối soát điện năng hay không.

**a) Bộ kiểm thử tự động — 45 trường hợp.**
Mỗi quy tắc nghiệp vụ và mỗi lỗi từng gặp đều được khoá lại bằng một trường hợp kiểm thử. Chạy dưới 1 giây, không cần dữ liệu thật. Có một trường hợp không đạt thì không được đưa vào sử dụng. Nhờ đó **mỗi lỗi chỉ có thể xảy ra đúng một lần**.

**b) Cảnh báo chủ động thay vì im lặng cho ra số sai.**
Hệ thống cảnh báo rõ khi gặp tình huống mà **máy không tự phân biệt được đúng/sai**:

| Cảnh báo | Ý nghĩa |
|---|---|
| Có lệnh 0-0 trong ngày | Tổ máy có thể đã ngừng do sự cố; theo quy tắc thì lệnh này không tính, nên Qdd đang cao hơn thực tế |
| Ngày không có lệnh hiệu lực nào | Có thể đúng (ngày không có lệnh), cũng có thể do quên nhập — chỉ người vận hành biết |
| Ngày được nối tiếp ramp từ hôm trước | Thông báo bình thường, để người dùng biết vì sao đầu ngày không phẳng |

Nguyên tắc: **chỉ cảnh báo, không tự sửa số liệu**. Quyền quyết định thuộc về người vận hành.

**c) Mặc định an toàn cho thao tác không hoàn tác được.**
Tuỳ chọn xoá dữ liệu nguồn sau khi tính **mặc định tắt**. Muốn xoá thì phải chủ động tick.

**d) Một luồng duy nhất cho mỗi công việc.**
Không có hai đường code làm cùng một việc (ví dụ "tính 1 ngày" tách khỏi "tính nhiều ngày"), vì đường ít dùng hơn sẽ là đường có lỗi mà không ai phát hiện.

**e) Đọc dữ liệu theo tên cột, không theo vị trí cột.**
File gốc đổi thứ tự cột hay thêm bớt cột phụ đều không ảnh hưởng. Đọc theo vị trí đã từng gây lệch cột và loại sạch lệnh mà không báo lỗi gì.

**f) Đối chiếu bắt buộc sau khi nhập lệnh.**
Sau mỗi lần nhập danh sách lệnh, hệ thống hiện **khoảng thời gian BĐTH thực tế đã ghi vào hệ thống** để người dùng đối chiếu ngay với file gốc.

---

## 7. Kết quả kiểm chứng

**Phương pháp**: đối chiếu kết quả hệ thống với **bảng tính tay độc lập** (do người làm nghiệp vụ tính, không phải xuất từ công cụ), so từng chu kỳ trong 48 chu kỳ.

**Hai đợt, 16 tổ hợp (ngày, tổ máy)** trên dữ liệu vận hành thật tháng 7/2026, nhà máy Duyên Hải 1.

| Ngày (tổ S1) | Sai lệch Qdd lớn nhất | Ghi chú |
|---|---|---|
| 17/07 | 0,0075 MW | mức làm tròn |
| 18/07 | 0,0538 MW | mức làm tròn |
| 19/07 | 0,0293 MW | mức làm tròn |
| 23/07 | 0,0359 MW | mức làm tròn |
| 24/07 | 0,1171 MW | ngày nhận chuyển tiếp qua nửa đêm |
| 25/07 | **0,0000 MW** | ngày không có lệnh nào |

**Qdư — con số dùng để đối soát — khớp tuyệt đối** ở toàn bộ 288 chu kỳ của 6 ngày.

**Các tình huống nghiệp vụ đã được kiểm chứng bằng dữ liệu thật:**

| Tình huống | Kiểm chứng bằng |
|---|---|
| Lệnh thay đổi công suất thông thường | Toàn bộ 6 ngày |
| Lệnh bị dừng (cả SO và MO) | 4 lệnh trong các ngày 17, 19, 23/07 |
| Lệnh mới cắt ngang khi đang tăng/giảm tải | Nhiều chu kỳ trong 6 ngày |
| Ngày không có lệnh nào | 25/07 |
| **Ramp vắt qua nửa đêm** | 23→24/07 |
| Lệnh 0-0 khi trip tổ máy | 07/07 (đợt 1) |

**Ngoài ra, quá trình đối chiếu phát hiện 3 ô sai trong chính bảng tính tay** (chênh ~1 MW). Đã tính tay độc lập lại từng chu kỳ để xác định hệ thống đúng. Điều này cho thấy hệ thống không chỉ nhanh hơn mà còn **phát hiện được sai sót của cách làm thủ công**.

Chi tiết đầy đủ: [15_Accuracy_Validation_2026-07.md](15_Accuracy_Validation_2026-07.md).

---

## 8. Giới hạn hiện tại

Nêu rõ để hồ sơ trung thực:

- **Mới kiểm chứng tổ S1.** Tổ S2 có mã công tơ và chuỗi công suất riêng, cần đối chiếu bổ sung khi có bảng tính tay tương ứng. Kết quả S1 khớp không tự động bảo đảm S2 khớp.
- **Quy tắc lệnh 0-0** hiện xử lý bằng cảnh báo, chưa tự động đưa công suất về 0 — đây là quy tắc nghiệp vụ cần cấp có thẩm quyền quyết định, không phải hạn chế kỹ thuật.
- **Phụ thuộc nền tảng Google**: cần tài khoản Google và kết nối mạng.
- **Chất lượng dữ liệu đầu vào quyết định kết quả.** Cụ thể: khi khởi động lại tổ máy sau sự cố, lệnh phải ghi đủ **CS ra lệnh = CS hoàn thành = tải thật** thì hệ thống mới nhận. Đây là việc cần điều chỉnh ở khâu nhập liệu vận hành.

## 9. Hướng phát triển

- Kiểm chứng tổ S2 và triển khai cho nhà máy thứ hai.
- Triển khai cho nhà máy thứ hai (hệ thống đã hỗ trợ số tổ máy bất kỳ — thêm tổ chỉ cần thêm dòng cấu hình).

---

## Phụ lục — Thuật ngữ

| Từ | Nghĩa |
|---|---|
| **Qdd** | Công suất điều độ trung bình của một chu kỳ 30 phút (MW), dựng lại từ lệnh điều độ |
| **Qdd_V** | Điện năng quy đổi từ Qdd (MWh) |
| **Qdc** | Điện năng đo được từ công tơ Qdc trong chu kỳ (MWh) |
| **P_Qdc** | Công suất tương ứng với Qdc (MW) |
| **Qmp** | Điện năng đo được từ công tơ Qmp trong chu kỳ (MWh) |
| **Qdư** | Điện năng dư — phần nằm ngoài dải dung sai ±3% (MWh) |
| **BĐTH** | Thời điểm bắt đầu thực hiện lệnh |
| **SO / MO** | Loại lệnh điều độ |
| **Ramp** | Quá trình tăng/giảm tải theo tốc độ quy định (MW/phút) |
| **P0** | Công suất đầu ngày, bằng công suất tại đúng 24:00 của ngày trước |
| **Chu kỳ** | Khoảng 30 phút; một ngày có 48 chu kỳ |
