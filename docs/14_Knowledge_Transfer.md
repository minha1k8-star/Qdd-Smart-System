# Knowledge Transfer

Ghi lại các lỗi và quyết định kỹ thuật đã thực sự gặp trong quá trình phát triển, để người/AI tiếp nhận dự án sau này không lặp lại.

Đây là tài liệu **kinh nghiệm**, không phải đặc tả. Đặc tả thuật toán ở [04_Algorithm_Specification.md](04_Algorithm_Specification.md), quy tắc nghiệp vụ ở [03_Business_Rules.md](03_Business_Rules.md).

---

## Lỗi âm thầm — nhóm nguy hiểm nhất

Đặc điểm chung: **kết quả sai nhưng không có lỗi nào hiện ra**. Người dùng tin vào con số sai. Cả ba lỗi dưới đây đều thuộc nhóm này và đều đã xảy ra thật.

### `instanceof` không dùng được qua ranh giới Apps Script Library

**Triệu chứng**: mọi ngày tính ra Qdd phẳng đúng bằng P0 suốt 48 chu kỳ, không lệnh nào được áp dụng, không báo lỗi gì.

**Nguyên nhân**: `CommandFilter.selectEffective` kiểm tra `c.bdth instanceof Date`. Khi Sheet gọi sang `QDD-Core-Library`, **mỗi scope có constructor `Date` riêng**, nên Date tạo ở script gọi KHÔNG thoả `instanceof Date` bên trong thư viện → toàn bộ lệnh bị loại âm thầm.

**Quy tắc rút ra**: không dùng `instanceof` cho bất kỳ dữ liệu nào truyền qua ranh giới Library (`Date`, `Array`, `Error`…). Dùng duck typing: `typeof v.getTime === 'function'`.

### Múi giờ khi chuyển đổi file Excel qua Drive

**Triệu chứng**: nhập danh sách lệnh từ file Excel xong, mọi lệnh lệch **đúng +14 giờ** — lệnh buổi tối nhảy sang ngày hôm sau, ngày 17/07 chỉ còn 2 lệnh thay vì 4.

**Nguyên nhân**: ngày-giờ trong `.xlsx` là giá trị "trần", không kèm múi giờ. Bản Google Sheets tạm do Drive tạo ra khi chuyển đổi lấy **múi giờ mặc định của tài khoản Google** (`America/Los_Angeles`, −07 giờ hè), trong khi Sheet đích dùng `Asia/Ho_Chi_Minh` (+07) → chênh đúng 14 giờ.

**Khắc phục**: `alignTimeZoneWithTargetSheet_` đặt lại múi giờ file tạm trước khi đọc. **Chặn tái diễn**: sau mỗi lần nhập, sidebar hiện khoảng BĐTH thực tế đã ghi vào `LENH` để đối chiếu bằng mắt với file gốc.

### P0 lấy nhầm Qdd chu kỳ 48

**Triệu chứng**: ngày 19/07 lệch đều +29,43 MW ở 36 chu kỳ đầu, dù ngày 17 và 18 khớp.

**Nguyên nhân**: P0 của ngày sau lấy **Qdd chu kỳ 48** của ngày trước — đó là công suất *trung bình* khoảng 23:30–24:00, không phải công suất *tại* 24:00. Ngày 18/07 lúc đó đang giảm tải 533,1 → 435,7 nên trung bình là 465,131 còn giá trị cuối mới là 435,7.

**Quy tắc rút ra**: P0 luôn là công suất **tại đúng thời điểm 24:00**, tính từ đoạn công suất cuối ngày. Không suy từ bất kỳ giá trị trung bình nào.

---

## Bài học về thuật toán

### Mốc bắt đầu ramp là BĐTH, không phải thời điểm hoàn thành

Kiểm thử từng cho kết quả ngày 18 đúng nhưng ngày 17 lệch lớn ở khung 19h30–20h00. Nguyên nhân: dùng sai cột làm mốc bắt đầu ramp — phải là **Thời điểm BĐTH**, không phải Thời điểm hoàn thành. Đây là lý do R04 nhấn mạnh điểm này.

**Cách chẩn đoán**: lệch tập trung ở **một khung giờ cụ thể** (không phải lệch cả ngày) thường là lỗi chọn sai mốc thời gian của một lệnh, không phải lỗi công thức tổng thể. Ngược lại, lệch **đều cả ngày** thường là sai P0.

### Ramp bị cắt giữa chừng phải lấy giá trị nội suy

Khi một đoạn ramp bị cắt (lệnh mới đến, hoặc hết ngày), công suất cuối đoạn phải là **giá trị nội suy tại điểm cắt**, không phải công suất mục tiêu. Ghi sai chỗ này làm sai độ dốc đoạn → sai diện tích → sai Qdd. Lỗi này từng tồn tại song song với lỗi P0 ở trên và bị che khuất bởi nó.

### Ranh giới trách nhiệm giữa Ramp Engine và tầng dựng đoạn

Ramp Engine tính "nếu không bị gì cản thì bao giờ ramp xong" — **không giới hạn trong phạm vi một ngày**. Việc cắt theo ranh giới 24:00 và mang phần dở dang sang ngày sau (R07) là trách nhiệm của tầng dựng đoạn + carry-over. Trộn lẫn hai việc này là nguồn của nhiều lỗi tinh vi.

---

## Bài học nghiệp vụ

### Lệnh "0-0" không được tính — và hệ quả vận hành

Khi tổ máy dừng do sự cố (trip) và được ghi bằng một lệnh có cả CS ra lệnh và CS hoàn thành đều bằng 0, hệ thống **chủ động loại bỏ** lệnh này. Đây là quy tắc nghiệp vụ có chủ đích, đã xác nhận với người phụ trách — không phải lỗi. Tương tự, khi khởi động lại sau sự cố, lệnh chỉ được coi là hoàn thành khi **CS ra lệnh và CS hoàn thành cùng bằng một giá trị tải thật** (ví dụ 435,7-435,7).

**Hệ quả**: nếu dữ liệu chỉ có đúng lệnh 0-0, hệ thống sẽ giữ nguyên công suất trước đó cho hết ngày thay vì đưa về 0. Vì vậy hệ thống **cảnh báo rõ ràng** khi phát hiện lệnh 0-0 trong ngày đang tính, để người vận hành biết mà xử lý tay — cảnh báo, không tự ý sửa số.

**Việc cần làm ở khâu vận hành**: ghi đủ CS hoàn thành khi khởi động lại tổ máy, để dữ liệu tương lai đủ điều kiện tính tự động.

### Lệnh bị dừng dùng CS hoàn thành

Người phụ trách nghiệp vụ xác nhận: *"Ở cột lý do dừng/hủy có ghi chú thì cột công suất ra lệnh bằng CS hoàn thành."* Đây chính là R03 (và R01 cho lệnh SO).

### Ngày không có lệnh nào là hợp lệ

Qdd giữ nguyên P0 suốt 48 chu kỳ. Nhưng tình huống này **nhìn giống hệt** trường hợp quên nhập lệnh, mà máy không phân biệt được — nên phải cảnh báo để người vận hành tự xác nhận.

### Âm/dương của Qdư

Cột "Qdư âm/dương" chỉ ghi **âm**, **dương**, hoặc để dấu `-` khi Qdư = 0 (nằm trong dải dung sai) — không hiển thị số 0 gây nhầm lẫn.

### Báo cáo phải có cả S1 và S2

Báo cáo phải có cả hai tổ máy cạnh nhau theo đúng mẫu gốc; tổ nào chưa có dữ liệu thì **để bảng trống**, không bỏ hẳn bảng và không báo lỗi.

---

## Bài học về vận hành hệ thống

### Sửa thư viện phải tạo version mới

Sửa code trong `QDD-Core-Library` mà quên `npx clasp version` thì Sheet vẫn chạy code cũ, không có dấu hiệu gì. Sau khi tạo version phải cập nhật số version trong `src/NhaMay-Mau-Template/appsscript.json` rồi push lại.

### Mặc định an toàn cho thao tác không hoàn tác được

Tuỳ chọn "dọn lệnh + CSV sau khi tính" ban đầu **mặc định bật**, làm dữ liệu nguồn bị xoá ngay sau mỗi lần tính nếu người dùng không để ý. Đã đổi thành mặc định tắt: xoá là không hoàn tác được, trong khi nhu cầu tính lại/đối chiếu là thường xuyên.

### Hiển thị làm tròn, giá trị giữ nguyên

Kết quả hiển thị 2 số thập phân bằng **định dạng ô**, không làm tròn giá trị thật. P0 ngày kế tiếp và hàng Tổng ngày đọc lại chính các ô này — làm tròn hẳn số gốc sẽ cộng dồn sai số qua chuỗi ngày liên tiếp.

### Chỉ giữ một luồng cho mỗi việc

Sidebar từng có "Lưu CSV 1 file" song song với "Tải CSV hàng loạt", và "Tính 1 ngày" song song với "Tính hàng loạt". Hai luồng làm cùng một việc nghĩa là hai đường code phải kiểm chứng, và đường ít dùng hơn sẽ là đường có lỗi. Đã gộp còn một luồng duy nhất cho mỗi việc.

### Đọc dữ liệu sheet theo TÊN cột, không theo vị trí

Từng có lần sheet `LENH` còn cấu trúc cột cũ trong khi code đọc theo vị trí cột mới → lệch cột → **mọi lệnh bị loại**, Qdd phẳng bằng P0. Mọi chỗ đọc dữ liệu hiện đều dò theo tên tiêu đề.
