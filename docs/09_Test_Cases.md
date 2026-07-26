# Test Cases (UAT) — QDD Smart System

Nguồn gốc: sheet `KIEM_THU_UAT` trong workbook chính thức (`legacy/CongCu_Tinh_Qdd_Qdu_v1_3_0_AllInOne.xlsm`) — bộ case này viết cho **bản Excel/VBA**.

> Theo [AGENTS.md](../AGENTS.md): bất kỳ thay đổi nào ở [03_Business_Rules.md](03_Business_Rules.md) hoặc [04_Algorithm_Specification.md](04_Algorithm_Specification.md) đều phải có Test Case tương ứng ở đây (hoặc case mới nếu chưa có).

## Đối chiếu với bản Google Sheets (cập nhật 2026-07-26)

Bộ 31 case trên viết cho bản VBA, nên **không áp dụng nguyên vẹn** cho bản Google Sheets. Phân loại lại:

| Nhóm | Số case | Case |
|---|---|---|
| ✅ **Đã kiểm chứng** trên bản Google Sheets bằng dữ liệu vận hành thật | 10 | UAT-01, 03, 04, 05, 06, 07, 08, 12, 14, 32 |
| ✅ **Đã triển khai** (case vốn ghi "chưa triển khai") | 1 | UAT-34 |
| ⬜ **Không còn áp dụng** — đặc thù Excel/VBA hoặc chức năng đã bỏ có chủ đích | 15 | UAT-09, 10, 11, 13, 16, 17, 18, 19, 24, 25, 26, 27, 28, 29, 30, 31 |
| 🟡 **Còn phải kiểm** trên bản Google Sheets | 3 | UAT-02, 15, 33 |

**Chi tiết nhóm "đã kiểm chứng"** — bằng chứng ở [15_Accuracy_Validation_2026-07.md](15_Accuracy_Validation_2026-07.md), đợt 2:

| Case | Kiểm chứng bằng |
|---|---|
| UAT-01 lọc đúng ngày/tổ | File lệnh tổng hợp 63 lệnh nhiều ngày, mỗi ngày lọc ra đúng lệnh của mình |
| UAT-03 P0 tự động | Chuỗi 17→18→19/07 và 23→24→25/07, P0 tự ghi sau mỗi lần tính |
| UAT-04 ramp qua 00:00 | **23→24/07/2026** — lệnh 23:50:46, lúc 24:00 còn 590,18 MW, ngày sau chạy tiếp tới 534 MW |
| UAT-05 SO hợp lệ | Lệnh SO bị dừng ngày 19/07 (3002, 3005, 2998) → dùng CS hoàn thành |
| UAT-06 MO bình thường | Phần lớn lệnh trong 6 ngày đã đối chiếu |
| UAT-07 MO dừng sớm | Lệnh 2983 (17/07), 3035 và 3037 (23/07) |
| UAT-08 lệnh ngắt ramp | Các chu kỳ ramp bị lệnh mới cắt ngang, khớp bảng tính tay trong mức làm tròn |
| UAT-12 đủ Qdc/Qmp | 12 file CSV của 6 ngày, đủ 48 giá trị mỗi file |
| UAT-14 ngày không có lệnh | **25/07** — Qdd phẳng 435,7 suốt 48 chu kỳ, khớp tuyệt đối bảng tính tay |
| UAT-32 lệnh 0-0 không tính | Đã đóng từ đợt 1, xác nhận nghiệp vụ |

**Vì sao 15 case không còn áp dụng:**

| Case | Lý do |
|---|---|
| UAT-09, 10 (CSV Mac 1 cột / Windows 50 cột) | Lỗi của Excel khi mở CSV theo Regional Settings từng máy. Apps Script đọc CSV bằng `Utilities.parseCsv`, không phụ thuộc hệ điều hành |
| UAT-11 (CSV sai ngày) | Bản mới **đọc ngày từ chính nội dung file** thay vì bắt người dùng chọn rồi so sánh — khái niệm "CSV sai ngày" không còn |
| UAT-13, 25 (đối chiếu "Qdd file gốc", nút 4) | Chức năng đối chiếu thủ công của bản VBA, không phải một phần thuật toán ([04_Algorithm_Specification.md](04_Algorithm_Specification.md) mục 7). Bản mới không có |
| UAT-16, 26, 27 (backup, nút 11) | Bỏ có chủ đích: Google Sheets có sẵn lịch sử phiên bản (`Tệp → Lịch sử phiên bản`), làm backup thủ công là thừa |
| UAT-17, 18, 19 (31 / 45 / 60 lệnh) | Giới hạn 60 lệnh của bản VBA sinh ra từ lưới công thức 121 đoạn cố định. Bản mới **không có giới hạn này** — đã chạy thử 31/45/60/80/**120** lệnh/ngày, đều ra đủ 48 chu kỳ hợp lệ, dưới 5 ms |
| UAT-24, 30, 31 (xuất báo cáo tháng lỗi Mac/Windows) | Lỗi `merged cell 1004` và `Cannot run the macro` là lỗi riêng của Excel/VBA. Bản mới xuất file qua Google, không có macro |
| UAT-28, 29 (compile module, kiểm tra module tháng) | Khái niệm "module VBA" không tồn tại ở Apps Script |

**Ba case còn phải kiểm** (đây là phần nợ thật sự, không phải thủ tục giấy tờ):

| Case | Vì sao còn nợ |
|---|---|
| **UAT-15 — tổ S2** | Toàn bộ kiểm chứng đến nay mới chạy **S1**. Mã công tơ, P0 và lịch sử của S2 độc lập với S1, chưa có ngày nào đối chiếu |
| **UAT-02 — giữ P0 nhập tay** | Code có xử lý (dòng P0 do người dùng nhập tay không bị ghi đè — xem `saveNextDayP0_`), nhưng **chưa kiểm thử thật** trên Sheet |
| **UAT-33 — khởi động lại sau sự cố** | Chưa có dữ liệu vận hành đúng mẫu "CS ra lệnh = CS hoàn thành = tải thật" để kiểm chứng |

> Ngoài ra, bản Google Sheets có các tình huống **bản VBA không có** nên chưa có mã UAT: nhập lệnh từ file Excel (lệch múi giờ, dò sheet/dòng tiêu đề), tải nhiều CSV cùng lúc, tính hàng loạt nhiều ngày, đặt tên file xuất. Hiện được bao bởi 45 test cục bộ và kiểm chứng dữ liệu thật; nếu cần đánh mã UAT thì bắt đầu từ **UAT-35** (tránh khoảng trống UAT-20…23 chưa rõ nguồn gốc).

---

## Bảng gốc (bản Excel/VBA)

Cột "Trạng thái" dưới đây phản ánh tình trạng trong workbook VBA, **không phải** tình trạng của bản Google Sheets — xem phần đối chiếu ở trên.

| ID | Tình huống | Tiền điều kiện | Thao tác | Kết quả mong đợi | Trạng thái |
|----|-----------|-----------------|----------|-------------------|-----------|
| UAT-01 | Danh sách lệnh nhiều ngày | `CAI_DAT` đã chọn ngày/tổ | Nhập file lệnh có nhiều ngày | Không đổi B4:B6; chỉ lọc đúng ngày/tổ | Chưa chạy |
| UAT-02 | Giữ P0 nhập tay | B6 có số; nguồn nhập tay | Nhập danh sách lệnh | B6 giữ nguyên, không hỏi lại | Chưa chạy |
| UAT-03 | P0 tự động | Đã tính ngày liền trước cùng tổ | Chọn ngày kế tiếp | B6 lấy P cuối ngày trước | Chưa chạy |
| UAT-04 | Ramp qua 00:00 | Lệnh cuối ngày chưa đạt mục tiêu | Tính ngày sau | Có AUTO_CARRY và ramp tiếp tục | Chưa chạy |
| UAT-05 | SO hợp lệ | SO hoàn thành, F>0 | Tính Qdd | P hiệu lực = F | Chưa chạy |
| UAT-06 | MO bình thường | MO hoàn thành, E>0 | Tính Qdd | P hiệu lực = E | Chưa chạy |
| UAT-07 | MO dừng sớm | MO, Dừng=TRUE, F>0 | Tính Qdd | P hiệu lực = F | Chưa chạy |
| UAT-08 | Lệnh ngắt ramp | Lệnh mới đến giữa ramp | Tính Qdd | Ramp mới từ P nội suy | Chưa chạy |
| UAT-09 | CSV Mac một cột | Excel mở dòng CSV dồn vào cột A | Nhập CSV | Tự tách 50 trường; đủ 48 số | Chưa chạy |
| UAT-10 | CSV Windows 50 cột | Excel tách sẵn A:AX | Nhập CSV | Đọc trực tiếp; đủ 48 số | Chưa chạy |
| UAT-11 | CSV sai ngày | CSV khác ngày trong CAI_DAT | Nhập CSV | Từ chối; không đổi B4:B6 | Chưa chạy |
| UAT-12 | Đủ Qdc/Qmp | Hai CSV hợp lệ | Tính và kiểm tra | 48/48/48 tại B15 | Chưa chạy |
| UAT-13 | Đối chiếu Qdd gốc | Nhập đủ 48 Qdd gốc | Tính | Hiện sai lệch lớn nhất | Chưa chạy |
| UAT-14 | Ngày không có lệnh | P0 và CSV đầy đủ | Tính | Qdd giữ theo P0; không lỗi | Chưa chạy |
| UAT-15 | Tổ máy S2 | Dữ liệu S2 cùng cấu trúc | Chọn S2 và nhập dữ liệu | Lọc đúng S2; lịch sử P độc lập với S1 | Chưa chạy |
| UAT-16 | Sao lưu và xuất báo cáo | Kết quả không lỗi | Bấm nút 11 và 12 | Tạo backup và báo cáo .xlsx | Chưa chạy |
| UAT-17 | 31 lệnh | 31 lệnh hợp lệ | Tính | Đủ 31 lệnh được xử lý | Chưa chạy |
| UAT-18 | 45 lệnh | 45 lệnh hợp lệ | Tính | Đủ 45 lệnh được xử lý | Chưa chạy |
| UAT-19 | 60 lệnh (giới hạn tối đa) | 60 lệnh hợp lệ | Tính | Đủ 60 lệnh được xử lý | Chưa chạy |
| UAT-24 | Xuất báo cáo tháng trên Mac | Có snapshot ngày | Bấm nút 12 | Không còn lỗi merged cell 1004 | Chưa chạy |
| UAT-25 | Chạy chính thức không cần Qdd gốc | O2:O49 trống hoặc chưa đủ | Chạy nút 7 | Tính và lưu lịch sử bình thường | Chưa chạy |
| UAT-26 | Không tự tạo backup | Theo dõi thư mục Backup_Qdu | Chạy nút 7 | Không tạo file backup mới | Chưa chạy |
| UAT-27 | Backup thủ công | Bấm nút 11 | Sao lưu | Tạo bản sao dự phòng bình thường | Chưa chạy |
| UAT-28 | Chỉ còn một module Qdu | Đã xoá module cũ | Compile | Không lỗi; chỉ còn module AllInOne | Chưa chạy |
| UAT-29 | Kiểm tra module tháng | Có lịch sử tháng | Bấm nút 13 | Báo module tháng ĐẠT | Chưa chạy |
| UAT-30 | Xuất báo cáo tháng trên Windows | Có snapshot ngày | Bấm nút 12 | Không còn lỗi "Cannot run the macro" | Chưa chạy |
| UAT-31 | Xuất báo cáo tháng trên Mac | Có snapshot ngày | Bấm nút 12 | Xuất được, hoặc tự lưu cạnh workbook | Chưa chạy |
| UAT-32 | Lệnh 0-0 (trip/dừng sự cố) không được tính | Lệnh MO "Ngừng tổ máy", CS ra lệnh=0, CS hoàn thành=0, Hoàn thành=1 | Tính Qdd | Lệnh bị loại khỏi `LENH_DIEU_DO`, không tính vào Qdd | **ĐÃ XÁC NHẬN ĐÚNG** — hành vi gốc (`>0`) đúng theo nghiệp vụ, không sửa. Xem [15_Accuracy_Validation_2026-07.md](15_Accuracy_Validation_2026-07.md) |
| UAT-33 | Khởi động lại tổ máy sau sự cố | Lệnh MO, CS ra lệnh = CS hoàn thành = tải thật (vd 435,7-435,7), Hoàn thành=1 | Tính Qdd | Lệnh được coi là hoàn thành, bắt đầu tính Qdu từ tải đó | Chưa chạy — chưa có dữ liệu thật đúng mẫu 2 giá trị bằng nhau để kiểm chứng (dữ liệu 07/07 chỉ có CS ra lệnh, thiếu CS hoàn thành) |
| UAT-34 | Cảnh báo lệnh 0-0 (kế hoạch, chưa triển khai) | Có lệnh CS ra lệnh=CS hoàn thành=0 trong ngày đang tính | Tính và kiểm tra | Hiển thị cảnh báo cho người vận hành biết có lệnh dừng máy chưa được tính — không tự ý sửa Qdd | Chưa triển khai — dự kiến Giai đoạn 2/3 (Apps Script), xem [ROADMAP.md](../ROADMAP.md) |

## Nghi vấn cần xác nhận: UAT-11 (CSV sai ngày)

Phân tích công thức/code thực tế cho [04_Algorithm_Specification.md](04_Algorithm_Specification.md#7-csv-parsing-vba--readcsvwithexcel50-và-chuỗi-hàm-liên-quan) cho thấy: trong bản v1.3.1, **cột ngày của CSV không hề được đọc hay so sánh** khi import — hàm từng làm việc này (`ParseCsvDate`) đã bị đánh dấu "legacy, không gọi từ v1.1.2". Nghĩa là hành vi "Từ chối; không đổi B4:B6" mô tả ở UAT-11 **có thể không còn đúng với code hiện tại** — cần chạy thử case này trên bản v1.3.1 thật để xác nhận, và cập nhật lại UAT-11 hoặc mở lại yêu cầu nghiệp vụ nếu hành vi từ chối này vẫn cần thiết.

## Nhận xét về khoảng trống mã số

Bộ UAT nhảy từ UAT-19 sang UAT-24 (thiếu UAT-20 đến UAT-23) trong workbook gốc — chưa rõ đây là case đã bị xoá hay chưa từng viết. Cần xác nhận với người phụ trách nghiệp vụ trước khi đánh số case mới, để tránh trùng ID với case có thể được khôi phục sau này.

## Việc cần làm

Toàn bộ 31 case đang ở trạng thái **"Chưa chạy"**. Trước khi tiếp tục phát triển tính năng mới (kể cả migrate sang Apps Script — xem [ROADMAP.md](../ROADMAP.md)), nên thực thi bộ UAT này trên bản v1.3.1 và cập nhật kết quả thật (Đạt/Không đạt) vào cả sheet `KIEM_THU_UAT` và bảng trên, vì đây là cơ sở duy nhất để xác nhận thuật toán mới (nếu viết lại) cho ra kết quả giống hệ thống cũ.
