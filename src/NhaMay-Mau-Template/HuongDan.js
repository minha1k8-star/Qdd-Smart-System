/**
 * Sheet HƯỚNG DẪN ngay trong file - để người dùng mới (hoặc người tiếp
 * nhận sau này) đọc được cách dùng mà không cần tìm tài liệu bên ngoài.
 * Tương đương sheet HUONG_DAN của bản VBA cũ.
 *
 * Nội dung do hệ thống quản lý: mỗi lần chạy "Thiết lập sheet" sẽ được
 * ghi lại từ đầu, nên đừng sửa tay trực tiếp vào sheet này.
 */

var HUONG_DAN_ROWS = [
  ['HƯỚNG DẪN SỬ DỤNG — QDD SMART SYSTEM', ''],
  ['', ''],
  ['Mọi thao tác nằm ở menu "QDD Smart System" → "Bảng điều khiển" (sidebar bên phải).', ''],
  ['Sheet này do hệ thống tự tạo lại mỗi lần chạy "Thiết lập sheet" — không nên sửa tay.', ''],
  ['', ''],

  ['A. CHUẨN BỊ (chỉ làm một lần)', ''],
  ['A1', 'Menu "QDD Smart System" → "Thiết lập sheet" để tạo đủ các sheet cần thiết.'],
  ['A2', 'Mở sheet CAI_DAT, điền đúng thông số nhà máy: tốc độ ramp, hệ số Qdd_V, dung sai, MÃ CÔNG TƠ Qdc/Qmp cho từng tổ máy, nhãn báo cáo.'],
  ['A3', 'Mã công tơ khác nhau giữa S1/S2 và giữa các nhà máy — giá trị có sẵn chỉ đúng cho Duyên Hải 1, phải sửa lại khi dùng cho nhà máy khác.'],
  ['A4', 'Mở sheet P0_NGAY, điền 1 dòng cho NGÀY ĐẦU TIÊN muốn tính: Ngày, Tổ máy, P0 (MW). Các ngày sau hệ thống tự ghi, không cần nhập nữa.'],
  ['', ''],

  ['B. QUY TRÌNH HÀNG NGÀY', ''],
  ['B1', 'Nhập lệnh điều độ: copy toàn bộ file "DanhSachLenhKetThuc" gốc rồi dán thẳng vào sheet LENH (từ dòng 2). Cấu trúc LENH giống hệt file gốc nên không lo lệch cột.'],
  ['B2', 'Cách khác (nên dùng khi nhập bổ sung): dán file gốc vào sheet LENH_STAGING rồi bấm "Nhập từ LENH_STAGING" ở sidebar mục 2 — lệnh trùng ID sẽ được cập nhật, không bị nhân đôi.'],
  ['B3', 'Lưu CSV công tơ: sidebar mục 1. Chọn tổ máy + loại dữ liệu (Qdc/Qmp) + file CSV. Hệ thống tự đọc ngày trong file. Cần đủ CẢ Qdc và Qmp mới tính được.'],
  ['B4', 'Tính: sidebar mục 3, chọn ngày + tổ máy → bấm Tính. Kết quả ghi vào sheet KET_QUA.'],
  ['B5', 'Đọc kỹ phần cảnh báo (nếu có) hiện trong sidebar sau khi tính — xem mục E bên dưới.'],
  ['', ''],

  ['C. LÀM NHIỀU NGÀY CÙNG LÚC', ''],
  ['C1', 'Nhập lệnh của nhiều ngày vào LENH (dán một lần, nhiều ngày cũng được).'],
  ['C2', 'Sidebar mục 4 "Tải CSV hàng loạt": chọn CÙNG LÚC nhiều file CSV. Hệ thống tự nhận diện tổ máy/loại dữ liệu theo tên file và tự đọc ngày.'],
  ['C3', 'Sidebar mục 5 "Tính hàng loạt": chọn khoảng ngày + tổ máy. Ngày nào thiếu dữ liệu sẽ được báo riêng, không chặn các ngày còn lại.'],
  ['', ''],

  ['D. BÁO CÁO', ''],
  ['D1', 'Sidebar mục 6 "Báo cáo tháng": cập nhật bảng tổng vào BAO_CAO_THANG và xuất 1 file gồm toàn bộ ngày đã tính trong tháng.'],
  ['D2', 'Sidebar mục 7 "Xuất báo cáo": xuất 1 ngày hoặc một khoảng ngày bất kỳ (chọn Từ ngày = Đến ngày nếu chỉ cần 1 ngày).'],
  ['D3', 'File xuất theo đúng mẫu "Kiểm tra Qdu": mỗi ngày 1 tab, S1 và S2 nằm cạnh nhau, tổ nào chưa có dữ liệu thì để bảng trống.'],
  ['D4', 'Chỉ xuất được ngày ĐÃ TÍNH XONG. Ngày chưa tính sẽ được báo riêng.'],
  ['', ''],

  ['E. CẢNH BÁO TỰ ĐỘNG — ĐỌC KỸ KHI XUẤT HIỆN', ''],
  ['Lệnh 0-0', 'Ngày có lệnh với CS ra lệnh = CS hoàn thành = 0 (thường là trip/ngừng sự cố). Theo quy tắc nghiệp vụ, lệnh này KHÔNG được tính. Nếu tổ máy thực sự đã ngừng thì Qdd giai đoạn đó đang CAO HƠN thực tế — cần kiểm tra và xử lý tay.'],
  ['Không có lệnh', 'Ngày tính không có lệnh hiệu lực nào → Qdd giữ nguyên P0 suốt 48 chu kỳ. Nếu ngày đó thực tế CÓ lệnh mà chưa nhập thì kết quả sai — kiểm tra lại sheet LENH.'],
  ['Nối tiếp ramp', 'Ngày trước còn tăng/giảm tải dở dang lúc 24:00 → ngày này tự chạy tiếp cho tới khi đạt mục tiêu (quy tắc R07). Đây là thông báo bình thường, không phải lỗi.'],
  ['', ''],

  ['F. CÁC SHEET TRONG FILE', ''],
  ['CAI_DAT', 'Cấu hình nhà máy: tốc độ ramp, hệ số, dung sai, mã công tơ, nhãn báo cáo. KHÔNG bị xoá khi dọn dữ liệu.'],
  ['LENH', 'Danh sách lệnh điều độ, cấu trúc giống hệt file gốc (25 cột).'],
  ['LENH_STAGING', 'Nơi dán tạm file gốc trước khi bấm nút nhập ở sidebar mục 2.'],
  ['CSV_DATA', 'Dữ liệu công tơ đã lưu, mỗi dòng là 1 cặp (ngày, mã công tơ) kèm 48 chu kỳ.'],
  ['P0_NGAY', 'Công suất đầu ngày. Hệ thống tự ghi cho ngày kế tiếp sau mỗi lần tính; dòng bạn tự nhập sẽ không bị ghi đè.'],
  ['KET_QUA', 'Kết quả tính, mỗi ngày 48 dòng. Ngày mới nhất nằm trên cùng.'],
  ['BAO_CAO_THANG', 'Bảng tổng hợp theo ngày của tháng được chọn.'],
  ['', ''],

  ['G. LƯU Ý QUAN TRỌNG', ''],
  ['G1', 'Tuỳ chọn "Dọn lệnh + CSV sau khi tính" (mục 3 và 5) mặc định BẬT. Khi đang dò/đối chiếu kết quả nên TẮT đi, để không phải nhập lại dữ liệu nếu cần tính lại.'],
  ['G2', 'Sidebar mục 8 "Dọn dữ liệu cũ" dùng khi bàn giao file hoặc khi dữ liệu tích luỹ nhiều tháng. Giữ CAI_DAT và kết quả N ngày gần nhất, xoá phần còn lại. Không hoàn tác được.'],
  ['G3', 'Dữ liệu gốc (file CSV, file danh sách lệnh) vẫn nằm trên máy bạn — nếu lỡ xoá trong file này thì nhập lại được.'],
  ['G4', 'Người dùng cần quyền CHỈNH SỬA mới thấy menu. Người chỉ có quyền Xem sẽ không chạy được chức năng nào.'],
  ['G5', 'Lần đầu mỗi người dùng bấm một chức năng, Google sẽ hỏi cấp quyền — chọn tài khoản, bấm Nâng cao (Advanced) rồi Cho phép (Allow).'],
  ['', ''],

  ['Tài liệu kỹ thuật đầy đủ: https://github.com/minha1k8-star/Qdd-Smart-System', ''],
];

/** Tạo/ghi lại sheet HƯỚNG DẪN và đưa lên vị trí đầu tiên. */
function buildHuongDanSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEETS.HUONG_DAN);
  if (!sh) sh = ss.insertSheet(SHEETS.HUONG_DAN);

  sh.clear();
  sh.getRange(1, 1, HUONG_DAN_ROWS.length, 2).setValues(HUONG_DAN_ROWS);

  // Tiêu đề chính
  sh.getRange(1, 1, 1, 2).merge().setFontSize(14).setFontWeight('bold');

  // In đậm các dòng tiêu đề mục (A., B., C....) và các nhãn ở cột A
  HUONG_DAN_ROWS.forEach(function (row, i) {
    var label = String(row[0]);
    var isSectionHeader = /^[A-G]\. /.test(label);
    if (isSectionHeader) {
      sh.getRange(i + 1, 1, 1, 2).merge().setFontWeight('bold').setBackground('#e8eaed');
    } else if (label && !row[1]) {
      sh.getRange(i + 1, 1, 1, 2).merge(); // dòng ghi chú trải ngang
    }
  });

  sh.setColumnWidth(1, 150);
  sh.setColumnWidth(2, 780);
  sh.getRange(1, 1, HUONG_DAN_ROWS.length, 2).setVerticalAlignment('top').setWrap(true);
  sh.setFrozenRows(1);

  ss.setActiveSheet(sh);
  ss.moveActiveSheet(1); // luôn nằm ngoài cùng bên trái để người mới thấy ngay
}
