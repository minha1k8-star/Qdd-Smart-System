/**
 * Tạo cấu trúc sheet cần thiết nếu chưa có - tương đương
 * GetOrCreateControlSheet/GetOrCreateStateSheet trong VBA, nhưng đơn giản
 * hoá vì mục đích khác nhau: ở đây lưu NHIỀU NGÀY trong cùng 1 sheet
 * (có cột Ngày), không phải 1 sheet = 1 ngày như CAI_DAT/LENH_GOC cũ.
 *
 * LƯU Ý: đây là schema RÚT GỌN, chỉ giữ đúng các trường mà
 * QDD-Core-Library cần (xem CommandFilter.js RawCommand) - KHÔNG đầy đủ
 * 25 cột như LENH_GOC gốc (người ra lệnh, lý do lệnh, AGC...). Nếu cần
 * đầy đủ để đối chiếu nghiệp vụ, có thể mở rộng thêm cột sau.
 */

var SHEETS = {
  CAI_DAT: 'CAI_DAT',
  LENH: 'LENH',
  CSV_DATA: 'CSV_DATA',
  P0_NGAY: 'P0_NGAY',
  KET_QUA: 'KET_QUA',
  BAO_CAO_THANG: 'BAO_CAO_THANG',
};

var LENH_HEADERS = [
  'ID Lệnh', 'Tổ máy', 'Nội dung lệnh', 'CS ra lệnh (MW)', 'CS hoàn thành (MW)',
  'Thời điểm BĐTH', 'Hoàn thành (1/0)', 'Dừng lệnh (TRUE/FALSE)', 'Nguồn lệnh (SO/MO)',
];

var CSV_DATA_HEADERS = ['Ngày', 'Mã công tơ'].concat(
  Array.from({ length: 48 }, function (_, i) { return 'Chu kỳ ' + (i + 1); })
);

var P0_NGAY_HEADERS = ['Ngày', 'Tổ máy', 'P0 (MW)', 'Ghi chú'];

var KET_QUA_HEADERS = [
  'Ngày', 'Tổ máy', 'Chu kỳ', 'Qdd (MW)', 'Qdd_V (MWh)', 'Qdc (MWh)',
  'P_Qdc (MW)', 'Ngưỡng dưới', 'Ngưỡng trên', 'Qmp (MWh)', 'Qdư (MWh)', 'Dấu hiệu',
];

var BAO_CAO_THANG_HEADERS = ['Ngày', 'Tổ máy', 'Tổng Qdc (MWh)', 'Tổng Qmp (MWh)', 'Tổng Qdd_V (MWh)', 'Tổng Qdư (MWh)'];

function ensureSheet_(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
  }
  if (headers && sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function setupAllSheets() {
  ensureSheet_(SHEETS.CAI_DAT);
  var caiDat = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CAI_DAT);
  if (caiDat.getLastRow() === 0) {
    caiDat.getRange('A1:B5').setValues([
      ['Tên nhà máy', 'Duyên Hải 1'],
      ['Tốc độ ramp (MW/phút)', 3.5],
      ['Hệ số Qdd_V', 0.9188],
      ['Dung sai (+-)', 0.03],
      ['Ghi chú', 'Sheet mẫu - copy cho từng nhà máy, chỉnh 4 giá trị trên cho đúng thực tế'],
    ]);
  }
  ensureSheet_(SHEETS.LENH, LENH_HEADERS);
  ensureSheet_(SHEETS.CSV_DATA, CSV_DATA_HEADERS);
  ensureSheet_(SHEETS.P0_NGAY, P0_NGAY_HEADERS);
  ensureSheet_(SHEETS.KET_QUA, KET_QUA_HEADERS);
  ensureSheet_(SHEETS.BAO_CAO_THANG, BAO_CAO_THANG_HEADERS);

  SpreadsheetApp.getUi().alert('Đã tạo đủ các sheet cần thiết. Mở menu "QDD Smart System → Bảng điều khiển" để bắt đầu nhập CSV/tính toán.');
}
