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

/**
 * Nhãn các dòng cấu hình trong CAI_DAT (cột A) - đọc bằng
 * getConfigValue_() theo NHÃN, không theo số dòng cố định, để chèn/xoá
 * dòng không làm hỏng việc đọc cấu hình.
 *
 * QUAN TRỌNG: mã công tơ Qdc/Qmp KHÁC NHAU giữa các tổ máy và giữa các
 * nhà máy - "6001"/"6303" chỉ là ví dụ mặc định của Duyên Hải 1 tổ S1,
 * KHÔNG được coi là cố định. Mỗi nhà máy khi copy Sheet này phải tự điền
 * đúng mã công tơ thật của mình vào 4 dòng "Mã công tơ ..." bên dưới.
 */
var CAI_DAT_LABELS = {
  TEN_NHA_MAY: 'Tên nhà máy',
  RAMP_RATE: 'Tốc độ ramp (MW/phút)',
  QDD_V_COEF: 'Hệ số Qdd_V',
  TOLERANCE: 'Dung sai (+-)',
  METER_QDC_S1: 'Mã công tơ Qdc - S1',
  METER_QMP_S1: 'Mã công tơ Qmp - S1',
  METER_QDC_S2: 'Mã công tơ Qdc - S2',
  METER_QMP_S2: 'Mã công tơ Qmp - S2',
};

function setupAllSheets() {
  ensureSheet_(SHEETS.CAI_DAT);
  var caiDat = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CAI_DAT);
  if (caiDat.getLastRow() === 0) {
    var L = CAI_DAT_LABELS;
    caiDat.getRange('A1:B10').setValues([
      [L.TEN_NHA_MAY, 'Duyên Hải 1'],
      [L.RAMP_RATE, 3.5],
      [L.QDD_V_COEF, 0.9188],
      [L.TOLERANCE, 0.03],
      [L.METER_QDC_S1, '6001'],
      [L.METER_QMP_S1, '6303'],
      [L.METER_QDC_S2, ''],
      [L.METER_QMP_S2, ''],
      ['', ''],
      ['Ghi chú', 'Sheet mẫu - copy cho từng nhà máy, chỉnh lại toàn bộ giá trị cột B cho đúng thực tế. Mã công tơ S1/S2 CHẮC CHẮN khác nhau, và khác nhau giữa các nhà máy - phải tự điền đúng, không dùng nguyên "6001/6303" cho nhà máy/tổ máy khác.'],
    ]);
  }
  migrateConfigLabels_(caiDat); // bù các dòng cấu hình mới nếu CAI_DAT đã có từ trước (bản cũ)

  ensureSheet_(SHEETS.LENH, LENH_HEADERS);
  ensureSheet_(SHEETS.CSV_DATA, CSV_DATA_HEADERS);
  ensureSheet_(SHEETS.P0_NGAY, P0_NGAY_HEADERS);
  ensureSheet_(SHEETS.KET_QUA, KET_QUA_HEADERS);
  ensureSheet_(SHEETS.BAO_CAO_THANG, BAO_CAO_THANG_HEADERS);

  SpreadsheetApp.getUi().alert('Đã tạo đủ các sheet cần thiết. Nhớ điền đúng "Mã công tơ Qdc/Qmp" cho từng tổ máy trong CAI_DAT (khác nhau giữa S1/S2 và giữa các nhà máy). Mở menu "QDD Smart System → Bảng điều khiển" để bắt đầu nhập CSV/tính toán.');
}

/** Thêm các dòng nhãn cấu hình còn thiếu vào CAI_DAT đã có sẵn (nâng cấp từ bản cũ). */
function migrateConfigLabels_(caiDat) {
  var lastRow = caiDat.getLastRow();
  var existingLabels = lastRow > 0 ? caiDat.getRange(1, 1, lastRow, 1).getValues().map(function (r) { return String(r[0]).trim(); }) : [];
  var toAdd = [];
  Object.keys(CAI_DAT_LABELS).forEach(function (key) {
    var label = CAI_DAT_LABELS[key];
    if (existingLabels.indexOf(label) === -1) {
      var defaultValue = (key === 'METER_QDC_S1') ? '6001' : (key === 'METER_QMP_S1') ? '6303' : '';
      toAdd.push([label, defaultValue]);
    }
  });
  if (toAdd.length > 0) {
    caiDat.getRange(lastRow + 1, 1, toAdd.length, 2).setValues(toAdd);
  }
}
