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
  LENH_STAGING: 'LENH_STAGING',
  CSV_DATA: 'CSV_DATA',
  P0_NGAY: 'P0_NGAY',
  KET_QUA: 'KET_QUA',
  BAO_CAO_THANG: 'BAO_CAO_THANG',
};

/**
 * Cấu trúc GIỐNG HỆT file "DanhSachLenhKetThuc" gốc (25 cột, đúng thứ tự)
 * để người dùng có thể copy TOÀN BỘ file gốc dán thẳng vào đây mà không
 * lệch cột. Hệ thống chỉ dùng 9 trường cho tính toán, các cột còn lại giữ
 * nguyên để đối chiếu nghiệp vụ.
 *
 * Dù vậy, việc đọc dữ liệu KHÔNG phụ thuộc thứ tự cột này - readAllCommands_
 * dò cột theo TÊN tiêu đề, nên sheet cũ (9/10 cột) hay sheet đã đổi thứ tự
 * cột vẫn đọc đúng.
 */
var LENH_HEADERS = [
  'ID Lệnh', 'Nhà máy', 'Tổ máy', 'Nội dung lệnh', 'CS ra lệnh (MW)', 'CS hoàn thành (MW)',
  'Thời điểm BĐTH', 'Thời điểm hoàn thành', 'Người ra lệnh', 'Người thực hiện', 'AGC',
  'Nhiên liệu', 'Lý do lệnh', 'Ghi chú ra lệnh', 'Ghi chú hoàn thành', 'Hoàn thành',
  'Dừng lệnh', 'Thời điểm dừng', 'Lý do dừng', 'Người dừng', 'Lý do hủy', 'Người hủy',
  'Lệnh cụm', 'Lệnh nhập lại', 'Nguồn lệnh',
];

var CSV_DATA_HEADERS = ['Ngày', 'Mã công tơ'].concat(
  Array.from({ length: 48 }, function (_, i) { return 'Chu kỳ ' + (i + 1); })
);

// Cột 'Ramp tiếp đến (MW)' phục vụ R07: nếu ramp còn dở dang lúc 24:00 của
// ngày trước, ghi mục tiêu công suất vào đây để ngày này chạy tiếp cho xong.
var P0_NGAY_HEADERS = ['Ngày', 'Tổ máy', 'P0 (MW)', 'Ghi chú', 'Ramp tiếp đến (MW)'];

var KET_QUA_HEADERS = [
  'Ngày', 'Tổ máy', 'Chu kỳ', 'Qdd (MW)', 'Qdd_V (MWh)', 'Qdc (MWh)',
  'P_Qdc (MW)', 'Ngưỡng dưới', 'Ngưỡng trên', 'Qmp (MWh)', 'Qdư (MWh)', 'Dấu hiệu',
];

/** vd chuKy=1 -> "00:00-00:30", chuKy=48 -> "23:30-24:00". */
function periodTimeRangeLabel_(chuKy) {
  var startMin = (chuKy - 1) * 30;
  var endMin = chuKy * 30;
  function fmt(totalMin) {
    var h = Math.floor(totalMin / 60), m = totalMin % 60;
    return (h < 10 ? '0' : '') + h + ':' + (m < 10 ? '0' : '') + m;
  }
  return fmt(startMin) + '-' + fmt(endMin);
}

/** Nhãn chu kỳ đầy đủ (số + khung giờ trong CÙNG 1 ô), khớp file "Kiểm tra Qdu" gốc: "01 [00:00-00:30]". */
function periodLabel_(chuKy) {
  return (chuKy < 10 ? '0' : '') + chuKy + ' [' + periodTimeRangeLabel_(chuKy) + ']';
}

/** Đọc ngược số chu kỳ từ ô "Chu kỳ" (chấp nhận cả số thuần của dữ liệu cũ lẫn nhãn "01 [00:00-00:30]"). */
function parsePeriodNumber_(value) {
  if (typeof value === 'number') return value;
  var m = String(value).trim().match(/^(\d+)/);
  return m ? Number(m[1]) : NaN;
}

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
  REPORT_LABEL_S1: 'Nhãn báo cáo - S1',
  REPORT_LABEL_S2: 'Nhãn báo cáo - S2',
};

/** Giá trị mặc định cho Duyên Hải 1 - nhà máy khác PHẢI tự điền lại. */
var CAI_DAT_DEFAULTS = {
  METER_QDC_S1: '6001',
  METER_QMP_S1: '6303',
  METER_QDC_S2: '6002',
  METER_QMP_S2: '6301',
  REPORT_LABEL_S1: 'S1DH1',
  REPORT_LABEL_S2: 'S2DH1',
};

function setupAllSheets() {
  ensureSheet_(SHEETS.CAI_DAT);
  var caiDat = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CAI_DAT);
  if (caiDat.getLastRow() === 0) {
    var L = CAI_DAT_LABELS, D = CAI_DAT_DEFAULTS;
    caiDat.getRange('A1:B12').setValues([
      [L.TEN_NHA_MAY, 'Duyên Hải 1'],
      [L.RAMP_RATE, 3.5],
      [L.QDD_V_COEF, 0.9188],
      [L.TOLERANCE, 0.03],
      [L.METER_QDC_S1, D.METER_QDC_S1],
      [L.METER_QMP_S1, D.METER_QMP_S1],
      [L.METER_QDC_S2, D.METER_QDC_S2],
      [L.METER_QMP_S2, D.METER_QMP_S2],
      [L.REPORT_LABEL_S1, D.REPORT_LABEL_S1],
      [L.REPORT_LABEL_S2, D.REPORT_LABEL_S2],
      ['', ''],
      ['Ghi chú', 'Sheet mẫu - copy cho từng nhà máy, chỉnh lại toàn bộ giá trị cột B cho đúng thực tế. Mã công tơ và nhãn báo cáo khác nhau giữa các nhà máy - phải tự điền đúng, không dùng nguyên giá trị mặc định của Duyên Hải 1.'],
    ]);
  }
  migrateConfigLabels_(caiDat); // bù các dòng cấu hình mới nếu CAI_DAT đã có từ trước (bản cũ)

  ensureSheet_(SHEETS.LENH, LENH_HEADERS);
  migrateLenhSheet_(); // chèn cột "Nhà máy" nếu LENH còn dùng cấu trúc 9 cột cũ
  ensureSheet_(SHEETS.LENH_STAGING);
  ensureSheet_(SHEETS.CSV_DATA, CSV_DATA_HEADERS);
  ensureSheet_(SHEETS.P0_NGAY, P0_NGAY_HEADERS);
  ensureSheet_(SHEETS.KET_QUA, KET_QUA_HEADERS);
  ensureSheet_(SHEETS.BAO_CAO_THANG, BAO_CAO_THANG_HEADERS);
  removeLeftoverSheet_('CSV_STAGING'); // sheet cũ từ thiết kế trước, không còn dùng (đọc CSV trực tiếp trong sidebar)

  SpreadsheetApp.getUi().alert('Đã tạo đủ các sheet cần thiết. Nhớ điền đúng "Mã công tơ Qdc/Qmp" cho từng tổ máy trong CAI_DAT (khác nhau giữa S1/S2 và giữa các nhà máy). Mở menu "QDD Smart System → Bảng điều khiển" để bắt đầu nhập CSV/tính toán.');
}

/** Xoá 1 sheet cũ không còn dùng nếu tồn tại (dọn dẹp khi nâng cấp từ bản cũ). */
function removeLeftoverSheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (sh) ss.deleteSheet(sh);
}

/** Thêm các dòng nhãn cấu hình còn thiếu vào CAI_DAT đã có sẵn (nâng cấp từ bản cũ). */
function migrateConfigLabels_(caiDat) {
  var lastRow = caiDat.getLastRow();
  var existingLabels = lastRow > 0 ? caiDat.getRange(1, 1, lastRow, 1).getValues().map(function (r) { return String(r[0]).trim(); }) : [];
  var toAdd = [];
  Object.keys(CAI_DAT_LABELS).forEach(function (key) {
    var label = CAI_DAT_LABELS[key];
    if (existingLabels.indexOf(label) === -1) {
      toAdd.push([label, CAI_DAT_DEFAULTS[key] || '']);
    }
  });
  if (toAdd.length > 0) {
    caiDat.getRange(lastRow + 1, 1, toAdd.length, 2).setValues(toAdd);
  }
}

/**
 * Dựng lại sheet LENH đúng cấu trúc 25 cột giống hệt file gốc (Nhà máy ở
 * cột B...), sắp xếp lại dữ liệu đang có theo TÊN cột nên không mất/không
 * lệch dữ liệu. Nhờ vậy người dùng có thể copy nguyên file gốc dán thẳng
 * vào LENH mà các cột vẫn trùng khớp 1-1.
 */
function migrateLenhSheet_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LENH);
  if (!sh) return;
  var lastCol = sh.getLastColumn();
  var lastRow = sh.getLastRow();

  if (lastCol === 0) {
    sh.getRange(1, 1, 1, LENH_HEADERS.length).setValues([LENH_HEADERS]);
    sh.setFrozenRows(1);
    return;
  }

  var currentHeaders = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(normalizeHeader_);
  var targetHeaders = LENH_HEADERS.map(normalizeHeader_);
  var alreadyCorrect = targetHeaders.every(function (h, i) { return currentHeaders[i] === h; });
  if (alreadyCorrect) {
    sh.setFrozenRows(1);
    return;
  }

  // Đọc dữ liệu cũ rồi ánh xạ sang đúng vị trí cột chuẩn theo tên.
  var oldData = lastRow >= 2 ? sh.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];
  var remapped = oldData
    .filter(function (r) { return r.join('') !== ''; })
    .map(function (r) {
      return targetHeaders.map(function (name) {
        var srcIdx = currentHeaders.indexOf(name);
        return srcIdx === -1 ? '' : r[srcIdx];
      });
    });

  sh.clear();
  sh.getRange(1, 1, 1, LENH_HEADERS.length).setValues([LENH_HEADERS]);
  if (remapped.length > 0) {
    sh.getRange(2, 1, remapped.length, LENH_HEADERS.length).setValues(remapped);
  }
  sh.setFrozenRows(1);
}
