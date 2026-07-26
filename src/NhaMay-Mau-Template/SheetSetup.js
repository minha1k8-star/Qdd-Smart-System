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
  HUONG_DAN: 'HUONG_DAN',
  CAI_DAT: 'CAI_DAT',
  LENH: 'LENH',
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
 * Các dòng cấu hình RIÊNG THEO TỔ MÁY (mã công tơ, nhãn báo cáo) không
 * liệt kê ở đây mà suy ra từ chính CAI_DAT - xem getConfiguredUnits_.
 */
var CAI_DAT_LABELS = {
  TEN_NHA_MAY: 'Tên nhà máy',
  RAMP_RATE: 'Tốc độ ramp (MW/phút)',
  QDD_V_COEF: 'Hệ số Qdd_V',
  TOLERANCE: 'Dung sai (+-)',
};

/**
 * Tổ máy mặc định khi tạo Sheet mới. Nhà máy có NHIỀU HƠN 2 TỔ MÁY chỉ
 * cần THÊM 3 DÒNG vào CAI_DAT cho mỗi tổ (không phải sửa code):
 *   Mã công tơ Qdc - S3 | <mã>
 *   Mã công tơ Qmp - S3 | <mã>
 *   Nhãn báo cáo - S3   | <nhãn>
 * Danh sách tổ máy được suy ra từ chính các dòng này - xem getConfiguredUnits_.
 *
 * MÃ CÔNG TƠ KHÔNG KÈM CHỮ SỐ NĂM: tên file CSV có dạng
 * <ngày><tháng><năm 1 chữ số><mã công tơ>, vd "17076001.CSV" = 17/07, năm
 * 2026 (số 6), công tơ 001. Ghi "6001" thì sang 2027 tên file thành
 * "17077001.CSV" và hệ thống sẽ không nhận ra file nào.
 *
 * Giá trị dưới đây là của Duyên Hải 1 - nhà máy khác PHẢI tự điền lại.
 */
var DEFAULT_UNITS = [
  { unit: 'S1', qdc: 'csv001', qmp: 'csv303', reportLabel: 'S1DH1' },
  { unit: 'S2', qdc: 'csv002', qmp: 'csv301', reportLabel: 'S2DH1' },
];

/**
 * Tiền tố bắt buộc của mã công tơ trong CAI_DAT.
 *
 * VÌ SAO PHẢI CÓ CHỮ: Google Sheets tự hiểu ô "001" là SỐ 1 rồi cắt mất
 * hai số 0 đầu - đặt định dạng văn bản cũng KHÔNG cứu được ô đã lỡ lưu
 * thành số (đã thử, ô vẫn hiển thị 1). Thêm chữ "csv" phía trước là cách
 * chắc chắn để ô luôn là văn bản, đồng thời người đọc hiểu ngay đó là mã
 * nằm trong TÊN FILE CSV.
 *
 * Việc so khớp chỉ lấy phần chữ số nên "csv001", "001" hay 1 đều khớp
 * đúng cùng một công tơ - tiền tố chỉ phục vụ hiển thị.
 */
var METER_CODE_PREFIX = 'csv';

/** Số chữ số tối thiểu của mã công tơ - dùng để bù lại các số 0 đầu đã bị Sheets cắt. */
var METER_CODE_MIN_DIGITS = 3;

function setupAllSheets() {
  ensureSheet_(SHEETS.CAI_DAT);
  var caiDat = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CAI_DAT);
  if (caiDat.getLastRow() === 0) {
    var L = CAI_DAT_LABELS;
    var rows = [
      [L.TEN_NHA_MAY, 'Duyên Hải 1'],
      [L.RAMP_RATE, 3.5],
      [L.QDD_V_COEF, 0.9188],
      [L.TOLERANCE, 0.03],
    ];
    DEFAULT_UNITS.forEach(function (u) {
      rows.push([meterLabel_(u.unit, 'Qdc'), u.qdc]);
      rows.push([meterLabel_(u.unit, 'Qmp'), u.qmp]);
      rows.push([reportLabelLabel_(u.unit), u.reportLabel]);
    });
    rows.push(['', '']);
    rows.push(['Ghi chú', 'Sheet mẫu - chỉnh lại toàn bộ giá trị cột B cho đúng nhà máy của bạn. ' +
      'MÃ CÔNG TƠ KHÔNG KÈM CHỮ SỐ NĂM (vd file 17076001.CSV thì mã công tơ là 001, số 6 là năm 2026). ' +
      'Nhà máy có thêm tổ máy: thêm 3 dòng "Mã công tơ Qdc - S3", "Mã công tơ Qmp - S3", "Nhãn báo cáo - S3".']);
    caiDat.getRange(1, 1, rows.length, 2).setValues(rows);
  }
  formatMeterCodeCellsAsText_(caiDat);
  migrateConfigLabels_(caiDat);      // bù các dòng cấu hình chung còn thiếu (nâng cấp từ bản cũ)
  var meterChanges = migrateMeterCodes_();  // bỏ chữ số năm khỏi mã công tơ của bản cũ

  ensureSheet_(SHEETS.LENH, LENH_HEADERS);
  migrateLenhSheet_(); // chèn cột "Nhà máy" nếu LENH còn dùng cấu trúc 9 cột cũ
  var csvSh = ensureSheet_(SHEETS.CSV_DATA, CSV_DATA_HEADERS);
  csvSh.getRange(2, 2, csvSh.getMaxRows() - 1, 1).setNumberFormat('@'); // mã công tơ: giữ số 0 đầu
  ensureSheet_(SHEETS.P0_NGAY, P0_NGAY_HEADERS);
  ensureSheet_(SHEETS.KET_QUA, KET_QUA_HEADERS);
  ensureSheet_(SHEETS.BAO_CAO_THANG, BAO_CAO_THANG_HEADERS);
  removeLeftoverSheet_('CSV_STAGING'); // sheet cũ từ thiết kế trước, không còn dùng (đọc CSV trực tiếp trong sidebar)
  removeLeftoverSheet_('NHAT_KY');   // nhật ký thao tác - đã bỏ theo yêu cầu nghiệp vụ
  removeLeftoverSheet_('LENH_STAGING'); // không còn dùng: lệnh nhập thẳng từ file Excel tải lên ở sidebar mục 2
  buildHuongDanSheet_(); // ghi lại sheet hướng dẫn, đặt ngoài cùng bên trái

  var units = getConfiguredUnits_();
  var msg = 'Đã tạo đủ các sheet cần thiết. Tổ máy đang cấu hình: ' + (units.join(', ') || '(chưa có)') + '.';
  if (meterChanges.length > 0) {
    msg += '\n\nĐã chuẩn hoá mã công tơ về dạng "csv001":\n' + meterChanges.join('\n') +
      '\n\nLý do: chữ số đầu trong "6001" là NĂM (2026), không thuộc mã công tơ; ' +
      'và Google Sheets cắt mất số 0 đầu nếu ô là số, nên thêm chữ "csv" cho ô luôn là văn bản.\n' +
      'Các dòng CSV_DATA cũ đã được đổi theo, KHÔNG cần tải lại CSV.';
  }
  msg += '\n\nXem sheet HUONG_DAN (ngoài cùng bên trái) để biết cách dùng. ' +
    'Nhớ kiểm tra "Mã công tơ Qdc/Qmp" cho từng tổ máy trong CAI_DAT trước khi tính.';
  SpreadsheetApp.getUi().alert(msg);
}

/** Xoá 1 sheet cũ không còn dùng nếu tồn tại (dọn dẹp khi nâng cấp từ bản cũ). */
function removeLeftoverSheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (sh) ss.deleteSheet(sh);
}

/** Thêm các dòng cấu hình CHUNG còn thiếu vào CAI_DAT đã có sẵn (nâng cấp từ bản cũ). */
function migrateConfigLabels_(caiDat) {
  var lastRow = caiDat.getLastRow();
  var existingLabels = lastRow > 0 ? caiDat.getRange(1, 1, lastRow, 1).getValues().map(function (r) { return String(r[0]).trim(); }) : [];
  var toAdd = [];
  Object.keys(CAI_DAT_LABELS).forEach(function (key) {
    var label = CAI_DAT_LABELS[key];
    if (existingLabels.indexOf(label) === -1) toAdd.push([label, '']);
  });
  if (toAdd.length > 0) {
    caiDat.getRange(lastRow + 1, 1, toAdd.length, 2).setValues(toAdd);
  }
}

/**
 * Đặt các ô mã công tơ thành ĐỊNH DẠNG VĂN BẢN THUẦN.
 *
 * VÌ SAO: Google Sheets tự hiểu "001" là SỐ 1 rồi cắt mất hai số 0 đầu,
 * nên mã công tơ hiển thị sai đi (001 -> 1). Định dạng văn bản giữ
 * nguyên đúng những gì người dùng gõ.
 *
 * Việc so khớp vẫn chịu được cả trường hợp ô đã bị mất số 0 (xem
 * sameMeterCode_/meterMatchesFilename_) - đây là lớp bảo vệ thứ hai, còn
 * định dạng văn bản là lớp thứ nhất để người dùng NHÌN thấy đúng mã.
 */
function formatMeterCodeCellsAsText_(caiDat) {
  var lastRow = caiDat.getLastRow();
  if (!lastRow) return;
  var labels = caiDat.getRange(1, 1, lastRow, 1).getValues();
  labels.forEach(function (r, i) {
    if (/^Mã công tơ (Qdc|Qmp) - /.test(String(r[0]).trim())) {
      caiDat.getRange(i + 1, 2).setNumberFormat('@');
    }
  });
}

/**
 * Chuẩn hoá mã công tơ trong CAI_DAT về dạng "csv001", và sửa luôn các
 * dòng CSV_DATA cũ theo.
 *
 * Xử lý cùng lúc ba vấn đề đã gặp thật:
 *
 * 1. CHỮ SỐ NĂM. "6001" = năm 2026 (số 6) + công tơ 001. Để nguyên thì
 *    sang 2027 tên file thành "17077001.CSV", không còn khớp -> mọi file
 *    CSV bị bỏ qua mà người dùng không hiểu vì sao.
 * 2. MẤT SỐ 0 ĐẦU. Google Sheets hiểu "001" là số 1. Bù lại đủ 3 chữ số.
 * 3. Ô lưu dạng SỐ nên bị cắt lại lần nữa. Thêm tiền tố "csv" để ô luôn
 *    là văn bản, và người đọc hiểu ngay đó là mã trong tên file CSV.
 *
 * Mọi thay đổi đều được liệt kê trong thông báo cuối "Thiết lập sheet" -
 * không đổi âm thầm.
 *
 * @returns {string[]} mô tả các thay đổi đã thực hiện (rỗng nếu không có gì)
 */
function migrateMeterCodes_() {
  var caiDat = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CAI_DAT);
  var lastRow = caiDat.getLastRow();
  if (!lastRow) return [];

  var rows = caiDat.getRange(1, 1, lastRow, 2).getValues();
  var changes = [];

  rows.forEach(function (r, i) {
    var label = String(r[0]).trim();
    if (!/^Mã công tơ (Qdc|Qmp) - /.test(label)) return;
    var oldRaw = String(r[1]).trim();
    if (!oldRaw) return;
    var newCode = normalizeMeterCode_(oldRaw);
    if (newCode === oldRaw) return;
    caiDat.getRange(i + 1, 2).setNumberFormat('@').setValue(newCode);
    changes.push(label + ': ' + oldRaw + ' \u2192 ' + newCode);
  });

  migrateCsvDataMeterCodes_();
  return changes;
}

/**
 * Đưa một mã công tơ bất kỳ về dạng chuẩn "csv001".
 * Nhận được: "6001" (kèm chữ số năm), "001", 1 (đã mất số 0), "csv001".
 */
function normalizeMeterCode_(raw) {
  var digits = String(raw).replace(/[^0-9]/g, '');
  if (!digits) return String(raw);
  // 4 chữ số = còn kèm chữ số năm ở đầu (vd 6001) -> bỏ chữ số đầu.
  if (digits.length === 4) digits = digits.slice(1);
  while (digits.length < METER_CODE_MIN_DIGITS) digits = '0' + digits;
  return METER_CODE_PREFIX + digits;
}

/**
 * Đổi mã công tơ trong các dòng CSV_DATA cũ sang đúng mã đang cấu hình.
 * Không có bước này thì dữ liệu CSV đã tải trước đây sẽ không tìm thấy
 * khi tính lại ngày cũ (báo "thiếu CSV" cho ngày đã có đủ dữ liệu).
 */
function migrateCsvDataMeterCodes_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CSV_DATA);
  if (!sh) return;
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  var configured = getAllConfiguredMeters_().map(function (m) { return m.code; });
  var range = sh.getRange(2, 2, lastRow - 1, 1);
  var values = range.getValues();
  var touched = false;

  for (var i = 0; i < values.length; i++) {
    var current = String(values[i][0]).trim();
    if (!current) continue;
    var normalized = normalizeMeterCode_(current);
    for (var j = 0; j < configured.length; j++) {
      if (sameMeterCode_(normalized, configured[j]) && current !== configured[j]) {
        values[i][0] = configured[j];
        touched = true;
        break;
      }
    }
  }
  if (touched) {
    range.setNumberFormat('@');
    range.setValues(values);
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
