/**
 * Đọc/ghi dữ liệu giữa Sheet và QDD-Core-Library. Toàn bộ tính toán thật
 * nằm ở Library (QDDCoreLibrary) - các hàm ở đây chỉ chuyển đổi định dạng.
 */

/** Đọc giá trị cột B của CAI_DAT theo NHÃN ở cột A (không theo số dòng cố định). */
function getConfigValue_(label) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CAI_DAT);
  var lastRow = sh.getLastRow();
  if (lastRow === 0) return null;
  var rows = sh.getRange(1, 1, lastRow, 2).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === label) return rows[i][1];
  }
  return null;
}

function getConfig_() {
  var L = CAI_DAT_LABELS;
  return {
    rampRate: Number(getConfigValue_(L.RAMP_RATE)),
    qddVCoef: Number(getConfigValue_(L.QDD_V_COEF)),
    tolerance: Number(getConfigValue_(L.TOLERANCE)),
  };
}

/**
 * Tra mã công tơ thật (vd "6001") theo (tổ máy, loại dữ liệu) từ CAI_DAT.
 * Mã công tơ KHÁC NHAU giữa các tổ máy và giữa các nhà máy - không có giá
 * trị mặc định cố định ngoài Sheet mẫu ban đầu (6001/6303 cho S1).
 * @param {string} unit "S1"|"S2"
 * @param {string} role "Qdc"|"Qmp"
 * @returns {string|null}
 */
function resolveMeterCode_(unit, role) {
  var L = CAI_DAT_LABELS;
  var key = 'METER_' + role.toUpperCase() + '_' + unit.toUpperCase(); // vd METER_QDC_S1
  var label = L[key];
  if (!label) return null;
  var value = getConfigValue_(label);
  return value ? String(value).trim() : null;
}

/** @returns {import('../QDD-Core-Library/CommandFilter').RawCommand[]} */
function readAllCommands_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LENH);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  var rows = sh.getRange(2, 1, lastRow - 1, LENH_HEADERS.length).getValues();
  return rows
    .filter(function (r) { return r[0] !== ''; })
    .map(function (r) {
      return {
        id: r[0],
        toMay: r[1],
        noiDungLenh: r[2],
        csRaLenh: r[3] === '' ? null : Number(r[3]),
        csHoanThanh: r[4] === '' ? null : Number(r[4]),
        bdth: r[5] instanceof Date ? r[5] : null,
        hoanThanh: r[6],
        dungLenh: r[7] === true || String(r[7]).toUpperCase() === 'TRUE',
        nguonLenh: r[8],
      };
    });
}

/** Ghi 1 lệnh mới vào cuối sheet LENH (dùng khi nhập tay từng lệnh). */
function appendCommand(command) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LENH);
  sh.appendRow([
    command.id, command.toMay, command.noiDungLenh, command.csRaLenh, command.csHoanThanh,
    command.bdth, command.hoanThanh, command.dungLenh, command.nguonLenh,
  ]);
}

/**
 * @returns {Array<{unit:string, role:string, code:string}>} 4 mã công tơ đã cấu hình (bỏ ô trống)
 */
function getAllConfiguredMeters_() {
  var out = [];
  ['S1', 'S2'].forEach(function (unit) {
    ['Qdc', 'Qmp'].forEach(function (role) {
      var code = resolveMeterCode_(unit, role);
      if (code) out.push({ unit: unit, role: role, code: code });
    });
  });
  return out;
}

/**
 * Dò (tổ máy, loại dữ liệu) từ TÊN FILE, dựa vào mã công tơ đã cấu hình
 * trong CAI_DAT. Tên file CSV thật thường có dạng <ngày><tháng><mã công
 * tơ>.CSV (vd "17076001.CSV" = ngày 17, tháng 07, công tơ 6001) - mã công
 * tơ luôn nằm ở CUỐI phần số của tên file, nên so khớp theo endsWith.
 * @param {string} filename
 * @returns {{unit:string, role:string, code:string}|null}
 */
function matchMeterFromFilename_(filename) {
  var base = String(filename).replace(/\.[^.]+$/, ''); // bỏ phần đuôi .csv
  var meters = getAllConfiguredMeters_();
  // Ưu tiên mã dài hơn trước, tránh mã ngắn khớp nhầm vào phần đuôi của mã dài hơn.
  meters.sort(function (a, b) { return b.code.length - a.code.length; });
  for (var i = 0; i < meters.length; i++) {
    if (base.indexOf(meters[i].code) !== -1 && base.slice(-meters[i].code.length) === meters[i].code) {
      return meters[i];
    }
  }
  return null;
}

/**
 * Đọc thử ngày ở cột A của dòng dữ liệu đầu tiên trong CSV (dd-mm-yy(yy)
 * hoặc dd/mm/yy(yy) hoặc yyyy-mm-dd). Dùng cho luồng tải hàng loạt phía
 * server (luồng tải 1 file dùng bản JS phía client trong Sidebar.html để
 * điền trực tiếp lên form, xem parseCsvDateGuess ở đó).
 * @param {string} csvText
 * @returns {Date|null}
 */
function guessDateFromCsvText_(csvText) {
  var lines = String(csvText).split(/\r?\n/);
  var firstLine = null;
  for (var i = 0; i < lines.length; i++) {
    if (lines[i].trim().length > 0) { firstLine = lines[i]; break; }
  }
  if (!firstLine) return null;
  var firstCell = firstLine.split(',')[0].trim().replace(/^"|"$/g, '');

  var m = firstCell.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));

  m = firstCell.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{2,4})$/);
  if (m) {
    var year = m[3];
    if (year.length === 2) year = (Number(year) <= 79 ? '20' : '19') + year;
    return new Date(Number(year), Number(m[2]) - 1, Number(m[1]));
  }
  return null;
}

/** Sắp xếp lại toàn bộ dữ liệu (từ hàng 2 trở đi) của 1 sheet theo các cột chỉ định - gọi lại sau mỗi lần thêm/xoá dòng để luôn xem theo thứ tự ngày, không theo thứ tự thao tác. */
function sortSheetRows_(sheet, sortSpecs) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 3) return; // cần ít nhất 2 dòng dữ liệu mới cần sắp xếp
  sheet.getRange(2, 1, lastRow - 1, lastCol).sort(sortSpecs);
}

/** Ghi/ghi đè 1 dòng CSV_DATA cho (ngày, mã công tơ) - dùng chung cho luồng lưu 1 file và lưu hàng loạt. Tự sắp xếp lại theo Ngày rồi Mã công tơ sau khi ghi. */
function saveCsvRow_(date, meterCode, kwhGiao) {
  var dataSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CSV_DATA);
  var lastRow = dataSh.getLastRow();
  if (lastRow >= 2) {
    var existing = dataSh.getRange(2, 1, lastRow - 1, 2).getValues();
    var dateStrFmt = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    for (var i = existing.length - 1; i >= 0; i--) {
      var rowDate = existing[i][0];
      var rowDateStr = rowDate instanceof Date
        ? Utilities.formatDate(rowDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(rowDate);
      if (rowDateStr === dateStrFmt && String(existing[i][1]) === meterCode) {
        dataSh.deleteRow(i + 2);
      }
    }
  }
  dataSh.appendRow([date, meterCode].concat(kwhGiao));
  sortSheetRows_(dataSh, [{ column: 1, ascending: true }, { column: 2, ascending: true }]);
}

/**
 * Đọc 48 giá trị KwhGiao theo (ngày, tổ máy, loại dữ liệu) - mã công tơ
 * thật được TRA TỪ CAI_DAT (khác nhau giữa các tổ máy/nhà máy, không cố
 * định "6001"/"6303").
 * @param {Date} date
 * @param {string} unit  "S1"|"S2"
 * @param {string} role  "Qdc"|"Qmp"
 * @returns {number[]|null}  48 giá trị, hoặc null nếu chưa có dữ liệu/chưa cấu hình mã công tơ
 */
function readCsv48_(date, unit, role) {
  var meterCode = resolveMeterCode_(unit, role);
  if (!meterCode) return null;
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CSV_DATA);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return null;
  var rows = sh.getRange(2, 1, lastRow - 1, 2 + 48).getValues();
  var dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  for (var i = 0; i < rows.length; i++) {
    var rowDate = rows[i][0];
    var rowDateStr = rowDate instanceof Date
      ? Utilities.formatDate(rowDate, Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : String(rowDate);
    if (rowDateStr === dateStr && String(rows[i][1]) === meterCode) {
      return rows[i].slice(2, 50);
    }
  }
  return null;
}

/**
 * Lấy P0 cho (ngày, tổ máy) - ưu tiên giá trị nhập tay trong P0_NGAY, nếu
 * không có thì SUY RA từ chu kỳ cuối (48) của KẾT QUẢ NGÀY HÔM TRƯỚC đã
 * tính (nếu có). Đây là XẤP XỈ, không phải carry-over R07 đầy đủ (không
 * mô phỏng ramp còn dở dang qua nửa đêm, chỉ lấy đúng Qdd cuối ngày trước
 * làm điểm bắt đầu) - đủ dùng cho phần lớn trường hợp thực tế (ramp
 * thường hoàn tất trong ngày), nhưng KHÔNG chính xác nếu ramp thật sự
 * đang dở dang lúc 24:00 (xem UAT-04, docs/09_Test_Cases.md).
 *
 * @returns {{value:number, source:string}|null}
 */
function readOrInferP0_(date, unit) {
  var manual = readP0_(date, unit);
  if (manual !== null) return { value: manual, source: 'nhập tay' };

  var prevDate = new Date(date);
  prevDate.setDate(prevDate.getDate() - 1);
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.KET_QUA);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return null;
  var rows = sh.getRange(2, 1, lastRow - 1, KET_QUA_HEADERS.length).getValues();
  var prevDateStr = Utilities.formatDate(prevDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var lastPeriodValue = null;
  rows.forEach(function (r) {
    var rowDate = r[0];
    var rowDateStr = rowDate instanceof Date
      ? Utilities.formatDate(rowDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(rowDate);
    if (rowDateStr === prevDateStr && String(r[1]).toUpperCase() === unit.toUpperCase() && Number(r[2]) === 48) {
      lastPeriodValue = Number(r[3]); // cột D = Qdd (MW)
    }
  });
  if (lastPeriodValue === null) return null;
  return { value: lastPeriodValue, source: 'suy ra từ chu kỳ cuối ngày ' + Utilities.formatDate(prevDate, Session.getScriptTimeZone(), 'dd/MM') + ' (xấp xỉ, xem README)' };
}

/** Lấy P0 đã nhập tay cho (ngày, tổ máy) từ sheet P0_NGAY - xem giới hạn ở README (chưa tự động carry-over R07). */
function readP0_(date, unit) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.P0_NGAY);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return null;
  var rows = sh.getRange(2, 1, lastRow - 1, 3).getValues();
  var dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  for (var i = 0; i < rows.length; i++) {
    var rowDate = rows[i][0];
    var rowDateStr = rowDate instanceof Date
      ? Utilities.formatDate(rowDate, Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : String(rowDate);
    if (rowDateStr === dateStr && String(rows[i][1]).toUpperCase() === unit.toUpperCase()) {
      return Number(rows[i][2]);
    }
  }
  return null;
}

/** Ghi kết quả 48 chu kỳ của 1 (ngày, tổ máy) vào KET_QUA (nối thêm, không xoá kết quả ngày khác). */
function appendResultToSheet_(date, unit, periods) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.KET_QUA);
  var rows = periods.map(function (p) {
    return [date, unit, p.chuKy, p.qdd, p.qddV, p.qdc, p.pQdc, p.nguongDuoi, p.nguongTren, p.qmp, p.qdu, p.dauHieu];
  });
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, KET_QUA_HEADERS.length).setValues(rows);
  sortSheetRows_(sh, [{ column: 1, ascending: true }, { column: 2, ascending: true }, { column: 3, ascending: true }]);
}

/** Xoá sạch KET_QUA trước khi tính lại (tránh trùng lặp khi chạy lại cùng ngày). */
function clearResultsForDate_(date, unit) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.KET_QUA);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  var data = sh.getRange(2, 1, lastRow - 1, 2).getValues();
  var dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var keepRows = [];
  for (var i = 0; i < data.length; i++) {
    var rowDate = data[i][0];
    var rowDateStr = rowDate instanceof Date
      ? Utilities.formatDate(rowDate, Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : String(rowDate);
    var isMatch = rowDateStr === dateStr && String(data[i][1]).toUpperCase() === unit.toUpperCase();
    if (!isMatch) keepRows.push(i + 2);
  }
  // Xoá toàn bộ rồi ghi lại các dòng cần giữ - đơn giản, đủ dùng cho quy mô 1 nhà máy.
  if (keepRows.length === data.length) return; // không có gì trùng để xoá
  var all = sh.getRange(2, 1, lastRow - 1, KET_QUA_HEADERS.length).getValues();
  var filtered = all.filter(function (_, idx) { return keepRows.indexOf(idx + 2) !== -1; });
  sh.getRange(2, 1, lastRow - 1, KET_QUA_HEADERS.length).clearContent();
  if (filtered.length > 0) {
    sh.getRange(2, 1, filtered.length, KET_QUA_HEADERS.length).setValues(filtered);
  }
}
