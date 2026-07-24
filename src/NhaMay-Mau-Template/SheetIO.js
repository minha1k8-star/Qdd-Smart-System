/**
 * Đọc/ghi dữ liệu giữa Sheet và QDD-Core-Library. Toàn bộ tính toán thật
 * nằm ở Library (QDDCoreLibrary) - các hàm ở đây chỉ chuyển đổi định dạng.
 */

function getConfig_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CAI_DAT);
  var values = sh.getRange('B2:B4').getValues();
  return {
    rampRate: Number(values[0][0]),
    qddVCoef: Number(values[1][0]),
    tolerance: Number(values[2][0]),
  };
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
 * @param {Date} date
 * @param {string} meterCode  "6001" | "6303"
 * @returns {number[]|null}  48 giá trị, hoặc null nếu chưa có dữ liệu ngày đó
 */
function readCsv48_(date, meterCode) {
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
    if (rowDateStr === dateStr && String(rows[i][1]) === String(meterCode)) {
      return rows[i].slice(2, 50);
    }
  }
  return null;
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
