/**
 * Xuất báo cáo cho 1 hoặc nhiều ngày, nhiều tổ máy, thành 1 file Excel
 * hoặc PDF riêng - đọc từ dữ liệu ĐÃ CÓ trong KET_QUA (không tính lại).
 *
 * Layout bám sát file "Kiểm tra Qdu" gốc:
 *   A1        : "MWh" (đơn vị)
 *   Hàng 2    : nhãn tổ máy (vd "S1DH1") gộp trên mỗi khối
 *   Hàng 3    : tên cột; A3 = "Chu kỳ"
 *   Hàng 4-51 : 48 chu kỳ, cột A dạng "01 [00:00-00:30]" (số + khung giờ CÙNG 1 ô)
 *   Hàng 52   : tổng ngày (các cột MWh)
 * Các khối tổ máy nằm CẠNH NHAU, liền cột (B:K = tổ 1, L:U = tổ 2).
 */

// Thứ tự cột khớp file gốc: Qdd, Qdd_V, Qdc, Qmp, Qdư, Qdư âm/dương, rồi các cột đối chiếu, cuối cùng là Ghi chú.
var EXPORT_METRIC_COLUMNS = ['Qdd', 'Qdd_V', 'Qdc', 'Qmp', 'Qdư', 'Qdư âm/dương', 'P_Qdc', 'Ngưỡng dưới', 'Ngưỡng trên', 'Ghi chú'];
var EXPORT_BLOCK_WIDTH = EXPORT_METRIC_COLUMNS.length; // 10
var EXPORT_BLOCK_STRIDE = EXPORT_BLOCK_WIDTH; // các khối liền nhau, không chừa cột trống (giống file gốc)
var EXPORT_FIRST_DATA_ROW = 4;
var EXPORT_TOTAL_ROW = EXPORT_FIRST_DATA_ROW + 48; // hàng 52
/** Vị trí (0-based) trong EXPORT_METRIC_COLUMNS của các cột cần cộng tổng ngày (chỉ các cột MWh). */
var EXPORT_TOTAL_COL_OFFSETS = [1, 2, 3, 4]; // Qdd_V, Qdc, Qmp, Qdư

/**
 * Chuyển 1 Google Sheets thành blob Excel/PDF bằng link xuất trực tiếp
 * của Google (docs.google.com/.../export) - KHÔNG dùng
 * DriveApp.getFileById().getAs() vì hàm đó không hỗ trợ chuyển Google
 * Sheets sang .xlsx (chỉ hỗ trợ vài kiểu chuyển đổi cố định).
 * @param {string} spreadsheetId
 * @param {string} format  "xlsx" | "pdf"
 * @returns {Blob}
 */
function exportSpreadsheetAsBlob_(spreadsheetId, format) {
  var url = 'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/export?format=' + format;
  var response = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true,
  });
  if (response.getResponseCode() !== 200) {
    throw new Error('Không xuất được file (' + format + '), mã lỗi ' + response.getResponseCode() + '.');
  }
  return response.getBlob();
}

/** Nhãn hiển thị của tổ máy trong báo cáo (vd "S1DH1"), lấy từ CAI_DAT; không có thì dùng "Tổ S1". */
function reportUnitLabel_(unit) {
  var key = 'REPORT_LABEL_' + unit.toUpperCase();
  var label = CAI_DAT_LABELS[key] ? getConfigValue_(CAI_DAT_LABELS[key]) : null;
  return label ? String(label).trim() : 'Tổ ' + unit;
}

/**
 * @param {Date} date
 * @param {string} unit
 * @returns {Array[]|null} 48 dòng theo đúng thứ tự EXPORT_METRIC_COLUMNS (không gồm cột Chu kỳ), hoặc null nếu chưa có
 */
function readKetQuaMetricsForExport_(date, unit) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.KET_QUA);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return null;
  var rows = sh.getRange(2, 1, lastRow - 1, KET_QUA_HEADERS.length).getValues();
  var dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var byChuKy = {};
  rows.forEach(function (r) {
    var rowDate = r[0];
    var rowDateStr = rowDate instanceof Date
      ? Utilities.formatDate(rowDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(rowDate);
    if (rowDateStr === dateStr && String(r[1]).toUpperCase() === unit.toUpperCase()) {
      // KET_QUA: 3=Qdd, 4=Qdd_V, 5=Qdc, 6=P_Qdc, 7=Ngưỡng dưới, 8=Ngưỡng trên, 9=Qmp, 10=Qdư, 11=Dấu hiệu
      byChuKy[parsePeriodNumber_(r[2])] = [r[3], r[4], r[5], r[9], r[10], exportSignLabel_(r[11]), r[6], r[7], r[8], ''];
    }
  });
  if (Object.keys(byChuKy).length === 0) return null;
  var out = [];
  for (var i = 1; i <= 48; i++) {
    out.push(byChuKy[i] || emptyMetricRow_());
  }
  return out;
}

/** 48 dòng trống - dùng cho tổ máy chưa có dữ liệu (vẫn giữ bảng trong báo cáo). */
function emptyMetricsForExport_() {
  var out = [];
  for (var i = 0; i < 48; i++) out.push(emptyMetricRow_());
  return out;
}

function emptyMetricRow_() {
  return EXPORT_METRIC_COLUMNS.map(function () { return ''; });
}

/**
 * Nhãn cột "Qdư âm/dương" trong file xuất, bám đúng bản "Kiểm tra Qdu" gốc:
 * chỉ ghi "âm"/"dương", còn khi Qdư = 0 (nằm trong dải dung sai) thì để dấu
 * "-" - không ghi chữ "trong ±3%" như sheet KET_QUA nội bộ.
 */
function exportSignLabel_(value) {
  var s = String(value == null ? '' : value).trim().toLowerCase();
  if (s === 'âm' || s === 'dương') return s;
  return '-';
}

function uniqueSheetName_(ss, baseName) {
  var name = baseName, n = 2;
  while (ss.getSheetByName(name)) {
    name = baseName + ' (' + n + ')';
    n++;
  }
  return name.slice(0, 100);
}

/**
 * @param {Date[]} dates
 * @param {string[]} units  vd ["S1","S2"] - áp dụng cho mọi ngày trong danh sách
 * @param {string} format  "xlsx" | "pdf"
 * @returns {{url:string, name:string, missing:string[]}}
 */
function buildAndExportReport_(dates, units, format) {
  var missing = [];
  var temp = SpreadsheetApp.create('QDD_export_tmp_' + new Date().getTime());
  var usedFirstSheet = false;

  dates.forEach(function (date) {
    var dateLabel = Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd-MM-yyyy');
    // Luôn dựng bảng cho MỌI tổ máy được chọn - tổ nào chưa có dữ liệu thì
    // để bảng trống, đúng như bản "Kiểm tra Qdu" gốc (S1 và S2 luôn nằm
    // cạnh nhau, bên nào chưa nhập thì bỏ trống chứ không mất bảng).
    var blocks = [];
    var hasAnyData = false;
    units.forEach(function (unit) {
      var rows = readKetQuaMetricsForExport_(date, unit);
      if (!rows) {
        missing.push(dateLabel + ' ' + unit + ': chưa có trong KET_QUA (để bảng trống trong báo cáo)');
        blocks.push({ unit: unit, rows: emptyMetricsForExport_(), empty: true });
        return;
      }
      hasAnyData = true;
      blocks.push({ unit: unit, rows: rows });
    });
    // Ngày mà KHÔNG tổ máy nào có dữ liệu thì bỏ hẳn, không tạo tab rỗng.
    if (!hasAnyData) return;

    var sheetName = uniqueSheetName_(temp, dateLabel);
    var sh;
    if (!usedFirstSheet) {
      sh = temp.getSheets()[0];
      sh.setName(sheetName);
      usedFirstSheet = true;
    } else {
      sh = temp.insertSheet(sheetName);
    }

    var lastCol = 1 + blocks.length * EXPORT_BLOCK_STRIDE;

    // Hàng 1: đơn vị + tiêu đề ngày (gốc chỉ có "MWh" ở A1, thêm ngày ở cột B cho rõ khi xuất nhiều ngày)
    sh.getRange(1, 1).setValue('MWh').setFontWeight('bold');
    sh.getRange(1, 2, 1, lastCol - 1).merge().setFontWeight('bold')
      .setValue('BÁO CÁO Qdd/Qdư - Ngày ' + Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd/MM/yyyy'));

    sh.getRange(3, 1).setValue('Chu kỳ').setFontWeight('bold');
    blocks.forEach(function (block, i) {
      var startCol = 2 + i * EXPORT_BLOCK_STRIDE;
      sh.getRange(2, startCol, 1, EXPORT_BLOCK_WIDTH).merge()
        .setValue(reportUnitLabel_(block.unit))
        .setFontWeight('bold').setHorizontalAlignment('center');
      sh.getRange(3, startCol, 1, EXPORT_BLOCK_WIDTH).setValues([EXPORT_METRIC_COLUMNS]).setFontWeight('bold');
      sh.getRange(EXPORT_FIRST_DATA_ROW, startCol, block.rows.length, EXPORT_BLOCK_WIDTH).setValues(block.rows);

      // Bảng trống thì không cộng tổng (tránh hiện 0 gây hiểu nhầm là đã tính ra 0).
      if (block.empty) return;
      EXPORT_TOTAL_COL_OFFSETS.forEach(function (offset) {
        var col = startCol + offset;
        var colLetter = columnLetter_(col);
        sh.getRange(EXPORT_TOTAL_ROW, col).setFormula(
          '=SUM(' + colLetter + EXPORT_FIRST_DATA_ROW + ':' + colLetter + (EXPORT_TOTAL_ROW - 1) + ')'
        ).setFontWeight('bold');
      });
    });

    var chuKyRows = [];
    for (var i = 1; i <= 48; i++) chuKyRows.push([periodLabel_(i)]);
    sh.getRange(EXPORT_FIRST_DATA_ROW, 1, 48, 1).setValues(chuKyRows);
    sh.getRange(EXPORT_TOTAL_ROW, 1).setValue('Tổng ngày').setFontWeight('bold');

    sh.setFrozenRows(3);
    // Không cố định cột: hàng 1 có ô gộp trải nhiều cột, cố định cột sẽ bị
    // Google Sheets từ chối vì cắt ngang ô hợp nhất.
    sh.autoResizeColumns(1, lastCol);
  });

  if (!usedFirstSheet) {
    DriveApp.getFileById(temp.getId()).setTrashed(true);
    throw new Error('Không có ngày/tổ máy nào trong danh sách đã được tính (xem KET_QUA). Không có gì để xuất.');
  }

  SpreadsheetApp.flush();
  var ext = format === 'pdf' ? '.pdf' : '.xlsx';
  var blob = exportSpreadsheetAsBlob_(temp.getId(), format);

  var fileName = 'BaoCao_QDD_' + new Date().getTime() + ext;
  var parents = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId()).getParents();
  var folder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  var outFile = folder.createFile(blob).setName(fileName);

  DriveApp.getFileById(temp.getId()).setTrashed(true);

  return { url: outFile.getUrl(), name: fileName, missing: missing };
}

/** 1 -> "A", 27 -> "AA" (dùng để dựng công thức SUM theo cột). */
function columnLetter_(col) {
  var letter = '';
  while (col > 0) {
    var rem = (col - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}
