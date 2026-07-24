/**
 * Xuất báo cáo cho 1 hoặc nhiều ngày, nhiều tổ máy, thành 1 file Excel
 * hoặc PDF riêng - đọc từ dữ liệu ĐÃ CÓ trong KET_QUA (không tính lại).
 *
 * Layout: MỖI NGÀY 1 TAB, các tổ máy đã chọn nằm CẠNH NHAU trong cùng tab
 * (vd cột B:J = S1, cột L:T = S2) - giống layout file báo cáo gốc
 * (BAO_CAO_QDU/"Kiểm tra Qdu" có S1DH1 và S2DH1 cạnh nhau), không phải
 * mỗi tổ máy 1 tab riêng.
 */

var EXPORT_METRIC_COLUMNS = ['Qdd (MW)', 'Qdd_V (MWh)', 'Qdc (MWh)', 'P_Qdc (MW)', 'Ngưỡng dưới', 'Ngưỡng trên', 'Qmp (MWh)', 'Qdư (MWh)', 'Dấu hiệu'];
var EXPORT_BLOCK_WIDTH = EXPORT_METRIC_COLUMNS.length; // 9
var EXPORT_BLOCK_STRIDE = EXPORT_BLOCK_WIDTH + 1; // +1 cột trống ngăn cách giữa các tổ máy

/**
 * @param {Date} date
 * @param {string} unit
 * @returns {Array[]|null} 48 dòng (đúng thứ tự EXPORT_METRIC_COLUMNS, KHÔNG có cột Chu kỳ), hoặc null nếu chưa có
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
      byChuKy[Number(r[2])] = [r[3], r[4], r[5], r[6], r[7], r[8], r[9], r[10], r[11]];
    }
  });
  var keys = Object.keys(byChuKy);
  if (keys.length === 0) return null;
  var out = [];
  for (var i = 1; i <= 48; i++) {
    out.push(byChuKy[i] || ['', '', '', '', '', '', '', '', '']);
  }
  return out;
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
    var blocks = [];
    units.forEach(function (unit) {
      var rows = readKetQuaMetricsForExport_(date, unit);
      if (!rows) {
        missing.push(dateLabel + ' ' + unit + ': chưa có trong KET_QUA (chưa tính ngày này)');
        return;
      }
      blocks.push({ unit: unit, rows: rows });
    });
    if (blocks.length === 0) return;

    var sheetName = uniqueSheetName_(temp, dateLabel);
    var sh;
    if (!usedFirstSheet) {
      sh = temp.getSheets()[0];
      sh.setName(sheetName);
      usedFirstSheet = true;
    } else {
      sh = temp.insertSheet(sheetName);
    }

    var lastCol = 1 + blocks.length * EXPORT_BLOCK_STRIDE - 1; // cột cuối cùng có dữ liệu (bỏ cột trống thừa sau khối cuối)
    sh.getRange(1, 1, 1, lastCol).merge().setFontWeight('bold')
      .setValue('BÁO CÁO Qdd/Qdư - Ngày ' + Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd/MM/yyyy'));

    sh.getRange(3, 1).setValue('Chu kỳ').setFontWeight('bold');
    blocks.forEach(function (block, i) {
      var startCol = 2 + i * EXPORT_BLOCK_STRIDE;
      sh.getRange(2, startCol, 1, EXPORT_BLOCK_WIDTH).merge().setFontWeight('bold')
        .setValue('Tổ ' + block.unit);
      sh.getRange(3, startCol, 1, EXPORT_BLOCK_WIDTH).setValues([EXPORT_METRIC_COLUMNS]).setFontWeight('bold');
      sh.getRange(4, startCol, block.rows.length, EXPORT_BLOCK_WIDTH).setValues(block.rows);
    });
    var chuKyRows = [];
    for (var i = 1; i <= 48; i++) chuKyRows.push([i]);
    sh.getRange(4, 1, 48, 1).setValues(chuKyRows);

    sh.setFrozenRows(3);
    sh.setFrozenColumns(1);
    sh.autoResizeColumns(1, lastCol);
  });

  if (!usedFirstSheet) {
    DriveApp.getFileById(temp.getId()).setTrashed(true);
    throw new Error('Không có ngày/tổ máy nào trong danh sách đã được tính (xem KET_QUA). Không có gì để xuất.');
  }

  SpreadsheetApp.flush();
  var mimeType = format === 'pdf' ? MimeType.PDF : MimeType.MICROSOFT_EXCEL;
  var ext = format === 'pdf' ? '.pdf' : '.xlsx';
  var blob = DriveApp.getFileById(temp.getId()).getAs(mimeType);

  var fileName = 'BaoCao_QDD_' + new Date().getTime() + ext;
  var parents = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId()).getParents();
  var folder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  var outFile = folder.createFile(blob).setName(fileName);

  DriveApp.getFileById(temp.getId()).setTrashed(true);

  return { url: outFile.getUrl(), name: fileName, missing: missing };
}
