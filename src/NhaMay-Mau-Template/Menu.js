/**
 * Menu điều khiển - tương đương 15 nút VBA cũ, nhưng rút gọn theo đúng
 * phần đã cài đặt ở QDD-Core-Library (chưa có carry-over R07, chưa có
 * cảnh báo lệnh 0-0 UAT-34, chưa xuất báo cáo dạng file riêng - xem
 * README.md trong thư mục này).
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('QDD Smart System')
    .addItem('1. Thiết lập sheet', 'setupAllSheets')
    .addSeparator()
    .addItem('2. Lưu CSV vừa nhập → công tơ 6001 (Qdc)', 'saveCsv6001FromStaging')
    .addItem('3. Lưu CSV vừa nhập → công tơ 6303 (Qmp)', 'saveCsv6303FromStaging')
    .addSeparator()
    .addItem('4. Tính 1 ngày', 'runSingleDay')
    .addItem('5. Tính nhiều ngày (khoảng thời gian)', 'runBatchRange')
    .addItem('6. Tổng hợp báo cáo tháng (từ KET_QUA đã có)', 'generateMonthlyReport')
    .addToUi();
}

function saveCsv6001FromStaging() { saveCsvFromStaging_('6001'); }
function saveCsv6303FromStaging() { saveCsvFromStaging_('6303'); }

/**
 * Đọc CSV_STAGING (Sheet đã Import CSV bằng File > Import trong Google
 * Sheets - tận dụng bộ đọc CSV có sẵn của Google, không tự viết lại logic
 * phát hiện Mac/Windows như VBA), trích dòng KwhGiao rồi lưu vào CSV_DATA
 * theo ngày người dùng nhập.
 */
function saveCsvFromStaging_(meterCode) {
  var ui = SpreadsheetApp.getUi();
  var staging = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CSV_STAGING);
  var values = staging.getDataRange().getValues();
  if (!values || values.length === 0) {
    ui.alert('CSV_STAGING đang trống. Dùng File > Import để nhập CSV vào sheet này trước.');
    return;
  }

  var kwhGiao;
  try {
    // QDDCoreLibrary expose toàn bộ namespace QDD (không chỉ các hàm ở Public.js)
    kwhGiao = QDDCoreLibrary.QDD.CsvParser.extractKwhGiao(values);
  } catch (e) {
    ui.alert('Lỗi đọc CSV: ' + e.message);
    return;
  }

  var dateResp = ui.prompt('Ngày của file CSV này (dd/mm/yyyy):');
  if (dateResp.getSelectedButton() !== ui.Button.OK) return;
  var date = parseDateInput_(dateResp.getResponseText());
  if (!date) { ui.alert('Ngày không hợp lệ.'); return; }

  var dataSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CSV_DATA);
  // Xoá dòng cũ cùng (ngày, mã công tơ) nếu có, tránh trùng khi nhập lại.
  var lastRow = dataSh.getLastRow();
  if (lastRow >= 2) {
    var existing = dataSh.getRange(2, 1, lastRow - 1, 2).getValues();
    var dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    for (var i = existing.length - 1; i >= 0; i--) {
      var rowDate = existing[i][0];
      var rowDateStr = rowDate instanceof Date
        ? Utilities.formatDate(rowDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(rowDate);
      if (rowDateStr === dateStr && String(existing[i][1]) === meterCode) {
        dataSh.deleteRow(i + 2);
      }
    }
  }
  dataSh.appendRow([date, meterCode].concat(kwhGiao));
  ui.alert('Đã lưu 48 giá trị KwhGiao (công tơ ' + meterCode + ') cho ngày ' +
    Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd/MM/yyyy') + '.');
}

function parseDateInput_(text) {
  var m = String(text).trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!m) return null;
  return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
}

function runSingleDay() {
  var ui = SpreadsheetApp.getUi();
  var dateResp = ui.prompt('Tính cho ngày nào? (dd/mm/yyyy)');
  if (dateResp.getSelectedButton() !== ui.Button.OK) return;
  var date = parseDateInput_(dateResp.getResponseText());
  if (!date) { ui.alert('Ngày không hợp lệ.'); return; }

  var unitResp = ui.prompt('Tổ máy? (S1 hoặc S2)');
  if (unitResp.getSelectedButton() !== ui.Button.OK) return;
  var unit = unitResp.getResponseText().trim().toUpperCase();

  var result = computeOneDay_(date, unit);
  if (result.error) {
    ui.alert('Không tính được: ' + result.error);
    return;
  }
  clearResultsForDate_(date, unit);
  appendResultToSheet_(date, unit, result.periods);
  ui.alert('Đã tính xong ngày ' + Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd/MM/yyyy') +
    ' - tổ ' + unit + '. Xem sheet KET_QUA.');
}

/**
 * @returns {{periods: Array}|{error: string}}
 */
function computeOneDay_(date, unit) {
  var config = getConfig_();
  var allCommands = readAllCommands_();
  var effectiveCommands = QDDCoreLibrary.selectEffectiveCommands(allCommands, date, unit);
  var p0 = readP0_(date, unit);
  if (p0 === null) {
    return { error: 'Chưa có P0 cho ngày/tổ máy này trong sheet P0_NGAY (chưa hỗ trợ tự động carry-over qua nửa đêm - xem README).' };
  }
  var qdc48 = readCsv48_(date, '6001');
  var qmp48 = readCsv48_(date, '6303');
  if (!qdc48 || !qmp48) {
    return { error: 'Thiếu dữ liệu CSV (6001 hoặc 6303) cho ngày này trong CSV_DATA.' };
  }
  try {
    var periods = QDDCoreLibrary.calculateDay({
      effectiveCommands: effectiveCommands, p0: p0, qdc48: qdc48, qmp48: qmp48,
      qddVCoef: config.qddVCoef, tolerance: config.tolerance,
    });
    return { periods: periods };
  } catch (e) {
    return { error: e.message };
  }
}

function runBatchRange() {
  var ui = SpreadsheetApp.getUi();
  var fromResp = ui.prompt('Tính hàng loạt - Từ ngày (dd/mm/yyyy):');
  if (fromResp.getSelectedButton() !== ui.Button.OK) return;
  var fromDate = parseDateInput_(fromResp.getResponseText());
  var toResp = ui.prompt('Đến ngày (dd/mm/yyyy):');
  if (toResp.getSelectedButton() !== ui.Button.OK) return;
  var toDate = parseDateInput_(toResp.getResponseText());
  if (!fromDate || !toDate || toDate < fromDate) { ui.alert('Khoảng ngày không hợp lệ.'); return; }

  var unitResp = ui.prompt('Tổ máy? (S1, S2, hoặc "S1,S2" để tính cả hai)');
  if (unitResp.getSelectedButton() !== ui.Button.OK) return;
  var units = unitResp.getResponseText().split(',').map(function (s) { return s.trim().toUpperCase(); }).filter(Boolean);

  var okCount = 0, errCount = 0, errDetails = [];
  for (var d = new Date(fromDate); d <= toDate; d.setDate(d.getDate() + 1)) {
    var date = new Date(d);
    units.forEach(function (unit) {
      var result = computeOneDay_(date, unit);
      var label = Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd/MM/yyyy') + ' ' + unit;
      if (result.error) {
        errCount++;
        errDetails.push(label + ': ' + result.error);
      } else {
        clearResultsForDate_(date, unit);
        appendResultToSheet_(date, unit, result.periods);
        okCount++;
      }
    });
  }

  var msg = 'Đã tính xong ' + okCount + ' (ngày, tổ máy). Lỗi/thiếu dữ liệu: ' + errCount + '.';
  if (errDetails.length > 0) {
    msg += '\n\nChi tiết:\n' + errDetails.slice(0, 20).join('\n');
    if (errDetails.length > 20) msg += '\n... và ' + (errDetails.length - 20) + ' dòng khác.';
  }
  ui.alert(msg);
}

/**
 * Tổng hợp báo cáo tháng TỪ KẾT QUẢ ĐÃ CÓ trong KET_QUA (không tính lại từ
 * đầu) - đúng theo thiết kế "không cần snapshot" (docs/05_System_Architecture.md):
 * mỗi ngày trong KET_QUA đã là kết quả đầy đủ, độc lập, tổng hợp bất cứ lúc nào.
 */
function generateMonthlyReport() {
  var ui = SpreadsheetApp.getUi();
  var monthResp = ui.prompt('Tổng hợp báo cáo tháng nào? (mm/yyyy)');
  if (monthResp.getSelectedButton() !== ui.Button.OK) return;
  var m = String(monthResp.getResponseText()).trim().match(/^(\d{1,2})\/(\d{4})$/);
  if (!m) { ui.alert('Định dạng tháng không hợp lệ, cần mm/yyyy.'); return; }
  var month = Number(m[1]) - 1, year = Number(m[2]);

  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.KET_QUA);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) { ui.alert('KET_QUA chưa có dữ liệu.'); return; }
  var rows = sh.getRange(2, 1, lastRow - 1, KET_QUA_HEADERS.length).getValues();

  var byDayUnit = {}; // key "yyyy-MM-dd|unit" -> {date, unit, periods:[]}
  rows.forEach(function (r) {
    var date = r[0];
    if (!(date instanceof Date) || date.getMonth() !== month || date.getFullYear() !== year) return;
    var unit = r[1];
    var key = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd') + '|' + unit;
    if (!byDayUnit[key]) byDayUnit[key] = { date: date, unit: unit, periods: [] };
    byDayUnit[key].periods.push({ qdc: r[5], qddV: r[4], qmp: r[9], qdu: r[10] });
  });

  var dayResults = Object.keys(byDayUnit).map(function (k) { return byDayUnit[k]; });
  if (dayResults.length === 0) { ui.alert('Không có dữ liệu KET_QUA cho tháng ' + m[1] + '/' + m[2] + '.'); return; }

  var report = QDDCoreLibrary.aggregateMonthlyReport(dayResults);

  var outSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.BAO_CAO_THANG);
  outSh.getRange(2, 1, Math.max(outSh.getLastRow() - 1, 0), BAO_CAO_THANG_HEADERS.length).clearContent();
  var outRows = report.perDay.map(function (r) {
    return [r.date, r.unit, r.tongQdc || '', r.tongQmp || '', r.tongQddV || '', r.tongQdu || ''];
  });
  if (outRows.length > 0) {
    outSh.getRange(2, 1, outRows.length, BAO_CAO_THANG_HEADERS.length).setValues(outRows);
  }
  ui.alert('Đã tổng hợp ' + report.tongHop.soNgay + ' (ngày, tổ máy) cho tháng ' + m[1] + '/' + m[2] +
    '.\nTổng Qdư: ' + report.tongHop.tongQdu.toFixed(3) + ' MWh.');
}
