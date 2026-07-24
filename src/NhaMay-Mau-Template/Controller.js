/**
 * Hàm server được gọi từ Sidebar.html qua google.script.run. Mỗi hàm trả
 * về 1 chuỗi text hiển thị trực tiếp lên sidebar (thành công) hoặc ném
 * lỗi (withFailureHandler bên client bắt lại).
 */

function sidebar_saveCsv(dateStr, unit, role, csvText) {
  var date = parseIsoDate_(dateStr);
  if (!date) throw new Error('Ngày không hợp lệ.');

  var meterCode = resolveMeterCode_(unit, role);
  if (!meterCode) {
    throw new Error('Chưa cấu hình mã công tơ ' + role + ' cho tổ ' + unit +
      ' trong sheet CAI_DAT (dòng "Mã công tơ ' + role + ' - ' + unit + '"). Điền mã công tơ thật rồi thử lại.');
  }

  var kwhGiao = QDDCoreLibrary.extractKwhGiaoFromCsv(csvText);

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
  return 'Đã lưu CSV ' + role + ' (công tơ ' + meterCode + ', tổ ' + unit + ') cho ngày ' +
    Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd/MM/yyyy') + '.';
}

function sidebar_calcOneDay(dateStr, unit) {
  var date = parseIsoDate_(dateStr);
  if (!date) throw new Error('Ngày không hợp lệ.');

  var result = computeOneDay_(date, unit);
  if (result.error) throw new Error(result.error);

  clearResultsForDate_(date, unit);
  appendResultToSheet_(date, unit, result.periods);
  return 'Đã tính xong ' + Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd/MM/yyyy') +
    ' - tổ ' + unit + ' (P0 ' + result.p0Source + '). Xem sheet KET_QUA.';
}

function sidebar_calcBatch(fromStr, toStr, units) {
  var fromDate = parseIsoDate_(fromStr);
  var toDate = parseIsoDate_(toStr);
  if (!fromDate || !toDate || toDate < fromDate) throw new Error('Khoảng ngày không hợp lệ.');

  var okCount = 0, errDetails = [];
  for (var d = new Date(fromDate); d <= toDate; d.setDate(d.getDate() + 1)) {
    var date = new Date(d);
    units.forEach(function (unit) {
      var result = computeOneDay_(date, unit);
      var label = Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd/MM') + ' ' + unit;
      if (result.error) {
        errDetails.push(label + ': ' + result.error);
      } else {
        clearResultsForDate_(date, unit);
        appendResultToSheet_(date, unit, result.periods);
        okCount++;
      }
    });
  }

  var msg = 'Đã tính xong ' + okCount + ' (ngày, tổ máy).';
  if (errDetails.length > 0) {
    msg += ' Lỗi/thiếu dữ liệu: ' + errDetails.length + '.\n' + errDetails.slice(0, 15).join('\n');
    if (errDetails.length > 15) msg += '\n... và ' + (errDetails.length - 15) + ' dòng khác.';
  }
  return msg;
}

function sidebar_monthlyReport(monthStr) {
  var m = String(monthStr).trim().match(/^(\d{4})-(\d{2})$/); // input type=month -> yyyy-MM
  if (!m) throw new Error('Định dạng tháng không hợp lệ.');
  var year = Number(m[1]), month = Number(m[2]) - 1;

  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.KET_QUA);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) throw new Error('KET_QUA chưa có dữ liệu.');
  var rows = sh.getRange(2, 1, lastRow - 1, KET_QUA_HEADERS.length).getValues();

  var byDayUnit = {};
  rows.forEach(function (r) {
    var date = r[0];
    if (!(date instanceof Date) || date.getMonth() !== month || date.getFullYear() !== year) return;
    var unit = r[1];
    var key = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd') + '|' + unit;
    if (!byDayUnit[key]) byDayUnit[key] = { date: date, unit: unit, periods: [] };
    byDayUnit[key].periods.push({ qdc: r[5], qddV: r[4], qmp: r[9], qdu: r[10] });
  });

  var dayResults = Object.keys(byDayUnit).map(function (k) { return byDayUnit[k]; });
  if (dayResults.length === 0) throw new Error('Không có dữ liệu KET_QUA cho tháng ' + m[2] + '/' + m[1] + '.');

  var report = QDDCoreLibrary.aggregateMonthlyReport(dayResults);

  var outSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.BAO_CAO_THANG);
  var oldLast = outSh.getLastRow();
  if (oldLast > 1) outSh.getRange(2, 1, oldLast - 1, BAO_CAO_THANG_HEADERS.length).clearContent();
  var outRows = report.perDay.map(function (r) {
    return [r.date, r.unit, r.tongQdc || '', r.tongQmp || '', r.tongQddV || '', r.tongQdu || ''];
  });
  if (outRows.length > 0) {
    outSh.getRange(2, 1, outRows.length, BAO_CAO_THANG_HEADERS.length).setValues(outRows);
  }
  return 'Đã tổng hợp ' + report.tongHop.soNgay + ' (ngày, tổ máy). Tổng Qdư: ' +
    report.tongHop.tongQdu.toFixed(3) + ' MWh. Xem sheet BAO_CAO_THANG.';
}
