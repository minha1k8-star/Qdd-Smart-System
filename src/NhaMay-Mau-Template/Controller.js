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
  saveCsvRow_(date, meterCode, kwhGiao);
  return 'Đã lưu CSV ' + role + ' (công tơ ' + meterCode + ', tổ ' + unit + ') cho ngày ' +
    Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd/MM/yyyy') + '.';
}

/**
 * Tải hàng loạt nhiều file CSV cùng lúc - tự dò (tổ máy, loại dữ liệu) từ
 * TÊN FILE (khớp với mã công tơ đã cấu hình trong CAI_DAT) và tự đọc
 * ngày từ NỘI DUNG file. File nào không tự dò được sẽ bị bỏ qua, báo lại
 * tên file + lý do để xử lý tay bằng mục "Lưu CSV" (1 file).
 *
 * @param {Array<{filename:string, content:string}>} files
 * @returns {string}
 */
function sidebar_saveCsvBulk(files) {
  var savedCount = 0;
  var skipped = [];

  files.forEach(function (f) {
    var meterMatch = matchMeterFromFilename_(f.filename);
    if (!meterMatch) {
      skipped.push(f.filename + ': không khớp mã công tơ nào đã cấu hình trong CAI_DAT');
      return;
    }
    var date = guessDateFromCsvText_(f.content);
    if (!date) {
      skipped.push(f.filename + ': không đọc được ngày trong nội dung file');
      return;
    }
    var kwhGiao;
    try {
      kwhGiao = QDDCoreLibrary.extractKwhGiaoFromCsv(f.content);
    } catch (e) {
      skipped.push(f.filename + ': ' + e.message);
      return;
    }
    saveCsvRow_(date, meterMatch.code, kwhGiao);
    savedCount++;
  });

  var msg = 'Đã lưu ' + savedCount + '/' + files.length + ' file.';
  if (skipped.length > 0) {
    msg += ' Bỏ qua ' + skipped.length + ' file:\n' + skipped.slice(0, 15).join('\n');
    if (skipped.length > 15) msg += '\n... và ' + (skipped.length - 15) + ' file khác.';
    msg += '\n(Lưu tay các file này ở mục "Lưu CSV" phía trên.)';
  }
  return msg;
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

/**
 * Xuất báo cáo cho khoảng ngày + tổ máy đã chọn (chỉ đọc dữ liệu ĐÃ CÓ
 * trong KET_QUA, không tính lại) thành 1 file Excel/PDF riêng.
 * @returns {{html: string}}  Trả về HTML nhỏ (có link tải) để sidebar hiển thị trực tiếp.
 */
function sidebar_exportReport(fromStr, toStr, units, format) {
  var fromDate = parseIsoDate_(fromStr);
  var toDate = parseIsoDate_(toStr);
  if (!fromDate || !toDate || toDate < fromDate) throw new Error('Khoảng ngày không hợp lệ.');
  if (!units || units.length === 0) throw new Error('Chọn ít nhất 1 tổ máy.');

  var dates = [];
  for (var d = new Date(fromDate); d <= toDate; d.setDate(d.getDate() + 1)) {
    dates.push(new Date(d));
  }

  var result = buildAndExportReport_(dates, units, format);
  var html = '✓ Đã xuất file: <a href="' + result.url + '" target="_blank">' + result.name + '</a>';
  if (result.missing.length > 0) {
    html += '<br><br>Bỏ qua (chưa có dữ liệu):<br>' + result.missing.join('<br>');
  }
  return { html: html };
}
