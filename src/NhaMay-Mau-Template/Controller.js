/**
 * Hàm server được gọi từ Sidebar.html qua google.script.run. Mỗi hàm trả
 * về 1 chuỗi text hiển thị trực tiếp lên sidebar (thành công) hoặc ném
 * lỗi (withFailureHandler bên client bắt lại).
 */

function sidebar_importCommands() {
  var result = importCommandsFromStaging_();
  var msg = 'Đã nhập ' + result.imported + ' lệnh mới, cập nhật ' + result.updated + ' lệnh đã có (trùng ID Lệnh).';
  if (result.skipped.length > 0) {
    msg += ' Bỏ qua ' + result.skipped.length + ' dòng lỗi:\n' +
      result.skipped.slice(0, 15).map(function (s) { return 'Dòng ' + s.row + ': ' + s.reason; }).join('\n');
    if (result.skipped.length > 15) msg += '\n... và ' + (result.skipped.length - 15) + ' dòng khác.';
  }
  return msg;
}

function sidebar_saveCsv(dateStr, unit, role, csvText, filename) {
  var date = parseIsoDate_(dateStr);
  if (!date) throw new Error('Ngày không hợp lệ.');

  var meterCode = resolveMeterCode_(unit, role);
  if (!meterCode) {
    throw new Error('Chưa cấu hình mã công tơ ' + role + ' cho tổ ' + unit +
      ' trong sheet CAI_DAT (dòng "Mã công tơ ' + role + ' - ' + unit + '"). Điền mã công tơ thật rồi thử lại.');
  }

  // Chặn lỗi chọn nhầm file: nếu tên file khớp một mã công tơ đã cấu hình
  // KHÁC với mã của (tổ máy, loại dữ liệu) đang chọn thì dừng lại. Lỗi này
  // từng xảy ra thật và gây sai số âm thầm (lưu dữ liệu Qmp vào ô Qdc).
  if (filename) {
    var fromName = matchMeterFromFilename_(filename);
    if (fromName && fromName.code !== meterCode) {
      throw new Error('Tên file "' + filename + '" ứng với công tơ ' + fromName.code +
        ' (' + fromName.role + ' - ' + fromName.unit + '), nhưng bạn đang chọn ' + role + ' - ' + unit +
        ' (công tơ ' + meterCode + '). Chọn lại cho khớp, hoặc dùng mục "Tải CSV hàng loạt" để hệ thống tự nhận diện.');
    }
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

function sidebar_calcOneDay(dateStr, unit, cleanupSource) {
  var date = parseIsoDate_(dateStr);
  if (!date) throw new Error('Ngày không hợp lệ.');

  var result = computeOneDay_(date, unit);
  if (result.error) throw new Error(result.error);

  clearResultsForDate_(date, unit);
  appendResultToSheet_(date, unit, result.periods);
  saveNextDayP0_(date, unit, result.endPower);

  var msg = 'Đã tính xong ' + Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd/MM/yyyy') +
    ' - tổ ' + unit + ' (P0 ' + result.p0Source + '). Xem sheet KET_QUA.';
  if (cleanupSource) {
    var removed = clearSourceDataForDate_(date, unit);
    msg += ' Đã dọn ' + removed.lenh + ' dòng lệnh và ' + removed.csv + ' dòng CSV của ngày này.';
  }
  if (result.warnings && result.warnings.length > 0) {
    msg += '\n\n⚠ CẢNH BÁO:\n' + result.warnings.join('\n');
  }
  return msg;
}

function sidebar_calcBatch(fromStr, toStr, units, cleanupSource) {
  var fromDate = parseIsoDate_(fromStr);
  var toDate = parseIsoDate_(toStr);
  if (!fromDate || !toDate || toDate < fromDate) throw new Error('Khoảng ngày không hợp lệ.');

  var okCount = 0, errDetails = [], allWarnings = [], cleaned = { lenh: 0, csv: 0 };
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
        saveNextDayP0_(date, unit, result.endPower);
        okCount++;
        if (result.warnings && result.warnings.length > 0) {
          result.warnings.forEach(function (w) { allWarnings.push(label + ': ' + w); });
        }
        if (cleanupSource) {
          var removed = clearSourceDataForDate_(date, unit);
          cleaned.lenh += removed.lenh;
          cleaned.csv += removed.csv;
        }
      }
    });
  }

  var msg = 'Đã tính xong ' + okCount + ' (ngày, tổ máy).';
  if (cleanupSource && (cleaned.lenh > 0 || cleaned.csv > 0)) {
    msg += ' Đã dọn ' + cleaned.lenh + ' dòng lệnh và ' + cleaned.csv + ' dòng CSV.';
  }
  if (errDetails.length > 0) {
    msg += ' Lỗi/thiếu dữ liệu: ' + errDetails.length + '.\n' + errDetails.slice(0, 15).join('\n');
    if (errDetails.length > 15) msg += '\n... và ' + (errDetails.length - 15) + ' dòng khác.';
  }
  if (allWarnings.length > 0) {
    msg += '\n\n⚠ CẢNH BÁO (' + allWarnings.length + '):\n' + allWarnings.slice(0, 10).join('\n');
    if (allWarnings.length > 10) msg += '\n... và ' + (allWarnings.length - 10) + ' cảnh báo khác.';
  }
  return msg;
}

/**
 * Tổng hợp số liệu tháng vào BAO_CAO_THANG VÀ xuất file gộp tất cả các
 * ngày đã tính trong tháng đó (mỗi ngày 1 tab, tổ máy cạnh nhau - giống
 * mục "Xuất báo cáo"), chỉ cho các tổ máy được chọn.
 * @returns {{html: string}}
 */
function sidebar_monthlyReport(monthStr, units, format) {
  var m = String(monthStr).trim().match(/^(\d{4})-(\d{2})$/); // input type=month -> yyyy-MM
  if (!m) throw new Error('Định dạng tháng không hợp lệ.');
  if (!units || units.length === 0) throw new Error('Chọn ít nhất 1 tổ máy.');
  var year = Number(m[1]), month = Number(m[2]) - 1;

  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.KET_QUA);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) throw new Error('KET_QUA chưa có dữ liệu.');
  var rows = sh.getRange(2, 1, lastRow - 1, KET_QUA_HEADERS.length).getValues();

  var byDayUnit = {};
  var dateSet = {};
  rows.forEach(function (r) {
    var date = r[0];
    if (!(date instanceof Date) || date.getMonth() !== month || date.getFullYear() !== year) return;
    var unit = r[1];
    if (units.indexOf(unit) === -1) return;
    var dateKey = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var key = dateKey + '|' + unit;
    if (!byDayUnit[key]) byDayUnit[key] = { date: date, unit: unit, periods: [] };
    byDayUnit[key].periods.push({ qdc: r[5], qddV: r[4], qmp: r[9], qdu: r[10] });
    dateSet[dateKey] = date;
  });

  var dayResults = Object.keys(byDayUnit).map(function (k) { return byDayUnit[k]; });
  if (dayResults.length === 0) throw new Error('Không có dữ liệu KET_QUA cho tháng ' + m[2] + '/' + m[1] + ' (với tổ máy đã chọn).');

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

  var dates = Object.keys(dateSet).sort().map(function (k) { return dateSet[k]; });
  var exportResult = buildAndExportReport_(dates, units, format);

  var html = 'Đã tổng hợp ' + report.tongHop.soNgay + ' (ngày, tổ máy). Tổng Qdư: ' +
    report.tongHop.tongQdu.toFixed(3) + ' MWh (xem sheet BAO_CAO_THANG).<br><br>' +
    '✓ File báo cáo tháng: <a href="' + exportResult.url + '" target="_blank">' + exportResult.name + '</a>';
  if (exportResult.missing.length > 0) {
    html += '<br><br>Bỏ qua (chưa có dữ liệu):<br>' + exportResult.missing.join('<br>');
  }
  return { html: html };
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
