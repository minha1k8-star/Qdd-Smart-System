/**
 * Menu điều khiển - chỉ 2 mục: mở bảng điều khiển (sidebar) và thiết lập
 * sheet lần đầu. Toàn bộ thao tác nhập liệu/tính toán nằm trong Sidebar.html
 * + Controller.js (thay cho các hộp thoại prompt() gõ tay trước đây).
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('QDD Smart System')
    .addItem('Bảng điều khiển', 'showSidebar')
    .addItem('Thiết lập sheet (chạy 1 lần đầu)', 'setupAllSheets')
    .addToUi();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('QDD Smart System');
  SpreadsheetApp.getUi().showSidebar(html);
}

/** "yyyy-MM-dd" (định dạng input type=date của HTML) -> Date */
function parseIsoDate_(str) {
  var m = String(str).trim().match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return null;
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}

/**
 * @returns {{periods: Array, p0Source?: string}|{error: string}}
 */
function computeOneDay_(date, unit) {
  var config = getConfig_();
  var allCommands = readAllCommands_();
  var effectiveCommands = QDDCoreLibrary.selectEffectiveCommands(allCommands, date, unit);
  var p0Info = readOrInferP0_(date, unit);
  if (p0Info === null) {
    return { error: 'Chưa có P0 cho ngày/tổ máy này trong sheet P0_NGAY. ' +
      'P0 được ghi tự động sau khi tính ngày liền trước - nếu đây là ngày đầu tiên (hoặc ngày trước chưa tính), hãy điền tay 1 dòng vào P0_NGAY.' };
  }
  var qdc48 = readCsv48_(date, unit, 'Qdc');
  var qmp48 = readCsv48_(date, unit, 'Qmp');
  if (!qdc48 || !qmp48) {
    return { error: 'Thiếu dữ liệu CSV Qdc/Qmp cho (ngày, tổ máy) này - kiểm tra đã lưu CSV chưa, và đã điền đúng mã công tơ trong CAI_DAT chưa.' };
  }
  try {
    var periods = QDDCoreLibrary.calculateDay({
      effectiveCommands: effectiveCommands, p0: p0Info.value, qdc48: qdc48, qmp48: qmp48,
      qddVCoef: config.qddVCoef, tolerance: config.tolerance,
      carryTarget: p0Info.carryTarget, // R07: ramp còn dở dang từ ngày trước
    });
    var warnings = detectZeroZeroCommands_(allCommands, date, unit);
    if (p0Info.carryTarget) {
      warnings.push('Ngày trước còn ramp dở dang lúc 24:00 -> đã tự chạy tiếp từ ' +
        p0Info.value + ' MW đến mục tiêu ' + p0Info.carryTarget + ' MW (quy tắc R07).');
    }
    if (effectiveCommands.length === 0 && !p0Info.carryTarget) {
      // Ngày không có lệnh nào vẫn tính được (Qdd giữ nguyên P0 cả ngày) và
      // đó có thể là đúng thực tế. Nhưng nếu chỉ vì QUÊN NHẬP LỆNH thì kết
      // quả sai mà trông vẫn "thành công" - lỗi này đã xảy ra thật với ngày
      // 19-20/07, nên phải cảnh báo rõ.
      warnings.push('Không có lệnh hiệu lực nào trong sheet LENH cho ngày/tổ máy này ' +
        '-> Qdd giữ nguyên P0 = ' + p0Info.value + ' MW suốt 48 chu kỳ. ' +
        'Kiểm tra lại xem đã nhập danh sách lệnh chưa (nếu ngày đó thực sự không có lệnh thì bỏ qua).');
    }
    return {
      periods: periods,
      p0Source: p0Info.source,
      endPower: periods.endPower,
      carry: periods.carry,
      warnings: warnings,
    };
  } catch (e) {
    return { error: e.message };
  }
}

/**
 * UAT-34: phát hiện lệnh "0-0" (CS ra lệnh = CS hoàn thành = 0, thường là
 * lệnh ngừng tổ máy do sự cố/trip). Theo quy tắc nghiệp vụ đã xác nhận,
 * loại lệnh này CHỦ ĐÍCH không được tính (xem docs/03_Business_Rules.md,
 * docs/15_Accuracy_Validation_2026-07.md) - nhưng nếu không cảnh báo,
 * công cụ sẽ âm thầm giữ nguyên công suất cũ đến hết ngày trong khi thực
 * tế tổ máy đã ngừng, gây sai số lớn mà người dùng không biết.
 *
 * Hàm này KHÔNG tự sửa số liệu, chỉ báo để người vận hành tự xử lý.
 * @returns {string[]} danh sách cảnh báo (rỗng nếu không có lệnh 0-0)
 */
function detectZeroZeroCommands_(allCommands, date, unit) {
  var tz = Session.getScriptTimeZone();
  var dateStr = Utilities.formatDate(date, tz, 'yyyy-MM-dd');
  var unitPrefix = unit.toUpperCase().slice(0, 2);
  var warnings = [];

  allCommands.forEach(function (c) {
    if (!(c.bdth instanceof Date)) return;
    if (Utilities.formatDate(c.bdth, tz, 'yyyy-MM-dd') !== dateStr) return;
    if (String(c.toMay || '').toUpperCase().slice(0, 2) !== unitPrefix) return;
    var completed = (c.hoanThanh === 1 || c.hoanThanh === true || String(c.hoanThanh).toUpperCase() === 'TRUE');
    if (!completed) return;
    if (Number(c.csRaLenh) === 0 && Number(c.csHoanThanh) === 0) {
      warnings.push('Lệnh 0-0 lúc ' + Utilities.formatDate(c.bdth, tz, 'HH:mm') +
        ' (' + (c.noiDungLenh || c.id) + ') KHÔNG được tính theo quy tắc nghiệp vụ. ' +
        'Nếu tổ máy thực sự đã ngừng, Qdd từ thời điểm này đang CAO HƠN thực tế - cần kiểm tra/điều chỉnh tay.');
    }
  });

  return warnings;
}
