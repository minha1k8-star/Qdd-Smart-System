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
    return { error: 'Chưa có P0 (không có trong P0_NGAY, và ngày trước đó chưa được tính). Cần điền tay P0_NGAY cho ngày đầu tiên.' };
  }
  var qdc48 = readCsv48_(date, '6001');
  var qmp48 = readCsv48_(date, '6303');
  if (!qdc48 || !qmp48) {
    return { error: 'Thiếu dữ liệu CSV (6001 hoặc 6303) cho ngày này trong CSV_DATA.' };
  }
  try {
    var periods = QDDCoreLibrary.calculateDay({
      effectiveCommands: effectiveCommands, p0: p0Info.value, qdc48: qdc48, qmp48: qmp48,
      qddVCoef: config.qddVCoef, tolerance: config.tolerance,
    });
    return { periods: periods, p0Source: p0Info.source };
  } catch (e) {
    return { error: e.message };
  }
}
