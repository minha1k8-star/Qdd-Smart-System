/**
 * Điểm vào công khai của thư viện, dùng khi gắn làm Apps Script Library
 * vào Sheet của từng nhà máy (xem docs/05_System_Architecture.md).
 * Ví dụ gọi từ script của nhà máy (giả sử đặt tên thư viện là QDDCoreLibrary):
 *
 *   var effective = QDDCoreLibrary.selectEffectiveCommands(commands, targetDate, 'S1');
 *   var results = QDDCoreLibrary.calculateDay({
 *     effectiveCommands: effective, p0: 435.7, qdc48: qdc, qmp48: qmp
 *   });
 */

/**
 * @param {import('./CommandFilter').RawCommand[]} commands
 * @param {Date} targetDate
 * @param {string} unit
 */
function selectEffectiveCommands(commands, targetDate, unit) {
  return QDD.CommandFilter.selectEffective(commands, targetDate, unit);
}

/**
 * @param {Object} input  Xem QDD.QddCalculator.calculateDay
 */
function calculateDay(input) {
  return QDD.QddCalculator.calculateDay(input);
}

/**
 * @param {string} csvText
 */
function extractKwhGiaoFromCsv(csvText) {
  return QDD.CsvParser.extractKwhGiaoFromText(csvText);
}

/**
 * Tính nhiều ngày/nhiều tổ máy cùng lúc - xem QDD.BatchCalculator.
 * @param {import('./BatchCalculator').DayInput[]} dayInputs
 */
function calculateMultipleDays(dayInputs) {
  return QDD.BatchCalculator.calculateMultiple(dayInputs);
}

/**
 * Lọc lệnh tích luỹ nhiều ngày thành đầu vào cho từng (ngày, tổ máy).
 * @param {import('./CommandFilter').RawCommand[]} allCommands
 * @param {Date[]} dates
 * @param {string[]} units
 */
function buildDayUnitInputs(allCommands, dates, units) {
  return QDD.BatchCalculator.buildDayUnitInputs(allCommands, dates, units);
}

/**
 * Tổng hợp báo cáo tháng từ kết quả nhiều ngày - xem QDD.MonthlyReport.
 * @param {import('./BatchCalculator').DayResult[]} dayResults
 */
function aggregateMonthlyReport(dayResults) {
  return QDD.MonthlyReport.aggregate(dayResults);
}

/**
 * Phiên bản của thư viện (đối chiếu với CHANGELOG.md khi có thay đổi thuật toán).
 */
function getLibraryVersion() {
  return '0.2.0';
}
