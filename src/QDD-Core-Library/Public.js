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
 * Lệnh này có thuộc tổ máy đang tính không (khớp cả khi cột "Tổ máy"
 * trong file lệnh ghi dài hơn, vd cấu hình "S1" - file ghi "S1DH1").
 * Tầng Sheet dùng để lọc cảnh báo, phải đi qua đây để dùng CHUNG một
 * quy tắc với CommandFilter, không tự lặp lại logic riêng.
 * @param {string} toMay
 * @param {string} unit
 */
function matchesUnit(toMay, unit) {
  return QDD.CommandFilter.matchesUnit(toMay, unit);
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
 * Công suất tại 24:00 của ngày - dùng làm P0 cho ngày kế tiếp.
 * KHÔNG dùng Qdd chu kỳ 48 thay cho giá trị này (xem Segments.endPowerOfDay).
 * @param {Array<{seconds:number,p:number}>} effectiveCommands
 * @param {number} p0
 */
function getEndOfDayPower(effectiveCommands, p0) {
  return QDD.AreaIntegration.endPowerOfDay(effectiveCommands, p0);
}

/**
 * R07 - thông tin ramp còn dở dang lúc 24:00 để ngày sau chạy tiếp.
 * @returns {{target:number, remainingSeconds:number}|null}
 */
function getCarryOver(effectiveCommands, p0) {
  return QDD.AreaIntegration.carryOverOfDay(effectiveCommands, p0);
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
