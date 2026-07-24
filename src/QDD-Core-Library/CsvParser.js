/**
 * Đọc dòng KwhGiao (48 giá trị) từ CSV công tơ 6001/6303 - R10/R12
 * (docs/04_Algorithm_Specification.md mục 7). Cột A (ngày) chủ động KHÔNG
 * đọc/so sánh - đúng hành vi thật của v1.3.1, không phải giản lược.
 *
 * CHƯA cài đặt: logic phát hiện CSV bị Excel-Mac dồn vào 1 cột (như VBA
 * ReadCsvWithExcel50) - vì trên Apps Script không đi qua Excel để mở CSV,
 * input luôn là văn bản CSV thô hoặc mảng 2 chiều đã tách sẵn, nên vấn đề
 * gốc (khác biệt do Excel tự tách cột) không áp dụng ở đây.
 */
var QDD = QDD || {};

QDD.CsvParser = (function () {
  /**
   * @param {string} csvText  Nội dung file CSV thô
   * @returns {string[][]}
   */
  function parseCsvText(csvText) {
    return Utilities.parseCsv(csvText);
  }

  /**
   * @param {string[][]} rows  CSV đã tách thành mảng 2 chiều
   * @returns {number[]} 48 giá trị KwhGiao
   */
  function extractKwhGiao(rows) {
    var row = rows.filter(function (r) {
      return r[1] && String(r[1]).trim().toUpperCase() === 'KWHGIAO';
    })[0];
    if (!row) {
      throw new Error('Không tìm thấy dòng KwhGiao trong CSV.');
    }
    var values = row.slice(2, 2 + QDD.Config.PERIOD_COUNT).map(function (v, idx) {
      var n = parseFloat(v);
      if (isNaN(n)) {
        throw new Error('Giá trị KwhGiao không hợp lệ ở vị trí chu kỳ ' + (idx + 1) + ': "' + v + '"');
      }
      return n;
    });
    if (values.length !== QDD.Config.PERIOD_COUNT) {
      throw new Error('Dòng KwhGiao có ' + values.length + ' giá trị, cần đủ ' + QDD.Config.PERIOD_COUNT + '.');
    }
    return values;
  }

  /**
   * @param {string} csvText
   * @returns {number[]} 48 giá trị KwhGiao
   */
  function extractKwhGiaoFromText(csvText) {
    return extractKwhGiao(parseCsvText(csvText));
  }

  return {
    parseCsvText: parseCsvText,
    extractKwhGiao: extractKwhGiao,
    extractKwhGiaoFromText: extractKwhGiaoFromText,
  };
})();

if (typeof module !== 'undefined' && module.exports) { module.exports = QDD; }
