/**
 * Tổng hợp kết quả 48 chu kỳ (R08-R14), bám sát TINH_TOAN
 * (docs/04_Algorithm_Specification.md mục 5).
 *
 * @typedef {Object} PeriodResult
 * @property {number} chuKy       1..48
 * @property {number} qdd         MW
 * @property {number} qddV        MWh
 * @property {number} qdc         MWh (từ CSV 6001)
 * @property {number} pQdc        MW
 * @property {number} nguongDuoi  MW
 * @property {number} nguongTren  MW
 * @property {number} qmp         MWh (từ CSV 6303)
 * @property {number} qdu         MWh
 * @property {string} dauHieu     "âm" | "dương" | "trong ±3%"
 */
var QDD = QDD || {};

QDD.QddCalculator = (function () {
  /**
   * @param {Object} input
   * @param {Array<{seconds:number,p:number}>} input.effectiveCommands  Kết quả QDD.CommandFilter.selectEffective
   * @param {number} input.p0        Công suất đầu ngày
   * @param {number[]} input.qdc48   48 giá trị KwhGiao CSV 6001 (chưa chia 1000)
   * @param {number[]} input.qmp48   48 giá trị KwhGiao CSV 6303 (chưa chia 1000)
   * @param {number} [input.qddVCoef]   mặc định QDD.Config.QDD_V_COEF
   * @param {number} [input.tolerance]  mặc định QDD.Config.TOLERANCE
   * @returns {PeriodResult[]}
   */
  function calculateDay(input) {
    var qddVCoef = input.qddVCoef || QDD.Config.QDD_V_COEF;
    var tolerance = (input.tolerance === undefined) ? QDD.Config.TOLERANCE : input.tolerance;

    if (input.qdc48.length !== QDD.Config.PERIOD_COUNT || input.qmp48.length !== QDD.Config.PERIOD_COUNT) {
      throw new Error('qdc48/qmp48 phải có đúng ' + QDD.Config.PERIOD_COUNT + ' giá trị.');
    }

    var qddArr = QDD.AreaIntegration.computeDay(input.effectiveCommands, input.p0);
    var endPower = QDD.AreaIntegration.endPowerOfDay(input.effectiveCommands, input.p0);
    var results = [];

    for (var i = 0; i < QDD.Config.PERIOD_COUNT; i++) {
      var qdd = qddArr[i];
      var qddV = qdd / 2 * qddVCoef;
      var qdc = input.qdc48[i] / 1000;
      var pQdc = qdc * 2;
      var qmp = input.qmp48[i] / 1000;
      var nguongDuoi = qdd * (1 - tolerance);
      var nguongTren = qdd * (1 + tolerance);

      var qdu, dauHieu;
      if (pQdc < nguongDuoi) {
        qdu = qmp - qddV;
        dauHieu = 'âm';
      } else if (pQdc > nguongTren) {
        qdu = qmp - qddV;
        dauHieu = 'dương';
      } else {
        qdu = 0;
        dauHieu = 'trong ±3%';
      }

      results.push({
        chuKy: i + 1,
        qdd: qdd,
        qddV: qddV,
        qdc: qdc,
        pQdc: pQdc,
        nguongDuoi: nguongDuoi,
        nguongTren: nguongTren,
        qmp: qmp,
        qdu: qdu,
        dauHieu: dauHieu,
      });
    }

    // Gắn công suất cuối ngày vào kết quả để tầng gọi lưu làm P0 ngày sau.
    // Dùng thuộc tính trên mảng thay vì đổi kiểu trả về, để không phá vỡ
    // các chỗ đang duyệt kết quả như một mảng 48 phần tử.
    results.endPower = endPower;
    return results;
  }

  return { calculateDay: calculateDay };
})();

if (typeof module !== 'undefined' && module.exports) { module.exports = QDD; }
