/**
 * Tích phân hình thang theo 48 chu kỳ 30 phút, bám sát DIEN_TICH/TINH_TOAN
 * (docs/04_Algorithm_Specification.md mục 4): Qdd = tổng diện tích MW.s
 * của chu kỳ / 1800.
 */
var QDD = QDD || {};

QDD.AreaIntegration = (function () {
  /**
   * @param {import('./Segments').Segment[]} segments
   * @param {number} cycleStart  Giây bắt đầu chu kỳ
   * @param {number} cycleEnd    Giây kết thúc chu kỳ
   * @returns {number} Qdd (MW) của chu kỳ
   */
  function cycleQdd(segments, cycleStart, cycleEnd) {
    var total = 0;
    segments.forEach(function (seg) {
      var s = seg[0], e = seg[1], pS = seg[2], pE = seg[3];
      if (e <= cycleStart || s >= cycleEnd || e === s) return;
      var ovStart = Math.max(cycleStart, s);
      var ovEnd = Math.min(cycleEnd, e);
      if (ovEnd <= ovStart) return;
      function pAt(t) { return pS + (pE - pS) * (t - s) / (e - s); }
      total += (pAt(ovStart) + pAt(ovEnd)) / 2 * (ovEnd - ovStart);
    });
    return total / QDD.Config.CYCLE_SECONDS;
  }

  /**
   * @param {Array<{seconds:number,p:number}>} effectiveCommands
   * @param {number} p0
   * @returns {number[]} Qdd (MW) của 48 chu kỳ
   */
  /**
   * Tính 1 lần, trả về đủ mọi thứ cần cho 1 ngày - tránh dựng lại
   * ramp/segment nhiều lần.
   * @returns {{qdd:number[], endPower:number, carry:{target:number, remainingSeconds:number}|null}}
   */
  function computeDayFull(effectiveCommands, p0) {
    var rampRows = QDD.RampEngine.buildRows(effectiveCommands, p0);
    var segments = QDD.Segments.build(rampRows, p0);

    var qdd = [];
    for (var i = 0; i < QDD.Config.PERIOD_COUNT; i++) {
      var cs = i * QDD.Config.CYCLE_SECONDS;
      var ce = (i + 1) * QDD.Config.CYCLE_SECONDS;
      qdd.push(cycleQdd(segments, cs, ce));
    }

    return {
      qdd: qdd,
      endPower: QDD.Segments.endPowerOfDay(segments, p0),
      carry: QDD.Segments.carryOverOf(rampRows),
    };
  }

  function computeDay(effectiveCommands, p0) {
    return computeDayFull(effectiveCommands, p0).qdd;
  }

  /**
   * Công suất tại 24:00 của ngày - dùng làm P0 cho ngày kế tiếp.
   * @returns {number}
   */
  function endPowerOfDay(effectiveCommands, p0) {
    return computeDayFull(effectiveCommands, p0).endPower;
  }

  /**
   * R07 - thông tin ramp còn dở dang lúc 24:00, để ngày sau chạy tiếp.
   * @returns {{target:number, remainingSeconds:number}|null}
   */
  function carryOverOfDay(effectiveCommands, p0) {
    return computeDayFull(effectiveCommands, p0).carry;
  }

  return {
    cycleQdd: cycleQdd,
    computeDay: computeDay,
    computeDayFull: computeDayFull,
    endPowerOfDay: endPowerOfDay,
    carryOverOfDay: carryOverOfDay,
  };
})();

if (typeof module !== 'undefined' && module.exports) { module.exports = QDD; }
