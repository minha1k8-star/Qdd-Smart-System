/**
 * Dựng danh sách đoạn công suất phủ kín 24h từ các dòng Ramp Engine, bám
 * sát DOAN_CONG_SUAT (docs/04_Algorithm_Specification.md mục 3): mỗi lệnh
 * sinh 1 đoạn RAMP + 1 đoạn HOLD, cộng thêm 1 đoạn HOLD đầu ngày ở P0.
 *
 * @typedef {[number, number, number, number]} Segment  [start, end, pStart, pEnd] (giây, MW)
 */
var QDD = QDD || {};

QDD.Segments = (function () {
  /**
   * @param {import('./RampEngine').RampRow[]} rampRows
   * @param {number} p0
   * @returns {Segment[]}
   */
  function build(rampRows, p0) {
    var DAY = QDD.Config.SECONDS_PER_DAY;
    var segs = [];
    var firstB = rampRows.length ? rampRows[0].B : DAY;
    segs.push([0, Math.min(firstB, DAY), p0, p0]);

    rampRows.forEach(function (rr, i) {
      var nextB = (i + 1 < rampRows.length) ? rampRows[i + 1].B : DAY;
      var rampEnd = Math.min(rr.I, nextB, DAY);
      if (rampEnd > rr.B) {
        segs.push([rr.B, rampEnd, rr.F, rr.D]);
      }
      var holdEnd = Math.min(nextB, DAY);
      if (holdEnd > rampEnd) {
        segs.push([rampEnd, holdEnd, rr.D, rr.D]);
      }
    });

    return segs;
  }

  return { build: build };
})();

if (typeof module !== 'undefined' && module.exports) { module.exports = QDD; }
