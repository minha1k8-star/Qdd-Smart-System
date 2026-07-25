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
        // Nếu ramp bị CẮT trước khi đạt mục tiêu (do lệnh mới đến, hoặc do
        // hết ngày), công suất cuối đoạn phải là giá trị NỘI SUY tại điểm
        // cắt - không phải mục tiêu D. Ghi nhầm D sẽ làm sai độ dốc của
        // đoạn, kéo theo sai diện tích (Qdd) và sai cả công suất cuối ngày.
        // Bám đúng công thức gốc DOAN_CONG_SUAT!F (xem
        // docs/04_Algorithm_Specification.md mục 3).
        var pEnd = (rampEnd >= rr.I)
          ? rr.D
          : rr.F + (rr.D - rr.F) * (rampEnd - rr.B) / Math.max(rr.I - rr.B, QDD.Config.EPS);
        segs.push([rr.B, rampEnd, rr.F, pEnd]);
      }
      var holdEnd = Math.min(nextB, DAY);
      if (holdEnd > rampEnd) {
        segs.push([rampEnd, holdEnd, rr.D, rr.D]);
      }
    });

    return segs;
  }

  /**
   * Công suất tại đúng thời điểm 24:00 (cuối ngày) - dùng làm P0 của ngày
   * kế tiếp.
   *
   * KHÔNG được lấy Qdd của chu kỳ 48 thay cho giá trị này: Qdd chu kỳ 48 là
   * công suất TRUNG BÌNH trong khoảng 23:30-24:00, khác hẳn công suất tại
   * thời điểm 24:00 khi tổ máy đang tăng/giảm tải. Nhầm hai đại lượng này
   * từng gây sai số 29,4 MW kéo dài cả ngày hôm sau (dữ liệu thật 19/07).
   *
   * @param {Segment[]} segments
   * @param {number} p0  dùng khi không có đoạn nào (ngày trống)
   * @returns {number}
   */
  function endPowerOfDay(segments, p0) {
    if (!segments || segments.length === 0) return p0;
    return segments[segments.length - 1][3];
  }

  /**
   * R07 - phát hiện ramp CHƯA hoàn tất lúc 24:00 (vắt qua nửa đêm).
   *
   * Tổ máy không dừng lại ở nửa đêm: nó tiếp tục tăng/giảm tải sang ngày
   * hôm sau cho tới khi đạt mục tiêu. Ngày kế tiếp phải chạy tiếp phần
   * còn lại, nếu không sẽ giữ nguyên công suất lúc 24:00 suốt đến khi có
   * lệnh mới - sai so với thực tế.
   *
   * @param {import('./RampEngine').RampRow[]} rampRows
   * @returns {{target:number, remainingSeconds:number}|null} null nếu ramp đã xong trong ngày
   */
  function carryOverOf(rampRows) {
    if (!rampRows || rampRows.length === 0) return null;
    var last = rampRows[rampRows.length - 1];
    var DAY = QDD.Config.SECONDS_PER_DAY;
    if (last.I <= DAY + QDD.Config.EPS) return null;
    return { target: last.D, remainingSeconds: last.I - DAY };
  }

  return { build: build, endPowerOfDay: endPowerOfDay, carryOverOf: carryOverOf };
})();

if (typeof module !== 'undefined' && module.exports) { module.exports = QDD; }
