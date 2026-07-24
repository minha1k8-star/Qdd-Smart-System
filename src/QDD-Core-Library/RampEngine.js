/**
 * Ramp Engine - bám sát công thức XU_LY_LENH!B..K thật (đã trích xuất từ
 * workbook gốc, xem docs/04_Algorithm_Specification.md mục 2). Cột F (P bắt
 * đầu) của lệnh sau là nội suy tuyến tính trên đường ramp của lệnh trước
 * (R06 - ngắt ramp) nếu lệnh mới đến giữa ramp đang chạy, hoặc bằng mục
 * tiêu của lệnh trước nếu ramp trước đã hoàn tất.
 *
 * @typedef {Object} RampRow
 * @property {number} B  Bắt đầu (giây)
 * @property {number} D  P mục tiêu
 * @property {number} F  P bắt đầu
 * @property {number} H  Thời lượng (giây)
 * @property {number} I  Kết thúc (giây)
 */
var QDD = QDD || {};

QDD.RampEngine = (function () {
  /**
   * @param {Array<{seconds:number,p:number}>} effectiveCommands  Đã sắp theo thời gian
   * @param {number} p0  Công suất đầu ngày (CAI_DAT!B6)
   * @param {number} [rampRateMwPerMin]  CAI_DAT!B7, mặc định QDD.Config.RAMP_RATE_MW_PER_MIN
   * @returns {RampRow[]}
   */
  function buildRows(effectiveCommands, p0, rampRateMwPerMin) {
    var rate = rampRateMwPerMin || QDD.Config.RAMP_RATE_MW_PER_MIN;
    var eps = QDD.Config.EPS;
    var rows = [];
    var prevF = p0, prevB = null, prevI = null, prevD = null;

    effectiveCommands.forEach(function (cmd, i) {
      var B = cmd.seconds;
      var D = cmd.p;
      var F;
      if (i === 0) {
        F = p0;
      } else if (B >= prevI) {
        F = prevD;
      } else {
        var denom = Math.max(prevI - prevB, eps);
        F = prevF + (prevD - prevF) * (B - prevB) / denom;
      }
      var H = Math.abs(D - F) < eps ? 0 : Math.abs(D - F) / rate * 60;
      var I = B + H;
      rows.push({ B: B, D: D, F: F, H: H, I: I });
      prevF = F; prevB = B; prevI = I; prevD = D;
    });

    return rows;
  }

  return { buildRows: buildRows };
})();

if (typeof module !== 'undefined' && module.exports) { module.exports = QDD; }
