/**
 * Tính nhiều ngày / nhiều tổ máy trong cùng một lần gọi - khác với bản
 * Excel/VBA (chỉ tính được 1 ngày tại một thời điểm vì CAI_DAT!B4 chỉ giữ
 * 1 ngày). Ở đây mỗi ngày+tổ máy độc lập hoàn toàn, không có "ô cấu hình
 * toàn cục" nào bị ghi đè - xem docs/05_System_Architecture.md.
 *
 * CHƯA làm: tự động nối P0 giữa các ngày theo đúng cơ chế carry-over R07
 * (ramp qua 00:00, xem docs/04_Algorithm_Specification.md mục 6). Hàm này
 * đòi hỏi P0 được truyền vào tường minh cho mỗi ngày - việc suy ra P0 từ
 * ngày trước là quyết định của tầng gọi (script Sheet), không phải của
 * thư viện, để giữ thư viện thuần và dễ kiểm chứng.
 */
var QDD = QDD || {};

QDD.BatchCalculator = (function () {
  /**
   * @typedef {Object} DayInput
   * @property {string|Date} date   Ngày tính (dùng làm khoá kết quả)
   * @property {string} unit        "S1" | "S2"
   * @property {Array<{seconds:number,p:number}>} effectiveCommands
   * @property {number} p0
   * @property {number[]} qdc48
   * @property {number[]} qmp48
   * @property {number} [qddVCoef]
   * @property {number} [tolerance]
   *
   * @typedef {Object} DayResult
   * @property {string|Date} date
   * @property {string} unit
   * @property {import('./QddCalculator').PeriodResult[]} periods
   * @property {Error} [error]  Có mặt nếu ngày này tính lỗi - các ngày khác vẫn tính tiếp
   */

  /**
   * @param {DayInput[]} dayInputs
   * @returns {DayResult[]}  Cùng thứ tự với dayInputs; ngày lỗi vẫn có mặt
   *                         trong kết quả (kèm trường `error`) thay vì làm
   *                         dừng toàn bộ lô - để 1 ngày dữ liệu xấu không
   *                         chặn các ngày còn lại.
   */
  function calculateMultiple(dayInputs) {
    return dayInputs.map(function (input) {
      try {
        var periods = QDD.QddCalculator.calculateDay({
          effectiveCommands: input.effectiveCommands,
          p0: input.p0,
          qdc48: input.qdc48,
          qmp48: input.qmp48,
          qddVCoef: input.qddVCoef,
          tolerance: input.tolerance,
        });
        return { date: input.date, unit: input.unit, periods: periods };
      } catch (e) {
        return { date: input.date, unit: input.unit, periods: null, error: e };
      }
    });
  }

  /**
   * Lọc danh sách lệnh tích luỹ (nhiều ngày) thành đầu vào cho từng
   * (ngày, tổ máy) trong một khoảng thời gian - dùng khi nhập hàng loạt.
   *
   * @param {import('./CommandFilter').RawCommand[]} allCommands  Toàn bộ lệnh đã tích luỹ (nhiều ngày)
   * @param {Date[]} dates
   * @param {string[]} units  vd ["S1","S2"]
   * @returns {Array<{date:Date, unit:string, effectiveCommands: Array}>}
   *          (chưa có p0/qdc48/qmp48 - tầng gọi cần bổ sung trước khi đưa vào calculateMultiple)
   */
  function buildDayUnitInputs(allCommands, dates, units) {
    var out = [];
    dates.forEach(function (date) {
      units.forEach(function (unit) {
        out.push({
          date: date,
          unit: unit,
          effectiveCommands: QDD.CommandFilter.selectEffective(allCommands, date, unit),
        });
      });
    });
    return out;
  }

  return {
    calculateMultiple: calculateMultiple,
    buildDayUnitInputs: buildDayUnitInputs,
  };
})();

if (typeof module !== 'undefined' && module.exports) { module.exports = QDD; }
