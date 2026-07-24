/**
 * Tổng hợp báo cáo tháng trực tiếp từ kết quả nhiều ngày (QDD.BatchCalculator),
 * KHÔNG cần cơ chế "snapshot đóng băng công thức" như VBA
 * (QduMonth_SaveCurrentDay/LICH_SU_THANG/LS_..., xem
 * docs/04_Algorithm_Specification.md mục 8) - vì mỗi ngày+tổ máy đã độc
 * lập hoàn toàn ngay từ khi tính (QDD.BatchCalculator), không có sheet
 * dùng chung nào bị ghi đè giữa các ngày cần phải "chốt" trước khi mất.
 *
 * Tương đương sheet TONG_HOP trong báo cáo tháng VBA.
 */
var QDD = QDD || {};

QDD.MonthlyReport = (function () {
  function sum(periods, field) {
    return periods.reduce(function (acc, p) { return acc + p[field]; }, 0);
  }

  /**
   * @param {import('./BatchCalculator').DayResult[]} dayResults
   * @returns {Object}
   *   perDay: [{date, unit, tongQdc, tongQmp, tongQddV, tongQdu, error}]
   *   perUnit: { [unit]: {tongQdc, tongQmp, tongQddV, tongQdu, soNgay} }
   *   tongHop: {tongQdc, tongQmp, tongQddV, tongQdu, soNgay, soNgayLoi}
   */
  function aggregate(dayResults) {
    var perDay = [];
    var perUnit = {};
    var grand = { tongQdc: 0, tongQmp: 0, tongQddV: 0, tongQdu: 0, soNgay: 0, soNgayLoi: 0 };

    dayResults.forEach(function (r) {
      if (r.error || !r.periods) {
        perDay.push({ date: r.date, unit: r.unit, error: r.error ? String(r.error) : 'Không có dữ liệu' });
        grand.soNgayLoi++;
        return;
      }
      var row = {
        date: r.date,
        unit: r.unit,
        tongQdc: sum(r.periods, 'qdc'),
        tongQmp: sum(r.periods, 'qmp'),
        tongQddV: sum(r.periods, 'qddV'),
        tongQdu: sum(r.periods, 'qdu'),
      };
      perDay.push(row);

      if (!perUnit[r.unit]) {
        perUnit[r.unit] = { tongQdc: 0, tongQmp: 0, tongQddV: 0, tongQdu: 0, soNgay: 0 };
      }
      perUnit[r.unit].tongQdc += row.tongQdc;
      perUnit[r.unit].tongQmp += row.tongQmp;
      perUnit[r.unit].tongQddV += row.tongQddV;
      perUnit[r.unit].tongQdu += row.tongQdu;
      perUnit[r.unit].soNgay += 1;

      grand.tongQdc += row.tongQdc;
      grand.tongQmp += row.tongQmp;
      grand.tongQddV += row.tongQddV;
      grand.tongQdu += row.tongQdu;
      grand.soNgay += 1;
    });

    return { perDay: perDay, perUnit: perUnit, tongHop: grand };
  }

  return { aggregate: aggregate };
})();

if (typeof module !== 'undefined' && module.exports) { module.exports = QDD; }
