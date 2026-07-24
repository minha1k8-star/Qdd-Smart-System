/**
 * R01-R03: chọn lệnh hiệu lực + công suất hiệu lực từ danh sách lệnh gốc.
 * Bám sát công thức thật LENH_DIEU_DO!T3/U3 (docs/04_Algorithm_Specification.md
 * mục 1). Điều kiện >0 cho cả SO và MO là CHỦ ĐÍCH, đã xác nhận nghiệp vụ
 * (xem docs/09_Test_Cases.md UAT-32) - KHÔNG nới thành >=0.
 *
 * @typedef {Object} RawCommand
 * @property {string} id
 * @property {string} toMay        "S1" | "S2"
 * @property {number} csRaLenh     CS ra lệnh (MW), có thể null/undefined
 * @property {number} csHoanThanh  CS hoàn thành (MW), có thể null/undefined
 * @property {Date}   bdth         Thời điểm BĐTH
 * @property {number|boolean} hoanThanh  Hoàn thành (1/true = đã hoàn thành)
 * @property {boolean} dungLenh    Dừng lệnh sớm
 * @property {string} nguonLenh    "SO" | "MO"
 *
 * @typedef {Object} EffectiveCommand
 * @property {string} id
 * @property {Date}   bdth
 * @property {number} seconds  Giây trong ngày của BĐTH
 * @property {number} p        Công suất hiệu lực (MW)
 * @property {string} nguon
 */
var QDD = QDD || {};

QDD.CommandFilter = (function () {
  function isCompleted(hoanThanh) {
    return hoanThanh === 1 || hoanThanh === true || hoanThanh === '1' || hoanThanh === 'TRUE';
  }

  function isStopped(dungLenh) {
    return dungLenh === true || dungLenh === 'TRUE' || dungLenh === 1;
  }

  function secondsOfDay(date) {
    return date.getHours() * 3600 + date.getMinutes() * 60 + date.getSeconds();
  }

  function sameDate(a, b) {
    return a.getFullYear() === b.getFullYear() &&
      a.getMonth() === b.getMonth() &&
      a.getDate() === b.getDate();
  }

  /**
   * @param {RawCommand[]} commands  Toàn bộ lệnh (nhiều ngày/nhiều tổ máy)
   * @param {Date} targetDate        Ngày đang tính
   * @param {string} unit            "S1" | "S2"
   * @returns {EffectiveCommand[]}   Đã lọc + sắp theo thời gian tăng dần
   */
  function selectEffective(commands, targetDate, unit) {
    var unitPrefix = unit.toUpperCase().slice(0, 2);
    var out = [];

    commands.forEach(function (c) {
      if (!c.bdth || !(c.bdth instanceof Date)) return;
      if (!sameDate(c.bdth, targetDate)) return;
      if ((c.toMay || '').toUpperCase().slice(0, 2) !== unitPrefix) return;
      if (!isCompleted(c.hoanThanh)) return;

      var nguon = (c.nguonLenh || '').toUpperCase();
      var valid = false;
      if (nguon === 'SO' && typeof c.csHoanThanh === 'number' && c.csHoanThanh > 0) {
        valid = true;
      }
      if (nguon === 'MO' && typeof c.csRaLenh === 'number' && c.csRaLenh > 0) {
        valid = true;
      }
      if (!valid) return;

      var pHieuLuc;
      if (nguon === 'SO') {
        pHieuLuc = c.csHoanThanh;
      } else if (nguon === 'MO' && isStopped(c.dungLenh) &&
        typeof c.csHoanThanh === 'number' && c.csHoanThanh > 0) {
        pHieuLuc = c.csHoanThanh; // R03: MO dừng sớm
      } else {
        pHieuLuc = c.csRaLenh; // R02: MO bình thường
      }

      out.push({
        id: c.id,
        bdth: c.bdth,
        seconds: secondsOfDay(c.bdth),
        p: pHieuLuc,
        nguon: nguon,
      });
    });

    out.sort(function (a, b) { return a.seconds - b.seconds; });
    return out;
  }

  return { selectEffective: selectEffective };
})();

if (typeof module !== 'undefined' && module.exports) { module.exports = QDD; }
