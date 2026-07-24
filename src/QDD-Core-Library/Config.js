/**
 * Hằng số nghiệp vụ - đối chiếu docs/04_Algorithm_Specification.md mục 9.
 * Giá trị mặc định khớp CAI_DAT của bản Excel/VBA v1.3.1. Khi dùng thật,
 * nên đọc các giá trị này từ Sheet cấu hình của từng nhà máy thay vì hằng
 * số cứng, vì tốc độ ramp/hệ số/dung sai có thể khác nhau giữa các nhà máy
 * (xem docs/05_System_Architecture.md).
 */
var QDD = QDD || {};

QDD.Config = {
  RAMP_RATE_MW_PER_MIN: 3.5,   // CAI_DAT!B7
  QDD_V_COEF: 0.9188,          // CAI_DAT!B8
  TOLERANCE: 0.03,             // CAI_DAT!B9 (+-3%)
  PERIOD_COUNT: 48,
  SECONDS_PER_DAY: 86400,
  CYCLE_SECONDS: 1800,
  EPS: 0.000001,
  MAX_COMMANDS: 199,
  MAX_EFFECTIVE_COMMANDS: 60,
};

if (typeof module !== 'undefined' && module.exports) { module.exports = QDD; }
