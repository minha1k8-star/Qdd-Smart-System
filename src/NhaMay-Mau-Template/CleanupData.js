/**
 * Dọn dữ liệu cũ - dùng khi bàn giao Sheet cho người khác dùng thử, hoặc
 * khi dữ liệu tích luỹ nhiều tháng làm sheet nặng và khó nhìn.
 *
 * LUÔN giữ nguyên: CAI_DAT (cấu hình nhà máy).
 * Giữ có chọn lọc: KET_QUA của N ngày gần nhất - để P0 của ngày kế tiếp
 * vẫn tự suy ra được, không phải nhập tay lại (xem readOrInferP0_).
 * Xoá sạch: LENH, LENH_STAGING, CSV_DATA, BAO_CAO_THANG, P0_NGAY và phần
 * KET_QUA cũ hơn.
 */

/** Chuẩn hoá 1 ô ngày về chuỗi 'yyyy-MM-dd' để so sánh; trả '' nếu không phải ngày. */
function dateKeyOf_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return '';
}

/**
 * @param {number} keepDays  Số ngày gần nhất muốn giữ lại kết quả (0 = xoá sạch cả KET_QUA)
 * @returns {{keptDates:string[], removed:Object}}
 */
function cleanupOldData_(keepDays) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var removed = {};

  // --- Xác định các ngày cần giữ trong KET_QUA ---
  var ketQuaSh = ss.getSheetByName(SHEETS.KET_QUA);
  var keepDateKeys = [];
  if (keepDays > 0 && ketQuaSh && ketQuaSh.getLastRow() >= 2) {
    var dateCol = ketQuaSh.getRange(2, 1, ketQuaSh.getLastRow() - 1, 1).getValues();
    var uniqueKeys = {};
    dateCol.forEach(function (r) {
      var k = dateKeyOf_(r[0]);
      if (k) uniqueKeys[k] = true;
    });
    keepDateKeys = Object.keys(uniqueKeys).sort().slice(-keepDays); // 'yyyy-MM-dd' sắp xếp chuỗi = sắp theo thời gian
  }

  // --- KET_QUA: chỉ giữ các ngày trong keepDateKeys ---
  if (ketQuaSh) {
    removed.ketQua = rewriteSheetRows_(ketQuaSh, function (row) {
      return keepDateKeys.indexOf(dateKeyOf_(row[0])) !== -1;
    });
  }

  // --- P0_NGAY: giữ đúng các ngày còn lại trong KET_QUA (nếu có), xoá phần còn lại ---
  var p0Sh = ss.getSheetByName(SHEETS.P0_NGAY);
  if (p0Sh) {
    removed.p0 = rewriteSheetRows_(p0Sh, function (row) {
      return keepDateKeys.indexOf(dateKeyOf_(row[0])) !== -1;
    });
  }

  // --- Các sheet dữ liệu nguồn/kết xuất: xoá sạch dữ liệu, giữ dòng tiêu đề ---
  [SHEETS.LENH, SHEETS.LENH_STAGING, SHEETS.CSV_DATA, SHEETS.BAO_CAO_THANG].forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh) return;
    var lastRow = sh.getLastRow();
    var lastCol = sh.getLastColumn();
    if (lastRow >= 2 && lastCol > 0) {
      removed[name] = lastRow - 1;
      sh.getRange(2, 1, lastRow - 1, lastCol).clearContent();
    } else {
      removed[name] = 0;
    }
  });

  return { keptDates: keepDateKeys, removed: removed };
}

/**
 * Hàm gọi từ sidebar. Việc xác nhận trước khi xoá do phía sidebar đảm nhận.
 * @param {number} keepDays
 * @returns {string}
 */
function sidebar_cleanupOldData(keepDays) {
  var n = Number(keepDays);
  if (isNaN(n) || n < 0) throw new Error('Số ngày giữ lại không hợp lệ.');

  var result = cleanupOldData_(n);
  var r = result.removed;

  var msg = 'Đã dọn xong. Đã xoá: ' +
    (r.ketQua || 0) + ' dòng KET_QUA, ' +
    (r[SHEETS.LENH] || 0) + ' dòng LENH, ' +
    (r[SHEETS.CSV_DATA] || 0) + ' dòng CSV_DATA, ' +
    (r[SHEETS.BAO_CAO_THANG] || 0) + ' dòng BAO_CAO_THANG, ' +
    (r.p0 || 0) + ' dòng P0_NGAY.';

  if (result.keptDates.length > 0) {
    var pretty = result.keptDates.map(function (k) {
      var p = k.split('-');
      return p[2] + '/' + p[1] + '/' + p[0];
    });
    msg += '\nGiữ lại kết quả ngày: ' + pretty.join(', ') + ' (để P0 ngày kế tiếp tự suy ra được).';
  } else {
    msg += '\nĐã xoá sạch KET_QUA - lần tính tiếp theo phải nhập lại P0 vào sheet P0_NGAY.';
  }
  msg += '\nCấu hình trong CAI_DAT được giữ nguyên.';
  return msg;
}
