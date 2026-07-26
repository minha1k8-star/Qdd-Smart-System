/**
 * Gộp danh sách lệnh (đọc từ file Excel người dùng tải lên) vào sheet LENH
 * - nhận diện cột theo TÊN TIÊU ĐỀ (không theo vị trí cột), để nhận NGUYÊN
 * bảng của file gốc (25 cột, bất kỳ thứ tự nào miễn còn dòng tiêu đề) mà
 * không cần cắt/sắp xếp lại đúng 9 cột của LENH - tránh lỗi lệch cột đã
 * gặp trong thực tế (xem docs/14_Knowledge_Transfer.md).
 */

var COMMAND_REQUIRED_FIELD_LABELS = {
  id: 'ID Lệnh',
  toMay: 'Tổ máy',
  noiDungLenh: 'Nội dung lệnh',
  csRaLenh: 'CS ra lệnh',
  csHoanThanh: 'CS hoàn thành',
  bdth: 'Thời điểm BĐTH',
  hoanThanh: 'Hoàn thành',
  dungLenh: 'Dừng lệnh',
  nguonLenh: 'Nguồn lệnh',
};

/** Cột không bắt buộc - nếu có thì lấy, không có thì để trống, không báo lỗi. */
var COMMAND_OPTIONAL_FIELD_LABELS = {
  nhaMay: 'Nhà máy',
};

/** Bỏ hậu tố dạng " (MW)"/" (1/0)"/" (TRUE/FALSE)" ở cuối tiêu đề trước khi so khớp, để chấp nhận cả tiêu đề gốc lẫn tiêu đề của LENH (có/không có hậu tố). */
function normalizeHeader_(h) {
  return String(h || '').replace(/\s*\([^)]*\)\s*$/, '').trim();
}

/**
 * Đọc ô BĐTH thành Date. Nhận cả ô đã là ngày-giờ thật (trường hợp thường
 * gặp sau khi Google chuyển đổi file Excel) lẫn ô dạng CHỮ "dd/MM/yyyy HH:mm"
 * - một số file gốc lưu cột này dưới dạng text nên nếu chỉ kiểm tra kiểu Date
 * thì toàn bộ lệnh bị loại mà người dùng không hiểu vì sao.
 *
 * Không dùng `instanceof Date` (xem AGENTS.md) mà kiểm tra theo đặc điểm.
 *
 * @returns {Date|null} null nếu không đọc được
 */
function coerceBdth_(value) {
  if (value && typeof value.getTime === 'function' && !isNaN(value.getTime())) {
    return value;
  }
  var s = String(value || '').trim();
  if (!s) return null;

  // dd/MM/yyyy hoặc dd-MM-yyyy, kèm giờ tuỳ chọn HH:mm[:ss]
  var m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})(?:[\sT]+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (!m) return null;
  var d = new Date(
    Number(m[3]), Number(m[2]) - 1, Number(m[1]),
    Number(m[4] || 0), Number(m[5] || 0), Number(m[6] || 0)
  );
  return isNaN(d.getTime()) ? null : d;
}

/**
 * @param {Array} headerRow
 * @returns {{indices: Object, missing: string[]}}
 */
function findCommandColumnIndices_(headerRow) {
  var normalized = headerRow.map(normalizeHeader_);
  var indices = {};
  var missing = [];
  Object.keys(COMMAND_REQUIRED_FIELD_LABELS).forEach(function (field) {
    var label = COMMAND_REQUIRED_FIELD_LABELS[field];
    var idx = normalized.indexOf(label);
    if (idx === -1) {
      missing.push(label);
    } else {
      indices[field] = idx;
    }
  });
  Object.keys(COMMAND_OPTIONAL_FIELD_LABELS).forEach(function (field) {
    var idx = normalized.indexOf(COMMAND_OPTIONAL_FIELD_LABELS[field]);
    if (idx !== -1) indices[field] = idx;
  });
  return { indices: indices, missing: missing };
}

/**
 * Gộp một bảng lệnh đã đọc được (dòng đầu là tiêu đề) vào sheet LENH.
 *
 * @param {Array[]} table  dòng 0 là tiêu đề, các dòng sau là dữ liệu
 * @param {number} [firstDataRowNumber]  số dòng thật của dòng dữ liệu đầu
 *   tiên trong file nguồn, chỉ dùng để báo lỗi cho đúng dòng người dùng thấy
 * @returns {{imported:number, updated:number, skipped:Array<{row:number, reason:string}>}}
 */
function importCommandTable_(table, firstDataRowNumber) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!table || table.length < 2) {
    throw new Error('Bảng lệnh trống (chỉ có dòng tiêu đề, không có dòng dữ liệu nào).');
  }
  var baseRowNumber = firstDataRowNumber || 2;

  var headerRow = table[0];
  var dataRows = table.slice(1);

  var found = findCommandColumnIndices_(headerRow);
  if (found.missing.length > 0) {
    throw new Error('Không tìm thấy cột: ' + found.missing.join(', ') +
      '. Kiểm tra lại dòng tiêu đề trong file nguồn có đúng tên như file gốc không.');
  }
  var idx = found.indices;

  // Sheet đích: dò cột theo TÊN để ghi đúng vị trí, bất kể LENH đang dùng
  // cấu trúc 25 cột giống file gốc hay bản rút gọn cũ.
  var lenhSh = ss.getSheetByName(SHEETS.LENH);
  var lenhLastCol = lenhSh.getLastColumn();
  if (lenhLastCol === 0) {
    lenhSh.getRange(1, 1, 1, LENH_HEADERS.length).setValues([LENH_HEADERS]);
    lenhSh.setFrozenRows(1);
    lenhLastCol = LENH_HEADERS.length;
  }
  var lenhHeaders = lenhSh.getRange(1, 1, 1, lenhLastCol).getValues()[0];
  var lenhNormalized = lenhHeaders.map(normalizeHeader_);
  var lenhFound = findCommandColumnIndices_(lenhHeaders);
  if (lenhFound.missing.length > 0) {
    throw new Error('Sheet LENH thiếu cột: ' + lenhFound.missing.join(', ') +
      '. Chạy "QDD Smart System → Thiết lập sheet" để tạo lại tiêu đề chuẩn.');
  }
  var lenhIdx = lenhFound.indices;

  // Bản đồ tên cột (đã chuẩn hoá) của bảng nguồn, để copy cả các cột phụ
  // (Người ra lệnh, Lý do lệnh...) khi LENH có sẵn cột cùng tên.
  var sourceByName = {};
  headerRow.forEach(function (h, i) {
    var key = normalizeHeader_(h);
    if (key && sourceByName[key] === undefined) sourceByName[key] = i;
  });

  var skipped = [];
  var parsed = [];
  dataRows.forEach(function (r, i) {
    var rowNum = i + baseRowNumber; // số dòng thật trong file nguồn
    var id = r[idx.id];
    if (id === '' || id === null) return; // dòng trống, bỏ qua âm thầm

    var bdth = coerceBdth_(r[idx.bdth]);
    if (!bdth) {
      skipped.push({ row: rowNum, reason: 'Thời điểm BĐTH không đọc được thành ngày-giờ (giá trị: "' + r[idx.bdth] + '"). Trong file gốc, ô này phải là ngày-giờ hoặc dạng chữ dd/MM/yyyy HH:mm.' });
      return;
    }

    // Dựng dòng theo đúng số cột của LENH: cột nào trùng tên với staging thì lấy nguyên giá trị.
    var row = new Array(lenhLastCol).fill('');
    lenhNormalized.forEach(function (name, col) {
      if (name && sourceByName[name] !== undefined) {
        row[col] = r[sourceByName[name]];
      }
    });

    // Chuẩn hoá riêng các trường thuật toán cần đúng kiểu dữ liệu.
    var csRaLenh = r[idx.csRaLenh];
    var csHoanThanh = r[idx.csHoanThanh];
    if (csRaLenh !== '' && typeof csRaLenh !== 'number') csRaLenh = Number(csRaLenh);
    if (csHoanThanh !== '' && typeof csHoanThanh !== 'number') csHoanThanh = Number(csHoanThanh);
    var hoanThanhRaw = r[idx.hoanThanh];
    var dungLenhRaw = r[idx.dungLenh];

    row[lenhIdx.id] = id;
    row[lenhIdx.toMay] = r[idx.toMay];
    row[lenhIdx.noiDungLenh] = r[idx.noiDungLenh];
    row[lenhIdx.csRaLenh] = (csRaLenh === '' || isNaN(csRaLenh)) ? '' : csRaLenh;
    row[lenhIdx.csHoanThanh] = (csHoanThanh === '' || isNaN(csHoanThanh)) ? '' : csHoanThanh;
    row[lenhIdx.bdth] = bdth;
    row[lenhIdx.hoanThanh] = (hoanThanhRaw === true || hoanThanhRaw === 1 || String(hoanThanhRaw).toUpperCase() === 'TRUE') ? 1 : 0;
    row[lenhIdx.dungLenh] = (dungLenhRaw === true || String(dungLenhRaw).toUpperCase() === 'TRUE');
    row[lenhIdx.nguonLenh] = String(r[idx.nguonLenh] || '').toUpperCase().trim();
    if (lenhIdx.nhaMay !== undefined && idx.nhaMay !== undefined) {
      row[lenhIdx.nhaMay] = r[idx.nhaMay];
    }

    parsed.push(row);
  });

  if (parsed.length === 0) {
    throw new Error('Không có dòng dữ liệu hợp lệ nào để nhập (xem chi tiết lỗi nếu có).');
  }

  var lenhLastRow = lenhSh.getLastRow();
  var existingIds = {};
  if (lenhLastRow >= 2) {
    var existingIdCol = lenhSh.getRange(2, lenhIdx.id + 1, lenhLastRow - 1, 1).getValues();
    existingIdCol.forEach(function (r, i) {
      if (r[0] !== '') existingIds[String(r[0])] = i + 2; // dòng thật trong LENH
    });
  }

  var imported = 0, updated = 0;
  parsed.forEach(function (row) {
    var idKey = String(row[lenhIdx.id]);
    if (existingIds[idKey]) {
      lenhSh.getRange(existingIds[idKey], 1, 1, row.length).setValues([row]);
      updated++;
    } else {
      lenhSh.appendRow(row);
      imported++;
    }
  });

  sortSheetRows_(lenhSh, [{ column: lenhIdx.bdth + 1, ascending: true }]);

  // Khoảng ngày BĐTH thực sự đã ghi vào LENH - để người vận hành đối chiếu
  // ngay với file gốc. Lệch ngày do múi giờ từng xảy ra và sai âm thầm
  // (xem alignTimeZoneWithTargetSheet_), nên luôn hiện con số này ra.
  var times = parsed.map(function (row) { return row[lenhIdx.bdth].getTime(); });
  var tz = Session.getScriptTimeZone();
  var range = {
    from: Utilities.formatDate(new Date(Math.min.apply(null, times)), tz, 'dd/MM/yyyy HH:mm'),
    to: Utilities.formatDate(new Date(Math.max.apply(null, times)), tz, 'dd/MM/yyyy HH:mm'),
  };

  return { imported: imported, updated: updated, skipped: skipped, range: range };
}
