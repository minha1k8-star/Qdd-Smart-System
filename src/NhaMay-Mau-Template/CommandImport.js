/**
 * Nhập danh sách lệnh từ sheet LENH_STAGING vào LENH - nhận diện cột theo
 * TÊN TIÊU ĐỀ (không theo vị trí cột), để người dùng có thể dán NGUYÊN
 * file gốc (25 cột, bất kỳ thứ tự nào miễn còn dòng tiêu đề) mà không
 * cần tự cắt/sắp xếp lại đúng 9 cột của LENH - tránh lỗi lệch cột đã gặp
 * trong thực tế (xem docs/14_Knowledge_Transfer.md).
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
 * @returns {{imported:number, updated:number, skipped:Array<{row:number, reason:string}>}}
 */
function importCommandsFromStaging_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var stagingSh = ss.getSheetByName('LENH_STAGING');
  if (!stagingSh || stagingSh.getLastRow() < 2) {
    throw new Error('LENH_STAGING đang trống. Dán dữ liệu (kèm dòng tiêu đề) vào sheet này trước.');
  }

  var lastRow = stagingSh.getLastRow();
  var lastCol = stagingSh.getLastColumn();
  var allValues = stagingSh.getRange(1, 1, lastRow, lastCol).getValues();
  var headerRow = allValues[0];
  var dataRows = allValues.slice(1);

  var found = findCommandColumnIndices_(headerRow);
  if (found.missing.length > 0) {
    throw new Error('Không tìm thấy cột: ' + found.missing.join(', ') +
      '. Kiểm tra lại dòng tiêu đề trong LENH_STAGING có đúng tên như file gốc không.');
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

  // Bản đồ tên cột (đã chuẩn hoá) của LENH_STAGING, để copy cả các cột phụ
  // (Người ra lệnh, Lý do lệnh...) khi LENH có sẵn cột cùng tên.
  var stagingByName = {};
  headerRow.forEach(function (h, i) {
    var key = normalizeHeader_(h);
    if (key && stagingByName[key] === undefined) stagingByName[key] = i;
  });

  var skipped = [];
  var parsed = [];
  dataRows.forEach(function (r, i) {
    var rowNum = i + 2; // số dòng thật trong LENH_STAGING
    var id = r[idx.id];
    if (id === '' || id === null) return; // dòng trống, bỏ qua âm thầm

    var bdth = r[idx.bdth];
    if (!(bdth instanceof Date)) {
      skipped.push({ row: rowNum, reason: 'Thời điểm BĐTH không phải kiểu ngày-giờ hợp lệ (dán "Paste values only"/"Giá trị" có thể làm mất định dạng ngày - thử dán bình thường Ctrl+V lại).' });
      return;
    }

    // Dựng dòng theo đúng số cột của LENH: cột nào trùng tên với staging thì lấy nguyên giá trị.
    var row = new Array(lenhLastCol).fill('');
    lenhNormalized.forEach(function (name, col) {
      if (name && stagingByName[name] !== undefined) {
        row[col] = r[stagingByName[name]];
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

  stagingSh.getRange(2, 1, lastRow - 1, lastCol).clearContent();

  return { imported: imported, updated: updated, skipped: skipped };
}
