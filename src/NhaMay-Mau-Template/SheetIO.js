/**
 * Đọc/ghi dữ liệu giữa Sheet và QDD-Core-Library. Toàn bộ tính toán thật
 * nằm ở Library (QDDCoreLibrary) - các hàm ở đây chỉ chuyển đổi định dạng.
 */

/** Đọc giá trị cột B của CAI_DAT theo NHÃN ở cột A (không theo số dòng cố định). */
function getConfigValue_(label) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CAI_DAT);
  var lastRow = sh.getLastRow();
  if (lastRow === 0) return null;
  var rows = sh.getRange(1, 1, lastRow, 2).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === label) return rows[i][1];
  }
  return null;
}

function getConfig_() {
  var L = CAI_DAT_LABELS;
  return {
    rampRate: Number(getConfigValue_(L.RAMP_RATE)),
    qddVCoef: Number(getConfigValue_(L.QDD_V_COEF)),
    tolerance: Number(getConfigValue_(L.TOLERANCE)),
  };
}

/** Nhãn dòng cấu hình mã công tơ, vd meterLabel_('S1','Qdc') -> "Mã công tơ Qdc - S1". */
function meterLabel_(unit, role) {
  return 'Mã công tơ ' + role + ' - ' + unit;
}

/** Nhãn dòng cấu hình nhãn báo cáo, vd "Nhãn báo cáo - S1". */
function reportLabelLabel_(unit) {
  return 'Nhãn báo cáo - ' + unit;
}

/**
 * Danh sách tổ máy của nhà máy này, SUY RA TỪ CHÍNH CAI_DAT - mỗi dòng
 * "Mã công tơ Qdc - <tổ máy>" là một tổ máy.
 *
 * Nhờ vậy nhà máy có 3 tổ máy trở lên chỉ cần THÊM DÒNG vào CAI_DAT
 * ("Mã công tơ Qdc - S3", "Mã công tơ Qmp - S3", "Nhãn báo cáo - S3"),
 * không phải sửa một dòng code nào. Thứ tự trong CAI_DAT quyết định thứ
 * tự hiển thị trên sidebar và thứ tự các khối trong file báo cáo.
 *
 * @returns {string[]} vd ["S1","S2"] hoặc ["S1","S2","S3"]
 */
function getConfiguredUnits_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CAI_DAT);
  var lastRow = sh ? sh.getLastRow() : 0;
  if (!lastRow) return [];
  var rows = sh.getRange(1, 1, lastRow, 1).getValues();
  var units = [];
  rows.forEach(function (r) {
    var m = String(r[0]).trim().match(/^Mã công tơ Qdc - (.+)$/);
    if (m) {
      var unit = m[1].trim();
      if (unit && units.indexOf(unit) === -1) units.push(unit);
    }
  });
  return units;
}

/** Sidebar gọi để dựng danh sách tổ máy động (không hard-code S1/S2). */
function sidebar_getUnits() {
  return getConfiguredUnits_();
}

/**
 * Tra mã công tơ thật theo (tổ máy, loại dữ liệu) từ CAI_DAT.
 *
 * MÃ CÔNG TƠ KHÔNG CHỨA CHỮ SỐ NĂM. Tên file CSV thật có dạng
 * <ngày><tháng><năm 1 chữ số><mã công tơ>, vd "17076001.CSV" = ngày 17,
 * tháng 07, năm 2026 (số 6), công tơ 001. Cấu hình chỉ ghi phần mã công
 * tơ ("001"), KHÔNG ghi "6001" - nếu ghi kèm chữ số năm thì sang năm sau
 * tên file đổi thành "17077001.CSV" và hệ thống sẽ không nhận ra file nào.
 *
 * @param {string} unit vd "S1"
 * @param {string} role "Qdc"|"Qmp"
 * @returns {string|null}
 */
function resolveMeterCode_(unit, role) {
  var value = getConfigValue_(meterLabel_(unit, role));
  return value ? String(value).trim() : null;
}

/**
 * Đọc danh sách lệnh từ sheet LENH. Dò cột theo TÊN TIÊU ĐỀ (không theo
 * vị trí), nên hoạt động đúng với cả:
 *   - sheet LENH chuẩn 25 cột giống file gốc (dán thẳng file gốc vào),
 *   - sheet LENH rút gọn của các bản cũ (9/10 cột),
 *   - sheet có thứ tự cột khác hoặc thừa cột.
 * @returns {import('../QDD-Core-Library/CommandFilter').RawCommand[]}
 */
function readAllCommands_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LENH);
  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol === 0) return [];

  var headerRow = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var found = findCommandColumnIndices_(headerRow);
  if (found.missing.length > 0) {
    throw new Error('Sheet LENH thiếu cột: ' + found.missing.join(', ') +
      '. Chạy "QDD Smart System → Thiết lập sheet" hoặc dán lại dữ liệu kèm dòng tiêu đề đúng tên.');
  }
  var idx = found.indices;

  var rows = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  return rows
    .filter(function (r) { return r[idx.id] !== '' && r[idx.id] !== null; })
    .map(function (r) {
      return {
        id: r[idx.id],
        nhaMay: idx.nhaMay !== undefined ? r[idx.nhaMay] : '',
        toMay: r[idx.toMay],
        noiDungLenh: r[idx.noiDungLenh],
        csRaLenh: r[idx.csRaLenh] === '' ? null : Number(r[idx.csRaLenh]),
        csHoanThanh: r[idx.csHoanThanh] === '' ? null : Number(r[idx.csHoanThanh]),
        bdth: r[idx.bdth] instanceof Date ? r[idx.bdth] : null,
        hoanThanh: r[idx.hoanThanh],
        dungLenh: r[idx.dungLenh] === true || String(r[idx.dungLenh]).toUpperCase() === 'TRUE',
        nguonLenh: r[idx.nguonLenh],
      };
    });
}

/**
 * @returns {Array<{unit:string, role:string, code:string}>} 4 mã công tơ đã cấu hình (bỏ ô trống)
 */
function getAllConfiguredMeters_() {
  var out = [];
  getConfiguredUnits_().forEach(function (unit) {
    ['Qdc', 'Qmp'].forEach(function (role) {
      var code = resolveMeterCode_(unit, role);
      if (code) out.push({ unit: unit, role: role, code: code });
    });
  });
  return out;
}

/**
 * So khớp hai mã công tơ, CHẤP NHẬN MẤT SỐ 0 ĐỨNG ĐẦU.
 *
 * LÝ DO: Google Sheets tự hiểu ô "001" là SỐ 1 và cắt mất hai số 0 đầu.
 * Nghĩa là cùng một công tơ có thể đang nằm trong Sheet dưới hai dạng
 * ("001" ở chỗ này, 1 ở chỗ kia) tuỳ ô đó có được định dạng văn bản hay
 * không. So chuỗi thuần sẽ coi chúng là hai công tơ khác nhau và âm thầm
 * báo "thiếu CSV" cho ngày đã có dữ liệu.
 *
 * Với mã toàn chữ số thì so theo GIÁ TRỊ SỐ; còn lại so chuỗi như cũ.
 */
function sameMeterCode_(a, b) {
  var x = String(a == null ? '' : a).trim();
  var y = String(b == null ? '' : b).trim();
  if (x === y) return true;
  var dx = x.replace(/\D/g, ''), dy = y.replace(/\D/g, '');
  if (dx && dy) return Number(dx) === Number(dy);
  return false;
}

/**
 * Tên file có khớp mã công tơ này không.
 *
 * Đường chính: tên file thật có dạng <ngày 2><tháng 2><năm 1><mã công
 * tơ>, nên bỏ 5 ký tự đầu là ra đúng phần mã công tơ - so theo giá trị
 * số nên "001" hay 1 đều khớp. Đây là cách chính xác nhất vì so đúng
 * TOÀN BỘ phần mã, không phải so phần đuôi.
 *
 * Đường dự phòng (tên file dạng khác): so phần đuôi như trước.
 */
function meterMatchesFilename_(base, code) {
  // Chấp nhận cả mã có tiền tố chữ ("csv001") - người dùng hay thêm chữ
  // để Google Sheets khỏi cắt mất số 0 đầu. Chỉ lấy phần chữ số.
  var digits = String(code).replace(/\D/g, '');
  if (/^\d{6,}$/.test(base) && digits) {
    code = digits;
    // Tên file chuẩn: so ĐÚNG TOÀN BỘ phần mã, không so phần đuôi. So phần
    // đuôi ở đây sẽ khớp nhầm: mã "1" (đã mất số 0 đầu) khớp cả file của
    // công tơ 301 vì "...6301" cũng kết thúc bằng "1".
    return Number(base.slice(5)) === Number(code);
  }
  // Tên file không theo dạng chuẩn: so phần đuôi, nhưng chỉ với mã đủ dài
  // để không khớp bừa.
  return code.length >= 3 && code.length <= base.length && base.slice(-code.length) === code;
}

/**
 * Dò (tổ máy, loại dữ liệu) từ TÊN FILE, dựa vào mã công tơ đã cấu hình
 * trong CAI_DAT. Tên file CSV thật thường có dạng <ngày><tháng><mã công
 * tơ>.CSV (vd "17076001.CSV" = ngày 17, tháng 07, năm 2026 (số 6), công
 * tơ 001) - mã công tơ luôn nằm ở CUỐI phần số của tên file, nên so khớp
 * theo endsWith. Vì mã cấu hình KHÔNG kèm chữ số năm, cách so khớp này
 * vẫn đúng khi sang năm mới (file "17077001.CSV" vẫn kết thúc bằng "001").
 * @param {string} filename
 * @returns {{unit:string, role:string, code:string}|null}
 */
function matchMeterFromFilename_(filename) {
  var base = String(filename).replace(/\.[^.]+$/, ''); // bỏ phần đuôi .csv
  var meters = getAllConfiguredMeters_();
  // Ưu tiên mã dài hơn trước, tránh mã ngắn khớp nhầm vào phần đuôi của mã dài hơn.
  meters.sort(function (a, b) { return b.code.length - a.code.length; });
  for (var i = 0; i < meters.length; i++) {
    if (meterMatchesFilename_(base, meters[i].code)) return meters[i];
  }
  return null;
}

/**
 * Đọc thử ngày ở cột A của dòng dữ liệu đầu tiên trong CSV (dd-mm-yy(yy)
 * hoặc dd/mm/yy(yy) hoặc yyyy-mm-dd). Đây là đường DUY NHẤT xác định ngày
 * của một file CSV - không còn chỗ nào cho người dùng chọn ngày bằng tay.
 * @param {string} csvText
 * @returns {Date|null}
 */
function guessDateFromCsvText_(csvText) {
  var lines = String(csvText).split(/\r?\n/);
  var firstLine = null;
  for (var i = 0; i < lines.length; i++) {
    if (lines[i].trim().length > 0) { firstLine = lines[i]; break; }
  }
  if (!firstLine) return null;
  var firstCell = firstLine.split(',')[0].trim().replace(/^"|"$/g, '');

  var m = firstCell.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));

  m = firstCell.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{2,4})$/);
  if (m) {
    var year = m[3];
    if (year.length === 2) year = (Number(year) <= 79 ? '20' : '19') + year;
    return new Date(Number(year), Number(m[2]) - 1, Number(m[1]));
  }
  return null;
}

/**
 * Giữ lại các dòng thoả `keepPredicate`, xoá phần còn lại - bằng cách đọc
 * hết một lần rồi ghi lại, thay vì gọi deleteRow từng dòng (với vài nghìn
 * dòng, cách deleteRow chậm hơn hàng chục lần và dễ timeout).
 * @param {Sheet} sheet
 * @param {function(Array, number): boolean} keepPredicate  (row, index) -> giữ hay không
 * @returns {number} số dòng đã xoá
 */
function rewriteSheetRows_(sheet, keepPredicate) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol === 0) return 0;

  var all = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var kept = all.filter(keepPredicate);
  var removedCount = all.length - kept.length;
  if (removedCount === 0) return 0;

  sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
  if (kept.length > 0) {
    sheet.getRange(2, 1, kept.length, lastCol).setValues(kept);
  }
  return removedCount;
}

/** Sắp xếp lại toàn bộ dữ liệu (từ hàng 2 trở đi) của 1 sheet theo các cột chỉ định - gọi lại sau mỗi lần thêm/xoá dòng để luôn xem theo thứ tự ngày, không theo thứ tự thao tác. */
function sortSheetRows_(sheet, sortSpecs) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 3) return; // cần ít nhất 2 dòng dữ liệu mới cần sắp xếp
  sheet.getRange(2, 1, lastRow - 1, lastCol).sort(sortSpecs);
}

/** Ghi/ghi đè 1 dòng CSV_DATA cho (ngày, mã công tơ) - dùng chung cho luồng lưu 1 file và lưu hàng loạt. Tự sắp xếp lại theo Ngày rồi Mã công tơ sau khi ghi. */
function saveCsvRow_(date, meterCode, kwhGiao) {
  var dataSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CSV_DATA);
  var lastRow = dataSh.getLastRow();
  if (lastRow >= 2) {
    var existing = dataSh.getRange(2, 1, lastRow - 1, 2).getValues();
    var dateStrFmt = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    for (var i = existing.length - 1; i >= 0; i--) {
      var rowDate = existing[i][0];
      var rowDateStr = rowDate instanceof Date
        ? Utilities.formatDate(rowDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(rowDate);
      if (rowDateStr === dateStrFmt && sameMeterCode_(existing[i][1], meterCode)) {
        dataSh.deleteRow(i + 2);
      }
    }
  }
  dataSh.appendRow([date, meterCode].concat(kwhGiao));
  sortSheetRows_(dataSh, [{ column: 1, ascending: false }, { column: 2, ascending: true }]); // ngày mới nhất lên đầu, giống KET_QUA
}

/**
 * Đọc 48 giá trị KwhGiao theo (ngày, tổ máy, loại dữ liệu) - mã công tơ
 * thật được TRA TỪ CAI_DAT (khác nhau giữa các tổ máy/nhà máy, không cố
 * định).
 * @param {Date} date
 * @param {string} unit  vd "S1"
 * @param {string} role  "Qdc"|"Qmp"
 * @returns {number[]|null}  48 giá trị, hoặc null nếu chưa có dữ liệu/chưa cấu hình mã công tơ
 */
function readCsv48_(date, unit, role) {
  var meterCode = resolveMeterCode_(unit, role);
  if (!meterCode) return null;
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CSV_DATA);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return null;
  var rows = sh.getRange(2, 1, lastRow - 1, 2 + 48).getValues();
  var dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  for (var i = 0; i < rows.length; i++) {
    var rowDate = rows[i][0];
    var rowDateStr = rowDate instanceof Date
      ? Utilities.formatDate(rowDate, Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : String(rowDate);
    if (rowDateStr === dateStr && sameMeterCode_(rows[i][1], meterCode)) {
      return rows[i].slice(2, 50);
    }
  }
  return null;
}

/**
 * Lấy P0 cho (ngày, tổ máy) từ sheet P0_NGAY.
 *
 * P0 của ngày kế tiếp được GHI TỰ ĐỘNG ngay sau mỗi lần tính (xem
 * saveNextDayP0_) bằng công suất tại đúng 24:00, nên thường chỉ ngày đầu
 * tiên dùng hệ thống mới phải nhập tay.
 *
 * KHÔNG suy P0 từ Qdd chu kỳ 48 của ngày trước: Qdd chu kỳ 48 là công suất
 * TRUNG BÌNH khoảng 23:30-24:00, khác công suất tại 24:00 khi tổ máy đang
 * tăng/giảm tải - từng gây sai 29,4 MW kéo dài cả ngày (dữ liệu thật 19/07).
 *
 * @returns {{value:number, source:string}|null}
 */
function readOrInferP0_(date, unit) {
  var row = readP0Row_(date, unit);
  if (row === null) return null;
  return { value: row.p0, carryTarget: row.carryTarget, source: 'từ sheet P0_NGAY' };
}

/**
 * Ghi P0 cho NGÀY KẾ TIẾP = công suất tại 24:00 của ngày vừa tính.
 * Không ghi đè dòng do người dùng tự nhập (chỉ ghi đè dòng có ghi chú bắt
 * đầu bằng "Tự động"), để giá trị nhập tay luôn được tôn trọng.
 */
function saveNextDayP0_(date, unit, endPower, carry) {
  if (typeof endPower !== 'number' || isNaN(endPower)) return;
  var carryTarget = (carry && typeof carry.target === 'number') ? carry.target : '';

  var tz = Session.getScriptTimeZone();
  var nextDate = new Date(date);
  nextDate.setDate(nextDate.getDate() + 1);
  var nextKey = Utilities.formatDate(nextDate, tz, 'yyyy-MM-dd');
  var note = 'Tự động từ cuối ngày ' + Utilities.formatDate(date, tz, 'dd/MM/yyyy');

  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.P0_NGAY);
  var lastRow = sh.getLastRow();
  if (lastRow >= 2) {
    var rows = sh.getRange(2, 1, lastRow - 1, P0_NGAY_HEADERS.length).getValues();
    for (var i = 0; i < rows.length; i++) {
      var rowDate = rows[i][0];
      var rowKey = rowDate instanceof Date ? Utilities.formatDate(rowDate, tz, 'yyyy-MM-dd') : String(rowDate);
      if (rowKey === nextKey && String(rows[i][1]).toUpperCase() === unit.toUpperCase()) {
        var existingNote = String(rows[i][3] || '');
        if (existingNote.indexOf('Tự động') !== 0) return; // người dùng nhập tay -> giữ nguyên
        sh.getRange(i + 2, 3, 1, 3).setValues([[endPower, note, carryTarget]]);
        applyP0NumberFormat_(sh);
        return;
      }
    }
  }
  sh.appendRow([nextDate, unit, endPower, note, carryTarget]);
  sortSheetRows_(sh, [{ column: 1, ascending: false }, { column: 2, ascending: true }]);
  applyP0NumberFormat_(sh);
}

/** P0 và mục tiêu ramp hiển thị 2 số thập phân (giá trị thật giữ nguyên - xem applyResultNumberFormat_). */
function applyP0NumberFormat_(sh) {
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  sh.getRange(2, 3, lastRow - 1, 1).setNumberFormat('0.00'); // P0 (MW)
  sh.getRange(2, 5, lastRow - 1, 1).setNumberFormat('0.00'); // Ramp tiếp đến (MW)
}

/**
 * Đọc dòng P0 của (ngày, tổ máy): công suất đầu ngày + mục tiêu ramp còn
 * dở dang từ ngày trước (R07, cột "Ramp tiếp đến (MW)").
 * @returns {{p0:number, carryTarget:number|null}|null}
 */
function readP0Row_(date, unit) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.P0_NGAY);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return null;
  var rows = sh.getRange(2, 1, lastRow - 1, P0_NGAY_HEADERS.length).getValues();
  var dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  for (var i = 0; i < rows.length; i++) {
    var rowDate = rows[i][0];
    var rowDateStr = rowDate instanceof Date
      ? Utilities.formatDate(rowDate, Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : String(rowDate);
    if (rowDateStr === dateStr && String(rows[i][1]).toUpperCase() === unit.toUpperCase()) {
      var carry = rows[i][4];
      return {
        p0: Number(rows[i][2]),
        carryTarget: (carry === '' || carry === null || isNaN(Number(carry))) ? null : Number(carry),
      };
    }
  }
  return null;
}

/** Ghi kết quả 48 chu kỳ của 1 (ngày, tổ máy) vào KET_QUA (nối thêm, không xoá kết quả ngày khác). */
function appendResultToSheet_(date, unit, periods) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.KET_QUA);
  var rows = periods.map(function (p) {
    return [date, unit, periodLabel_(p.chuKy), p.qdd, p.qddV, p.qdc, p.pQdc, p.nguongDuoi, p.nguongTren, p.qmp, p.qdu, p.dauHieu];
  });
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, KET_QUA_HEADERS.length).setValues(rows);
  applyResultNumberFormat_(sh);
  // Ngày MỚI NHẤT lên đầu (người dùng thường xem kết quả gần nhất trước),
  // nhưng trong cùng 1 ngày thì tổ máy/chu kỳ vẫn tăng dần cho dễ đọc.
  sortSheetRows_(sh, [{ column: 1, ascending: false }, { column: 2, ascending: true }, { column: 3, ascending: true }]);
}

/**
 * Hiển thị mọi cột số của KET_QUA với ĐÚNG 2 SỐ THẬP PHÂN.
 *
 * Chỉ đổi ĐỊNH DẠNG HIỂN THỊ, KHÔNG làm tròn giá trị thật trong ô - báo
 * cáo tháng và P0 ngày kế tiếp vẫn đọc giá trị đầy đủ, nên sai số không
 * bị tích luỹ qua từng ngày. Nếu làm tròn hẳn số gốc, mỗi ngày sẽ nhích
 * đi một ít ở P0 và cộng dồn dần qua chuỗi ngày liên tiếp.
 */
function applyResultNumberFormat_(sh) {
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  var FIRST_NUMERIC_COL = 4;  // Qdd (MW)
  var NUMERIC_COL_COUNT = 8;  // Qdd -> Qdư (cột "Dấu hiệu" là chữ, không tính)
  sh.getRange(2, FIRST_NUMERIC_COL, lastRow - 1, NUMERIC_COL_COUNT).setNumberFormat('0.00');
}

/** Xoá sạch KET_QUA trước khi tính lại (tránh trùng lặp khi chạy lại cùng ngày). */
function clearResultsForDate_(date, unit) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.KET_QUA);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  var data = sh.getRange(2, 1, lastRow - 1, 2).getValues();
  var dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var keepRows = [];
  for (var i = 0; i < data.length; i++) {
    var rowDate = data[i][0];
    var rowDateStr = rowDate instanceof Date
      ? Utilities.formatDate(rowDate, Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : String(rowDate);
    var isMatch = rowDateStr === dateStr && String(data[i][1]).toUpperCase() === unit.toUpperCase();
    if (!isMatch) keepRows.push(i + 2);
  }
  // Xoá toàn bộ rồi ghi lại các dòng cần giữ - đơn giản, đủ dùng cho quy mô 1 nhà máy.
  if (keepRows.length === data.length) return; // không có gì trùng để xoá
  var all = sh.getRange(2, 1, lastRow - 1, KET_QUA_HEADERS.length).getValues();
  var filtered = all.filter(function (_, idx) { return keepRows.indexOf(idx + 2) !== -1; });
  sh.getRange(2, 1, lastRow - 1, KET_QUA_HEADERS.length).clearContent();
  if (filtered.length > 0) {
    sh.getRange(2, 1, filtered.length, KET_QUA_HEADERS.length).setValues(filtered);
  }
}

/**
 * Dọn dữ liệu nguồn của đúng (ngày, tổ máy) VỪA TÍNH XONG: xoá các lệnh
 * của tổ máy đó trong ngày đó khỏi LENH, và các dòng CSV_DATA ứng với
 * mã công tơ của tổ máy đó trong ngày đó.
 *
 * Chỉ đụng đến đúng (ngày, tổ máy) đã tính - dữ liệu của tổ máy còn lại
 * và các ngày khác giữ nguyên. Kết quả trong KET_QUA đã là giá trị tĩnh
 * nên không bị ảnh hưởng; nếu cần tính lại thì nhập lại từ file gốc.
 * @returns {{lenh:number, csv:number}} số dòng đã xoá
 */
function clearSourceDataForDate_(date, unit) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = Session.getScriptTimeZone();
  var dateStr = Utilities.formatDate(date, tz, 'yyyy-MM-dd');
  var unitUpper = unit.toUpperCase();
  var removed = { lenh: 0, csv: 0 };

  // --- LENH: xoá lệnh của (ngày, tổ máy) ---
  var lenhSh = ss.getSheetByName(SHEETS.LENH);
  var lenhLastCol = lenhSh.getLastColumn();
  if (lenhSh.getLastRow() >= 2 && lenhLastCol > 0) {
    var lenhFound = findCommandColumnIndices_(lenhSh.getRange(1, 1, 1, lenhLastCol).getValues()[0]);
    if (lenhFound.missing.length === 0) {
      var li = lenhFound.indices;
      removed.lenh = rewriteSheetRows_(lenhSh, function (row) {
        var bdth = row[li.bdth];
        if (!(bdth instanceof Date)) return true; // dòng không đọc được ngày -> giữ lại cho an toàn
        var sameDay = Utilities.formatDate(bdth, tz, 'yyyy-MM-dd') === dateStr;
        var sameUnit = String(row[li.toMay]).toUpperCase().slice(0, 2) === unitUpper.slice(0, 2);
        return !(sameDay && sameUnit);
      });
    }
  }

  // --- CSV_DATA: xoá các mã công tơ của tổ máy đó trong ngày đó ---
  var meterCodes = ['Qdc', 'Qmp']
    .map(function (role) { return resolveMeterCode_(unit, role); })
    .filter(function (c) { return !!c; });
  if (meterCodes.length > 0) {
    var csvSh = ss.getSheetByName(SHEETS.CSV_DATA);
    removed.csv = rewriteSheetRows_(csvSh, function (row) {
      var rowDate = row[0];
      var rowDateStr = rowDate instanceof Date
        ? Utilities.formatDate(rowDate, tz, 'yyyy-MM-dd') : String(rowDate);
      return !(rowDateStr === dateStr && meterCodes.indexOf(String(row[1])) !== -1);
    });
  }

  return removed;
}
