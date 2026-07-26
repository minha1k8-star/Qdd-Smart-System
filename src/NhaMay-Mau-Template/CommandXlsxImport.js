/**
 * Nhập danh sách lệnh TRỰC TIẾP TỪ FILE EXCEL (.xlsx/.xls) do người dùng
 * tải lên từ sidebar - không phải copy/dán tay nữa.
 *
 * Cách làm: file được tải lên Google Drive kèm yêu cầu chuyển đổi sang
 * Google Sheets (Drive API), đọc xong thì XOÁ file tạm ngay. Không dùng
 * thư viện đọc .xlsx nào - Apps Script không có sẵn, còn Drive thì chuyển
 * đổi được chính xác cả ô ngày-giờ (giữ nguyên kiểu Date, không thành chữ).
 *
 * Bảng đọc được đưa thẳng vào importCommandTable_() để gộp vào LENH -
 * không còn sheet trung gian LENH_STAGING. Mọi quy tắc dò cột theo TÊN
 * tiêu đề, gộp theo ID Lệnh, sắp xếp theo BĐTH vẫn giữ nguyên.
 */

var XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
var GOOGLE_SHEET_MIME = 'application/vnd.google-apps.spreadsheet';

/** Số dòng đầu mỗi sheet sẽ dò tìm dòng tiêu đề (file gốc hay có dòng tên đơn vị/tiêu đề báo cáo phía trên). */
var XLSX_HEADER_SCAN_ROWS = 25;

/**
 * Tải blob lên Drive, yêu cầu chuyển sang Google Sheets.
 * Dùng UrlFetchApp thay vì Advanced Drive Service để nhà máy khác chỉ cần
 * copy file/clasp push là chạy được, không phải bật thêm dịch vụ nào.
 * @returns {string} id của file Google Sheets tạm vừa tạo
 */
function uploadAndConvertToSheet_(blob, name) {
  var boundary = '-------QddSmartSystemBoundary';
  var metadata = { name: name, mimeType: GOOGLE_SHEET_MIME };
  var head = '--' + boundary + '\r\n' +
    'Content-Type: application/json; charset=UTF-8\r\n\r\n' +
    JSON.stringify(metadata) + '\r\n' +
    '--' + boundary + '\r\n' +
    'Content-Type: ' + (blob.getContentType() || XLSX_MIME) + '\r\n\r\n';
  var tail = '\r\n--' + boundary + '--';

  var payload = Utilities.newBlob(head).getBytes()
    .concat(blob.getBytes())
    .concat(Utilities.newBlob(tail).getBytes());

  var response = UrlFetchApp.fetch(
    'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&supportsAllDrives=true',
    {
      method: 'post',
      contentType: 'multipart/related; boundary=' + boundary,
      payload: payload,
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true,
    }
  );

  if (response.getResponseCode() !== 200) {
    throw new Error('Google Drive không đọc được file Excel này (mã lỗi ' +
      response.getResponseCode() + '). Kiểm tra file có đúng định dạng .xlsx/.xls và không đặt mật khẩu bảo vệ.');
  }
  var id = JSON.parse(response.getContentText()).id;
  if (!id) throw new Error('Không nhận được file sau khi chuyển đổi. Thử lại, hoặc dán trực tiếp vào sheet LENH.');
  return id;
}

/**
 * Đặt múi giờ của bản Google Sheets tạm bằng đúng múi giờ của Sheet đích
 * TRƯỚC KHI đọc dữ liệu.
 *
 * LỖI ĐÃ GẶP THẬT (nghiêm trọng, sai âm thầm - 07/2026): file Excel lưu
 * ngày-giờ dạng "trần", không kèm múi giờ. Bản Sheets tạm do Drive tạo ra
 * lấy múi giờ MẶC ĐỊNH CỦA TÀI KHOẢN GOOGLE (thường America/Los_Angeles),
 * trong khi Sheet đích dùng Asia/Ho_Chi_Minh. Apps Script đọc ô "18:36"
 * thành 18:36 giờ Los Angeles rồi ghi sang Sheet đích thành 08:36 HÔM SAU
 * - toàn bộ lệnh bị dịch đúng +14 giờ (chênh lệch +07 so với -07 giờ hè),
 * lệnh của ngày này nhảy sang ngày kia, kết quả sai dây chuyền mà KHÔNG
 * có lỗi nào được báo.
 *
 * setSpreadsheetTimeZone chỉ đổi cách DIỄN GIẢI giá trị ngày-giờ, không
 * sửa số liệu gốc - nên sau khi đặt đúng múi giờ, ô "18:36" trong file
 * Excel được đọc lại thành đúng 18:36 giờ Việt Nam.
 */
function alignTimeZoneWithTargetSheet_(tempSs) {
  var targetTz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  if (tempSs.getSpreadsheetTimeZone() !== targetTz) {
    tempSs.setSpreadsheetTimeZone(targetTz);
    SpreadsheetApp.flush();
  }
}

/**
 * Tìm sheet + dòng tiêu đề chứa đủ các cột bắt buộc trong file đã chuyển đổi.
 * File gốc thường có nhiều sheet (danh sách lệnh, phụ lục...) và vài dòng
 * tiêu đề/logo phía trên bảng, nên phải dò chứ không mặc định sheet đầu/dòng 1.
 * @returns {{values: Array[], headerRowIndex: number, sheetName: string}}
 */
function locateCommandTable_(ss) {
  var sheets = ss.getSheets();
  var triedInfo = [];

  for (var s = 0; s < sheets.length; s++) {
    var sh = sheets[s];
    var lastRow = sh.getLastRow();
    var lastCol = sh.getLastColumn();
    if (lastRow < 2 || lastCol < 1) continue;

    var scanRows = Math.min(lastRow, XLSX_HEADER_SCAN_ROWS);
    var head = sh.getRange(1, 1, scanRows, lastCol).getValues();
    var bestMissing = null;

    for (var r = 0; r < scanRows; r++) {
      var found = findCommandColumnIndices_(head[r]);
      if (found.missing.length === 0) {
        var values = sh.getRange(r + 1, 1, lastRow - r, lastCol).getValues();
        return { values: values, headerRowIndex: r + 1, sheetName: sh.getName() };
      }
      if (bestMissing === null || found.missing.length < bestMissing.length) {
        bestMissing = found.missing;
      }
    }
    if (bestMissing) {
      triedInfo.push('sheet "' + sh.getName() + '" thiếu: ' + bestMissing.join(', '));
    }
  }

  throw new Error('Không tìm thấy bảng danh sách lệnh trong file (đã dò ' + sheets.length +
    ' sheet, ' + XLSX_HEADER_SCAN_ROWS + ' dòng đầu mỗi sheet).\n' +
    triedInfo.slice(0, 3).join('\n') +
    '\nFile phải có dòng tiêu đề với đủ các cột: ' +
    Object.keys(COMMAND_REQUIRED_FIELD_LABELS).map(function (k) {
      return COMMAND_REQUIRED_FIELD_LABELS[k];
    }).join(', ') + '.');
}

/**
 * Toàn bộ luồng: file Excel (base64) -> LENH.
 * @param {string} base64  nội dung file, đã bỏ tiền tố "data:...;base64,"
 * @param {string} filename
 * @returns {{imported:number, updated:number, skipped:Array, sheetName:string, headerRowIndex:number, totalRows:number}}
 */
function importCommandsFromXlsx_(base64, filename) {
  var name = String(filename || 'danh-sach-lenh.xlsx');
  if (!/\.(xlsx|xlsm|xls)$/i.test(name)) {
    throw new Error('File "' + name + '" không phải Excel (.xlsx/.xlsm/.xls). Nếu là file CSV công tơ thì dùng mục 1.');
  }

  var blob = Utilities.newBlob(Utilities.base64Decode(base64), XLSX_MIME, name);
  var tempId = uploadAndConvertToSheet_(blob, 'TAM - ' + name + ' - ' + new Date().getTime());

  try {
    var tempSs = SpreadsheetApp.openById(tempId);
    alignTimeZoneWithTargetSheet_(tempSs);
    var located = locateCommandTable_(tempSs);
    // Dòng dữ liệu đầu tiên trong file = ngay dưới dòng tiêu đề, để báo lỗi
    // đúng số dòng người dùng nhìn thấy khi mở file gốc.
    var result = importCommandTable_(located.values, located.headerRowIndex + 1);
    result.sheetName = located.sheetName;
    result.headerRowIndex = located.headerRowIndex;
    result.totalRows = located.values.length - 1;
    return result;
  } finally {
    // Luôn xoá file tạm, kể cả khi nhập lỗi - không để rác tích tụ trên Drive.
    try {
      DriveApp.getFileById(tempId).setTrashed(true);
    } catch (e) {
      // Không chặn kết quả chỉ vì dọn file tạm thất bại.
    }
  }
}
