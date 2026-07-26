/**
 * Nhật ký thao tác - ghi lại AI làm GÌ, LÚC NÀO, KẾT QUẢ RA SAO.
 *
 * Dùng khi nhiều người cùng dùng một Sheet: lúc số liệu có vấn đề thì
 * truy được ai đã tính/nhập/xoá gì. Ghi cả thao tác LỖI - đó thường là
 * dòng hữu ích nhất khi truy vết.
 *
 * VÌ SAO PHẢI TỰ KHAI TÊN: `Session.getActiveUser().getEmail()` chỉ trả
 * về email khi người dùng cùng miền Google Workspace với chủ Sheet. Nhà
 * máy dùng Gmail cá nhân nên hàm đó luôn trả chuỗi rỗng. Vì vậy mỗi
 * người tự khai tên một lần, lưu trong UserProperties (riêng theo từng
 * tài khoản Google, không lẫn giữa các người dùng).
 *
 * GIỚI HẠN CẦN BIẾT: tên này là do người dùng TỰ KHAI, không phải danh
 * tính đã xác thực - dùng để phối hợp công việc, KHÔNG dùng làm bằng
 * chứng quy trách nhiệm. Muốn biết chắc ai sửa gì thì xem
 * "Tệp → Lịch sử phiên bản" của Google Sheets (Google tự ghi, không sửa được).
 */

var NHAT_KY_HEADERS = ['Thời gian', 'Người dùng', 'Thao tác', 'Chi tiết', 'Kết quả'];

/** Giữ tối đa ngần này dòng nhật ký; cũ hơn thì xoá bớt để sheet không phình vô hạn. */
var NHAT_KY_MAX_ROWS = 2000;

var USER_NAME_PROPERTY = 'QDD_TEN_NGUOI_DUNG';

/** Tên người dùng đã khai; '' nếu chưa khai. */
function getUserName_() {
  try {
    return String(PropertiesService.getUserProperties().getProperty(USER_NAME_PROPERTY) || '').trim();
  } catch (e) {
    return '';
  }
}

/** Sidebar gọi để hiển thị tên hiện tại. */
function sidebar_getUserName() {
  return getUserName_();
}

/** Sidebar gọi khi người dùng khai tên lần đầu hoặc đổi tên. */
function sidebar_setUserName(name) {
  var clean = String(name || '').trim().slice(0, 50);
  if (!clean) throw new Error('Tên không được để trống.');
  PropertiesService.getUserProperties().setProperty(USER_NAME_PROPERTY, clean);
  logAction_('Đặt tên người dùng', clean, 'OK');
  return clean;
}

/**
 * Ghi 1 dòng nhật ký. Dòng mới nhất nằm TRÊN CÙNG (chèn vào hàng 2) để
 * mở sheet là thấy ngay việc vừa làm, không phải cuộn xuống cuối.
 *
 * Hàm này KHÔNG được phép làm hỏng thao tác chính: mọi lỗi khi ghi log
 * đều nuốt lặng, vì mất một dòng nhật ký nhẹ hơn nhiều so với việc làm
 * hỏng một lần tính đã chạy xong.
 *
 * @param {string} thaoTac  vd "Tính", "Nhập lệnh"
 * @param {string} chiTiet  tham số người dùng đã chọn
 * @param {string} ketQua   tóm tắt kết quả, hoặc "✗ <lỗi>" nếu thất bại
 */
function logAction_(thaoTac, chiTiet, ketQua) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(SHEETS.NHAT_KY);
    if (!sh) {
      sh = ss.insertSheet(SHEETS.NHAT_KY);
      sh.getRange(1, 1, 1, NHAT_KY_HEADERS.length).setValues([NHAT_KY_HEADERS]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }

    sh.insertRowBefore(2);
    sh.getRange(2, 1, 1, NHAT_KY_HEADERS.length).setValues([[
      new Date(),
      getUserName_() || '(chưa khai tên)',
      String(thaoTac || ''),
      String(chiTiet || '').slice(0, 500),
      String(ketQua || '').replace(/\n/g, ' | ').slice(0, 500),
    ]]);
    sh.getRange(2, 1).setNumberFormat('dd/MM/yyyy HH:mm:ss');

    var lastRow = sh.getLastRow();
    if (lastRow > NHAT_KY_MAX_ROWS + 1) {
      sh.deleteRows(NHAT_KY_MAX_ROWS + 2, lastRow - NHAT_KY_MAX_ROWS - 1);
    }
  } catch (e) {
    // Cố tình bỏ qua: ghi nhật ký thất bại không được làm hỏng thao tác chính.
  }
}

/**
 * Chạy một thao tác của sidebar và ghi nhật ký cho cả trường hợp thành
 * công lẫn thất bại. Lỗi vẫn được ném tiếp ra ngoài để sidebar hiển thị
 * như cũ - hàm này chỉ quan sát, không nuốt lỗi.
 *
 * @param {string} thaoTac
 * @param {string} chiTiet
 * @param {function(): {value: *, tomTat: string}} body
 * @returns {*} giá trị `value` do body trả về
 */
function runLogged_(thaoTac, chiTiet, body) {
  var out;
  try {
    out = body();
  } catch (e) {
    logAction_(thaoTac, chiTiet, '✗ ' + e.message);
    throw e;
  }
  logAction_(thaoTac, chiTiet, out.tomTat);
  return out.value;
}
