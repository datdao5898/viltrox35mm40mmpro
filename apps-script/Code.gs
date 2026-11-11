/**
 * Viltrox Launch Event 05.12.2026 – nhận form đăng ký từ landing page, ghi vào Google Sheet.
 * Sheet: https://docs.google.com/spreadsheets/d/1c5lF2XymVFrLkosZt56UT6PHKmP81Kpaioe3_NI-9B4
 *
 * Cài đặt (1 lần):
 * 1. Mở Sheet → Tiện ích mở rộng → Apps Script → xóa code mặc định → dán file này → Lưu.
 * 2. Triển khai → Tùy chọn triển khai mới → Loại: Ứng dụng web
 *      Thực thi dưới dạng: Tôi   |   Ai có quyền truy cập: Bất kỳ ai
 *    → Triển khai → Cấp quyền → copy "URL ứng dụng web" (…/exec) dán vào FORM_ENDPOINT trong index.html.
 * Khi sửa code: Triển khai → Quản lý triển khai → ✏️ → Phiên bản mới (URL giữ nguyên).
 */

const SPREADSHEET_ID = '1c5lF2XymVFrLkosZt56UT6PHKmP81Kpaioe3_NI-9B4';
const SHEET_NAME = 'Đăng ký';
const TZ = 'Asia/Ho_Chi_Minh';
const HEADERS = ['Thời gian đăng ký', 'Họ và tên', 'Số điện thoại', 'Địa chỉ', 'Ngành nghề', 'Nguồn'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (d.website) return json_({ ok: true });                 // bẫy spam

    const name = clean_(d.ho_ten, 100);
    const phone = normalizePhone_(d.so_dien_thoai);
    const address = clean_(d.dia_chi, 250);
    const job = clean_(d.nganh_nghe, 120);
    if (!name || !phone || !address || !job) return json_({ ok: false, error: 'invalid' });

    const sheet = getSheet_();
    const last = sheet.getLastRow();
    if (last > 1) {                                             // mỗi SĐT chỉ ghi 1 lần
      const phones = sheet.getRange(2, 3, last - 1, 1).getDisplayValues().map(r => r[0]);
      if (phones.indexOf(phone) !== -1) return json_({ ok: true, duplicate: true });
    }
    sheet.appendRow([
      Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm:ss'),
      name, phone, address, job, clean_(d.nguon, 200)
    ]);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Mở URL /exec trên trình duyệt để kiểm tra script đang chạy
function doGet() {
  return json_({ ok: true, service: 'viltrox-launch-event-registration' });
}

function getSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#111111').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  sheet.getRange('C2:C').setNumberFormat('@');                  // giữ số 0 đầu của SĐT
  return sheet;
}

function normalizePhone_(v) {
  let p = String(v || '').replace(/[^\d+]/g, '');
  if (p.indexOf('+84') === 0) p = '0' + p.slice(3);
  else if (p.indexOf('84') === 0 && p.length === 11) p = '0' + p.slice(2);
  return /^0[35789]\d{8}$/.test(p) ? p : '';
}

function clean_(v, max) {
  let s = String(v || '').replace(/\s+/g, ' ').trim().slice(0, max);
  if (/^[=+\-@]/.test(s)) s = "'" + s;                          // chặn chèn công thức
  return s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
