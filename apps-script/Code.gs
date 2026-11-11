/**
 * Viltrox Launch Event 05.12.2026 – nhận form đăng ký từ landing page và ghi vào Google Sheet.
 *
 * Cài đặt (làm 1 lần):
 * 1. Mở Google Sheet nhận dữ liệu → Tiện ích mở rộng (Extensions) → Apps Script.
 * 2. Xóa code mặc định, dán toàn bộ file này vào, bấm Lưu.
 * 3. Triển khai (Deploy) → Tùy chọn triển khai mới (New deployment) → loại "Ứng dụng web" (Web app):
 *      - Thực thi dưới dạng (Execute as): Tôi (Me)
 *      - Ai có quyền truy cập (Who has access): Bất kỳ ai (Anyone)
 *    → Triển khai → cấp quyền → copy "URL ứng dụng web" (dạng https://script.google.com/macros/s/.../exec).
 * 4. Dán URL đó vào hằng FORM_ENDPOINT trong index.html.
 *
 * Khi sửa code này: Deploy → Quản lý triển khai → sửa → Phiên bản: "Phiên bản mới" (giữ nguyên URL).
 */

const SHEET_NAME = 'Đăng ký';
const TZ = 'Asia/Ho_Chi_Minh';
const HEADERS = ['Thời gian đăng ký', 'Họ và tên', 'Số điện thoại', 'Địa chỉ', 'Ngành nghề', 'Nguồn', 'Trình duyệt'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const data = JSON.parse((e && e.postData && e.postData.contents) || '{}');

    // Bẫy spam: ô ẩn "website" chỉ bot mới điền
    if (data.website) return json_({ ok: true });

    const name = clean_(data.name, 100);
    const phone = normalizePhone_(data.phone);
    const address = clean_(data.address, 250);
    const job = clean_(data.job, 120);

    if (!name || !phone || !address || !job) return json_({ ok: false, error: 'invalid' });

    const sheet = getSheet_();
    const last = sheet.getLastRow();
    if (last > 1) {
      const phones = sheet.getRange(2, 3, last - 1, 1).getDisplayValues().map(r => r[0]);
      if (phones.indexOf(phone) !== -1) return json_({ ok: true, duplicate: true });
    }

    sheet.appendRow([
      Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm:ss'),
      name, phone, address, job,
      clean_(data.source, 200), clean_(data.ua, 300)
    ]);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Mở URL web app trên trình duyệt để kiểm tra script đang chạy
function doGet() {
  return json_({ ok: true, service: 'viltrox-launch-event-registration' });
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#111111').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.getRange('C:C').setNumberFormat('@'); // giữ số 0 đầu của số điện thoại
    sheet.setColumnWidths(1, HEADERS.length, 180);
  }
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
  if (/^[=+\-@]/.test(s)) s = "'" + s; // chặn chèn công thức vào Sheet
  return s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
