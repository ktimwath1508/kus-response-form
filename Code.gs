// วางโค้ดนี้ใน Google Sheet → ส่วนขยาย → Apps Script
const SHEET_NAME = 'responses';
const ADMIN_KEY = 'p4-2569';   // ← เปลี่ยนเป็นรหัสของคุณ (ใช้เปิดหน้าสรุป)
const HEADERS = ['เวลา', 'ห้อง', 'เลขที่', 'เพศ', 'ชื่อนักเรียน', 'คำนำหน้าผปค.', 'ชื่อผู้ปกครอง', 'รับทราบ/แนบเงิน', 'ไซส์เสื้อ', 'กางเกงขาสั้น(ชาย)', 'กางเกงขายาว(หญิง)', 'ค่าชุด'];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
  }
  if (sh.getRange(1, 12).getValue() !== 'ค่าชุด') sh.getRange(1, 12).setValue('ค่าชุด');
  return sh;
}
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const d = JSON.parse(e.postData.contents);
    const boy = d.gender === 'ด.ช.';
    sheet_().appendRow([
      new Date(), Number(d.room), Number(d.no), d.gender, d.student, d.parentTitle, d.parent,
      d.ack ? '✓' : '', "'" + d.shirt, boy ? "'" + d.pants : '', boy ? '' : "'" + d.pants, d.pay || ''
    ]);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  if ((e.parameter.key || '') !== ADMIN_KEY) return json_({ ok: false, error: 'unauthorized' });
  const v = sheet_().getDataRange().getValues();
  v.shift();
  const rows = v.map(r => ({
    ts: r[0], room: r[1], no: r[2], gender: r[3], student: r[4], parentTitle: r[5], parent: r[6],
    ack: r[7] === '✓', shirt: String(r[8]), pants: String(r[9] || r[10] || ''), pay: r[11] || ''
  }));
  return json_({ ok: true, rows });
}

// กด Run ฟังก์ชันนี้ 1 ครั้ง: ขอสิทธิ์ + สร้างแท็บ responses พร้อมหัวคอลัมน์
function setup() {
  sheet_();
}
