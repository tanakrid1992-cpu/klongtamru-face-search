/** ระบบค้นหารูปด้วยใบหน้า (Face Photo Finder) — Google Apps Script */
const CONFIG = {
  ROOT_FOLDER_ID: '1JPcHUl3zoTDhX5de43RsinUeqHzTEjYO',   // โฟลเดอร์หลัก (โฟลเดอร์ย่อย = อัลบั้ม/กิจกรรม)
  SHEET_ID: '',                                    // เว้นว่าง = สร้างชีตดัชนีให้อัตโนมัติ
  ADMIN_PASSWORD: 'admin1234',                     // เปลี่ยนก่อนใช้งานจริง
  ORG_NAME: 'เทศบาลตำบลคลองตำหรุ',
  THUMB_SIZE: 2048,                                // ขนาดภาพที่ใช้สแกนใบหน้า (px)
  MAX_DIST: 0.65                                   // ระยะห่างสูงสุดที่ส่งกลับไปให้ผู้ใช้
};

/* ---------- API (เรียกจาก GitHub Pages) ---------- */
const ACTIONS = { searchFaces, listPhotos, adminListPending, adminGetImage, adminSave, adminReset };
function doGet() { return json_({ ok: true, org: CONFIG.ORG_NAME }); }
function doPost(e) {
  try {
    const b = JSON.parse(e.postData.contents);
    if (!ACTIONS[b.action]) throw new Error('unknown action');
    return json_({ ok: true, data: ACTIONS[b.action].apply(null, b.args || []) });
  } catch (err) { return json_({ ok: false, error: err.message || String(err) }); }
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/* ---------- Storage ---------- */
function sheets_() {
  const p = PropertiesService.getScriptProperties();
  const id = CONFIG.SHEET_ID || p.getProperty('SHEET_ID');
  let ss;
  if (id) ss = SpreadsheetApp.openById(id);
  else { ss = SpreadsheetApp.create('FaceSearch_Index'); p.setProperty('SHEET_ID', ss.getId()); }
  let faces = ss.getSheetByName('faces');
  if (!faces) {
    faces = ss.insertSheet('faces');
    faces.appendRow(['fileId', 'fileName', 'album', 'faceNo', 'descriptor', 'indexedAt']);
    faces.setFrozenRows(1);
    faces.getRange('E:E').setNumberFormat('@');
  }
  let files = ss.getSheetByName('files');
  if (!files) {
    files = ss.insertSheet('files');
    files.appendRow(['fileId', 'fileName', 'album', 'faceCount', 'indexedAt']);
    files.setFrozenRows(1);
  }
  const s1 = ss.getSheetByName('Sheet1') || ss.getSheetByName('ชีต1');
  if (s1 && ss.getSheets().length > 2) ss.deleteSheet(s1);
  return { faces, files };
}

function auth_(pw) { if (pw !== CONFIG.ADMIN_PASSWORD) throw new Error('รหัสผ่านไม่ถูกต้อง'); }

/* ---------- Public: smart search ----------
 * รับใบหน้าอ้างอิงได้หลายรูป (หลายมุม) + Query Expansion:
 * รอบ 1 หาใบหน้าที่ตรงมาก → นำมาเฉลี่ยเป็น "โปรไฟล์ใบหน้า" → รอบ 2 ค้นซ้ำเพื่อจับรูปมุมอื่น/แสงต่าง */
function searchFaces(input, opt) {
  opt = opt || {}; const confirm = new Set(opt.confirm || []), reject = new Set(opt.reject || []);
  let qs = Array.isArray(input) && Array.isArray(input[0]) ? input : [input];
  qs = qs.filter(q => Array.isArray(q) && q.length === 128).map(q => q.map(Number));
  if (!qs.length) throw new Error('ข้อมูลใบหน้าไม่ถูกต้อง');
  const sh = sheets_().faces, n = sh.getLastRow();
  if (n < 2) return { indexed: 0, results: [] };
  const lim = CONFIG.MAX_DIST * CONFIG.MAX_DIST;
  const dist = (a, b) => { let s = 0; for (let i = 0; i < 128; i++) { const x = a[i] - b[i]; s += x * x; if (s > lim) return 9; } return Math.sqrt(s); };
  const faces = sh.getRange(2, 1, n - 1, 5).getValues()
    .map(r => ({ id: r[0], name: r[1], album: r[2] || 'ทั่วไป', d: String(r[4]).split(' ').map(Number) }))
    .filter(f => f.d.length === 128);
  // Relevance feedback: ใบหน้าในรูปที่ผู้ใช้ยืนยันว่า "ใช่ฉัน" → ใช้เป็นใบหน้าอ้างอิงเพิ่ม
  if (confirm.size) {
    const raw = (a, b) => { let s = 0; for (let i = 0; i < 128; i++) { const x = a[i] - b[i]; s += x * x; } return Math.sqrt(s); };
    const pick = {};
    faces.forEach(f => { if (!confirm.has(f.id)) return; const v = Math.min.apply(null, qs.map(q => raw(f.d, q))); if (!pick[f.id] || v < pick[f.id].v) pick[f.id] = { v, d: f.d }; });
    const base = qs.slice();
    Object.values(pick).forEach(p => { if (p.v < 0.75) base.push(p.d); });
    qs = base.slice(0, 40);
  }
  const strong = [];
  faces.forEach(f => { f.m = Math.min.apply(null, qs.map(q => dist(f.d, q))); if (f.m <= 0.42) strong.push(f); });
  strong.sort((a, b) => a.m - b.m);
  const pool = qs.concat(strong.slice(0, 12).map(f => f.d));
  if (pool.length > qs.length) {
    const cen = new Array(128).fill(0);
    pool.forEach(p => p.forEach((v, i) => cen[i] += v / pool.length));
    faces.forEach(f => { const v = dist(f.d, cen) + 0.03; if (v < f.m) f.m = v; });
  }
  const best = {};
  faces.forEach(f => { if (reject.has(f.id)) return; if (confirm.has(f.id)) f.m = 0; if (f.m <= CONFIG.MAX_DIST && (!best[f.id] || f.m < best[f.id].dist))
    best[f.id] = { id: f.id, name: f.name, album: f.album, dist: +f.m.toFixed(4) }; });
  return { indexed: faces.length, results: Object.values(best).sort((a, b) => a.dist - b.dist).slice(0, 500) };
}

/* ---------- Public: รายการรูปทั้งหมด (ที่สแกนแล้ว) ใหม่สุดก่อน ---------- */
function listPhotos() {
  const sh = sheets_().files, n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, 5).getValues()
    .map(r => ({ id: r[0], name: r[1], album: r[2] || 'ทั่วไป', t: +new Date(r[4]) || 0 }))
    .sort((a, b) => b.t - a.t);
}

/* ---------- Admin ---------- */
function adminListPending(pw) {
  auth_(pw);
  const s = sheets_(), done = new Set(), n = s.files.getLastRow();
  if (n > 1) s.files.getRange(2, 1, n - 1, 1).getValues().forEach(r => done.add(r[0]));
  const out = [];
  (function walk(folder, album) {
    const it = folder.getFiles();
    while (it.hasNext()) {
      const f = it.next();
      if (/^image\//.test(f.getMimeType()) && !done.has(f.getId()))
        out.push({ id: f.getId(), name: f.getName(), album: album || 'ทั่วไป' });
    }
    const fs = folder.getFolders();
    while (fs.hasNext()) { const sub = fs.next(); walk(sub, album ? album + ' / ' + sub.getName() : sub.getName()); }
  })(DriveApp.getFolderById(CONFIG.ROOT_FOLDER_ID), '');
  return { pending: out, indexedFiles: done.size, faces: Math.max(s.faces.getLastRow() - 1, 0) };
}

function adminGetImage(pw, id) {
  auth_(pw);
  const h = { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true };
  try {
    const meta = JSON.parse(UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + id +
      '?fields=thumbnailLink&supportsAllDrives=true', h).getContentText());
    if (meta.thumbnailLink) {
      const r = UrlFetchApp.fetch(meta.thumbnailLink.replace(/=s\d+$/, '=s' + CONFIG.THUMB_SIZE), h);
      if (r.getResponseCode() === 200) {
        const b = r.getBlob();
        return 'data:' + (b.getContentType() || 'image/jpeg') + ';base64,' + Utilities.base64Encode(b.getBytes());
      }
    }
  } catch (e) { /* fallback below */ }
  const b = DriveApp.getFileById(id).getBlob();
  return 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes());
}

function adminSave(pw, file, faces) {
  auth_(pw);
  const s = sheets_(), lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const now = new Date();
    if (faces && faces.length)
      s.faces.getRange(s.faces.getLastRow() + 1, 1, faces.length, 6)
        .setValues(faces.map((d, i) => [file.id, file.name, file.album, i + 1, d, now]));
    s.files.appendRow([file.id, file.name, file.album, faces ? faces.length : 0, now]);
  } finally { lock.releaseLock(); }
  return true;
}

function adminReset(pw) {
  auth_(pw);
  const s = sheets_();
  [s.faces, s.files].forEach(sh => { if (sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1); });
  return true;
}
