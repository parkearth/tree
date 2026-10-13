/**
 * 空樹穴巡查缺失表 — Google Apps Script 後端
 * 手機 PWA 填報 → 寫入 Google 試算表「空樹穴巡查缺失表」，照片存到雲端硬碟資料夾。
 *
 * 部署：擴充功能 → Apps Script → 貼上本檔 → 修改下方 CONFIG
 *       → 部署 → 新增部署作業 → 類型「網頁應用程式」
 *         執行身分：我　｜　誰可以存取：所有人
 *       → 複製「網頁應用程式網址」貼到 PWA 的設定頁。
 * 修改程式後要「管理部署作業 → 編輯 → 版本：新版本」，網址才會吃到新程式。
 */
const CONFIG = {
  SHEET_ID: '請貼上試算表ID',          // 試算表網址 /d/ 與 /edit 中間那串
  SHEET_NAME: '空樹穴巡查缺失表',       // 工作表名稱
  FOLDER_ID: '',                        // 照片資料夾 ID；留空會在試算表同一層自動建立「空樹穴巡查照片」
  TOKEN: 'treepit2026',                 // 通行碼，要與 PWA 設定相同
  FIRST_DATA_ROW: 4,                    // 第 4 列開始是資料（第 3 列是標題）
  PHOTO_ROW_HEIGHT: 120,
};

const COL = { SERIAL: 1, DIST: 2, ADDR: 3, TYPE: 4, PHOTO_BAD: 5, PHOTO_FIX: 6, REMARK: 7 };
const TYPES = ['空樹穴且未立告示牌', '廢樹頭且未立告牌', '舊告示牌未撤除'];

/* ---------------- 入口 ---------------- */
function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    if (p.action === 'ping') return json_({ ok: true, sheet: sheet_().getName(), time: new Date().toISOString() });
    checkToken_(p.token);
    if (p.action === 'reverse') return json_(reverseGeocode_(Number(p.lat), Number(p.lng)));
    if (p.action === 'list') return json_(listRows_(Number(p.limit) || 50));
    return json_({ ok: false, error: '未知的 action' });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); }
  catch (err) { return json_({ ok: false, error: '資料格式錯誤' }); }
  const lock = LockService.getScriptLock();
  try {
    checkToken_(body.token);
    lock.waitLock(30000);
    if (body.action === 'submit') return json_(submit_(body));
    if (body.action === 'improve') return json_(improve_(body));
    return json_({ ok: false, error: '未知的 action' });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

/* ---------------- 新增一筆缺失 ---------------- */
function submit_(b) {
  if (!b.uid) throw new Error('缺少 uid');
  const props = PropertiesService.getScriptProperties();
  const done = props.getProperty('uid_' + b.uid);
  if (done) { const d = JSON.parse(done); return { ok: true, duplicate: true, serial: d.serial, row: d.row }; }

  if (!b.dist) throw new Error('請選擇行政區');
  if (!b.address) throw new Error('請填寫位置');
  if (TYPES.indexOf(b.type) < 0) throw new Error('缺失種類不正確');

  const sh = sheet_();
  const row = targetRow_(sh);
  const serial = sh.getRange(row, COL.SERIAL).getValue();

  sh.getRange(row, COL.DIST, 1, 3).setValues([[b.dist, b.address, b.type]]);
  sh.getRange(row, COL.REMARK).setValue(remark_(b));

  const stamp = Utilities.formatDate(new Date(b.createdAt || Date.now()), 'Asia/Taipei', 'yyyyMMdd-HHmm');
  if (b.photoBad) putPhoto_(sh, row, COL.PHOTO_BAD, b.photoBad, `${stamp}_${serial}_缺失`);
  if (b.photoFix) putPhoto_(sh, row, COL.PHOTO_FIX, b.photoFix, `${stamp}_${serial}_改善`);
  if (b.photoBad || b.photoFix) sh.setRowHeight(row, CONFIG.PHOTO_ROW_HEIGHT);

  SpreadsheetApp.flush();
  props.setProperty('uid_' + b.uid, JSON.stringify({ serial: serial, row: row }));
  return { ok: true, serial: serial, row: row };
}

/* ---------------- 補上改善照片 ---------------- */
function improve_(b) {
  if (!b.photoFix) throw new Error('沒有收到改善照片');
  const sh = sheet_();
  const row = findRowBySerial_(sh, b.serial);
  if (!row) throw new Error('找不到序號 ' + b.serial);
  const stamp = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMdd-HHmm');
  putPhoto_(sh, row, COL.PHOTO_FIX, b.photoFix, `${stamp}_${b.serial}_改善`);
  sh.setRowHeight(row, CONFIG.PHOTO_ROW_HEIGHT);
  if (b.note) {
    const r = sh.getRange(row, COL.REMARK);
    r.setValue([r.getValue(), b.note].filter(String).join('；'));
  }
  SpreadsheetApp.flush();
  return { ok: true, serial: b.serial, row: row };
}

/* ---------------- 清單（給「補改善照片」用） ---------------- */
function listRows_(limit) {
  const sh = sheet_();
  const last = sh.getLastRow();
  if (last < CONFIG.FIRST_DATA_ROW) return { ok: true, rows: [] };
  const n = last - CONFIG.FIRST_DATA_ROW + 1;
  const vals = sh.getRange(CONFIG.FIRST_DATA_ROW, 1, n, 7).getValues();
  const fx = sh.getRange(CONFIG.FIRST_DATA_ROW, COL.PHOTO_FIX, n, 1).getFormulas();
  const rows = [];
  vals.forEach((v, i) => {
    if (!v[COL.DIST - 1] && !v[COL.ADDR - 1]) return;
    rows.push({ serial: v[0], dist: v[1], address: v[2], type: v[3],
                hasFix: !!(fx[i][0] || v[COL.PHOTO_FIX - 1]), remark: v[6] });
  });
  return { ok: true, rows: rows.slice(-limit).reverse() };
}

/* ---------------- 地址反查 ---------------- */
function reverseGeocode_(lat, lng) {
  if (!isFinite(lat) || !isFinite(lng)) throw new Error('座標不正確');
  const res = Maps.newGeocoder().setLanguage('zh-TW').setRegion('tw').reverseGeocode(lat, lng);
  if (!res.results || !res.results.length) return { ok: true, address: '' };
  let addr = res.results[0].formatted_address || '';
  addr = addr.replace(/^\d{3,6}\s*/, '').replace(/^台灣/, '').replace(/^臺灣/, '').trim();
  addr = addr.replace(/^台北市/, '臺北市');
  return { ok: true, address: addr };
}

/* ---------------- 工具 ---------------- */
function sheet_() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sh = ss.getSheetByName(CONFIG.SHEET_NAME) || ss.getSheets()[0];
  return sh;
}

function checkToken_(t) {
  if (CONFIG.TOKEN && t !== CONFIG.TOKEN) throw new Error('通行碼錯誤，請到設定頁確認');
}

/** 找第一個「行政區、位置、缺失種類」都空白的列；沒有就在最後新增一列並沿用格式、下拉選單、序號+1 */
function targetRow_(sh) {
  const first = CONFIG.FIRST_DATA_ROW;
  const last = Math.max(sh.getLastRow(), first);
  const vals = sh.getRange(first, 1, last - first + 1, 4).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (!vals[i][1] && !vals[i][2] && !vals[i][3]) {
      const row = first + i;
      if (vals[i][0] === '' || vals[i][0] === null) sh.getRange(row, 1).setValue(prevSerial_(vals, i) + 1);
      return row;
    }
  }
  const row = last + 1;
  sh.insertRowAfter(last);
  const src = sh.getRange(last, 1, 1, 7), dst = sh.getRange(row, 1, 1, 7);
  src.copyTo(dst, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  src.copyTo(dst, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
  dst.clearContent();
  sh.setRowHeight(row, sh.getRowHeight(last) > 40 ? 21 : sh.getRowHeight(last));
  sh.getRange(row, 1).setValue(prevSerial_(vals, vals.length) + 1);
  return row;
}

function prevSerial_(vals, i) {
  for (let k = i - 1; k >= 0; k--) { const n = Number(vals[k][0]); if (isFinite(n) && n > 0) return n; }
  return 0;
}

function findRowBySerial_(sh, serial) {
  const first = CONFIG.FIRST_DATA_ROW, last = sh.getLastRow();
  if (last < first) return 0;
  const vals = sh.getRange(first, 1, last - first + 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) if (String(vals[i][0]) === String(serial)) return first + i;
  return 0;
}

function remark_(b) {
  const parts = [];
  if (b.remark) parts.push(b.remark);
  if (b.holeId) parts.push('樹穴編號 ' + b.holeId);
  if (b.lat && b.lng) parts.push(`座標 ${Number(b.lat).toFixed(6)},${Number(b.lng).toFixed(6)}`);
  if (b.reporter) parts.push('填報：' + b.reporter);
  const t = Utilities.formatDate(new Date(b.createdAt || Date.now()), 'Asia/Taipei', 'MM/dd HH:mm');
  parts.push('巡查 ' + t);
  return parts.join('；');
}

function folder_() {
  if (CONFIG.FOLDER_ID) return DriveApp.getFolderById(CONFIG.FOLDER_ID);
  const props = PropertiesService.getScriptProperties();
  const saved = props.getProperty('photoFolderId');
  if (saved) { try { return DriveApp.getFolderById(saved); } catch (e) {} }
  const file = DriveApp.getFileById(CONFIG.SHEET_ID);
  const parents = file.getParents();
  const parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  const it = parent.getFoldersByName('空樹穴巡查照片');
  const f = it.hasNext() ? it.next() : parent.createFolder('空樹穴巡查照片');
  props.setProperty('photoFolderId', f.getId());
  return f;
}

/** dataUrl: "data:image/jpeg;base64,...." → 存檔 → 儲存格放 IMAGE()，註解放原圖連結 */
function putPhoto_(sh, row, col, dataUrl, name) {
  const m = /^data:(image\/[\w+.-]+);base64,(.+)$/.exec(dataUrl);
  if (!m) throw new Error('照片格式錯誤');
  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], name + '.jpg');
  const file = folder_().createFile(blob);
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) { /* 網域禁止公開分享時，圖片仍可用連結開啟 */ }
  const id = file.getId();
  const cell = sh.getRange(row, col);
  cell.setFormula(`=IMAGE("https://drive.google.com/thumbnail?id=${id}&sz=w800")`);
  cell.setNote('原圖：' + file.getUrl());
  return id;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/** 部署前可在編輯器執行一次，完成授權並檢查設定 */
function testSetup() {
  const sh = sheet_();
  Logger.log('工作表：' + sh.getName() + '，最後一列：' + sh.getLastRow());
  Logger.log('照片資料夾：' + folder_().getUrl());
  Logger.log(JSON.stringify(reverseGeocode_(25.0563, 121.5155)));
}
