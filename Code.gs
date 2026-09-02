/**
 * 우리집 가계부 — Google Apps Script backend
 * 시트 구성: 거래내역 / 자산 / 설정  (열 순서는 아래 상수와 같아야 합니다)
 */
const SHEET_TX = '거래내역';
const SHEET_ASSETS = '자산';
const SHEET_SETTINGS = '설정';
// 거래내역 열 순서
const TX_COLS = ['id', '날짜', '구분', '대분류', '소분류', '내용', '금액(원)', '통화', '외화금액', '환율', '고정', '결제수단', '계좌', '입력자', '등록시각', '수정시각'];

function doGet(e) {
  if (e && e.parameter && e.parameter.ping) {
    return ContentService.createTextOutput(JSON.stringify({ ok: true, sheet: ss_().getName() })).setMimeType(ContentService.MimeType.JSON);
  }
  let out = null;
  ['Index', 'index', 'INDEX'].some(name => {
    try { out = HtmlService.createHtmlOutputFromFile(name); return true; } catch (e) { return false; }
  });
  if (!out) {
    return HtmlService.createHtmlOutput(
      '<div style="font-family:sans-serif;padding:24px;line-height:1.6">' +
      '<h3>Index.html 파일이 아직 없어요</h3>' +
      '<p>Apps Script 편집기에서 파일 목록 옆 <b>＋ → HTML</b>을 누르고 이름을 <b>Index</b>로 정한 뒤, ' +
      'Index.html 내용을 붙여넣고 저장하세요. 그다음 <b>배포 → 배포 관리 → 새 버전 → 배포</b>를 누르면 됩니다.</p></div>');
  }
  return out.setTitle('우리집 가계부')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ---------------- JSON API (GitHub Pages 등 외부 화면에서 호출)
const API_FNS_ = { getAll, addTx, updateTx, deleteTx, saveSettings, saveAssets, addAccount, deleteAccount };
function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents || '{}');
    const fn = API_FNS_[req.fn];
    if (!fn) throw new Error('알 수 없는 요청: ' + req.fn);
    out = { result: fn.apply(null, req.args || []) };
  } catch (err) {
    out = { error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sheet_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('시트가 없습니다: ' + name);
  return sh;
}
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try { return fn(); } finally { lock.releaseLock(); }
}
function dateStr_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(v || '').slice(0, 10);
}
function ym_(d) { return String(d).slice(0, 7); }

// ---------------- 설정
function getSettings_() {
  const sh = sheet_(SHEET_SETTINGS);
  const rows = sh.getDataRange().getValues();
  const s = {};
  for (let i = 1; i < rows.length; i++) {
    const k = rows[i][0], v = rows[i][1];
    if (!k) continue;
    try { s[k] = JSON.parse(v); } catch (e) { s[k] = v; }
  }
  return s;
}
function saveSettings(patch) {
  return withLock_(() => {
    const sh = sheet_(SHEET_SETTINGS);
    const rows = sh.getDataRange().getValues();
    const idx = {};
    for (let i = 1; i < rows.length; i++) if (rows[i][0]) idx[rows[i][0]] = i + 1;
    Object.keys(patch).forEach(k => {
      const val = JSON.stringify(patch[k]);
      if (idx[k]) sh.getRange(idx[k], 2).setValue(val);
      else sh.appendRow([k, val]);
    });
    return getSettings_();
  });
}

// ---------------- 자산 (행 = 계좌, 열 = 월)
function getAssets_() {
  const sh = sheet_(SHEET_ASSETS);
  const rows = sh.getDataRange().getValues();
  if (!rows.length) return { months: [], accounts: [] };
  const header = rows[0]; // 그룹 | 계좌 | 2025-12 | 2026-01 | ...
  const months = header.slice(2).map(h => h instanceof Date ? Utilities.formatDate(h, Session.getScriptTimeZone(), 'yyyy-MM') : String(h)).filter(Boolean);
  const accounts = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[1]) continue;
    const bal = {};
    months.forEach((m, j) => { const v = r[j + 2]; if (v !== '' && v !== null && v !== undefined) bal[m] = Number(v) || 0; });
    accounts.push({ id: 'r' + (i + 1), row: i + 1, group: String(r[0]), name: String(r[1]), bal });
  }
  return { months, accounts };
}
function saveAssets(month, values) { // values: {accountId: number}
  return withLock_(() => {
    const sh = sheet_(SHEET_ASSETS);
    const header = sh.getRange(1, 1, 1, Math.max(2, sh.getLastColumn())).getValues()[0];
    const months = header.slice(2).map(h => h instanceof Date ? Utilities.formatDate(h, Session.getScriptTimeZone(), 'yyyy-MM') : String(h));
    let col = months.indexOf(month);
    if (col < 0) {
      // insert keeping chronological order
      let insertAt = months.length;
      for (let i = 0; i < months.length; i++) if (months[i] && months[i] > month) { insertAt = i; break; }
      sh.insertColumnAfter(insertAt + 2);
      col = insertAt;
      sh.getRange(1, col + 3).setValue(month).setNumberFormat('@');
    }
    Object.keys(values).forEach(id => {
      const row = Number(id.replace('r', ''));
      if (row > 1) sh.getRange(row, col + 3).setValue(Number(values[id]) || 0);
    });
    return getAssets_();
  });
}
function addAccount(group, name) {
  return withLock_(() => {
    sheet_(SHEET_ASSETS).appendRow([group, name]);
    return getAssets_();
  });
}
function deleteAccount(id) {
  return withLock_(() => {
    const sh = sheet_(SHEET_ASSETS);
    const row = Number(id.replace('r', ''));
    if (row > 1) sh.deleteRow(row);
    return getAssets_();
  });
}

// ---------------- 거래내역
function rowToTx_(r) {
  return {
    id: String(r[0]), d: dateStr_(r[1]), t: String(r[2]), c: String(r[3]), s: String(r[4]), m: String(r[5] || ''),
    a: Number(r[6]) || 0, cur: String(r[7] || 'KRW'), fx: Number(r[8]) || 0, rate: Number(r[9]) || 1,
    f: r[10] === 'Y' || r[10] === true, p: String(r[11] || ''), acc: String(r[12] || ''), u: String(r[13] || ''),
    ts: Number(r[14]) || 0, edited: Number(r[15]) || 0
  };
}
function txToRow_(t) {
  return [t.id, t.d, t.t, t.c, t.s, t.m || '', Number(t.a) || 0, t.cur || 'KRW', Number(t.fx) || 0, Number(t.rate) || 1,
    t.f ? 'Y' : '', t.p || '', t.acc || '', t.u || '', Number(t.ts) || Date.now(), Number(t.edited) || ''];
}
function getTx_(year) {
  const sh = sheet_(SHEET_TX);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const rows = sh.getRange(2, 1, last - 1, TX_COLS.length).getValues();
  const out = [];
  rows.forEach(r => { if (!r[0]) return; const t = rowToTx_(r); if (!year || t.d.slice(0, 4) === String(year)) out.push(t); });
  return out;
}
function findRow_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}
function addTx(t) {
  return withLock_(() => {
    const sh = sheet_(SHEET_TX);
    if (!t.id) t.id = Utilities.getUuid().slice(0, 8);
    if (!t.ts) t.ts = Date.now();
    sh.appendRow(txToRow_(t));
    sh.getRange(sh.getLastRow(), 2).setNumberFormat('@');
    return t;
  });
}
function updateTx(t) {
  return withLock_(() => {
    const sh = sheet_(SHEET_TX);
    const row = findRow_(sh, t.id);
    t.edited = Date.now();
    if (row < 0) { sh.appendRow(txToRow_(t)); return t; }
    sh.getRange(row, 1, 1, TX_COLS.length).setValues([txToRow_(t)]);
    return t;
  });
}
function deleteTx(id) {
  return withLock_(() => {
    const sh = sheet_(SHEET_TX);
    const row = findRow_(sh, id);
    if (row > 0) sh.deleteRow(row);
    return { ok: true };
  });
}

// ---------------- 한 번에 읽기
function getAll(year) {
  return { settings: getSettings_(), assets: getAssets_(), tx: getTx_(year), now: Date.now(), tz: Session.getScriptTimeZone() };
}
