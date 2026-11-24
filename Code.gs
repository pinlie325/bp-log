/**
 * 血壓血糖手冊 — Google Apps Script 後端
 * 資料存在你 Google 雲端硬碟的試算表「血壓血糖紀錄資料」
 * 第一次請先執行 setup()，在「執行項目」記錄裡會印出試算表網址
 */
const APP_NAME = '血壓血糖紀錄資料';
const MAX_FAILS = 5;          // 連續錯幾次就鎖
const LOCK_MINUTES = 5;       // 鎖幾分鐘

/* ---------- API（給 GitHub Pages 網頁呼叫） ---------- */
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function doGet() {
  return json_({ ok: true, app: '血壓血糖手冊 API' });
}
function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    const fns = { register, login, whoAmI, logout, getRecords, addRecord, deleteRecord };
    const f = fns[req.fn];
    if (!f) throw new Error('未知的動作');
    return json_({ ok: true, data: f.apply(null, Array.isArray(req.args) ? req.args : []) });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

/* ---------- 試算表 ---------- */
function setup() {
  const ss = ss_();
  Logger.log('資料試算表：' + ss.getUrl());
}

function ss_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SSID');
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (e) { /* 被刪掉就重建 */ }
  }
  const ss = SpreadsheetApp.create(APP_NAME);
  const u = ss.getSheets()[0].setName('使用者');
  u.appendRow(['代碼', '稱呼', '鹽', '驗證碼雜湊', '建立時間', '錯誤次數', '鎖定到']);
  const r = ss.insertSheet('紀錄');
  r.appendRow(['編號', '代碼', '類型', '量測時間', '收縮壓', '舒張壓', '脈搏', '血糖', '時段', '建立時間']);
  const s = ss.insertSheet('登入');
  s.appendRow(['憑證', '代碼', '建立時間']);
  u.getRange('A:A').setNumberFormat('@');
  r.getRange('A:B').setNumberFormat('@');
  s.getRange('A:B').setNumberFormat('@');
  [u, r, s].forEach(sh => sh.setFrozenRows(1));
  u.hideColumns(3, 2); // 隱藏鹽與雜湊，避免誤改
  s.hideSheet();
  props.setProperty('SSID', ss.getId());
  return ss;
}
const sh_ = name => ss_().getSheetByName(name);

/* ---------- 工具 ---------- */
function hash_(salt, pin) {
  const b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':' + pin, Utilities.Charset.UTF_8);
  return b.map(x => ('0' + (x & 255).toString(16)).slice(-2)).join('');
}
function normCode_(code) {
  const c = String(code || '').trim().toLowerCase();
  if (c.length < 2 || c.length > 20 || /\s/.test(c)) throw new Error('代碼要 2 到 20 個字，不能有空白');
  return c;
}
function checkPin_(pin) {
  if (!/^\d{4}$/.test(String(pin))) throw new Error('驗證碼要 4 個數字');
  return String(pin);
}
function findUser_(code) {
  const sh = sh_('使用者');
  const v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (String(v[i][0]) === code) return { sh, row: i + 1, v: v[i] };
  return null;
}
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try { return fn(); } finally { lock.releaseLock(); }
}
function newSession_(code, name) {
  const token = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
  sh_('登入').appendRow([token, code, new Date()]);
  CacheService.getScriptCache().put('t_' + token, code, 21600);
  return { token, name };
}
function auth_(token) {
  if (!token) throw new Error('AUTH');
  const cache = CacheService.getScriptCache();
  const hit = cache.get('t_' + token);
  if (hit) return hit;
  const v = sh_('登入').getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (String(v[i][0]) === token) {
    cache.put('t_' + token, String(v[i][1]), 21600);
    return String(v[i][1]);
  }
  throw new Error('AUTH');
}

/* ---------- 給網頁呼叫的函式 ---------- */
function register(code, name, pin) {
  code = normCode_(code); pin = checkPin_(pin);
  name = String(name || '').trim().slice(0, 12) || code;
  return withLock_(() => {
    if (findUser_(code)) throw new Error('這個代碼已經有人用了，請換一個');
    const salt = Utilities.getUuid();
    sh_('使用者').appendRow([code, name, salt, hash_(salt, pin), new Date(), 0, '']);
    return newSession_(code, name);
  });
}

function login(code, pin) {
  code = normCode_(code); pin = checkPin_(pin);
  return withLock_(() => {
    const u = findUser_(code);
    if (!u) throw new Error('找不到這個代碼，請再確認一次');
    const lockUntil = u.v[6] ? new Date(u.v[6]).getTime() : 0;
    if (lockUntil > Date.now()) throw new Error('錯太多次了，請 ' + LOCK_MINUTES + ' 分鐘後再試');
    if (hash_(u.v[2], pin) !== u.v[3]) {
      const fails = Number(u.v[5] || 0) + 1;
      if (fails >= MAX_FAILS) u.sh.getRange(u.row, 6, 1, 2).setValues([[0, new Date(Date.now() + LOCK_MINUTES * 60000)]]);
      else u.sh.getRange(u.row, 6).setValue(fails);
      throw new Error('驗證碼不對，請再試一次');
    }
    u.sh.getRange(u.row, 6, 1, 2).setValues([[0, '']]);
    return newSession_(code, String(u.v[1] || code));
  });
}

function whoAmI(token) {
  const code = auth_(token);
  const u = findUser_(code);
  if (!u) throw new Error('AUTH');
  return { name: String(u.v[1] || code) };
}

function logout(token) {
  CacheService.getScriptCache().remove('t_' + token);
  withLock_(() => {
    const sh = sh_('登入'), v = sh.getDataRange().getValues();
    for (let i = v.length - 1; i >= 1; i--) if (String(v[i][0]) === token) sh.deleteRow(i + 1);
  });
  return true;
}

function getRecords(token) {
  const code = auth_(token);
  const v = sh_('紀錄').getDataRange().getValues();
  const out = [];
  for (let i = 1; i < v.length; i++) {
    const r = v[i];
    if (String(r[1]) !== code) continue;
    const t = r[3] instanceof Date ? r[3].getTime() : Number(r[3]);
    if (!t) continue;
    out.push(r[2] === 'bp'
      ? { id: String(r[0]), k: 'bp', t, s: Number(r[4]), d: Number(r[5]), p: r[6] === '' ? null : Number(r[6]) }
      : { id: String(r[0]), k: 'gl', t, v: Number(r[7]), m: String(r[8]) });
  }
  out.sort((a, b) => b.t - a.t);
  return out;
}

function addRecord(token, rec) {
  const code = auth_(token);
  const t = Number(rec && rec.t) || Date.now();
  const id = Utilities.getUuid().slice(0, 13);
  let row;
  if (rec.k === 'bp') {
    const s = Math.round(rec.s), d = Math.round(rec.d), p = rec.p ? Math.round(rec.p) : '';
    if (!(s >= 50 && s <= 260 && d >= 30 && d <= 160 && s > d)) throw new Error('血壓數字不合理，請再確認');
    row = [id, code, 'bp', new Date(t), s, d, p, '', '', new Date()];
  } else if (rec.k === 'gl') {
    const v = Math.round(rec.v), m = ['f', 'a', 'r'].indexOf(rec.m) >= 0 ? rec.m : 'r';
    if (!(v >= 20 && v <= 600)) throw new Error('血糖數字不合理，請再確認');
    row = [id, code, 'gl', new Date(t), '', '', '', v, m, new Date()];
  } else throw new Error('資料格式錯誤');
  withLock_(() => sh_('紀錄').appendRow(row));
  return getRecords(token);
}

function deleteRecord(token, id) {
  const code = auth_(token);
  withLock_(() => {
    const sh = sh_('紀錄'), v = sh.getDataRange().getValues();
    for (let i = v.length - 1; i >= 1; i--) if (String(v[i][0]) === String(id) && String(v[i][1]) === code) { sh.deleteRow(i + 1); break; }
  });
  return getRecords(token);
}
