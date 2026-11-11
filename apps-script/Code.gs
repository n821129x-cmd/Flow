/**
 * 圖文資料驗收進度管理 — Google 試算表後端（v4：須登入才能讀取；管理者／編輯者／檢視者）
 *
 * 部署：部署 → 新增部署作業（或管理部署作業 → 編輯 → 新版本）→ 網頁應用程式
 *       執行身分「我」、誰可以存取「所有人」
 * 帳號：預設管理者 n821129x / n821129x（第一次登入後請到選單「驗收系統 → 新增／修改帳號」改密碼）
 *
 * 工作表「docs」：A 路徑、B 內容(JSON)、C 更新時間、D 更新者
 */
var SHEET = 'docs';
var DEFAULT_ADMIN = 'n821129x';
var TOKEN_DAYS = 30;

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sheet_() {
  var sh = ss_().getSheetByName(SHEET);
  if (!sh) {
    sh = ss_().insertSheet(SHEET);
    sh.getRange(1, 1, 1, 4).setValues([['path', 'json', 'updatedAt', 'updatedBy']]).setFontWeight('bold');
    sh.setFrozenRows(1); sh.setColumnWidth(1, 320); sh.setColumnWidth(2, 520);
  }
  return sh;
}
function props_() { return PropertiesService.getScriptProperties(); }
function rev_() { return Number(props_().getProperty('REV') || 0); }
function bump_() { var r = rev_() + 1; props_().setProperty('REV', String(r)); return r; }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/* ---------- 帳號 ---------- */
function hash_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)
    .map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join('');
}
function salt_() {
  var s = props_().getProperty('SALT');
  if (!s) { s = Utilities.getUuid(); props_().setProperty('SALT', s); }
  return s;
}
function users_() {
  var raw = props_().getProperty('USERS'), u = raw ? JSON.parse(raw) : {};
  if (!Object.keys(u).length) {
    u[DEFAULT_ADMIN] = { h: hash_(salt_() + DEFAULT_ADMIN), role: 'admin' };
    props_().setProperty('USERS', JSON.stringify(u));
  }
  return u;
}
function saveUsers_(u) { props_().setProperty('USERS', JSON.stringify(u)); }
function adminCount_(u) { return Object.keys(u).filter(function (k) { return u[k].role === 'admin'; }).length; }
function dropTokens_(name) { var t = tokens_(); Object.keys(t).forEach(function (k) { if (t[k].u === name) delete t[k]; }); saveTokens_(t); }
function tokens_() { var raw = props_().getProperty('TOKENS'); return raw ? JSON.parse(raw) : {}; }
function saveTokens_(t) {
  var now = Date.now();
  Object.keys(t).forEach(function (k) { if (t[k].exp < now) delete t[k]; });
  props_().setProperty('TOKENS', JSON.stringify(t));
}
function login_(name, pw) {
  name = String(name || '').trim().toLowerCase();
  var u = users_()[name];
  if (!u || u.h !== hash_(salt_() + String(pw || ''))) return null;
  var t = tokens_(), tok = Utilities.getUuid().replace(/-/g, '');
  t[tok] = { u: name, exp: Date.now() + TOKEN_DAYS * 864e5 };
  saveTokens_(t);
  return { token: tok, user: name, role: u.role || 'editor' };
}
function who_(tok) {
  var t = tokens_()[String(tok || '')];
  if (!t || t.exp < Date.now()) return null;
  var u = users_()[t.u];
  return u ? { user: t.u, role: u.role || 'editor' } : null;
}

/* ---------- 讀取 ---------- */
function readAll_() {
  var sh = sheet_(), n = sh.getLastRow() - 1, docs = [];
  if (n < 1) return docs;
  var v = sh.getRange(2, 1, n, 2).getValues();
  for (var i = 0; i < v.length; i++) {
    if (!v[i][0]) continue;
    try { docs.push({ path: String(v[i][0]), data: JSON.parse(v[i][1] || '{}') }); } catch (e) {}
  }
  return docs;
}
function doGet(e) {
  var p = (e && e.parameter) || {}, op = p.op || 'all', w = who_(p.token);
  if (op === 'ping') return out_({ ok: true });
  if (op === 'me') return out_({ ok: true, login: !!w, editor: !!w && w.role !== 'viewer', user: w ? w.user : '', role: w ? w.role : '' });
  if (!w) return out_({ ok: false, error: 'not_login' });   // 未登入不提供任何資料
  if (op === 'rev') return out_({ ok: true, rev: rev_() });
  return out_({ ok: true, rev: rev_(), docs: readAll_() });
}

/* ---------- 寫入 ---------- */
function doPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'bad_json' }); }
  if (body.op === 'login') {
    var r = login_(body.user, body.password);
    return out_(r ? { ok: true, token: r.token, user: r.user, role: r.role } : { ok: false, error: 'bad_login' });
  }
  var w = who_(body.token);
  if (!w) return out_({ ok: false, error: 'not_editor' });
  if (body.op === 'logout') { var t = tokens_(); delete t[body.token]; saveTokens_(t); return out_({ ok: true }); }
  if (body.op === 'password') {
    var us = users_(), me = us[w.user];
    if (!me || me.h !== hash_(salt_() + String(body.old || ''))) return out_({ ok: false, error: 'bad_login' });
    if (String(body.password || '').length < 4) return out_({ ok: false, error: 'short_password' });
    me.h = hash_(salt_() + String(body.password)); saveUsers_(us);
    return out_({ ok: true });
  }
  if (body.op === 'users' || body.op === 'setUser' || body.op === 'delUser') {
    if (w.role !== 'admin') return out_({ ok: false, error: 'not_admin' });
    var us2 = users_();
    if (body.op === 'setUser') {
      var nm = String(body.name || '').trim().toLowerCase(), role = (body.role === 'admin' || body.role === 'viewer') ? body.role : 'editor';
      if (!/^[a-z0-9_.@-]{2,40}$/.test(nm)) return out_({ ok: false, error: 'bad_name' });
      var ex = us2[nm];
      if ((!ex || body.password) && String(body.password || '').length < 4) return out_({ ok: false, error: 'short_password' });
      if (ex && ex.role === 'admin' && role !== 'admin' && adminCount_(us2) <= 1) return out_({ ok: false, error: 'last_admin' });
      us2[nm] = { h: body.password ? hash_(salt_() + String(body.password)) : ex.h, role: role,
        createdAt: ex && ex.createdAt || new Date().toISOString(), updatedAt: new Date().toISOString(), by: w.user };
      saveUsers_(us2);
      if (body.password && ex) dropTokens_(nm);
    } else if (body.op === 'delUser') {
      var dn = String(body.name || '').trim().toLowerCase();
      if (!us2[dn]) return out_({ ok: false, error: 'no_user' });
      if (dn === w.user) return out_({ ok: false, error: 'self' });
      if (us2[dn].role === 'admin' && adminCount_(us2) <= 1) return out_({ ok: false, error: 'last_admin' });
      delete us2[dn]; saveUsers_(us2); dropTokens_(dn);
    }
    var list = Object.keys(us2).sort().map(function (k) {
      return { user: k, role: us2[k].role || 'editor', createdAt: us2[k].createdAt || '', updatedAt: us2[k].updatedAt || '', by: us2[k].by || '' }; });
    return out_({ ok: true, users: list });
  }
  var writes = body.writes || [];
  if (w.role === 'viewer') return out_({ ok: false, error: 'viewer' });
  if (!writes.length) return out_({ ok: true, rev: rev_() });
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = sheet_(), n = sh.getLastRow() - 1, idx = {};
    var paths = n > 0 ? sh.getRange(2, 1, n, 1).getValues() : [];
    for (var i = 0; i < paths.length; i++) if (paths[i][0]) idx[paths[i][0]] = i + 2;
    var now = new Date().toISOString(), append = [], dels = [];
    writes.forEach(function (x) {
      var p = String(x.path || '');
      if (!p) return;
      if (x.op === 'delete') { if (idx[p] > 0) dels.push(idx[p]); return; }
      var row = [p, JSON.stringify(x.data || {}), now, w.user];
      if (idx[p] > 0) sh.getRange(idx[p], 1, 1, 4).setValues([row]);
      else if (idx[p] === -1) { for (var j = 0; j < append.length; j++) if (append[j][0] === p) append[j] = row; }
      else { append.push(row); idx[p] = -1; }
    });
    if (append.length) sh.getRange(sh.getLastRow() + 1, 1, append.length, 4).setValues(append);
    dels.sort(function (a, b) { return b - a; }).forEach(function (r) { sh.deleteRow(r); });
    return out_({ ok: true, rev: bump_() });
  } finally { lock.releaseLock(); }
}

/* ---------- 試算表選單：帳號管理 ---------- */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('驗收系統')
    .addItem('新增／修改帳號', 'menuSetUser')
    .addItem('刪除帳號', 'menuDelUser')
    .addItem('列出帳號', 'menuListUsers')
    .addSeparator()
    .addItem('重新整理網站資料（匯入 docs 後執行）', 'menuBump')
    .addToUi();
  sheet_(); users_();
}
function menuSetUser() {
  var ui = SpreadsheetApp.getUi();
  var a = ui.prompt('新增／修改帳號', '帳號（英數字）：', ui.ButtonSet.OK_CANCEL);
  if (a.getSelectedButton() !== ui.Button.OK) return;
  var name = a.getResponseText().trim().toLowerCase();
  if (!/^[a-z0-9_.@-]{2,40}$/.test(name)) { ui.alert('帳號只能用英數字與 _ . @ -'); return; }
  var b = ui.prompt('設定密碼', '「' + name + '」的密碼（至少 4 個字）：', ui.ButtonSet.OK_CANCEL);
  if (b.getSelectedButton() !== ui.Button.OK) return;
  var pw = b.getResponseText();
  if (pw.length < 4) { ui.alert('密碼至少 4 個字'); return; }
  var c = ui.prompt('權限', '輸入 1＝管理者、2＝編輯者、3＝檢視者（只能看）', ui.ButtonSet.OK_CANCEL);
  if (c.getSelectedButton() !== ui.Button.OK) return;
  var rl = { '1': 'admin', '2': 'editor', '3': 'viewer' }[c.getResponseText().trim()] || 'editor';
  var u = users_(); u[name] = { h: hash_(salt_() + pw), role: rl, createdAt: new Date().toISOString(), by: 'sheet' };
  saveUsers_(u); ui.alert('已儲存帳號「' + name + '」');
}
function menuDelUser() {
  var ui = SpreadsheetApp.getUi();
  var a = ui.prompt('刪除帳號', '要刪除的帳號：', ui.ButtonSet.OK_CANCEL);
  if (a.getSelectedButton() !== ui.Button.OK) return;
  var name = a.getResponseText().trim().toLowerCase(), u = users_();
  if (!u[name]) { ui.alert('找不到這個帳號'); return; }
  delete u[name];
  if (!Object.keys(u).length) { ui.alert('至少要保留一個帳號'); return; }
  saveUsers_(u);
  var t = tokens_(); Object.keys(t).forEach(function (k) { if (t[k].u === name) delete t[k]; }); saveTokens_(t);
  ui.alert('已刪除');
}
function menuListUsers() {
  var u = users_();
  SpreadsheetApp.getUi().alert('目前帳號', Object.keys(u).map(function (k) {
    return k + '（' + ({ admin: '管理者', editor: '編輯者', viewer: '檢視者' }[u[k].role] || '編輯者') + '）'; }).join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);
}
function menuBump() { bump_(); SpreadsheetApp.getUi().alert('已通知網站重新讀取資料'); }
