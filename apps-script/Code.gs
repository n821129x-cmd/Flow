/**
 * 圖文資料驗收進度管理 — Google 試算表後端
 * 部署：部署 → 新增部署作業 → 網頁應用程式；執行身分「我」，存取權「所有人」。
 * 編輯密碼：先執行一次 setEditKey()（會跳出輸入框），或在「專案設定 → 指令碼屬性」新增 EDIT_KEY。
 *
 * 資料表「docs」：A 路徑、B 內容(JSON)、C 更新時間、D 更新者
 */
var SHEET = 'docs';

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.getRange(1, 1, 1, 4).setValues([['path', 'json', 'updatedAt', 'updatedBy']]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 320); sh.setColumnWidth(2, 520);
  }
  return sh;
}
function props_() { return PropertiesService.getScriptProperties(); }
function rev_() { return Number(props_().getProperty('REV') || 0); }
function bump_() { var r = rev_() + 1; props_().setProperty('REV', String(r)); return r; }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function keyOk_(k) { var real = props_().getProperty('EDIT_KEY'); return !!real && String(k || '') === real; }

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
  var op = (e && e.parameter && e.parameter.op) || 'all';
  if (op === 'rev') return out_({ ok: true, rev: rev_() });
  if (op === 'check') return out_({ ok: true, editor: keyOk_(e.parameter.key) });
  return out_({ ok: true, rev: rev_(), docs: readAll_() });
}

function doPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'bad_json' }); }
  if (!keyOk_(body.key)) return out_({ ok: false, error: 'not_editor' });
  var writes = body.writes || [];
  if (!writes.length) return out_({ ok: true, rev: rev_() });
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = sheet_(), n = sh.getLastRow() - 1, idx = {};
    var paths = n > 0 ? sh.getRange(2, 1, n, 1).getValues() : [];
    for (var i = 0; i < paths.length; i++) if (paths[i][0]) idx[paths[i][0]] = i + 2;
    var now = new Date().toISOString(), who = String(body.user || ''), append = [], dels = [];
    writes.forEach(function (w) {
      var p = String(w.path || '');
      if (!p) return;
      if (w.op === 'delete') { if (idx[p]) dels.push(idx[p]); return; }
      var row = [p, JSON.stringify(w.data || {}), now, who];
      if (idx[p]) sh.getRange(idx[p], 1, 1, 4).setValues([row]);
      else { append.push(row); idx[p] = -1; }
    });
    if (append.length) sh.getRange(sh.getLastRow() + 1, 1, append.length, 4).setValues(append);
    dels.sort(function (a, b) { return b - a; }).forEach(function (r) { sh.deleteRow(r); });
    var r = bump_();
    return out_({ ok: true, rev: r });
  } finally { lock.releaseLock(); }
}

/** 設定編輯密碼（在編輯器中執行一次） */
function setEditKey() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('設定編輯密碼', '網站上要修改資料時需輸入這組密碼：', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var k = res.getResponseText().trim();
  if (k.length < 4) { ui.alert('密碼至少 4 個字'); return; }
  props_().setProperty('EDIT_KEY', k);
  ui.alert('已設定');
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('驗收系統').addItem('設定編輯密碼', 'setEditKey').addToUi();
  sheet_();
}
