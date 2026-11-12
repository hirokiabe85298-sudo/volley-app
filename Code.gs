/**
 * バレー試合分析（Google版）サーバー側のプログラム
 * このファイルの中身を Apps Script の「コード.gs」にすべて貼り付けてください。
 * 記録はこのスプレッドシートの「data」シートに保存されます（手で編集しないでください）。
 */
var SHEET = 'data';
var CH = 40000;          // 1つのセルに入れる文字数（セルの上限は50,000文字）
var APP_CACHE = 600;     // アプリの画面を使い回す秒数（10分）
var APP_PART = 30000;    // キャッシュ1件あたりの文字数

// アプリの画面はGitHubから読み込みます（つながらないときは、このプロジェクトの「index」を使います）
var APP_URL = 'https://raw.githubusercontent.com/hirokiabe85298-sudo/volley-app/main/index.html';

function doGet() {
  var html = appHtml_();
  var out = html ? HtmlService.createHtmlOutput(html) : HtmlService.createHtmlOutputFromFile('index');
  return out
    .setTitle('バレー試合分析')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ---------- A1: アプリの画面をキャッシュして速く開く ---------- */
function appHtml_() {
  var c = CacheService.getScriptCache();
  try {
    var n = Number(c.get('app_n') || 0);
    if (n > 0) {
      var keys = [];
      for (var i = 0; i < n; i++) keys.push('app_' + i);
      var got = c.getAll(keys), parts = [];
      for (var j = 0; j < n; j++) { if (got['app_' + j] == null) { parts = null; break; } parts.push(got['app_' + j]); }
      if (parts) return parts.join('');
    }
  } catch (e) {}
  var html = null;
  try {
    var r = UrlFetchApp.fetch(APP_URL + '?t=' + Math.floor(Date.now() / 60000), { muteHttpExceptions: true });
    if (r.getResponseCode() === 200) {
      var h = r.getContentText('UTF-8');
      if (h.indexOf('id="main"') > 0) html = h;
    }
  } catch (e) {}
  if (html) {
    try {
      var obj = {}, k = 0;
      for (var p = 0; p < html.length; p += APP_PART) { obj['app_' + k] = html.slice(p, p + APP_PART); k++; }
      obj.app_n = String(k);
      c.putAll(obj, APP_CACHE);
    } catch (e) {}
  }
  return html;
}

/** 設定画面の「アプリを最新版にする」から呼ばれる */
function vbRefresh() {
  CacheService.getScriptCache().remove('app_n');
  return true;
}

/* ---------- 記録の保存場所 ---------- */
function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName(SHEET);
  if (!s) {
    s = ss.insertSheet(SHEET);
    s.getRange(1, 1, 1, 4).setValues([['path', 'ts', 'parts', 'data（アプリ用。編集しないでください）']]);
    s.setFrozenRows(1);
  }
  return s;
}

function index_(s) {
  var n = s.getLastRow() - 1;
  if (n < 1) return [];
  return s.getRange(2, 1, n, 3).getValues(); // [path, ts, parts]
}

/* ---------- A2: 一覧（path→時刻・行）をキャッシュして、読み込みを軽くする ---------- */
function map_() {
  var j = CacheService.getScriptCache().get('idx');
  if (j) { try { return JSON.parse(j); } catch (e) {} }
  return rebuildMap_();
}
function rebuildMap_() {
  var rows = index_(sheet_()), map = {};
  for (var i = 0; i < rows.length; i++) if (rows[i][0]) map[String(rows[i][0])] = [Number(rows[i][1]) || 0, i + 2, Number(rows[i][2]) || 0];
  saveMap_(map);
  return map;
}
function saveMap_(map) {
  try { CacheService.getScriptCache().put('idx', JSON.stringify(map), 21600); } catch (e) {}
}

function readRow_(s, row, parts) {
  if (!parts) return null;
  var vals = s.getRange(row, 4, 1, parts).getValues()[0];
  var str = vals.map(function (v) { return String(v).replace(/^~/, ''); }).join('');
  try { return JSON.parse(str); } catch (e) { return null; }
}

/** 1つの記録（ドキュメント）を保存する。保存した時刻を返す。 */
function vbSet(path, json) {
  if (typeof path !== 'string' || !/^[A-Za-z0-9_\-.:@+~\/]+$/.test(path)) throw new Error('bad path');
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var s = sheet_();
    var map = map_(), hit = map[path], row = -1, oldParts = 0;
    if (hit && String(s.getRange(hit[1], 1).getValue()) === path) { row = hit[1]; oldParts = hit[2]; }
    else if (hit) { map = rebuildMap_(); hit = map[path]; if (hit) { row = hit[1]; oldParts = hit[2]; } }
    if (row < 0) row = s.getLastRow() + 1;
    var chunks = [];
    for (var k = 0; k < json.length; k += CH) chunks.push('~' + json.slice(k, k + CH));
    if (!chunks.length) chunks.push('~');
    var ts = Date.now();
    s.getRange(row, 1, 1, 3).setValues([[path, ts, chunks.length]]);
    s.getRange(row, 4, 1, chunks.length).setValues([chunks]);
    if (oldParts > chunks.length) s.getRange(row, 4 + chunks.length, 1, oldParts - chunks.length).clearContent();
    map[path] = [ts, row, chunks.length];
    saveMap_(map);
    return ts;
  } finally {
    lock.releaseLock();
  }
}

/**
 * 変更のあった記録だけを返す。変更がなければシートを開かない。
 * wants: [{p: パス, kind: 'doc' か 'col', since: 前回の時刻}]
 */
function vbPull(wants) {
  var map = map_(), hits = [], seen = {};
  (wants || []).forEach(function (w) {
    var depth = w.p.split('/').length;
    for (var p in map) {
      var e = map[p];
      if (e[0] <= (w.since || 0) || seen[p]) continue;
      var ok = w.kind === 'doc' ? p === w.p : (p.indexOf(w.p + '/') === 0 && p.split('/').length === depth + 1);
      if (!ok) continue;
      seen[p] = 1;
      hits.push([p, e]);
    }
  });
  var docs = [];
  if (hits.length) {
    var s = sheet_();
    hits.forEach(function (h) {
      var p = h[0], e = h[1];
      if (String(s.getRange(e[1], 1).getValue()) !== p) { // 行がずれていたら作り直す
        var m2 = rebuildMap_(); e = m2[p]; if (!e) return;
      }
      docs.push({ p: p, ts: e[0], d: readRow_(s, e[1], e[2]) });
    });
  }
  return { now: Date.now(), docs: docs };
}

/** バックアップやCSVを、ドライブの「バレー分析」フォルダに保存してURLを返す。 */
function vbSaveFile(name, content, mime) {
  var f = folder_().createFile(name, content, mime || 'text/plain');
  return f.getUrl();
}
function folder_() {
  var it = DriveApp.getFoldersByName('バレー分析');
  return it.hasNext() ? it.next() : DriveApp.createFolder('バレー分析');
}

/* ---------- A4: 毎日の自動バックアップ（スプレッドシートを丸ごとコピー、14日分） ---------- */
function dailyBackup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var root = folder_(), it = root.getFoldersByName('自動バックアップ');
  var bf = it.hasNext() ? it.next() : root.createFolder('自動バックアップ');
  var name = 'バックアップ_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd_HHmm');
  var f = DriveApp.getFileById(ss.getId()).makeCopy(name, bf);
  var files = [], fi = bf.getFiles();
  while (fi.hasNext()) files.push(fi.next());
  files.sort(function (a, b) { return b.getDateCreated().getTime() - a.getDateCreated().getTime(); });
  for (var i = 14; i < files.length; i++) files[i].setTrashed(true);
  var pr = PropertiesService.getScriptProperties();
  pr.setProperty('lastBackup', Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'));
  pr.setProperty('lastBackupUrl', bf.getUrl());
  return bf.getUrl();
}
function vbBackupStatus() {
  var on = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'dailyBackup'; });
  var pr = PropertiesService.getScriptProperties();
  return { on: on, last: pr.getProperty('lastBackup') || '', url: pr.getProperty('lastBackupUrl') || '' };
}
function vbSetBackup(on) {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'dailyBackup') ScriptApp.deleteTrigger(t); });
  if (on) {
    ScriptApp.newTrigger('dailyBackup').timeBased().everyDays(1).atHour(3).inTimezone('Asia/Tokyo').create();
    dailyBackup();
  }
  return vbBackupStatus();
}
function vbBackupNow() {
  dailyBackup();
  return vbBackupStatus();
}
