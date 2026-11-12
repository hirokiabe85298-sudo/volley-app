/**
 * バレー試合分析（Google版）サーバー側のプログラム
 * このファイルの中身を Apps Script の「コード.gs」にすべて貼り付けてください。
 * 記録はこのスプレッドシートの「data」シートに保存されます（手で編集しないでください）。
 */
var SHEET = 'data';
var CH = 40000; // 1つのセルに入れる文字数（セルの上限は50,000文字）

// アプリの画面はGitHubから毎回読み込みます（つながらないときは、このプロジェクトの「index」を使います）
var APP_URL = 'https://raw.githubusercontent.com/hirokiabe85298-sudo/volley-app/main/index.html';

function doGet() {
  var out = null;
  try {
    var r = UrlFetchApp.fetch(APP_URL + '?t=' + Math.floor(Date.now() / 60000), { muteHttpExceptions: true });
    if (r.getResponseCode() === 200) {
      var html = r.getContentText('UTF-8');
      if (html.indexOf('id="main"') > 0) out = HtmlService.createHtmlOutput(html);
    }
  } catch (e) {}
  if (!out) out = HtmlService.createHtmlOutputFromFile('index');
  return out
    .setTitle('バレー試合分析')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

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
    var idx = index_(s), row = -1, oldParts = 0;
    for (var i = 0; i < idx.length; i++) if (idx[i][0] === path) { row = i + 2; oldParts = Number(idx[i][2]) || 0; break; }
    if (row < 0) row = s.getLastRow() + 1;
    var chunks = [];
    for (var k = 0; k < json.length; k += CH) chunks.push('~' + json.slice(k, k + CH));
    if (!chunks.length) chunks.push('~');
    var ts = Date.now();
    s.getRange(row, 1, 1, 3).setValues([[path, ts, chunks.length]]);
    s.getRange(row, 4, 1, chunks.length).setValues([chunks]);
    if (oldParts > chunks.length) s.getRange(row, 4 + chunks.length, 1, oldParts - chunks.length).clearContent();
    return ts;
  } finally {
    lock.releaseLock();
  }
}

/**
 * 変更のあった記録だけを返す。
 * wants: [{p: パス, kind: 'doc' か 'col', since: 前回の時刻}]
 */
function vbPull(wants) {
  var s = sheet_();
  var idx = index_(s), docs = [], seen = {};
  (wants || []).forEach(function (w) {
    var depth = w.p.split('/').length;
    for (var i = 0; i < idx.length; i++) {
      var p = String(idx[i][0]), ts = Number(idx[i][1]) || 0;
      if (ts <= (w.since || 0) || seen[p]) continue;
      var hit = w.kind === 'doc' ? p === w.p : (p.indexOf(w.p + '/') === 0 && p.split('/').length === depth + 1);
      if (!hit) continue;
      seen[p] = 1;
      docs.push({ p: p, ts: ts, d: readRow_(s, i + 2, Number(idx[i][2]) || 0) });
    }
  });
  return { now: Date.now(), docs: docs };
}

/** バックアップやCSVを、ドライブの「バレー分析」フォルダに保存してURLを返す。 */
function vbSaveFile(name, content, mime) {
  var it = DriveApp.getFoldersByName('バレー分析');
  var folder = it.hasNext() ? it.next() : DriveApp.createFolder('バレー分析');
  var f = folder.createFile(name, content, mime || 'text/plain');
  return f.getUrl();
}
