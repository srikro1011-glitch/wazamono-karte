/* =========================================================
   棚卸し 棚マスタ — Apps Script 側（tanaoroshi.html と対）
   ・doGet  : tanaoroshi.html をそのまま表示（ウェブアプリとして開く）
   ・doPost : 別サイトとして開いた tanaoroshi.html からの fetch を受ける
   ・rpc    : google.script.run からの呼び出し口（doPost も同じ関数を通す）
   ・保存先 : スクリプトプロパティ SHEET_ID のスプレッドシート。無ければ初回に作って記録する
   ・写真   : Drive の「棚卸し写真」フォルダ。シートにはURLを書く
   ・認証   : スクリプトプロパティ PW を設定した場合だけ、パスワード一致を要求する
   ========================================================= */
var SHEET_TITLE = '棚卸し 棚マスタ';
var PHOTO_FOLDER = '棚卸し写真';
var H_DATA   = ['店舗', 'masterRev', 'master', 'counts', '更新日時', '更新者'];
var H_MASTER = ['店舗', '場所記号', '場所', '棚番号', '棚名', '棚の写真', '品目', '数える単位', '1単位の換算', '換算単位', '仕入れ値(円/単位)', '風袋(g)', '仕入れ先', '重点', '写真', 'メモ', '更新日時'];
var H_COUNT  = ['店舗', '棚卸し日', '数えた人', '在庫金額', 'うち重点', '期間の仕入れ', '期間の売上', '材料費', '材料費率(%)', '締め日時', 'id'];
var H_LINES  = ['店舗', '棚卸し日', '棚番号', '品目', '数える単位', '数量(正味)', 'はかり表示', '風袋(g)', '仕入れ値(円/単位)', '金額', '重点', 'countId'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('tanaoroshi')
    .setTitle('棚卸し 棚マスタ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function doPost(e) {
  var body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) {}
  return ContentService.createTextOutput(JSON.stringify(rpc(body.fn, body.pw, body.arg)))
    .setMimeType(ContentService.MimeType.JSON);
}

function rpc(fn, pw, arg) {
  try {
    var need = PropertiesService.getScriptProperties().getProperty('PW') || '';
    if (need && pw !== need) return { ok: false, error: 'auth' };
    arg = arg || {};
    var fns = { version: version, pull: pull, pushMaster: pushMaster, pushCounts: pushCounts, deleteCount: deleteCount, uploadPhoto: uploadPhoto };
    if (!fns[fn]) return { ok: false, error: 'unknown fn: ' + fn };
    return { ok: true, data: fns[fn](arg) };
  } catch (err) {
    return { ok: false, error: String(err && err.message || err) };
  }
}

/* ---------- スプレッドシート ---------- */
function ss_() {
  var p = PropertiesService.getScriptProperties();
  var id = p.getProperty('SHEET_ID'), ss = null;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) {} }
  if (!ss) { ss = SpreadsheetApp.create(SHEET_TITLE); p.setProperty('SHEET_ID', ss.getId()); }
  return ss;
}
function sheet_(ss, name, header) {
  var sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(header); sh.setFrozenRows(1); }
  return sh;
}
function dataRow_(ss, store) {
  var sh = sheet_(ss, '_data', H_DATA);
  var vals = sh.getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) if (vals[i][0] === store) return { sh: sh, row: i + 1, vals: vals[i] };
  sh.appendRow([store, 0, JSON.stringify({ note: '', locations: [] }), '[]', new Date(), '']);
  return { sh: sh, row: sh.getLastRow(), vals: [store, 0, JSON.stringify({ note: '', locations: [] }), '[]', new Date(), ''] };
}
function parse_(s, d) { try { return s ? JSON.parse(s) : d; } catch (e) { return d; } }

function version(arg) {
  var r = dataRow_(ss_(), arg.store);
  return { masterRev: Number(r.vals[1]) || 0, countsRev: parse_(r.vals[3], []).length };
}
function pull(arg) {
  var ss = ss_(), r = dataRow_(ss, arg.store);
  return { masterRev: Number(r.vals[1]) || 0, master: parse_(r.vals[2], { note: '', locations: [] }), counts: parse_(r.vals[3], []), sheetUrl: ss.getUrl() };
}
function pushMaster(arg) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var ss = ss_(), r = dataRow_(ss, arg.store);
    var rev = Number(r.vals[1]) || 0;
    if (Number(arg.base) !== rev) return { conflict: true, masterRev: rev, master: parse_(r.vals[2], { note: '', locations: [] }) };
    rev++;
    r.sh.getRange(r.row, 2, 1, 5).setValues([[rev, JSON.stringify(arg.master || {}), r.vals[3], new Date(), arg.editor || '']]);
    writeMaster_(ss, arg.store, arg.master || {});
    return { masterRev: rev };
  } finally { lock.releaseLock(); }
}
function pushCounts(arg) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var ss = ss_(), r = dataRow_(ss, arg.store);
    var counts = parse_(r.vals[3], []);
    (arg.counts || []).forEach(function (c) {
      delete c.synced;
      var i = -1; counts.forEach(function (x, k) { if (x.id === c.id) i = k; });
      if (i >= 0) counts[i] = c; else counts.push(c);
    });
    counts.sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });
    r.sh.getRange(r.row, 4, 1, 3).setValues([[JSON.stringify(counts), new Date(), arg.editor || '']]);
    writeCounts_(ss, arg.store, counts);
    return { n: counts.length };
  } finally { lock.releaseLock(); }
}
function deleteCount(arg) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var ss = ss_(), r = dataRow_(ss, arg.store);
    var counts = parse_(r.vals[3], []).filter(function (c) { return c.id !== arg.id; });
    r.sh.getRange(r.row, 4, 1, 2).setValues([[JSON.stringify(counts), new Date()]]);
    writeCounts_(ss, arg.store, counts);
    return true;
  } finally { lock.releaseLock(); }
}

/* 人が読む表：その店の行を入れ替える */
function replaceRows_(sh, store, rows) {
  var vals = sh.getDataRange().getValues();
  var keep = vals.slice(1).filter(function (v) { return v[0] !== store; });
  var out = keep.concat(rows);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent();
  if (out.length) sh.getRange(2, 1, out.length, out[0].length).setValues(out);
}
function shelfNo_(loc, sh) { var n = String(sh.no || 0); while (n.length < 2) n = '0' + n; return loc.code + '-' + n; }
function writeMaster_(ss, store, master) {
  var now = new Date(), rows = [];
  (master.locations || []).forEach(function (loc) {
    (loc.shelves || []).forEach(function (sh) {
      (sh.items || []).forEach(function (it) {
        rows.push([store, loc.code, loc.name, shelfNo_(loc, sh), sh.name || '', sh.photo && sh.photo.url || '', it.name, it.unit || '', it.conv == null ? '' : it.conv, it.convUnit || '', it.price == null ? '' : it.price, it.tare == null ? '' : it.tare, it.vendor || '', it.star ? '★' : '', it.photo && it.photo.url || '', it.memo || '', now]);
      });
    });
  });
  replaceRows_(sheet_(ss, '品目マスタ', H_MASTER), store, rows);
}
function writeCounts_(ss, store, counts) {
  var heads = [], lines = [];
  counts.forEach(function (c) {
    heads.push([store, c.date, c.by || '', c.total || 0, c.starTotal || 0, c.purchases == null ? '' : c.purchases, c.sales == null ? '' : c.sales, c.cogs == null ? '' : c.cogs, c.mr == null ? '' : Math.round(c.mr * 10) / 10, c.closedAt || '', c.id]);
    (c.lines || []).forEach(function (l) {
      lines.push([store, c.date, l.shelf, l.item, l.unit || '', l.qty, l.raw == null ? '' : l.raw, l.tare == null ? '' : l.tare, l.price == null ? '' : l.price, l.amount == null ? '' : l.amount, l.star ? '★' : '', c.id]);
    });
  });
  replaceRows_(sheet_(ss, '棚卸し', H_COUNT), store, heads);
  replaceRows_(sheet_(ss, '棚卸し明細', H_LINES), store, lines);
}

/* ---------- 写真 ---------- */
function folder_() {
  var it = DriveApp.getFoldersByName(PHOTO_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(PHOTO_FOLDER);
}
function uploadPhoto(arg) {
  var m = String(arg.dataUrl || '').match(/^data:(image\/\w+);base64,(.+)$/);
  if (!m) throw new Error('画像の形式が違います');
  var blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], arg.name || ('photo_' + Date.now() + '.jpg'));
  var file = folder_().createFile(blob);
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  return { url: file.getUrl(), id: file.getId() };
}
