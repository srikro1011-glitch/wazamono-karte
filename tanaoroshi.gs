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
var H_MASTER = ['店舗', '場所', '棚', '棚の補足', '棚の写真', '商品', '仕込み品', '材料', '数える単位', '1単位の換算', '換算単位', '仕入れ単位', '入数', '仕入れ価格', '税', '価格の適用日', '円/数える単位', '風袋(g)', '仕入れ先', '重点', '写真', 'りそげん名', 'メモ', '更新日時'];
var H_PRICE  = ['店舗', '棚', '商品', '数える単位', '適用日', '仕入れ単位', '入数', '仕入れ価格', '税', '円/数える単位', '仕入れ先', 'メモ', 'id'];
var H_LINES  = ['店舗', '棚卸し日', '棚', '商品', '仕込み品', '数える単位', '数量(正味)', 'はかり表示', '風袋(g)', '円/数える単位', '金額', '重点', 'countId'];
var H_COUNT  = ['店舗', '棚卸し日', '数えた人', '在庫金額', 'うち重点', 'うち仕込み品', '期間の仕入れ', '期間の売上', '材料費', '材料費率(%)', '締め日時', 'id'];

/* 画面（tanaoroshi.html）は GitHub から読む。GAS 側に HTML を貼らなくてよく、GitHub を更新すれば数分で画面も変わる。
   別の場所に置きたいときはスクリプトプロパティ HTML_URL に URL を入れる。GAS 内に tanaoroshi という HTML ファイルがあればそちらを優先する */
var HTML_URL_DEFAULT = 'https://raw.githubusercontent.com/srikro1011-glitch/wazamono-karte/ccr-d62f97e1-3qs475/tanaoroshi.html';
function loadHtml_() {
  try { return HtmlService.createHtmlOutputFromFile('tanaoroshi').getContent(); } catch (e) {}
  var cache = CacheService.getScriptCache(), key = 'html_v1';
  var hit = cache.get(key);
  if (hit) return hit;
  var url = PropertiesService.getScriptProperties().getProperty('HTML_URL') || HTML_URL_DEFAULT;
  var html = UrlFetchApp.fetch(url, { muteHttpExceptions: true }).getContentText('UTF-8');
  if (html.length < 1000 || html.indexOf('<script>') < 0) throw new Error('画面を読み込めませんでした: ' + url);
  try { cache.put(key, html, 300); } catch (e) {}   // 5分キャッシュ（100KB 以内なら入る）
  return html;
}
function doGet(e) {
  return HtmlService.createHtmlOutput(loadHtml_())
    .setTitle('棚卸し 棚マスタ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
/* GitHub を更新した直後に反映させたいとき：エディタでこの関数を実行する */
function clearHtmlCache() { CacheService.getScriptCache().remove('html_v1'); }
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
    writeStoreSheet_(ss, arg.store, parse_(r.vals[2], { locations: [] }));
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
function shelfLabel_(sh) { return sh.label != null && sh.label !== '' ? String(sh.label) : String(sh.no == null ? '' : sh.no); }
function shelfNo_(loc, sh) { return loc.code ? loc.code + '-' + shelfLabel_(sh) : loc.name + ' ' + shelfLabel_(sh); }
function latestPrice_(it) {
  var ps = (it.prices || []).slice().sort(function (a, b) { return (a.date || '') < (b.date || '') ? -1 : (a.date || '') > (b.date || '') ? 1 : 0; });
  return ps.length ? ps[ps.length - 1] : null;
}
function unitPrice_(p, fallback) {
  if (!p) return fallback == null ? '' : fallback;
  var pack = Number(p.pack) > 0 ? Number(p.pack) : 1;
  return p.price == null ? '' : Math.round(Number(p.price) / pack * 100) / 100;
}
function writeMaster_(ss, store, master) {
  var now = new Date(), rows = [], prices = [], names = {};
  (master.locations || []).forEach(function (loc) { (loc.shelves || []).forEach(function (sh) { (sh.items || []).forEach(function (it) { names[it.id] = it.name; }); }); });
  (master.locations || []).forEach(function (loc) {
    (loc.shelves || []).forEach(function (sh) {
      (sh.items || []).forEach(function (it) {
        var no = shelfNo_(loc, sh), lp = latestPrice_(it);
        var made = (it.madeFrom || []).map(function (id) { return names[id] || ''; }).filter(String).join('、');
        rows.push([store, loc.name, no, sh.name || '', sh.photo && sh.photo.url || '', it.name, it.prep ? '仕込み品' : '', made, it.unit || '', it.conv == null ? '' : it.conv, it.convUnit || '',
          lp ? lp.buyUnit || '' : (it.buyUnit || ''), lp ? lp.pack : (it.pack == null ? '' : it.pack), lp ? lp.price : '', lp ? lp.tax || '' : '', lp ? lp.date || '' : '', unitPrice_(lp, it.price),
          it.tare == null ? '' : it.tare, it.vendor || '', it.star ? '★' : '', it.photo && it.photo.url || '', it.fName || '', it.memo || '', now]);
        (it.prices || []).forEach(function (p) {
          prices.push([store, no, it.name, it.unit || '', p.date || '', p.buyUnit || '', p.pack == null ? '' : p.pack, p.price == null ? '' : p.price, p.tax || '', unitPrice_(p), p.vendor || '', p.memo || '', p.id || '']);
        });
      });
    });
  });
  replaceRows_(sheet_(ss, '品目マスタ', H_MASTER), store, rows);
  replaceRows_(sheet_(ss, '価格履歴', H_PRICE), store, prices);
  writeStoreSheet_(ss, store, master);
}

/* 店別の在庫表：棚の順に商品を並べ、「今回」の列を空けておく。紙に印刷して数えても、画面で数えても同じ順番 */
function writeStoreSheet_(ss, store, master) {
  var name = '在庫表_' + store;
  var sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.setFrozenRows(2); }
  var head1 = [store + ' 在庫表', '', '', '', '', '', '', '', '更新 ' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm'), '', ''];
  var head2 = ['棚', '商品', '仕込み品', '★', '数える単位', '風袋(g)', '円/単位', '前回の数量', '今回の数量', '金額', '写真'];
  var lastQty = {}, cs = [];
  try { var r = dataRow_(ss, store); cs = parse_(r.vals[3], []); } catch (e) {}
  if (cs.length) lastQty = cs[cs.length - 1].qty || {};
  var rows = [];
  (master.locations || []).forEach(function (loc) {
    (loc.shelves || []).forEach(function (s2) {
      (s2.items || []).forEach(function (it) {
        var lp = latestPrice_(it), up = unitPrice_(lp, it.price);
        rows.push([shelfNo_(loc, s2), it.name, it.prep ? '○' : '', it.star ? '★' : '', it.unit || '', it.tare == null ? '' : it.tare, up, lastQty[it.id] == null ? '' : lastQty[it.id], '', '', it.photo && it.photo.url || '']);
      });
    });
  });
  sh.clearContents();
  sh.getRange(1, 1, 2, head2.length).setValues([head1, head2]);
  if (rows.length) {
    sh.getRange(3, 1, rows.length, head2.length).setValues(rows);
    // 金額 ＝ 今回の数量 × 円/単位（今回の数量を手で入れると出る）
    var f = []; for (var i = 0; i < rows.length; i++) { var n = i + 3; f.push(['=IF(I' + n + '="","",I' + n + '*G' + n + ')']); }
    sh.getRange(3, 10, rows.length, 1).setFormulas(f);
    sh.getRange(3, 9, rows.length, 1).setBackground('#fff8e1');
  }
  sh.getRange(1, 1).setFontWeight('bold'); sh.getRange(2, 1, 1, head2.length).setFontWeight('bold');
}

function writeCounts_(ss, store, counts) {
  var heads = [], lines = [];
  counts.forEach(function (c) {
    heads.push([store, c.date, c.by || '', c.total || 0, c.starTotal || 0, c.prepTotal || 0, c.purchases == null ? '' : c.purchases, c.sales == null ? '' : c.sales, c.cogs == null ? '' : c.cogs, c.mr == null ? '' : Math.round(c.mr * 10) / 10, c.closedAt || '', c.id]);
    (c.lines || []).forEach(function (l) {
      lines.push([store, c.date, l.shelf, l.item, l.prep ? '仕込み品' : '', l.unit || '', l.qty, l.raw == null ? '' : l.raw, l.tare == null ? '' : l.tare, l.price == null ? '' : l.price, l.amount == null ? '' : l.amount, l.star ? '★' : '', c.id]);
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
