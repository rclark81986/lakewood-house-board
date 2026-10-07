/**
 * Lakewood house board - Google Sheets backend (container-bound Apps Script).
 *
 * Paste this whole file into the "House Board Data" Sheet (Drive > Cozyla folder):
 *   Extensions > Apps Script > replace Code.gs > Save.
 * Then set TOKEN below, and Deploy > New deployment > Web app
 *   (Execute as: Me, Who has access: Anyone).
 *
 * The board (GitHub Pages) reads   GET  <web app URL>?k=TOKEN
 * and writes                        POST <web app URL>  body: {"k":TOKEN,"ops":[...]}
 * (POST is sent as text/plain so the browser skips the CORS preflight.)
 *
 * No secrets live in the public GitHub repo: the web app URL and token are only
 * in the Fully Kiosk Start URL (?api=...&k=...).
 */

// ---- Settings -------------------------------------------------------------
// Shared secret. Put the token here (or set a Script Property named TOKEN,
// which wins if present: Project Settings > Script Properties).
var TOKEN = 'PASTE_TOKEN_HERE';

// Drive folder that holds the recipe category folders.
var RECIPES_FOLDER_ID = '162vi0Rb-sBXwOBSCZJLL7YnJnSz4iLZj';

// Subfolders that never show on the board (matched by name, case-insensitive, or by ID).
var HIDDEN_FOLDER_NAMES = ['originals', 'pdfs'];
var HIDDEN_FOLDER_IDS = ['1Z4rNDI-PK_6KufjGziI3_ARdupQ2Mp1S', '1tTJylP7p-aT8BGVwd72m6fJeCsY2hW84'];

// Category order on the board. Any new folder you add shows up after these, A-Z.
var CATEGORY_ORDER = ['Baked Goods', 'Main Dishes', 'Desserts', 'Snacks', 'Preserves & Pickles', 'Cured & Smoked Meats'];

// How long (seconds) the recipe folder listing is cached. New PDFs appear within this time.
var RECIPE_CACHE_SECONDS = 120;

// ---- Sheet layout ---------------------------------------------------------
var TABS = {
  brew:   { name: 'Cold Brew',     headers: ['Key', 'Value', 'What it is'] },
  meat:   { name: 'Meat Freezer',  headers: ['ID', 'Category', 'Item', 'Qty', 'Note', 'Added'] },
  other:  { name: 'Other Freezer', headers: ['ID', 'Item', 'Qty', 'Note', 'Added'] },
  notes:  { name: 'Notes',         headers: ['Notes'] },
  links:  { name: 'Recipe Links',  headers: ['ID', 'Category', 'Name', 'URL', 'Note'] }
};
var BREW_ROWS = [
  ['started', '', 'Date the current batch started (YYYY-MM-DD). Blank = no batch.'],
  ['gallons', 5, 'Batch size in gallons'],
  ['pintsPerDay', 2, 'How many pints you drink per day']
];
var MEAT_CATEGORIES = ['beef', 'chicken', 'pork', 'elk', 'pheasant'];

// ---- Web app entry points -------------------------------------------------
function doGet(e) {
  var p = (e && e.parameter) || {};
  if (!authorized_(p.k)) return json_({ ok: false, error: 'unauthorized' });
  try {
    var withRecipes = p.recipes !== '0';
    return json_(snapshot_(withRecipes));
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function doPost(e) {
  var body = {};
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); }
  catch (err) { return json_({ ok: false, error: 'bad json' }); }
  if (!authorized_(body.k)) return json_({ ok: false, error: 'unauthorized' });

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return json_({ ok: false, error: 'busy, try again' });
  var results = [];
  try {
    ensureSheets_();
    var cache = CacheService.getScriptCache();
    (body.ops || []).forEach(function (op) {
      if (op && op.opId && cache.get('op_' + op.opId)) { results.push('dup'); return; }
      applyOp_(op || {});
      if (op && op.opId) cache.put('op_' + op.opId, '1', 21600);
      results.push('ok');
    });
    SpreadsheetApp.flush();
  } catch (err) {
    lock.releaseLock();
    return json_({ ok: false, error: String(err && err.message || err), results: results });
  }
  lock.releaseLock();
  var snap = snapshot_(body.recipes === true);
  snap.results = results;
  return json_(snap);
}

/** Run once from the editor (optional) to create any missing tabs and check access. */
function setup() {
  ensureSheets_();
  Logger.log(JSON.stringify(snapshot_(true)).slice(0, 2000));
}

// ---- Auth / output --------------------------------------------------------
function token_() {
  var prop = PropertiesService.getScriptProperties().getProperty('TOKEN');
  return String(prop || TOKEN || '');
}
function authorized_(k) {
  var t = token_();
  if (!t || t === 'PASTE_TOKEN_HERE' || t.length < 12) return false;  // refuse until a real token is set
  k = String(k || '');
  if (k.length !== t.length) return false;
  var diff = 0;
  for (var i = 0; i < t.length; i++) diff |= t.charCodeAt(i) ^ k.charCodeAt(i);
  return diff === 0;
}
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ---- Read -----------------------------------------------------------------
function snapshot_(withRecipes) {
  var lock = LockService.getScriptLock();
  var locked = lock.tryLock(20000);
  var data;
  try {
    ensureSheets_();
    data = readData_();
  } finally {
    if (locked) lock.releaseLock();
  }
  var out = { ok: true, serverTime: new Date().toISOString(), data: data };
  if (withRecipes) out.recipeFolders = recipeFolders_();
  return out;
}

function readData_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = ss.getSpreadsheetTimeZone();

  // Cold brew (key/value)
  var brew = { started: '', gallons: 5, pintsPerDay: 2 };
  rows_(TABS.brew.name).forEach(function (r) {
    var key = String(r[0] || '').trim();
    if (key === 'started') brew.started = dateStr_(r[1], tz);
    if (key === 'gallons') brew.gallons = num_(r[1], 5);
    if (key === 'pintsPerDay') brew.pintsPerDay = num_(r[1], 2);
  });

  var freezer = itemRows_(TABS.meat.name, true, tz);
  var otherFreezer = itemRows_(TABS.other.name, false, tz);

  var notesSheet = ss.getSheetByName(TABS.notes.name);
  var notes = String(notesSheet.getRange(2, 1).getValue() || '');

  var linkSheet = ss.getSheetByName(TABS.links.name);
  var recipeLinks = [];
  withIds_(linkSheet).forEach(function (r) {
    if (!r[2] && !r[3]) return;
    recipeLinks.push({ id: String(r[0]), category: String(r[1] || 'Other'), name: String(r[2] || r[3]), url: String(r[3] || ''), note: String(r[4] || '') });
  });

  return { brew: brew, freezer: freezer, otherFreezer: otherFreezer, notes: notes, recipeLinks: recipeLinks };
}

function itemRows_(tabName, isMeat, tz) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tabName);
  var out = [];
  withIds_(sh).forEach(function (r) {
    var name = isMeat ? r[2] : r[1];
    if (!String(name || '').trim()) return;
    if (isMeat) {
      out.push({ id: String(r[0]), category: String(r[1] || 'beef').trim().toLowerCase(), name: String(name), qty: num_(r[3], 1), note: String(r[4] || ''), added: dateStr_(r[5], tz) });
    } else {
      out.push({ id: String(r[0]), category: 'other', name: String(name), qty: num_(r[2], 1), note: String(r[3] || ''), added: dateStr_(r[4], tz) });
    }
  });
  return out;
}

/** Data rows (no header). Rows typed into the Sheet by hand get an ID filled in. */
function withIds_(sh) {
  var last = sh.getLastRow();
  if (last < 2) return [];
  var width = sh.getLastColumn();
  var range = sh.getRange(2, 1, last - 1, width);
  var values = range.getValues();
  var changed = false;
  values.forEach(function (r) {
    var hasContent = r.slice(1).some(function (v) { return String(v).trim() !== ''; });
    if (hasContent && !String(r[0]).trim()) { r[0] = uid_(); changed = true; }
  });
  if (changed) sh.getRange(2, 1, values.length, 1).setValues(values.map(function (r) { return [r[0]]; }));
  return values;
}

function rows_(tabName) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tabName);
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, Math.max(2, sh.getLastColumn())).getValues();
}

// ---- Recipes (Drive) ------------------------------------------------------
function recipeFolders_() {
  var cache = CacheService.getScriptCache();
  var hit = cache.get('recipeFolders');
  if (hit) { try { return JSON.parse(hit); } catch (e) {} }

  var folders = [];
  var it = DriveApp.getFolderById(RECIPES_FOLDER_ID).getFolders();
  while (it.hasNext()) {
    var f = it.next();
    var name = f.getName();
    if (HIDDEN_FOLDER_IDS.indexOf(f.getId()) >= 0) continue;
    if (HIDDEN_FOLDER_NAMES.indexOf(name.trim().toLowerCase()) >= 0) continue;
    if (f.isTrashed && f.isTrashed()) continue;
    var files = [];
    var fit = f.getFiles();
    while (fit.hasNext()) {
      var file = fit.next();
      if (file.isTrashed && file.isTrashed()) continue;
      var mime = file.getMimeType();
      var id = file.getId();
      files.push({
        id: id,
        name: file.getName().replace(/\.(pdf|docx?|txt)$/i, ''),
        mimeType: mime,
        url: viewUrl_(id, mime),
        updated: file.getLastUpdated().toISOString()
      });
    }
    files.sort(function (a, b) { return a.name.localeCompare(b.name); });
    folders.push({ id: f.getId(), name: name, url: 'https://drive.google.com/drive/folders/' + f.getId(), files: files });
  }
  folders.sort(function (a, b) {
    var ia = CATEGORY_ORDER.indexOf(a.name), ib = CATEGORY_ORDER.indexOf(b.name);
    if (ia < 0) ia = 999; if (ib < 0) ib = 999;
    return ia !== ib ? ia - ib : a.name.localeCompare(b.name);
  });
  try { cache.put('recipeFolders', JSON.stringify(folders), RECIPE_CACHE_SECONDS); } catch (e) {}
  return folders;
}

function viewUrl_(id, mime) {
  if (mime === 'application/vnd.google-apps.document') return 'https://docs.google.com/document/d/' + id + '/view';
  if (mime === 'application/vnd.google-apps.spreadsheet') return 'https://docs.google.com/spreadsheets/d/' + id + '/view';
  return 'https://drive.google.com/file/d/' + id + '/view';
}

// ---- Write ----------------------------------------------------------------
function applyOp_(op) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  switch (op.op) {
    case 'brew':
      setBrew_(op);
      break;
    case 'notes':
      ss.getSheetByName(TABS.notes.name).getRange(2, 1).setValue(String(op.text || ''));
      break;
    case 'addMeat':
      addItem_(TABS.meat.name, op.item || {}, true);
      break;
    case 'addOther':
      addItem_(TABS.other.name, op.item || {}, false);
      break;
    case 'qty':
      changeQty_(op.list === 'other' ? TABS.other.name : TABS.meat.name, op.list === 'other' ? 3 : 4, String(op.id), Number(op.delta) || 0);
      break;
    case 'addLink':
      var l = op.link || {};
      if (!l.url) break;
      if (findRow_(ss.getSheetByName(TABS.links.name), String(l.id || '')) > 0) break;
      appendRow_(ss.getSheetByName(TABS.links.name), [String(l.id || uid_()), String(l.category || 'Other'), String(l.name || l.url), String(l.url), String(l.note || 'Added')]);
      break;
    case 'removeLink':
      var lsh = ss.getSheetByName(TABS.links.name);
      var row = findRow_(lsh, String(op.id));
      if (row > 0) lsh.deleteRow(row);
      break;
    case 'replaceAll':   // used by "Import backup"
      replaceAll_(op.state || {});
      break;
    case 'mergeLocal':   // first sync from a device that already had data saved locally
      mergeLocal_(op.state || {});
      break;
    default:
      throw new Error('unknown op: ' + op.op);
  }
}

function setBrew_(b) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TABS.brew.name);
  var vals = rows_(TABS.brew.name);
  ['started', 'gallons', 'pintsPerDay'].forEach(function (key) {
    if (!(key in b)) return;
    var value = key === 'started' ? String(b.started || '') : (Number(b[key]) || (key === 'gallons' ? 5 : 2));
    var idx = -1;
    for (var i = 0; i < vals.length; i++) if (String(vals[i][0]).trim() === key) { idx = i; break; }
    if (idx < 0) {
      var def = BREW_ROWS.filter(function (r) { return r[0] === key; })[0];
      appendRow_(sh, [key, value, def ? def[2] : '']);
      vals = rows_(TABS.brew.name);
    } else {
      sh.getRange(idx + 2, 2).setValue(value);
    }
  });
}

function addItem_(tabName, item, isMeat) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tabName);
  var id = String(item.id || uid_());
  if (findRow_(sh, id) > 0) return;  // already added (retry)
  var qty = Math.max(1, Number(item.qty) || 1);
  var added = String(item.added || today_());
  var row = isMeat
    ? [id, String(item.category || 'beef').toLowerCase(), String(item.name || ''), qty, String(item.note || ''), added]
    : [id, String(item.name || ''), qty, String(item.note || ''), added];
  insertTopRow_(sh, row);
}

function changeQty_(tabName, qtyCol, id, delta) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tabName);
  var row = findRow_(sh, id);
  if (row < 1) return;
  var cell = sh.getRange(row, qtyCol);
  var next = (Number(cell.getValue()) || 0) + delta;
  if (next <= 0) sh.deleteRow(row); else cell.setValue(next);
}

function replaceAll_(s) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (s.brew) setBrew_(s.brew);
  ss.getSheetByName(TABS.notes.name).getRange(2, 1).setValue(String(s.notes || ''));
  writeTable_(ss.getSheetByName(TABS.meat.name), (s.freezer || []).filter(function (i) { return i && i.name; }).map(function (i) {
    return [String(i.id || uid_()), String(i.category || 'beef').toLowerCase(), String(i.name), Math.max(1, Number(i.qty) || 1), String(i.note || ''), String(i.added || '')];
  }));
  writeTable_(ss.getSheetByName(TABS.other.name), (s.otherFreezer || []).filter(function (i) { return i && i.name; }).map(function (i) {
    return [String(i.id || uid_()), String(i.name), Math.max(1, Number(i.qty) || 1), String(i.note || ''), String(i.added || '')];
  }));
  writeTable_(ss.getSheetByName(TABS.links.name), (s.recipeLinks || []).filter(function (l) { return l && l.url; }).map(function (l) {
    return [String(l.id || uid_()), String(l.category || 'Other'), String(l.name || l.url), String(l.url), String(l.note || '')];
  }));
}

function mergeLocal_(s) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var current = readData_();
  if (s.brew && s.brew.started && !current.brew.started) setBrew_(s.brew);
  var localNotes = String(s.notes || '').trim();
  if (localNotes && current.notes.indexOf(localNotes) < 0) {
    ss.getSheetByName(TABS.notes.name).getRange(2, 1).setValue(current.notes.trim() ? current.notes.replace(/\s+$/, '') + '\n\n' + localNotes : localNotes);
  }
  (s.freezer || []).slice().reverse().forEach(function (i) { if (i && i.name) addItem_(TABS.meat.name, i, true); });
  (s.otherFreezer || []).slice().reverse().forEach(function (i) { if (i && i.name) addItem_(TABS.other.name, i, false); });
  var urls = current.recipeLinks.map(function (l) { return l.url; });
  (s.recipeLinks || []).forEach(function (l) {
    if (!l || !l.url || urls.indexOf(l.url) >= 0) return;
    applyOp_({ op: 'addLink', link: l });
    urls.push(l.url);
  });
}

// ---- Sheet helpers --------------------------------------------------------
function ensureSheets_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(TABS).forEach(function (k) {
    var t = TABS[k];
    var sh = ss.getSheetByName(t.name);
    if (!sh) {
      sh = ss.insertSheet(t.name);
      sh.getRange(1, 1, 1, t.headers.length).setValues([t.headers]).setFontWeight('bold');
      sh.setFrozenRows(1);
      if (k === 'brew') sh.getRange(2, 1, BREW_ROWS.length, 3).setValues(BREW_ROWS);
      if (k === 'notes') sh.getRange(2, 1).setValue('');
    }
  });
}

function findRow_(sh, id) {
  if (!id) return -1;
  var last = sh.getLastRow();
  if (last < 2) return -1;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === id) return i + 2;
  return -1;
}

function insertTopRow_(sh, row) {
  // Newest first, like the board.
  if (sh.getLastRow() >= 2) sh.insertRowBefore(2);
  sh.getRange(2, 1, 1, row.length).setFontWeight('normal').setBackground(null).setNumberFormat('@').setValues([row.map(String)]);
  // keep Qty numeric
  var qtyCol = row.length === 6 ? 4 : 3;
  sh.getRange(2, qtyCol).setNumberFormat('0').setValue(Number(row[qtyCol - 1]));
}

function appendRow_(sh, row) {
  var r = sh.getLastRow() + 1;
  sh.getRange(r, 1, 1, row.length).setValues([row]);
}

function writeTable_(sh, rows) {
  var last = sh.getLastRow();
  if (last >= 2) sh.getRange(2, 1, last - 1, sh.getLastColumn()).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

function dateStr_(v, tz) {
  if (Object.prototype.toString.call(v) === '[object Date]') return isNaN(v.getTime()) ? '' : Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  var s = String(v || '').trim();
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);  // 10/7/2026 typed as text
  if (m) return m[3] + '-' + ('0' + m[1]).slice(-2) + '-' + ('0' + m[2]).slice(-2);
  return s;
}
function num_(v, dflt) { var n = Number(v); return isFinite(n) && String(v).trim() !== '' ? n : dflt; }
function uid_() { return Math.random().toString(36).slice(2, 9); }
function today_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
}
