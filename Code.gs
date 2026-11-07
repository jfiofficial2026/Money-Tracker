/**
 * Money Tracker - Google Sheets Cloud Sync
 * Paste this in: Google Sheet > Extensions > Apps Script
 * 1) Change SECRET_KEY below to your own password (letters/numbers).
 * 2) Deploy > New deployment > Web app
 *    Execute as: Me   |   Who has access: Anyone
 * 3) Copy the Web App URL and paste it in the app: Settings > Cloud sync.
 */
const SECRET_KEY = 'change-this-to-your-own-secret';

function doGet(e) {
  if (!e || !e.parameter || e.parameter.token !== SECRET_KEY) return out_({ ok: false, error: 'Wrong secret key' });
  return out_({ ok: true, data: loadData_() });
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'Bad request' }); }
  if (body.token !== SECRET_KEY) return out_({ ok: false, error: 'Wrong secret key' });
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const merged = merge_(loadData_() || {}, body.data || {});
    if (body.sheets) writeSheets_(body.sheets);
    saveData_(JSON.stringify(merged));
    return out_({ ok: true, updated: new Date().toISOString() });
  } finally { lock.releaseLock(); }
}

/* Keeps the newest version of every record, so 2 devices never overwrite each other */
function merge_(a, b) {
  const arr = (x, y) => {
    const m = {};
    (x || []).forEach(i => m[i.id] = i);
    (y || []).forEach(i => { if (!m[i.id] || (i.u || 0) > (m[i.id].u || 0)) m[i.id] = i; });
    return Object.keys(m).map(k => m[k]);
  };
  const memo = Object.assign({}, a.memo || {});
  Object.keys(b.memo || {}).forEach(k => { if (!memo[k] || (b.memo[k].u || 0) > (memo[k].u || 0)) memo[k] = b.memo[k]; });
  const s = ((b.settings || {}).u || 0) >= ((a.settings || {}).u || 0) ? b.settings : a.settings;
  return { ver: 1, accounts: arr(a.accounts, b.accounts), tx: arr(a.tx, b.tx), debts: arr(a.debts, b.debts), memo: memo, settings: s || {} };
}

/* Readable tabs: Summary, Transactions, Accounts, Debts */
function writeSheets_(sheets) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(sheets).forEach(name => {
    const rows = sheets[name];
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    sh.clearContents();
    if (rows && rows.length) {
      sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
      sh.getRange(1, 1, 1, rows[0].length).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
  });
}

/* Full app data is stored in a hidden tab "_data" (do not edit it) */
function saveData_(str) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('_data') || ss.insertSheet('_data');
  sh.clearContents();
  const chunks = [];
  for (let i = 0; i < str.length; i += 40000) chunks.push(['~' + str.slice(i, i + 40000)]);
  if (chunks.length) {
    const r = sh.getRange(1, 1, chunks.length, 1);
    r.setNumberFormat('@');
    r.setValues(chunks);
  }
  try { sh.hideSheet(); } catch (err) {}
}

function loadData_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('_data');
  if (!sh || sh.getLastRow() < 1) return null;
  const txt = sh.getRange(1, 1, sh.getLastRow(), 1).getValues().map(r => String(r[0]).slice(1)).join('');
  try { return JSON.parse(txt); } catch (err) { return null; }
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
