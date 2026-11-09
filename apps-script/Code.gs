/*
  TFTG Roster storage for the TFTGSched website.

  One-time setup (about 3 minutes, on a computer):
  1. Open the "TFTG Roster Data" Google Sheet in the Schedtftg folder.
  2. Extensions > Apps Script. Delete everything in Code.gs and paste this whole file. Click Save.
  3. Deploy > New deployment > gear icon > Web app.
     Execute as: Me.  Who has access: Anyone.  Click Deploy.
  4. Click Authorize access, pick your Google account. If Google says it hasn't verified the app,
     click Advanced > Go to (unsafe) > Allow. It is your own script, working on your own sheet.
  5. Copy the Web app URL (it ends in /exec).

  The "Data" tab holds the roster as JSON, one row per item: key, data, saved at.
  Each month's finished schedule is also written to its own tab, e.g. "2026-11".
*/

function doGet(e) { return handle_(e && e.parameter ? e.parameter : {}); }

function doPost(e) {
  let p = {};
  try { p = JSON.parse(e.postData.contents); } catch (err) {}
  return handle_(p);
}

function handle_(p) {
  const ss = SpreadsheetApp.getActive();
  const data = ss.getSheetByName('Data') || ss.insertSheet('Data');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (p.action === 'ping') return out_({ ok: true });
    if (p.action === 'loadAll') return out_({ ok: true, config: read_(data, 'config'), months: readMonths_(data) });
    if (p.action === 'load') return out_({ ok: true, config: read_(data, 'config'), month: read_(data, 'month:' + p.month) });
    if (p.action === 'saveConfig') {
      if (!p.config || !Array.isArray(p.config.stores)) return out_({ ok: false, error: 'Bad roster.' });
      write_(data, 'config', p.config);
      return out_({ ok: true });
    }
    if (p.action === 'saveMonth') {
      if (!/^\d{4}-\d{2}$/.test(String(p.month))) return out_({ ok: false, error: 'Bad month.' });
      write_(data, 'month:' + p.month, { entries: p.entries || {}, site: true });
      if (Array.isArray(p.table) && p.table.length) writeTab_(ss, String(p.month), p.table);
      return out_({ ok: true });
    }
    return out_({ ok: false, error: 'Unknown action.' });
  } finally {
    lock.releaseLock();
  }
}

function findRow_(sh, key) {
  const last = sh.getLastRow();
  if (!last) return 0;
  const keys = sh.getRange(1, 1, last, 1).getValues();
  for (let i = 0; i < keys.length; i++) if (String(keys[i][0]) === key) return i + 1;
  return 0;
}

function read_(sh, key) {
  const r = findRow_(sh, key);
  if (!r) return null;
  try { return JSON.parse(sh.getRange(r, 2).getValue()); } catch (e) { return null; }
}

function readMonths_(sh) {
  const out = {}, last = sh.getLastRow();
  if (!last) return out;
  sh.getRange(1, 1, last, 2).getValues().forEach(function (r) {
    const k = String(r[0]);
    if (k.indexOf('month:') !== 0) return;
    try { out[k.slice(6)] = JSON.parse(r[1]); } catch (e) {}
  });
  return out;
}

function write_(sh, key, obj) {
  let r = findRow_(sh, key);
  if (!r) r = sh.getLastRow() + 1;
  const range = sh.getRange(r, 1, 1, 3);
  range.setNumberFormat('@');
  range.setValues([[key, JSON.stringify(obj), new Date().toISOString()]]);
}

function writeTab_(ss, name, table) {
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.clear();
  const width = Math.max.apply(null, table.map(function (r) { return r.length; }));
  const rows = table.map(function (r) {
    const a = r.map(function (v) { return v == null ? '' : String(v); });
    while (a.length < width) a.push('');
    return a;
  });
  sh.getRange(1, 1, rows.length, width).setNumberFormat('@').setValues(rows);
  sh.setFrozenColumns(1);
  sh.setColumnWidth(1, 170);
  if (width > 1) sh.setColumnWidths(2, width - 1, 42);
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
