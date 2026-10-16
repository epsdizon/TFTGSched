const PIN = 'CHANGE-ME'; // Change this to your own PIN before deploying.

/*
  TFTG Roster storage. Paste this into the "TFTG Roster Data" Google Sheet
  (Extensions > Apps Script), then Deploy > New deployment > Web app,
  Execute as: Me, Who has access: Anyone.

  The "Data" tab holds the roster as JSON (one row per item).
  Each month's finished schedule is also written to its own tab, e.g. "2026-11".
*/

function doGet(e) { return handle_(e && e.parameter ? e.parameter : {}); }

function doPost(e) {
  let p = {};
  try { p = JSON.parse(e.postData.contents); } catch (err) {}
  return handle_(p);
}

function handle_(p) {
  if (String(p.pin || '') !== PIN) return out_({ ok: false, error: 'Wrong PIN.' });
  const ss = SpreadsheetApp.getActive();
  const data = ss.getSheetByName('Data') || ss.insertSheet('Data');
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    if (p.action === 'load') {
      return out_({ ok: true, config: read_(data, 'config'), month: read_(data, 'month:' + p.month) });
    }
    if (p.action === 'saveConfig') {
      write_(data, 'config', p.config);
      return out_({ ok: true });
    }
    if (p.action === 'saveMonth') {
      if (!/^\d{4}-\d{2}$/.test(String(p.month))) return out_({ ok: false, error: 'Bad month.' });
      write_(data, 'month:' + p.month, { entries: p.entries || {} });
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
