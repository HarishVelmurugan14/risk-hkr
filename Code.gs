/**
 * Risk HKR — Google Sheets backend
 * Deploy: Apps Script → Deploy → New deployment → Web app
 *   Execute as: Me
 *   Who has access: Anyone
 * Then paste the /exec URL into index.html — the secret does NOT go in the page.
 *
 * "Anyone" means the URL is reachable by anyone who has it, so the token is
 * what actually gates it. Don't ever put account numbers or credentials in
 * this sheet.
 */

// The sheet this backend writes to. Pinning the id means the script works whether
// it is bound to the sheet or standalone. Not a secret: opening it still needs
// Google permissions, and the web app is gated by the token below.
const SS_ID = '1BJIVdk5yWxSCTyhlizBTW6Kbti4V_SvHSGAu5vh64tU';

function getSpreadsheet_() {
  return SpreadsheetApp.openById(SS_ID);
}

// ── Token gate ───────────────────────────────────────────────────────────────
// Run TEST_setApiToken('your-long-random-secret') once from the Apps Script
// editor. Pass '' to remove the gate. The secret lives in Script Properties —
// never in this file, never in the published page. That is what lets the
// GitHub repo be public without leaking anything.

function checkToken_(token) {
  const stored = PropertiesService.getScriptProperties().getProperty('api_token');
  if (!stored) return true;              // gate not set — allow all, but see gateOpen_
  return String(token || '').trim() === stored;
}

// Reported to the page on every load so an unset gate is loud rather than silent.
function gateOpen_() {
  return !PropertiesService.getScriptProperties().getProperty('api_token');
}

function TEST_setApiToken(token) {
  const props = PropertiesService.getScriptProperties();
  if (token) {
    props.setProperty('api_token', String(token).trim());
    Logger.log('api_token set');
    return { set: true };
  }
  props.deleteProperty('api_token');
  Logger.log('api_token cleared');
  return { set: false };
}

const FIELDS = [
  'id','createdAt','date','ticker','kind','side',
  'hisQty','hisPrice','myPrice','myShares','myValue',
  'hisThesis','myView',
  'decision','skipReason'
];
const NUMS = ['hisQty','hisPrice','myPrice','myShares','myValue'];

function sheet_(name, header) {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
const calls_  = () => sheet_('calls', FIELDS);
const config_ = () => sheet_('config', ['key', 'value']);

function readCalls_() {
  const sh = calls_();
  if (sh.getLastRow() < 2) return [];
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, FIELDS.length).getValues();
  return rows.filter(r => r[0]).map(r => {
    const o = {};
    FIELDS.forEach((f, i) => {
      let v = r[i];
      if (NUMS.indexOf(f) >= 0) v = Number(v) || 0;
      else if (f === 'date' && v instanceof Date) v = Utilities.formatDate(v, 'Asia/Kolkata', 'yyyy-MM-dd');
      else v = v === '' ? '' : String(v);
      o[f] = v;
    });
    return o;
  });
}

function readConfig_() {
  const sh = config_();
  const o = {};
  if (sh.getLastRow() < 2) return o;
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues()
    .forEach(([k, v]) => { if (k) o[String(k)] = Number(v) || 0; });
  return o;
}

function rowOf_(id) {
  const sh = calls_();
  if (sh.getLastRow() < 2) return -1;
  const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}

function upsert_(c) {
  const sh = calls_();
  const row = FIELDS.map(f => c[f] === undefined ? '' : c[f]);
  const at = rowOf_(c.id);
  if (at > 0) sh.getRange(at, 1, 1, FIELDS.length).setValues([row]);
  else sh.appendRow(row);
  return { id: c.id };
}

function delete_(id) {
  const at = rowOf_(id);
  if (at > 0) calls_().deleteRow(at);
  return { id: id };
}

function config_write_(cfg) {
  const sh = config_();
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).clearContent();
  const rows = Object.keys(cfg).map(k => [k, cfg[k]]);
  if (rows.length) sh.getRange(2, 1, rows.length, 2).setValues(rows);
  return cfg;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const req = JSON.parse(e.postData.contents);
    if (!checkToken_(req.token)) return json_({ ok: false, error: 'unauthorized' });

    switch (req.action) {
      case 'load':   return json_({ ok: true, data: { calls: readCalls_(), cfg: readConfig_(), gateOpen: gateOpen_() } });
      case 'upsert': return json_({ ok: true, data: upsert_(req.payload) });
      case 'delete': return json_({ ok: true, data: delete_(req.payload.id) });
      case 'config': return json_({ ok: true, data: config_write_(req.payload) });
      default:       return json_({ ok: false, error: 'unknown action' });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

function doGet() {
  return ContentService.createTextOutput('Risk HKR backend is running. POST only.');
}
