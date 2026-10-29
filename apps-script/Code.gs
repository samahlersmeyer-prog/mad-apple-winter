/**
 * Mad Apple Winter Workouts — Google Sheet backend (v2, general purpose)
 *
 * This version is deliberately generic: it reads every tab and saves rows by key.
 * New features, stats and columns are added from the website and the sheet,
 * so this script should not need to change again.
 *
 * Deploy: Extensions → Apps Script → paste → Save →
 *   Deploy → Manage deployments → pencil → Version: New version → Deploy
 *   (first time only: Deploy → New deployment → Web app, Execute as Me, Access Anyone)
 *
 * Config tab (column A = setting, column B = value):
 *   Coach PIN          full access
 *   Player code        players: read + the write rules below
 *   Player hidden tabs comma list of tabs players can't see        (default: Notes)
 *   Player write rules Tab:mode, comma list                         (default: Tests:fill,Bullpens:create)
 *       modes: fill = only empty cells, create = new rows only, merge = non-blank values, overwrite = all given values
 */

const SS = SpreadsheetApp.getActiveSpreadsheet();
const TZ = 'America/Indiana/Indianapolis';

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function config() {
  const v = SS.getSheetByName('Config').getDataRange().getValues();
  const c = {};
  v.forEach(r => { if (r[0] !== '') c[String(r[0]).trim()] = String(r[1]).trim(); });
  return c;
}

function roleFor(pin, c) {
  pin = String(pin || '').trim();
  if (!pin) return null;
  if (c['Coach PIN'] && pin === c['Coach PIN']) return 'coach';
  if (c['Player code'] && pin === c['Player code']) return 'player';
  return null;
}

function list(s, dflt) { return String(s === undefined || s === '' ? dflt : s).split(',').map(x => x.trim()).filter(Boolean); }

function norm(v) {
  return v instanceof Date ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd') : String(v).trim();
}

function keyName(tab) { return tab.toLowerCase().replace(/[^a-z0-9]+/g, '_').replace(/^_|_$/g, ''); }

function rows(sh) {
  const v = sh.getDataRange().getValues();
  const head = v.shift() || [];
  return v.filter(r => r.some(c => c !== '')).map(r => {
    const o = {};
    head.forEach((h, i) => { if (h !== '') o[h] = r[i] instanceof Date ? norm(r[i]) : r[i]; });
    return o;
  });
}

function allData(role, c) {
  const hidden = role === 'player' ? list(c['Player hidden tabs'], 'Notes') : [];
  const data = {};
  SS.getSheets().forEach(sh => {
    const n = sh.getName();
    if (n === 'Config' || hidden.indexOf(n) >= 0) return;
    data[keyName(n)] = rows(sh);
  });
  return { ok: true, role: role, data: data };
}

function doGet(e) {
  const c = config();
  const role = roleFor((e.parameter || {}).pin, c);
  if (!role) return json({ ok: false, error: 'pin' });
  return json(allData(role, c));
}

function headers(sh) {
  if (sh.getLastColumn() < 1) return [];
  return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
}

function playerName(num) {
  const sh = SS.getSheetByName('Roster');
  const r = rows(sh).find(p => String(p.Number) === String(num));
  return r ? (r.First + ' ' + r.Last) : '';
}

// Save one row. Returns null on success or an error code.
function save(op, role, c) {
  const name = String(op.sheet || '');
  if (!name || name === 'Config') return 'bad_sheet';
  let sh = SS.getSheetByName(name);
  if (!sh) {
    if (role !== 'coach') return 'bad_sheet';
    sh = SS.insertSheet(name);           // coaches can start a new tracking tab from the site
  }
  let mode = op.mode || 'merge';
  if (role === 'player') {
    const rules = {};
    list(c['Player write rules'], 'Tests:fill,Bullpens:create').forEach(x => { const p = x.split(':'); rules[p[0].trim()] = (p[1] || 'fill').trim(); });
    if (!rules[name]) return 'coach_only';
    mode = rules[name];
  }
  const keys = Array.isArray(op.keys) && op.keys.length ? op.keys : null;
  const rec = Object.assign({}, op.rec || {});
  if (!keys) return 'bad_request';

  let head = headers(sh);
  if (rec.Number !== undefined && rec.Number !== '') {
    if (head.indexOf('Player') >= 0 && rec.Player === undefined) rec.Player = playerName(rec.Number);
    if (head.indexOf('Pitcher') >= 0 && name === 'Bullpens' && rec.Pitcher === undefined) rec.Pitcher = playerName(rec.Number);
  }
  // New fields become new columns (coach only), placed before "Logged by"/"Updated" when present.
  const missing = Object.keys(rec).concat(keys).filter((k, i, a) => a.indexOf(k) === i && head.indexOf(k) < 0);
  if (missing.length) {
    if (role !== 'coach') return 'coach_only';
    missing.forEach(k => {
      const tail = ['Logged by', 'Updated'].map(t => head.indexOf(t)).filter(i => i >= 0);
      const at = tail.length ? Math.min.apply(null, tail) + 1 : head.length + 1;
      if (head.length && at <= head.length) sh.insertColumnBefore(at);
      sh.getRange(1, at).setValue(k).setFontWeight('bold');
      head = headers(sh);
    });
  }
  if (head.indexOf('Updated') < 0) { sh.getRange(1, head.length + 1).setValue('Updated').setFontWeight('bold'); head = headers(sh); }

  // find existing row
  const idx = keys.map(k => head.indexOf(k));
  const last = sh.getLastRow();
  let r = 0;
  if (last >= 2) {
    const v = sh.getRange(2, 1, last - 1, head.length).getValues();
    for (let i = 0; i < v.length; i++) {
      if (idx.every((ci, j) => norm(v[i][ci]) === norm(rec[keys[j]]))) { r = i + 2; break; }
    }
  }
  if (r && mode === 'create') return 'exists';
  const isNew = !r;
  if (isNew) r = sh.getLastRow() + 1;
  const cur = isNew ? head.map(() => '') : sh.getRange(r, 1, 1, head.length).getValues()[0];
  const out = head.map((h, i) => {
    if (h === 'Updated') return new Date();
    if (h === 'Logged by') return role;
    if (!Object.prototype.hasOwnProperty.call(rec, h)) return cur[i];
    const val = rec[h];
    const blank = val === '' || val === null || val === undefined;
    if (mode === 'overwrite') return blank ? '' : val;
    if (mode === 'fill') return (cur[i] === '' && !blank) ? val : cur[i];
    return blank ? cur[i] : val;                         // merge / create
  });
  sh.getRange(r, 1, 1, head.length).setValues([out]);
  return null;
}

function doPost(e) {
  let b = {};
  try { b = JSON.parse(e.postData.contents || '{}'); } catch (err) { return json({ ok: false, error: 'bad_request' }); }
  const c = config();
  const role = roleFor(b.pin, c);
  if (!role) return json({ ok: false, error: 'pin' });
  const ops = b.action === 'batch' ? (b.ops || []) : [b];
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let err = null;
  try {
    for (let i = 0; i < ops.length && !err; i++) err = save(ops[i], role, c);
  } finally {
    lock.releaseLock();
  }
  const res = allData(role, c);
  if (err) { res.ok = false; res.error = err; }
  return json(res);
}
