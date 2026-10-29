/**
 * Mad Apple Winter Workouts — Google Sheet backend
 *
 * Setup (one time, ~3 minutes):
 *  1. Open the "Mad Apple Winter Workouts — Data" sheet.
 *  2. Extensions → Apps Script. Delete what's there, paste this whole file, Save.
 *  3. Deploy → New deployment → type: Web app.
 *       Execute as: Me
 *       Who has access: Anyone
 *     Deploy, approve the Google prompts, copy the Web app URL (ends in /exec).
 *  4. Paste that URL to Claude (or into config.js in the GitHub repo).
 *
 * The Coach PIN lives in the Config tab (cell B2). Change it there any time.
 * Every read and write needs the PIN, so names and numbers stay private.
 */

const SS = SpreadsheetApp.getActiveSpreadsheet();

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function pinOk(pin) {
  const real = String(SS.getSheetByName('Config').getRange('B2').getValue()).trim();
  return real && String(pin || '').trim() === real;
}

function rows(name) {
  const sh = SS.getSheetByName(name);
  const v = sh.getDataRange().getValues();
  const head = v.shift();
  return v.filter(r => r.some(c => c !== '')).map(r => {
    const o = {};
    head.forEach((h, i) => { o[h] = r[i] instanceof Date ? Utilities.formatDate(r[i], 'America/Indiana/Indianapolis', 'yyyy-MM-dd') : r[i]; });
    return o;
  });
}

function allData() {
  return {
    ok: true,
    roster: rows('Roster'),
    attendance: rows('Attendance'),
    bullpens: rows('Bullpens'),
    tests: rows('Tests'),
    notes: rows('Notes')
  };
}

function doGet(e) {
  const p = e.parameter || {};
  if (!pinOk(p.pin)) return json({ ok: false, error: 'pin' });
  return json(allData());
}

function norm(v) {
  return v instanceof Date ? Utilities.formatDate(v, 'America/Indiana/Indianapolis', 'yyyy-MM-dd') : String(v).trim();
}

function headers(sh) { return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]; }

// Find the row where every key column (by header name) matches rec; returns 1-based row or 0.
function findRow(sh, head, keyCols, rec) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const idx = keyCols.map(k => head.indexOf(k));
  const v = sh.getRange(2, 1, last - 1, head.length).getValues();
  for (let i = 0; i < v.length; i++) {
    if (idx.every((c, j) => norm(v[i][c]) === norm(rec[keyCols[j]]))) return i + 2;
  }
  return 0;
}

// Write a record (object keyed by header) into the row matching `keyCols`, creating it if needed.
// Blank values in `rec` don't overwrite existing cells unless `overwrite` is true.
function upsert(name, keyCols, rec, overwrite) {
  const sh = SS.getSheetByName(name);
  const head = headers(sh);
  let r = findRow(sh, head, keyCols, rec);
  const isNew = !r;
  if (isNew) r = sh.getLastRow() + 1;
  const cur = isNew ? head.map(() => '') : sh.getRange(r, 1, 1, head.length).getValues()[0];
  const out = head.map((h, i) => {
    if (h === 'Updated') return new Date();
    if (Object.prototype.hasOwnProperty.call(rec, h)) {
      const val = rec[h];
      if (overwrite || (val !== '' && val !== null && val !== undefined)) return val;
    }
    return cur[i];
  });
  sh.getRange(r, 1, 1, head.length).setValues([out]);
}

function playerName(num) {
  const r = rows('Roster').find(p => String(p.Number) === String(num));
  return r ? (r.First + ' ' + r.Last) : '';
}

function doPost(e) {
  let b = {};
  try { b = JSON.parse(e.postData.contents || '{}'); } catch (err) { return json({ ok: false, error: 'bad_request' }); }
  if (!pinOk(b.pin)) return json({ ok: false, error: 'pin' });

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    switch (b.action) {
      case 'attendance':
        upsert('Attendance', ['Week', 'Number'], {
          Week: b.week, Date: b.date, Number: b.num, Player: playerName(b.num), Present: !!b.present
        }, true);
        break;

      case 'pitcher': {
        const sh = SS.getSheetByName('Roster');
        const r = findRow(sh, headers(sh), ['Number'], { Number: b.num });
        if (r) sh.getRange(r, 4).setValue(!!b.pitcher);
        break;
      }

      case 'bullpen': {
        const pct = b.pitches ? Math.round((b.strikes / b.pitches) * 100) : '';
        upsert('Bullpens', ['Week', 'Number'], {
          Week: b.week, Date: b.date, Number: b.num, Pitcher: playerName(b.num),
          Pitches: b.pitches, Strikes: b.strikes, 'Strike %': pct,
          'Batters faced': b.batters, '1st-pitch strikes': b.fps, 'Ahead after 3': b.ahead,
          Changeups: b.ch, 'CH strikes': b.chStrikes, Cue: b.cue
        }, true);
        break;
      }

      case 'test':
        upsert('Tests', ['Week', 'Number'], {
          Week: b.week, Date: b.date, Number: b.num, Player: playerName(b.num),
          'Max exit velo (mph)': b.ev, 'Line-drive %': b.ld, '30-yd sprint (sec)': b.sprint,
          'Broad jump (in)': b.broad, 'Plank (sec)': b.plank
        }, false);
        break;

      case 'note':
        upsert('Notes', ['Week'], { Week: b.week, Date: b.date, Notes: b.text }, true);
        break;

      default:
        return json({ ok: false, error: 'unknown_action' });
    }
  } finally {
    lock.releaseLock();
  }
  return json(allData());
}
