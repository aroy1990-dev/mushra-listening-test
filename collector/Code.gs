/**
 * Response collector for the MUSHRA listening test (test/index.html).
 *
 * Setup: paste this file into Extensions > Apps Script of the Google Sheet, Save, then
 * Deploy > Manage deployments > (pencil) Edit > Version: New version > Deploy, so the /exec URL stays the same.
 * (First time: Deploy > New deployment > Web app, Execute as: Me, Who has access: Anyone.)
 *
 * One row per listener session in the tab "responses". The page autosaves after every rating, page change and tab
 * close, so a row appears as soon as someone starts and is updated in place (matched on sessionId) until
 * status = "completed".
 *
 * Columns: sessionId, status, participant, startedAt, submittedAt, elapsedMs, pagesDone, userAgent, order,
 * then one column per clip (e.g. q01_164) with its 0-100 score, then q01_slots..q10_slots (on-screen order A..G),
 * playCounts (JSON) and lastUpdate.
 */
const SHEET_NAME = 'responses';
const BASE = ['sessionId', 'status', 'participant', 'startedAt', 'submittedAt', 'elapsedMs', 'pagesDone',
              'userAgent', 'order'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const p = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

    if (sh.getLastRow() === 0) {
      const qids = Object.keys(p.slots || {}).sort();
      const header = BASE.concat(p.columns || [], qids.map(q => q + '_slots'), ['playCounts', 'lastUpdate']);
      sh.appendRow(header);
      sh.setFrozenRows(1);
    }
    const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];

    const r = p.ratings || {}, s = p.slots || {};
    const row = header.map(h => {
      if (h === 'playCounts') return JSON.stringify(p.playCounts || {});
      if (h === 'lastUpdate') return new Date().toISOString();
      if (BASE.indexOf(h) >= 0) return p[h] === undefined || p[h] === null ? '' : p[h];
      if (/_slots$/.test(h)) return s[h.replace(/_slots$/, '')] || '';
      return r[h] === undefined ? '' : r[h];
    });

    // find this session's row
    let target = -1;
    const last = sh.getLastRow();
    if (last > 1) {
      const ids = sh.getRange(2, 1, last - 1, 1).getValues();
      for (let i = 0; i < ids.length; i++) if (ids[i][0] === p.sessionId) { target = i + 2; break; }
    }
    // never let a late in-progress autosave overwrite a completed row
    if (target > 0 && p.status !== 'completed' && sh.getRange(target, 2).getValue() === 'completed') {
      return ContentService.createTextOutput('ok');
    }
    if (target > 0) sh.getRange(target, 1, 1, row.length).setValues([row]);
    else sh.appendRow(row);
    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
}

/**
 * GET ?participant=<name> returns that name's most recent session as JSON so the page can resume it on any device:
 * {found, status, sessionId, startedAt, order, slots: {q01: "..."}, ratings: {clip: score}, playCounts}.
 * Names are matched ignoring case and surrounding spaces. Without a name it is a health check.
 */
function doGet(e) {
  const name = e && e.parameter && e.parameter.participant;
  if (!name) return ContentService.createTextOutput('listening test collector v3 is running');
  const json = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sh || sh.getLastRow() < 2) return json({ found: false });

  const values = sh.getDataRange().getValues();
  const header = values[0], col = h => header.indexOf(h);
  const want = String(name).trim().toLowerCase();
  // Sheets may turn ISO strings into Dates, so compare times, not strings
  const time = v => v instanceof Date ? v.getTime() : (Date.parse(v) || 0);
  let best = null;
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (String(r[col('participant')]).trim().toLowerCase() !== want) continue;
    if (!best || time(r[col('lastUpdate')]) >= time(best[col('lastUpdate')])) best = r;
  }
  if (!best) return json({ found: false });

  const ratings = {}, slots = {};
  header.forEach((h, j) => {
    if (/^q\d+_\d+$/.test(h) && best[j] !== '') ratings[h] = Number(best[j]);
    if (/_slots$/.test(h)) slots[h.replace(/_slots$/, '')] = String(best[j]);
  });
  let playCounts = {};
  try { playCounts = JSON.parse(best[col('playCounts')] || '{}'); } catch (err) {}
  const iso = v => v instanceof Date ? v.toISOString() : String(v);
  return json({ found: true, status: best[col('status')], sessionId: best[col('sessionId')],
                participant: best[col('participant')], startedAt: iso(best[col('startedAt')]),
                order: String(best[col('order')]), slots, ratings, playCounts });
}
