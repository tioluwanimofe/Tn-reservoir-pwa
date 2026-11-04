/**
 * TN-Reservoir — Google Apps Script backend
 *
 * SETUP
 * 1. Open the TN-Reservoir Google Sheet (SHEET_ID below).
 * 2. Extensions > Apps Script. Delete what's there and paste this whole file.
 * 3. Change ADMIN_PIN below to your own PIN.
 * 4. Pick the "setup" function in the toolbar and click Run. Approve the permissions.
 *    This creates the "Buildings" (locations) and "Bookings" tabs and adds your five locations.
 * 5. Deploy > New deployment > type "Web app".
 *      Execute as: Me
 *      Who has access: Anyone
 *    Click Deploy and copy the Web app URL (ends in /exec).
 * 6. Paste that URL into config.js in the app folder.
 *
 * After editing this code later: Deploy > Manage deployments > edit (pencil) >
 * Version: New version > Deploy. The URL stays the same.
 */

const ADMIN_PIN = 'change-me';

// Your TN-Reservoir Google Sheet. The script always saves here, even if it's
// pasted into a standalone Apps Script project instead of the Sheet itself.
const SHEET_ID = '1lC11JRkFHdwg7ahJGhX-ocjQMc6s5OnZ8awk7SGRRkY';

// Email alerts. They go to the Google account that owns this script.
// To send them somewhere else too, add addresses here, e.g. 'a@x.com,b@y.com'
// (this file is on GitHub, so keep real addresses in Apps Script only).
const EXTRA_ALERT_EMAILS = '';
const EMAIL_ON_NEW_REQUEST = true;
const DAILY_SUMMARY_HOUR = 7; // 7 AM in the script's time zone
const APP_URL = 'https://tioluwanimofe.github.io/Tn-reservoir-pwa/';

const TEAMS = [
  'YA Department',
  'NF1', 'NF2', 'NF3', 'NF4',
  'TN1', 'TN2', 'TN3', 'TN4',
  '15.0GI', '17GI', '18GI', '19GI', '20GI', // Center Workers
  'MWG'
];

const SHEET_BUILDINGS = 'Buildings';
const SHEET_BOOKINGS = 'Bookings';
const SHEET_SERIES = 'Series';
const CENTER_WORKERS = ['15.0GI', '17GI', '18GI', '19GI', '20GI'];
const HORIZON_DAYS = 365; // how far ahead repeating bookings are shown and checked
const S_COLS = ['id', 'createdAt', 'name', 'team', 'buildingId', 'buildingName', 'days', 'startDate', 'endDate', 'start', 'end', 'purpose', 'status', 'adminNote', 'updatedAt', 'skips', 'secret'];
const B_COLS = ['id', 'name', 'capacity', 'notes', 'active', 'createdAt', 'group'];

// Added automatically the first time you run setup().
const STARTING_LOCATIONS = [
  ['Vandy Buttrick Hall', 'Vanderbilt'],
  ['Vandy Calhoun', 'Vanderbilt'],
  ['Vandy Student Center', 'Vanderbilt'],
  ['Vandy Wilson Hall', 'Vanderbilt'],
  ['Vandy Peabody', 'Vanderbilt'],
  ['Vandy Featheringill Hall', 'Vanderbilt'],
  ['Belmont Barbara Massey', 'Belmont']
];
const R_COLS = ['id', 'createdAt', 'name', 'team', 'buildingId', 'buildingName', 'date', 'start', 'end', 'purpose', 'status', 'adminNote', 'updatedAt', 'secret'];
const STATUSES = ['Pending', 'Approved', 'Declined', 'Cancelled'];
const OPEN = ['Pending', 'Approved'];

/* ---------- one-time setup ---------- */
function setup() {
  const ss = ss_();
  ensureSheet_(ss, SHEET_BUILDINGS, B_COLS);
  ensureSheet_(ss, SHEET_BOOKINGS, R_COLS);
  ensureSheet_(ss, SHEET_SERIES, S_COLS);
  if (readAll_(SHEET_BUILDINGS).length === 0) {
    STARTING_LOCATIONS.forEach(l => addBuilding_({ name: l[0], group: l[1] }));
  }
  const first = ss.getSheetByName('Sheet1');
  if (first && first.getLastRow() === 0 && ss.getSheets().length > 2) ss.deleteSheet(first);
}

function ensureSheet_(ss, name, cols) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
  sh.setFrozenRows(1);
  // Keep dates and times as plain text so Sheets doesn't reformat them.
  sh.getRange(1, 1, sh.getMaxRows(), cols.length).setNumberFormat('@');
}

/* ---------- web app entry points ---------- */
function doGet() {
  return json_({ ok: true, data: 'TN-Reservoir API is running.' });
}

function doPost(e) {
  try {
    const p = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    return json_({ ok: true, data: route_(p.action, p) });
  } catch (err) {
    return json_({ ok: false, error: String((err && err.message) || err) });
  }
}

function route_(action, p) {
  switch (action) {
    case 'listPublic':     return listPublic_();
    case 'create':         return createBooking_(p);        // locks only around the save
    case 'cancelMine':     return locked_(() => cancelMine_(p));
    case 'adminCheck':     requirePin_(p); return true;
    case 'adminList':      requirePin_(p); return adminList_();
    case 'createSeries':   return createSeries_(p);         // locks around the clash check + save
    case 'stopSeries':     return locked_(() => stopSeries_(p));
    case 'skipDate':       requirePin_(p); return locked_(() => skipDate_(p));
    case 'endSeries':      requirePin_(p); return locked_(() => endSeriesAdmin_(p));
    case 'setStatus':      requirePin_(p); return locked_(() => setStatus_(p));
    case 'addBuilding':    requirePin_(p); return locked_(() => addBuilding_(p));
    case 'updateBuilding': requirePin_(p); return locked_(() => updateBuilding_(p));
    case 'deleteBuilding': requirePin_(p); return locked_(() => deleteBuilding_(p));
    default: throw new Error('Unknown action.');
  }
}

/* ---------- public ---------- */
function listPublic_() {
  const cached = cacheGet_();
  if (cached) return cached;
  const from = isoAdd_(today_(), -30), to = isoAdd_(today_(), HORIZON_DAYS);
  const series = readAll_(SHEET_SERIES);
  const out = {
    buildings: readAll_(SHEET_BUILDINGS).filter(b => b.active),
    bookings: readAll_(SHEET_BOOKINGS).filter(r => r.date >= from).map(publicBooking_).concat(expandAll_(series, from, to)),
    series: series.map(publicBooking_)
  };
  cachePut_(out);
  return out;
}

function adminList_() {
  const from = isoAdd_(today_(), -30), to = isoAdd_(today_(), HORIZON_DAYS);
  const series = readAll_(SHEET_SERIES);
  return {
    buildings: readAll_(SHEET_BUILDINGS),
    bookings: readAll_(SHEET_BOOKINGS).map(publicBooking_).concat(expandAll_(series, from, to)),
    series: series.map(publicBooking_)
  };
}

// Approved one-off bookings plus every repeating booking that falls on this date.
function approvedOn_(date) {
  return readAll_(SHEET_BOOKINGS).filter(r => r.status === 'Approved').concat(expandAll_(readAll_(SHEET_SERIES), date, date));
}

function validateRequest_(p, buildings) {
  const name = String(p.name || '').trim();
  if (!name) throw new Error('Enter your name.');
  if (name.length > 80) throw new Error('Keep the name under 80 characters.');
  if (TEAMS.indexOf(p.team) === -1) throw new Error('Choose your team.');
  const b = buildings.find(x => x.id === p.buildingId && x.active);
  if (!b) throw new Error('Choose a location that is open for booking.');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(p.date || '')) throw new Error('Choose a date.');
  if (p.date < today_()) throw new Error('That date has already passed.');
  if (!/^\d{2}:\d{2}$/.test(p.start || '') || !/^\d{2}:\d{2}$/.test(p.end || '')) throw new Error('Choose a start and end time.');
  if (p.start >= p.end) throw new Error('The end time needs to be after the start time.');
  return { name: name, b: b };
}

function createBooking_(p) {
  const v = validateRequest_(p, readAll_(SHEET_BUILDINGS));
  const name = v.name, b = v.b;

  const c = clash_(approvedOn_(p.date), p, 'Approved');
  if (c) throw new Error(b.name + ' is already booked ' + c.start + '–' + c.end + ' that day. Pick another time.');

  const now = new Date().toISOString();
  const r = {
    id: Utilities.getUuid().slice(0, 8), createdAt: now, name: name, team: p.team,
    buildingId: b.id, buildingName: b.name, date: p.date, start: p.start, end: p.end,
    purpose: String(p.purpose || '').trim().slice(0, 300), status: 'Pending', adminNote: '',
    updatedAt: now, secret: Utilities.getUuid()
  };
  // A pending request doesn't block anyone (approval re-checks under the lock),
  // so only the write itself needs the lock.
  locked_(() => appendRow_(SHEET_BOOKINGS, R_COLS, r));
  notifyNew_(r, false);
  return { id: r.id, secret: r.secret };
}

function cancelMine_(p) {
  const r = readAll_(SHEET_BOOKINGS).find(x => x.id === p.id);
  if (!r) throw new Error('Request not found.');
  if (!p.secret || r.secret !== p.secret) throw new Error('This request can only be cancelled from the device that sent it.');
  if (OPEN.indexOf(r.status) === -1) throw new Error('This request is already closed.');
  updateRow_(SHEET_BOOKINGS, R_COLS, p.id, { status: 'Cancelled', adminNote: 'Cancelled by requester', updatedAt: new Date().toISOString() });
  return true;
}

/* ---------- repeating bookings (Center Workers) ---------- */
function parseDays_(v) {
  const seen = {};
  String(v === undefined || v === null ? '' : v).split(',').forEach(x => { const n = Number(x); if (x !== '' && n >= 0 && n <= 6) seen[n] = true; });
  return Object.keys(seen).map(Number).sort();
}

function expandSeries_(s, from, to) {
  const out = [];
  if (s.status === 'Cancelled') return out;
  const days = parseDays_(s.days);
  const skips = {};
  String(s.skips || '').split(',').forEach(d => { if (d) skips[d] = true; });
  let d = s.startDate > from ? s.startDate : from;
  const last = s.endDate && s.endDate < to ? s.endDate : to;
  while (d <= last) {
    if (days.indexOf(dow_(d)) !== -1 && !skips[d]) {
      out.push({ id: s.id + '~' + d, seriesId: s.id, name: s.name, team: s.team, buildingId: s.buildingId, buildingName: s.buildingName,
        date: d, start: s.start, end: s.end, purpose: s.purpose, status: 'Approved', adminNote: '', days: s.days, createdAt: s.createdAt });
    }
    d = isoAdd_(d, 1);
  }
  return out;
}

function expandAll_(list, from, to) {
  return list.reduce((acc, s) => acc.concat(expandSeries_(s, from, to)), []);
}

function seriesActive_(s) {
  return s.status !== 'Cancelled' && (!s.endDate || s.endDate >= today_());
}

function createSeries_(p) {
  if (CENTER_WORKERS.indexOf(p.team) === -1) throw new Error('Only Center Workers teams can make repeating bookings.');
  const v = validateRequest_(Object.assign({}, p, { date: p.startDate }), readAll_(SHEET_BUILDINGS));
  const days = parseDays_(p.days);
  if (!days.length) throw new Error('Pick at least one day for the booking to repeat on.');
  const endDate = String(p.endDate || '');
  if (endDate) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(endDate)) throw new Error('Choose a last date, or set it to keep repeating.');
    if (endDate < p.startDate) throw new Error('The last date needs to be on or after the start date.');
  }
  const now = new Date().toISOString();
  const s = {
    id: 'S' + Utilities.getUuid().slice(0, 7), createdAt: now, name: v.name, team: p.team,
    buildingId: v.b.id, buildingName: v.b.name, days: days.join(','), startDate: p.startDate, endDate: endDate,
    start: p.start, end: p.end, purpose: String(p.purpose || '').trim().slice(0, 300),
    status: 'Active', adminNote: '', updatedAt: now, skips: '', secret: Utilities.getUuid()
  };
  const to = endDate || isoAdd_(p.startDate, HORIZON_DAYS);
  const occ = expandSeries_(s, p.startDate, to);
  if (!occ.length) throw new Error('None of the chosen days fall between the start and last date.');
  // A repeating booking is approved on save, so the clash check and the save must happen together.
  locked_(() => {
    const approved = readAll_(SHEET_BOOKINGS).filter(r => r.status === 'Approved').concat(expandAll_(readAll_(SHEET_SERIES), p.startDate, to));
    for (let i = 0; i < occ.length; i++) {
      const c = clash_(approved, occ[i], 'Approved');
      if (c) throw new Error(v.b.name + ' is already booked on ' + occ[i].date + ' (' + c.start + '–' + c.end + ', ' + c.team + '). Change the days or time.');
    }
    appendRow_(SHEET_SERIES, S_COLS, s);
  });
  notifyNew_(s, true);
  return { id: s.id, secret: s.secret };
}

function findSeries_(id) {
  const s = readAll_(SHEET_SERIES).find(x => x.id === id);
  if (!s) throw new Error('Repeating booking not found.');
  return s;
}

function endSeries_(s, note) {
  if (!seriesActive_(s)) throw new Error('This repeating booking has already ended.');
  const y = isoAdd_(today_(), -1);
  const ch = { adminNote: note, updatedAt: new Date().toISOString() };
  if (s.startDate > y) ch.status = 'Cancelled';
  else { ch.status = 'Ended'; if (!s.endDate || s.endDate > y) ch.endDate = y; }
  updateRow_(SHEET_SERIES, S_COLS, s.id, ch);
  return true;
}

function stopSeries_(p) {
  const s = findSeries_(p.id);
  if (!p.secret || s.secret !== p.secret) throw new Error('This repeating booking can only be stopped from the device that made it.');
  return endSeries_(s, 'Stopped by requester');
}

function endSeriesAdmin_(p) {
  return endSeries_(findSeries_(p.id), String(p.note || '').trim().slice(0, 300) || 'Ended by admin');
}

function skipDate_(p) {
  const s = findSeries_(p.id);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(p.date || '')) throw new Error('Pick the date to cancel.');
  if (!expandSeries_(s, p.date, p.date).length) throw new Error("That date isn't one of this booking's repeat days.");
  const skips = String(s.skips || '').split(',').filter(Boolean);
  if (skips.indexOf(p.date) === -1) skips.push(p.date);
  const ch = { skips: skips.sort().join(','), updatedAt: new Date().toISOString() };
  if (p.note) ch.adminNote = String(p.note).trim().slice(0, 300);
  updateRow_(SHEET_SERIES, S_COLS, s.id, ch);
  return true;
}

/* ---------- admin ---------- */
function requirePin_(p) {
  if (String(p.pin || '') !== String(ADMIN_PIN)) throw new Error('That PIN is not right.');
}

function setStatus_(p) {
  if (STATUSES.indexOf(p.status) === -1) throw new Error('Unknown status.');
  const all = readAll_(SHEET_BOOKINGS);
  const r = all.find(x => x.id === p.id);
  if (!r) throw new Error('Request not found.');
  if (p.status === 'Approved') {
    const c = clash_(approvedOn_(r.date), r, 'Approved', r.id);
    if (c) throw new Error('This clashes with an approved booking for ' + c.team + ' (' + c.start + '–' + c.end + '). Cancel that one first, or decline this one.');
  }
  const changes = { status: p.status, updatedAt: new Date().toISOString() };
  if (p.note) changes.adminNote = String(p.note).trim().slice(0, 300);
  updateRow_(SHEET_BOOKINGS, R_COLS, p.id, changes);
  return true;
}

function addBuilding_(p) {
  const name = String(p.name || '').trim();
  if (!name) throw new Error('Give the location a name.');
  const existing = readAll_(SHEET_BUILDINGS);
  if (existing.some(b => b.name.toLowerCase() === name.toLowerCase())) throw new Error('A location with that name already exists.');
  const b = {
    id: Utilities.getUuid().slice(0, 8), name: name, group: String(p.group || '').trim(),
    capacity: p.capacity ? Number(p.capacity) : '', notes: String(p.notes || '').trim(),
    active: true, createdAt: new Date().toISOString()
  };
  appendRow_(SHEET_BUILDINGS, B_COLS, b);
  return b;
}

function updateBuilding_(p) {
  const changes = {};
  if ('active' in p) changes.active = !!p.active;
  if (p.name) changes.name = String(p.name).trim();
  if ('group' in p) changes.group = String(p.group || '').trim();
  if ('capacity' in p) changes.capacity = p.capacity ? Number(p.capacity) : '';
  if ('notes' in p) changes.notes = String(p.notes || '').trim();
  updateRow_(SHEET_BUILDINGS, B_COLS, p.id, changes);
  return true;
}

function deleteBuilding_(p) {
  const t = today_();
  const upcoming = readAll_(SHEET_BOOKINGS).some(r => r.buildingId === p.id && r.date >= t && OPEN.indexOf(r.status) !== -1);
  if (upcoming) throw new Error('This location has upcoming bookings. Hide it instead, or cancel those bookings first.');
  const sh = sheet_(SHEET_BUILDINGS);
  const row = findRow_(sh, p.id);
  if (!row) throw new Error('Location not found.');
  sh.deleteRow(row);
  return true;
}

/* ---------- email alerts ---------- */
const DOW_NAMES_ = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];

function alertTo_() {
  const me = Session.getEffectiveUser().getEmail();
  return [me].concat(String(EXTRA_ALERT_EMAILS || '').split(',')).map(x => x.trim()).filter(Boolean).join(',');
}

function fmtTime_(t) {
  const h = Number(t.slice(0, 2)), m = t.slice(3, 5);
  return (h % 12 || 12) + (m === '00' ? '' : ':' + m) + (h >= 12 ? ' PM' : ' AM');
}
function fmtDate_(iso) {
  const p = iso.split('-').map(Number);
  const d = new Date(Date.UTC(p[0], p[1] - 1, p[2]));
  return DOW_NAMES_[d.getUTCDay()] + ', ' + ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][p[1] - 1] + ' ' + p[2];
}
function esc_(v) { return String(v === undefined || v === null ? '' : v).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }

// Sent after a request is saved. A failed email never fails the booking.
function notifyNew_(r, repeating) {
  if (!EMAIL_ON_NEW_REQUEST) return;
  try {
    const when = repeating
      ? 'Every ' + parseDays_(r.days).map(i => DOW_NAMES_[i]).join(', ') + ', ' + fmtTime_(r.start) + '–' + fmtTime_(r.end) + ', from ' + fmtDate_(r.startDate) + (r.endDate ? ' to ' + fmtDate_(r.endDate) : ' (no end date)')
      : fmtDate_(r.date) + ', ' + fmtTime_(r.start) + '–' + fmtTime_(r.end);
    let overlap = '';
    if (!repeating) {
      const other = readAll_(SHEET_BOOKINGS).find(x => x.id !== r.id && x.status === 'Pending' && x.buildingId === r.buildingId && x.date === r.date && x.start < r.end && r.start < x.end);
      if (other) overlap = 'Overlaps a pending request from ' + other.team + ' (' + fmtTime_(other.start) + '–' + fmtTime_(other.end) + ').';
    }
    const subject = repeating
      ? 'TN-Reservoir: ' + r.team + ' booked ' + r.buildingName + ' (repeating)'
      : 'TN-Reservoir: new request from ' + r.team + ' for ' + r.buildingName;
    const lines = [
      ['Who', r.name + ' · ' + r.team], ['Where', r.buildingName], ['When', when],
      ['What for', r.purpose || '—'],
      ['Status', repeating ? 'Already blocking the location (repeating bookings don\'t need approval)' : 'Waiting for approval']
    ];
    const html = '<div style="font-family:Arial,sans-serif;font-size:14px;color:#18201D">'
      + '<p style="margin:0 0 12px">' + (repeating ? 'A Center Workers team made a repeating booking.' : 'A new request is waiting for you.') + '</p>'
      + '<table style="border-collapse:collapse">' + lines.map(l => '<tr><td style="padding:4px 14px 4px 0;color:#5A645F">' + l[0] + '</td><td style="padding:4px 0"><b>' + esc_(l[1]) + '</b></td></tr>').join('') + '</table>'
      + (overlap ? '<p style="color:#94590A;margin:12px 0 0">' + esc_(overlap) + '</p>' : '')
      + '<p style="margin:18px 0 0"><a href="' + APP_URL + '#admin" style="background:#4A6B64;color:#fff;padding:10px 16px;border-radius:8px;text-decoration:none">Open Admin</a></p></div>';
    const text = lines.map(l => l[0] + ': ' + l[1]).join('\n') + (overlap ? '\n' + overlap : '') + '\n\nOpen Admin: ' + APP_URL + '#admin';
    MailApp.sendEmail({ to: alertTo_(), subject: subject, body: text, htmlBody: html, name: 'TN-Reservoir' });
  } catch (err) {
    console.warn('Alert email failed: ' + err);
  }
}

// Emails a list of requests still waiting for approval. Runs every morning once turned on.
function sendDailySummary() {
  const t = today_();
  const waiting = readAll_(SHEET_BOOKINGS).filter(r => r.status === 'Pending' && r.date >= t).sort((a, b) => (a.date + a.start).localeCompare(b.date + b.start));
  if (!waiting.length) return 'Nothing waiting; no email sent.';
  const rows = waiting.map(r => '<tr><td style="padding:4px 12px 4px 0">' + esc_(fmtDate_(r.date)) + '</td><td style="padding:4px 12px 4px 0">' + esc_(fmtTime_(r.start) + '–' + fmtTime_(r.end)) + '</td><td style="padding:4px 12px 4px 0">' + esc_(r.buildingName) + '</td><td style="padding:4px 0">' + esc_(r.team + ' · ' + r.name) + '</td></tr>').join('');
  const html = '<div style="font-family:Arial,sans-serif;font-size:14px;color:#18201D"><p>' + waiting.length + ' request' + (waiting.length === 1 ? ' is' : 's are') + ' waiting for approval:</p>'
    + '<table style="border-collapse:collapse">' + rows + '</table>'
    + '<p style="margin:18px 0 0"><a href="' + APP_URL + '#admin" style="background:#4A6B64;color:#fff;padding:10px 16px;border-radius:8px;text-decoration:none">Open Admin</a></p></div>';
  const text = waiting.map(r => fmtDate_(r.date) + ' ' + fmtTime_(r.start) + '–' + fmtTime_(r.end) + ' · ' + r.buildingName + ' · ' + r.team + ' · ' + r.name).join('\n') + '\n\nOpen Admin: ' + APP_URL + '#admin';
  MailApp.sendEmail({ to: alertTo_(), subject: 'TN-Reservoir: ' + waiting.length + ' waiting for approval', body: text, htmlBody: html, name: 'TN-Reservoir' });
  return 'Sent summary of ' + waiting.length + '.';
}

// Run once from the Apps Script editor to start the morning email (safe to run again).
function turnOnDailySummary() {
  ScriptApp.getProjectTriggers().filter(tr => tr.getHandlerFunction() === 'sendDailySummary').forEach(tr => ScriptApp.deleteTrigger(tr));
  ScriptApp.newTrigger('sendDailySummary').timeBased().everyDays(1).atHour(DAILY_SUMMARY_HOUR).create();
  return 'Daily summary on at ' + DAILY_SUMMARY_HOUR + ':00 (' + Session.getScriptTimeZone() + ').';
}

// Run once from the editor to stop the morning email.
function turnOffDailySummary() {
  ScriptApp.getProjectTriggers().filter(tr => tr.getHandlerFunction() === 'sendDailySummary').forEach(tr => ScriptApp.deleteTrigger(tr));
  return 'Daily summary off.';
}

/* ---------- one-off maintenance ---------- */
// Permanently removes the rows created by the load test (location "LOAD TEST").
// Run it yourself from the editor; it does nothing else.
function deleteLoadTestRows() {
  return locked_(() => {
    const isTest = (r) => r.buildingName === 'LOAD TEST' || /^(Load tester|Repeat tester) \d+$/.test(r.name || '');
    const removeRows = (name, keep) => {
      const sh = sheet_(name), rows = readAll_(name); let removed = 0;
      for (let i = rows.length - 1; i >= 0; i--) if (!keep(rows[i])) { sh.deleteRow(i + 2); removed++; }
      return removed;
    };
    const b = removeRows(SHEET_BOOKINGS, r => !isTest(r));
    const s = removeRows(SHEET_SERIES, r => !isTest(r));
    const l = removeRows(SHEET_BUILDINGS, r => r.name !== 'LOAD TEST');
    const msg = 'Removed ' + b + ' test bookings, ' + s + ' test repeating bookings and ' + l + ' test location.';
    console.log(msg);
    return msg;
  });
}

/* ---------- speed: short-lived copy of the public schedule ---------- */
// Apps Script caches hold up to 100 KB per entry, so the schedule is split into pieces.
const CACHE_KEY_ = 'pub_v1';
function cacheGet_() {
  try {
    const c = CacheService.getScriptCache(), n = Number(c.get(CACHE_KEY_ + ':n'));
    if (!n) return null;
    const keys = []; for (let i = 0; i < n; i++) keys.push(CACHE_KEY_ + ':' + i);
    const got = c.getAll(keys); let str = '';
    for (let i = 0; i < keys.length; i++) { if (got[keys[i]] === undefined || got[keys[i]] === null) return null; str += got[keys[i]]; }
    return JSON.parse(str);
  } catch (e) { return null; }
}
function cachePut_(obj) {
  try {
    const str = JSON.stringify(obj), size = 45000, parts = {}; let n = 0;
    for (let i = 0; i < str.length; i += size) parts[CACHE_KEY_ + ':' + (n++)] = str.slice(i, i + size);
    if (n > 100) return; // too big to be worth caching
    CacheService.getScriptCache().putAll(parts, 60);
    CacheService.getScriptCache().put(CACHE_KEY_ + ':n', String(n), 60);
  } catch (e) {}
}
function cacheClear_() {
  try { CacheService.getScriptCache().remove(CACHE_KEY_ + ':n'); } catch (e) {}
}

/* ---------- helpers ---------- */
let ssCache_ = null;
function ss_() { return ssCache_ || (ssCache_ = SpreadsheetApp.openById(SHEET_ID)); }

function clash_(list, p, status, excludeId) {
  return list.find(r => r.id !== excludeId && r.status === status && r.buildingId === p.buildingId &&
    r.date === p.date && r.start < p.end && p.start < r.end);
}

function publicBooking_(r) {
  const out = {};
  Object.keys(r).forEach(k => { if (k !== 'secret') out[k] = r[k]; });
  return out;
}

function sheet_(name) {
  let sh = ss_().getSheetByName(name);
  if (!sh && name === SHEET_SERIES) { ensureSheet_(ss_(), SHEET_SERIES, S_COLS); sh = ss_().getSheetByName(name); }
  if (!sh) throw new Error('The "' + name + '" tab is missing. Run setup() in Apps Script first.');
  return sh;
}

// Pure date math on yyyy-MM-dd strings (no time zone involved).
function isoAdd_(iso, n) {
  const p = iso.split('-').map(Number);
  return new Date(Date.UTC(p[0], p[1] - 1, p[2] + n)).toISOString().slice(0, 10);
}
function dow_(iso) {
  const p = iso.split('-').map(Number);
  return new Date(Date.UTC(p[0], p[1] - 1, p[2])).getUTCDay();
}

// Looked up once per request (each lookup is a call to Google Sheets).
let tzCache_ = null, todayCache_ = null;
function tz_() { return tzCache_ || (tzCache_ = ss_().getSpreadsheetTimeZone()); }
function today_() { return todayCache_ || (todayCache_ = Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd')); }
function addDays_(iso, n) {
  const parts = iso.split('-').map(Number);
  const d = new Date(parts[0], parts[1] - 1, parts[2] + n, 12);
  return Utilities.formatDate(d, tz_(), 'yyyy-MM-dd');
}

function normalize_(col, v) {
  const tz = tz_();
  if (v instanceof Date) {
    if (col === 'date' || col === 'startDate' || col === 'endDate') return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
    if (col === 'start' || col === 'end') return Utilities.formatDate(v, tz, 'HH:mm');
    return v.toISOString();
  }
  if (col === 'days' && typeof v === 'number') return String(v);
  if (col === 'active') return v === true || String(v).toUpperCase() === 'TRUE';
  if (col === 'capacity') return v === '' || v === null ? '' : Number(v);
  if ((col === 'start' || col === 'end') && /^\d:\d{2}$/.test(String(v))) return '0' + v;
  return v === null || v === undefined ? '' : String(v);
}

function readAll_(name) {
  const values = sheet_(name).getDataRange().getValues();
  const head = values.shift() || [];
  return values.filter(row => row[0] !== '').map(row => {
    const o = {};
    head.forEach((h, i) => { o[h] = normalize_(h, row[i]); });
    return o;
  });
}

function appendRow_(name, cols, obj) {
  const sh = sheet_(name);
  const row = sh.getLastRow() + 1;
  const range = sh.getRange(row, 1, 1, cols.length);
  range.setNumberFormat('@');
  range.setValues([cols.map(c => {
    const v = obj[c];
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    return v === undefined || v === null ? '' : String(v);
  })]);
}

function findRow_(sh, id) {
  const ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  for (let i = 1; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 1;
  return 0;
}

function updateRow_(name, cols, id, changes) {
  const sh = sheet_(name);
  const row = findRow_(sh, id);
  if (!row) throw new Error('Not found.');
  Object.keys(changes).forEach(k => {
    const c = cols.indexOf(k);
    if (c === -1) return;
    const v = changes[k];
    const cell = sh.getRange(row, c + 1);
    cell.setNumberFormat('@');
    cell.setValue(typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v));
  });
}

// One writer at a time. flush() makes sure the change is really in the Sheet
// before the next request (waiting on the lock) reads it.
function locked_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('The booking system is busy right now. Please try again in a few seconds.');
  try { const out = fn(); SpreadsheetApp.flush(); cacheClear_(); return out; } finally { lock.releaseLock(); }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
