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

const TEAMS = [
  'NF1', 'NF2', 'NF3', 'NF4',
  'TN1', 'TN2', 'TN3', 'TN4',
  'Center Workers 15', 'Center Workers 17', 'Center Workers 18', 'Center Workers 19', 'Center Workers 20',
  'MWG'
];

const SHEET_BUILDINGS = 'Buildings';
const SHEET_BOOKINGS = 'Bookings';
const B_COLS = ['id', 'name', 'capacity', 'notes', 'active', 'createdAt', 'group'];

// Added automatically the first time you run setup().
const STARTING_LOCATIONS = [
  ['Vandy Buttrick Hall', 'Vanderbilt'],
  ['Vandy Calhoun', 'Vanderbilt'],
  ['Vandy Student Center', 'Vanderbilt'],
  ['Vandy Wilson Hall', 'Vanderbilt'],
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
    case 'create':         return locked_(() => createBooking_(p));
    case 'cancelMine':     return locked_(() => cancelMine_(p));
    case 'adminCheck':     requirePin_(p); return true;
    case 'adminList':      requirePin_(p); return { buildings: readAll_(SHEET_BUILDINGS), bookings: readAll_(SHEET_BOOKINGS).map(publicBooking_) };
    case 'setStatus':      requirePin_(p); return locked_(() => setStatus_(p));
    case 'addBuilding':    requirePin_(p); return locked_(() => addBuilding_(p));
    case 'updateBuilding': requirePin_(p); return locked_(() => updateBuilding_(p));
    case 'deleteBuilding': requirePin_(p); return locked_(() => deleteBuilding_(p));
    default: throw new Error('Unknown action.');
  }
}

/* ---------- public ---------- */
function listPublic_() {
  const cutoff = addDays_(today_(), -30);
  return {
    buildings: readAll_(SHEET_BUILDINGS).filter(b => b.active),
    bookings: readAll_(SHEET_BOOKINGS).filter(r => r.date >= cutoff).map(publicBooking_)
  };
}

function createBooking_(p) {
  const buildings = readAll_(SHEET_BUILDINGS);
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

  const c = clash_(readAll_(SHEET_BOOKINGS), p, 'Approved');
  if (c) throw new Error(b.name + ' is already booked ' + c.start + '–' + c.end + ' that day. Pick another time.');

  const now = new Date().toISOString();
  const r = {
    id: Utilities.getUuid().slice(0, 8), createdAt: now, name: name, team: p.team,
    buildingId: b.id, buildingName: b.name, date: p.date, start: p.start, end: p.end,
    purpose: String(p.purpose || '').trim().slice(0, 300), status: 'Pending', adminNote: '',
    updatedAt: now, secret: Utilities.getUuid()
  };
  appendRow_(SHEET_BOOKINGS, R_COLS, r);
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
    const c = clash_(all, r, 'Approved', r.id);
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
  const sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('The "' + name + '" tab is missing. Run setup() in Apps Script first.');
  return sh;
}

function tz_() { return ss_().getSpreadsheetTimeZone(); }
function today_() { return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd'); }
function addDays_(iso, n) {
  const parts = iso.split('-').map(Number);
  const d = new Date(parts[0], parts[1] - 1, parts[2] + n, 12);
  return Utilities.formatDate(d, tz_(), 'yyyy-MM-dd');
}

function normalize_(col, v) {
  const tz = tz_();
  if (v instanceof Date) {
    if (col === 'date') return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
    if (col === 'start' || col === 'end') return Utilities.formatDate(v, tz, 'HH:mm');
    return v.toISOString();
  }
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

function locked_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
