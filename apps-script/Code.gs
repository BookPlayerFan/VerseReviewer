/**
 * Bible Verse Spaced Repetition System
 * Tab: "Verses"
 * Columns:
 *   B = Next Review Date
 *   C = Last Review Date
 *   D = Rep
 *   E = Level
 *   F = Reference
 *   G = Version (e.g. ESV, NIV, CSB)
 *   H = Verse Text
 *
 * DAILY SESSION: the day's deck, your position in it, same-day retries and today's
 * stats are stored on the server (Script Properties), so every device shows the
 * same session. The deck is fixed when the day's first device opens the app and
 * stays until it is finished. Each new day starts a fresh session.
 * To inspect it, run debugSession() from the editor and open View -> Logs.
 *
 * LEARNING VERSES: Level 0 (new) verses are not part of the daily deck. They are returned
 * separately as `learning`, and the app remembers whether you were last on the review
 * screen or the learning screen (`mode`), synced across devices.
 *
 * TIME ZONE: set BOTH of these to America/Los_Angeles so "today" means the same thing everywhere:
 *   1) Apps Script: Project Settings -> Time zone
 *   2) Spreadsheet: File -> Settings -> Time zone
 *
 * OPTIONAL ACCESS KEY: add a Script Property named VR_KEY (Project Settings -> Script Properties).
 * When set, the web app only loads if the URL includes ?key=<that value>. When not set, no check is done.
 */

const MAX_GROUP = 8;
const UNDO_WINDOW_MS = 2 * 60 * 1000; // how long the last review can be undone
const BUILD = 'sync-27'; // shown in the app's status readout (tap the "Card x of y" label)

// Intervals (days) are used only here. The client mirrors names and repsNeeded only.
const GROUP_CONFIG = {
  0: { name: "New",               interval: 0,   repsNeeded: 1 },
  1: { name: "Bronze",             interval: 1,   repsNeeded: 4 },
  2: { name: "Silver",             interval: 2,   repsNeeded: 3 },
  3: { name: "Gold",             interval: 5,   repsNeeded: 2 },
  4: { name: "Platinum",            interval: 12,  repsNeeded: 2 },
  5: { name: "Sapphire",           interval: 30,  repsNeeded: 2 },
  6: { name: "Emerald",         interval: 75,  repsNeeded: 1 },
  7: { name: "Ruby",         interval: 180, repsNeeded: 1 },
  8: { name: "Diamond",           interval: 360, repsNeeded: 1 }
};

const VIEWPORT = 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no, viewport-fit=cover';

// ───────────────────────── Web app / sidebar entry points ─────────────────────────

function doGet(e) {
  const requiredKey = PropertiesService.getScriptProperties().getProperty('VR_KEY');
  if (requiredKey && !(e && e.parameter && e.parameter.key === requiredKey)) {
    return accessDeniedPage_();
  }
  return buildTemplate_().evaluate()
    .setTitle('Verse Review')
    .addMetaTag('viewport', VIEWPORT)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function accessDeniedPage_() {
  const html =
    '<!DOCTYPE html><html style="color-scheme: light dark;"><head><base target="_top"></head>' +
    '<body style="font-family:-apple-system,sans-serif;text-align:center;padding:48px 24px;">' +
    '<p>Access key required.</p>' +
    '<script>try{window.top.postMessage({type:"VR_AUTH_REQUIRED"},"*");}catch(e){}</script>' +
    '</body></html>';
  return HtmlService.createHtmlOutput(html)
    .setTitle('Verse Review')
    .addMetaTag('viewport', VIEWPORT)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Verse Review')
    .addItem('Open Flashcard Review', 'showFlashcardSidebar')
    .addToUi();
}

function showFlashcardSidebar() {
  const html = buildTemplate_().evaluate()
    .setTitle('Verse Review')
    .setWidth(360);
  SpreadsheetApp.getUi().showSidebar(html);
}

// Builds the page with today's session pre-injected. If loading fails, the page
// receives null and fetches the session itself (with a visible error and retry).
function buildTemplate_() {
  const template = HtmlService.createTemplateFromFile('Sidebar');
  let session = null;
  try { session = loadSession(); } catch (err) { session = null; }
  // Escape "<" so verse text containing "</script>" can't break out of the script block.
  template.initialSessionJson = JSON.stringify(session)
    .replace(/</g, '\\u003c')
    .replace(new RegExp(String.fromCharCode(0x2028), 'g'), '\\u2028')
    .replace(new RegExp(String.fromCharCode(0x2029), 'g'), '\\u2029');
  return template;
}

// ───────────────────────────────── Helpers ─────────────────────────────────

function today_() {
  const d = new Date();
  d.setHours(0, 0, 0, 0);
  return d;
}

function todayKey_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

// Returns a Date at local midnight, or null if the value isn't a usable date.
function toMidnight_(raw) {
  let d = null;
  if (raw instanceof Date) d = new Date(raw.getTime());
  else if (typeof raw === 'string' && raw.trim() !== '') d = new Date(raw);
  if (!d || isNaN(d.getTime())) return null;
  d.setHours(0, 0, 0, 0);
  return d;
}

function cleanRef_(raw) {
  if (raw === '' || raw === null || raw === undefined) return '';
  return String(raw).trim();
}

// Dates travel through session storage as milliseconds (Date objects can't be returned to the page).
function dateToMs_(v) {
  const d = toMidnight_(v);
  return d ? d.getTime() : '';
}

function msToDate_(ms) {
  return (ms === '' || ms === null || ms === undefined) ? '' : new Date(ms);
}

function readSheetData_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Verses');
  const lastRow = sheet ? sheet.getLastRow() : 0;
  return lastRow >= 2 ? sheet.getRange(1, 1, lastRow, 8).getValues() : [];
}

// "Day 1" is the day a verse was added to Learning; column C holds that date while the verse is Level 0.
function learningDay_(addedRaw) {
  const added = toMidnight_(addedRaw);
  if (!added) return null;
  return Math.max(1, Math.round((today_().getTime() - added.getTime()) / 86400000) + 1);
}

// How many Level 1+ verses will be in tomorrow's deck (everything due by the end of tomorrow).
function dueTomorrowCount_(data) {
  const limit = today_();
  limit.setDate(limit.getDate() + 1);
  let n = 0;
  for (let i = 1; i < data.length; i++) {
    if (!cleanRef_(data[i][5])) continue;
    const groupStr = String(data[i][4] === null || data[i][4] === undefined ? '' : data[i][4]).trim();
    if (groupStr === '' || isNaN(Number(groupStr)) || Number(groupStr) < 1) continue;
    const nextDue = toMidnight_(data[i][1]);
    if (!nextDue || nextDue <= limit) n++;
  }
  return n;
}

// Verses added before "Day N" existed have no date; stamp them with today so counting starts.
function stampLearningDates_(preRead) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Verses');
  if (!sheet) return;
  const data = preRead || readSheetData_();
  const today = today_();
  for (let i = 1; i < data.length; i++) {
    if (!cleanRef_(data[i][5])) continue;
    const groupStr = String(data[i][4] === null || data[i][4] === undefined ? '' : data[i][4]).trim();
    if (groupStr === '' || Number(groupStr) !== 0) continue;
    if (!toMidnight_(data[i][2])) { sheet.getRange(i + 1, 3).setValue(today); data[i][2] = today; }
  }
}

// ───────────────────────────────── Reads ─────────────────────────────────

function getDueCards(includeNew) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Verses');
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(1, 1, lastRow, 8).getValues();
  const today = today_();
  const dueCards = [];

  for (let i = 1; i < data.length; i++) {
    const nextDueRaw = data[i][1]; // Col B
    const repRaw     = data[i][3]; // Col D
    const groupRaw   = data[i][4]; // Col E
    const ref        = cleanRef_(data[i][5]); // Col F
    const versionRaw = data[i][6]; // Col G
    const text       = data[i][7]; // Col H

    if (!ref) continue;

    const groupStr = String(groupRaw === null || groupRaw === undefined ? '' : groupRaw).trim();
    if (groupStr === '') continue;               // unassigned verse
    const groupNum = Number(groupStr);
    if (isNaN(groupNum) || groupNum < 0) continue;
    const group = Math.min(Math.floor(groupNum), MAX_GROUP);

    const rep = Number(repRaw) || 1;

    let isDue;
    if (group === 0) {
      if (!includeNew) continue;
      isDue = true;
    } else {
      // A blank or unreadable date means "due", so a card can never silently disappear.
      const nextDue = toMidnight_(nextDueRaw);
      isDue = !nextDue || nextDue <= today;
    }

    if (isDue) {
      const groupInfo = GROUP_CONFIG[group];
      dueCards.push({
        rowIndex: i + 1,
        reference: ref,
        version: versionRaw ? String(versionRaw).trim() : "",
        text: text ? String(text) : "(No verse text found in Column H)",
        group: group,
        groupName: groupInfo.name,
        rep: rep,
        repsNeeded: groupInfo.repsNeeded
      });
    }
  }

  dueCards.sort((a, b) => a.group - b.group || a.rowIndex - b.rowIndex);
  return dueCards;
}

// ─────────────────────────── Daily session storage ───────────────────────────
//
// The session is stored compactly (no verse text) in Script Properties, split into
// small chunks because each property value is limited to about 9 KB.
//
//   { date, version, baseCount, currentIndex, processed,
//     stats: { reviewed, passed, failed, graduated },
//     queue: [ { r: row, f: reference, g: level, p: rep, l: 1 if same-day retry, h: 1 if a held repeat } ] }
//     (held repeats are failed verses waiting at the end of the queue until the Practice button starts them)

const SESSION_KEY = 'VR_SESSION';
const SESSION_VER_KEY = 'VR_SESSION_V'; // tiny "date|version" marker so polling needs only one read

function store_() {
  return PropertiesService.getScriptProperties();
}
const SESSION_CHUNK = 3000;
const MAX_QUEUE = 2000;

function readState_() {
  const props = store_();
  const n = Number(props.getProperty(SESSION_KEY + '_N'));
  if (!n) return null;
  const all = props.getProperties();
  let json = '';
  for (let i = 0; i < n; i++) {
    const part = all[SESSION_KEY + '_' + i];
    if (part === null || part === undefined) return null;
    json += part;
  }
  try {
    const s = JSON.parse(json);
    return (s && Array.isArray(s.queue)) ? s : null;
  } catch (err) {
    return null;
  }
}

function writeState_(s) {
  const props = store_();
  const json = JSON.stringify(s);
  const oldCount = Number(props.getProperty(SESSION_KEY + '_N')) || 0;

  const batch = {};
  let count = 0;
  for (let i = 0; i < json.length; i += SESSION_CHUNK) {
    batch[SESSION_KEY + '_' + count] = json.substring(i, i + SESSION_CHUNK);
    count++;
  }
  batch[SESSION_KEY + '_N'] = String(count);
  batch[SESSION_VER_KEY] = s.date + '|' + s.version;
  const laterTimes = (Array.isArray(s.later) ? s.later : []).map(function(e) { return e.at; });
  batch[LATER_KEY] = laterTimes.length ? String(Math.min.apply(null, laterTimes)) : '0';
  props.setProperties(batch);
  for (let i = count; i < oldCount; i++) props.deleteProperty(SESSION_KEY + '_' + i);
}

// ── Evening practice card ──
// After the first review of a Bronze verse (Level 1, rep 1) is passed, one extra practice card is held
// back and added to the deck later that day: at 5 pm, or an hour after the pass if that is already past 5 pm.
// It behaves like a same-day repeat: never written to the sheet, no rep, no demotion.
const LATER_KEY = 'VR_LATER_AT'; // earliest release time (ms), so polling only needs one tiny read
const LATER_HOUR = 17;
const LATER_DELAY_MS = 60 * 60 * 1000;

function nowMs_() { return Date.now(); }

// When the evening card should appear, or 0 if that would be after today ends.
function laterAt_() {
  const tz = Session.getScriptTimeZone();
  const now = nowMs_();
  const nowD = new Date(now);
  const secs = Number(Utilities.formatDate(nowD, tz, 'H')) * 3600 + Number(Utilities.formatDate(nowD, tz, 'm')) * 60 + Number(Utilities.formatDate(nowD, tz, 's'));
  const at = secs < LATER_HOUR * 3600 ? now + (LATER_HOUR * 3600 - secs) * 1000 : now + LATER_DELAY_MS;
  return Utilities.formatDate(new Date(at), tz, 'yyyy-MM-dd') === Utilities.formatDate(nowD, tz, 'yyyy-MM-dd') ? at : 0;
}

// Moves any evening cards whose time has come onto the end of the deck. Returns true if it changed anything.
function releaseLater_(s, now) {
  if (!s || !Array.isArray(s.later) || s.later.length === 0) return false;
  const due = s.later.filter(function(e) { return e.at <= now; });
  if (due.length === 0) return false;
  // Evening cards go in front of any held repeats, which stay at the very end until practice is started.
  let at = s.queue.findIndex(function(q) { return q.h; });
  if (at < 0) at = s.queue.length;
  due.forEach(function(e, i) { s.queue.splice(at + i, 0, { r: e.r, f: e.f, g: 1, p: 2, l: 1 }); });
  s.later = s.later.filter(function(e) { return e.at > now; });
  delete s.undo; // the deck changed, so the old undo no longer lines up
  s.version++;
  return true;
}

function compactCard_(c) {
  return { r: c.rowIndex, f: c.reference, g: c.group, p: c.rep, l: c.isLapse ? 1 : 0 };
}

// Where the active deck ends: held repeats (h) sit after it and are not part of the deck until practice starts.
function activeEnd_(s) {
  const i = s.queue.findIndex(function(q) { return q.h; });
  return i < 0 ? s.queue.length : i;
}

function emptyStats_() {
  return { reviewed: 0, passed: 0, failed: 0, graduated: 0, failedRefs: [], gradRefs: [] };
}

function createState_(dateKey) {
  const cards = getDueCards();
  const s = {
    date: dateKey,
    version: 1,
    baseCount: cards.length,
    currentIndex: 0,
    processed: 0,
    mode: 'review',
    stats: emptyStats_(),
    queue: cards.map(compactCard_)
  };
  writeState_(s);
  return s;
}

function findRowByRef_(sheet, ref) {
  const last = sheet.getLastRow();
  if (last < 2) return 0;
  const vals = sheet.getRange(2, 6, last - 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (cleanRef_(vals[i][0]) === ref) return i + 2;
  }
  return 0;
}

// Turns the compact stored session into full cards (verse text etc.) from the sheet.
// If rows were sorted or inserted, cards are matched back up by reference.
// Also returns the current Level 0 "learning" verses and the last screen you were on.
function seenFor_(s, learning) {
  const live = sanitizeSeen_(s.seen);
  const out = {};
  learning.forEach(function(c) { if (live[c.reference]) out[c.reference] = live[c.reference]; });
  return out;
}

function hydrate_(s, preRead) {
  const data = preRead || readSheetData_();

  const byRef = {};
  const learning = [];
  for (let i = 1; i < data.length; i++) {
    const ref = cleanRef_(data[i][5]);
    if (!ref) continue;
    if (!Object.prototype.hasOwnProperty.call(byRef, ref)) byRef[ref] = i;

    const groupStr = String(data[i][4] === null || data[i][4] === undefined ? '' : data[i][4]).trim();
    if (groupStr !== '' && Number(groupStr) === 0) {
      learning.push({
        rowIndex: i + 1,
        reference: ref,
        version: data[i][6] ? String(data[i][6]).trim() : "",
        text: data[i][7] ? String(data[i][7]) : "(No verse text found in Column H)",
        group: 0,
        groupName: GROUP_CONFIG[0].name,
        rep: Number(data[i][3]) || 1,
        repsNeeded: GROUP_CONFIG[0].repsNeeded,
        isLapse: false,
        day: learningDay_(data[i][2])
      });
    }
  }

  let idx = s.currentIndex;
  let base = s.baseCount;
  const cards = [];
  s.queue.forEach(function(e, qi) {
    let i = e.r - 1;
    if (!(i >= 1 && i < data.length && cleanRef_(data[i][5]) === e.f)) {
      i = Object.prototype.hasOwnProperty.call(byRef, e.f) ? byRef[e.f] : -1;
    }
    // Drop verses no longer in the sheet, and Level 0 verses left over from before
    // learning verses moved to their own screen.
    if (i < 1 || e.g === 0) {
      if (qi < s.currentIndex) idx--;
      if (qi < s.baseCount) base--;
      return;
    }
    const row = data[i];
    const info = GROUP_CONFIG[e.g] || GROUP_CONFIG[1];
    cards.push({
      rowIndex: i + 1,
      reference: e.f,
      version: row[6] ? String(row[6]).trim() : "",
      text: row[7] ? String(row[7]) : "(No verse text found in Column H)",
      group: e.g,
      groupName: info.name,
      rep: e.p,
      repsNeeded: info.repsNeeded,
      isLapse: !!e.l,
      isHeld: !!e.h
    });
  });

  return {
    build: BUILD,
    version: s.version,
    date: s.date,
    mode: s.mode === 'learning' ? 'learning' : 'review',
    baseCount: Math.max(0, Math.min(base, cards.length)),
    currentIndex: Math.max(0, Math.min(idx, cards.length)),
    processed: s.processed,
    stats: s.stats,
    cards: cards,
    learning: learning,
    seen: seenFor_(s, learning),
    tomorrow: dueTomorrowCount_(data)
  };
}

/**
 * Returns today's session, creating it the first time any device asks today.
 */
function loadSession() {
  const t0 = Date.now();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const today = todayKey_();
    let s = readState_();
    let data = readSheetData_();
    stampLearningDates_(data);
    if (!s || s.date !== today) { if (sortVerses_()) data = readSheetData_(); s = createState_(today); }
    else if (releaseLater_(s, nowMs_())) writeState_(s);
    const out = hydrate_(s, data);
    out.ms = Date.now() - t0;
    return out;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Used by "Check for Due Cards". If today's deck is finished, starts a new batch of
 * whatever is due now (for example a newly added verse) while keeping today's stats.
 * If the deck is not finished, it just returns the current session.
 */
function refreshSession() {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const today = todayKey_();
    let s = readState_();
    stampLearningDates_();
    let fresh = false;
    if (!s || s.date !== today) {
      sortVerses_();
      s = createState_(today);
      fresh = true;
    } else if (releaseLater_(s, nowMs_())) {
      writeState_(s); // an evening practice card is now due
    }
    if (!fresh && s.currentIndex >= activeEnd_(s)) {
      const cards = getDueCards();
      const held = s.queue.filter(function(q) { return q.h; }); // unpracticed repeats carry over to the new batch
      s.queue = cards.map(compactCard_).concat(held);
      s.baseCount = cards.length;
      s.currentIndex = 0;
      delete s.undo; // the old queue is gone, so there is nothing to undo
      s.version++;
      writeState_(s);
    }
    return hydrate_(s);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Cheap poll so a device can tell whether another device has moved the session along.
 */
function checkSession(clientVersion, clientDate) {
  const laterAt = Number(store_().getProperty(LATER_KEY) || 0);
  if (laterAt && nowMs_() >= laterAt) {
    const lock = LockService.getScriptLock();
    let got = true;
    try { lock.waitLock(5000); } catch (e) { got = false; }
    if (got) {
      try {
        const st = readState_();
        if (st && st.date === todayKey_() && releaseLater_(st, nowMs_())) writeState_(st);
        else store_().setProperty(LATER_KEY, '0'); // nothing (more) to release; stop checking
      } finally {
        lock.releaseLock();
      }
    }
  }
  const marker = store_().getProperty(SESSION_VER_KEY); // "date|version"
  const today = todayKey_();
  const changed = !marker ||
                  marker !== clientDate + '|' + Number(clientVersion) ||
                  marker.split('|')[0] !== today;
  return { changed: changed, build: BUILD, marker: marker || '', today: today };
}

// Run this from the Apps Script editor to see what the server has stored (View -> Logs).
function debugSession() {
  const s = readState_();
  if (!s) { Logger.log('No session stored. Today is ' + todayKey_()); return; }
  Logger.log(JSON.stringify({
    build: BUILD, today: todayKey_(), storedDate: s.date, version: s.version,
    mode: s.mode || 'review', currentIndex: s.currentIndex, queueLength: s.queue.length, baseCount: s.baseCount,
    processed: s.processed, stats: s.stats, marker: store_().getProperty(SESSION_VER_KEY)
  }));
}

// ───────────────────────────────── Writes ─────────────────────────────────

// Writes one review result to the sheet. The server computes all dates, so the phone's
// clock and time zone don't matter. If rows moved, the verse is found again by reference.
// On success, also returns the row's previous values so the review can be undone.
function writeReview_(review) {
  const rowIndex = Number(review.rowIndex);
  const newRep = Number(review.rep);
  const newGroup = Number(review.group);

  if (!Number.isInteger(rowIndex) || rowIndex < 2 ||
      !Number.isInteger(newGroup) || newGroup < 0 || newGroup > MAX_GROUP ||
      !Number.isInteger(newRep) || newRep < 1) {
    return { ok: false, reason: 'invalid', message: 'Invalid review data.' };
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Verses');
  if (!sheet) return { ok: false, reason: 'nosheet', message: 'Sheet "Verses" not found.' };

  const ref = cleanRef_(review.reference);
  let row = rowIndex;
  const rowOk = row <= sheet.getLastRow() && cleanRef_(sheet.getRange(row, 6).getValue()) === ref;
  if (!rowOk) {
    row = findRowByRef_(sheet, ref);
    if (!row) return { ok: false, reason: 'moved', message: 'That verse is no longer in the sheet.' };
  }

  const old = sheet.getRange(row, 2, 1, 4).getValues()[0];
  const before = [dateToMs_(old[0]), dateToMs_(old[1]), old[2], old[3]];

  let nextDateVal = "";
  let lastDateVal = "";
  if (newGroup > 0) {
    const today = today_();
    const next = new Date(today.getTime());
    next.setDate(next.getDate() + GROUP_CONFIG[newGroup].interval);
    nextDateVal = next;
    lastDateVal = today;
  }

  sheet.getRange(row, 2, 1, 4).setValues([[nextDateVal, lastDateVal, newRep, newGroup]]);
  return { ok: true, row: { r: row, f: ref, before: before } };
}

function sanitizeQueue_(q) {
  if (!Array.isArray(q) || q.length > MAX_QUEUE) return null;
  const out = [];
  for (let i = 0; i < q.length; i++) {
    const e = q[i] || {};
    const r = Number(e.r), g = Number(e.g), p = Number(e.p);
    if (!Number.isInteger(r) || r < 2 || !Number.isInteger(g) || g < 0 || g > MAX_GROUP ||
        !Number.isInteger(p) || p < 1 || typeof e.f !== 'string' || !e.f) return null;
    const o = { r: r, f: e.f, g: g, p: p, l: e.l ? 1 : 0 };
    if (e.h) o.h = 1;
    out.push(o);
  }
  return out;
}

// New-verse "Reviewed" marks: { reference: time (ms) until which that verse is counted as looked at }.
function sanitizeSeen_(obj) {
  const out = {};
  if (!obj || typeof obj !== 'object') return out;
  const now = Date.now();
  let n = 0;
  Object.keys(obj).forEach(function(k) {
    const t = Number(obj[k]);
    if (n < 200 && k && k.length <= 120 && isFinite(t) && t > now && t < now + 24 * 3600 * 1000) { out[k] = t; n++; }
  });
  return out;
}

function sanitizeStats_(st) {
  st = st || {};
  const n = function(v) { v = Number(v); return (Number.isInteger(v) && v >= 0) ? v : 0; };
  // Verses behind the Repeats / Leveled Up numbers: { f: reference, g: level it now sits at }.
  // Plain strings (older sessions) are accepted.
  const cleanRefs = function(list) {
    const out = [];
    const seen = Object.create(null);
    (Array.isArray(list) ? list : []).forEach(function(r) {
      const f = String((r && typeof r === 'object') ? (r.f || '') : (r || '')).slice(0, 120);
      if (!f || seen[f] || out.length >= 100) return;
      seen[f] = true;
      const g = (r && typeof r === 'object' && Number.isInteger(r.g) && r.g >= 0 && r.g <= 8) ? r.g : null;
      out.push(g === null ? { f: f } : { f: f, g: g });
    });
    return out;
  };
  return { reviewed: n(st.reviewed), passed: n(st.passed), failed: n(st.failed), graduated: n(st.graduated), failedRefs: cleanRefs(st.failedRefs), gradRefs: cleanRefs(st.gradRefs) };
}

/**
 * Saves one scored card: writes the review to the sheet (review is null for same-day
 * retries, which are never written) and stores the session's new state.
 *
 * `expectedVersion` is the session version this device last saw. If another device has
 * moved the session since then, nothing is written and the current session is returned
 * so this device can catch up.
 *
 * A scored card also records what is needed to undo it (see undoLast). Switching between
 * the Review and Learning screens does not.
 */
// ───────────── Review log ─────────────
// Every scored card is added as a row on a "Log" tab (created automatically on first use).
// Nothing in the app reads it yet; it is for looking back on later. A logging problem never blocks a review.
const LOG_HEADERS = ['Time', 'Reference', 'Result', 'Level Before', 'Level After', 'Rep Before', 'Rep After', 'Try Today', 'Same-Day Repeat'];
const LOG_SHEET = 'Log';

function getLogSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(LOG_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(LOG_SHEET);
    sheet.appendRow(LOG_HEADERS);
    try { sheet.setFrozenRows(1); } catch (e) {}
    try { sheet.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm:ss'); } catch (e) {}
  }
  return sheet;
}

// info: { ref, result, lb, la, rb, ra, tries, repeat }. Returns { t, f } so Undo can remove the row.
function logEvent_(info) {
  try {
    const sheet = getLogSheet_();
    const now = new Date();
    sheet.appendRow([
      now, info.ref, info.result,
      info.lb, info.la, info.rb, info.ra,
      info.tries, info.repeat ? 'Yes' : ''
    ]);
    return { t: now.getTime(), f: info.ref };
  } catch (e) {
    console.error('Log write failed: ' + e);
    return null;
  }
}

function unlogEvent_(logRef) {
  try {
    if (!logRef) return;
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET);
    if (!sheet) return;
    const last = sheet.getLastRow();
    const first = Math.max(2, last - 40);
    if (last < first) return;
    const vals = sheet.getRange(first, 1, last - first + 1, 2).getValues();
    for (let i = vals.length - 1; i >= 0; i--) {
      const t = vals[i][0] instanceof Date ? vals[i][0].getTime() : Number(vals[i][0]);
      if (t === logRef.t && cleanRef_(vals[i][1]) === logRef.f) {
        sheet.deleteRow(first + i);
        return;
      }
    }
  } catch (e) {
    console.error('Log undo failed: ' + e);
  }
}

function submitReview(expectedVersion, review, next) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000); // throws on timeout; the client keeps the review and offers a retry
  try {
    const today = todayKey_();
    const s = readState_();

    if (!s || s.date !== today || s.version !== Number(expectedVersion)) {
      const current = (s && s.date === today) ? s : createState_(today);
      return { ok: false, reason: 'conflict', session: hydrate_(current) };
    }

    const queue = sanitizeQueue_(next && next.queue);
    const idx = Number(next && next.currentIndex);
    if (!queue || !Number.isInteger(idx) || idx < 0 || idx > queue.length) {
      return { ok: false, reason: 'invalid', message: 'Invalid session data.' };
    }

    let written = null;
    if (review) {
      const res = writeReview_(review);
      if (!res.ok) return res;
      written = res.row;
    }

    const prev = { idx: s.currentIndex, qlen: s.queue.length, stats: s.stats, processed: s.processed };
    const isScore = !!review || idx !== prev.idx || queue.length !== prev.qlen;
    const prevCard = s.queue[s.currentIndex] || null;

    s.queue = queue;
    s.currentIndex = idx;
    s.baseCount = Math.max(0, Math.min(Number(next.baseCount) || 0, queue.length));
    s.processed = Math.max(0, Number(next.processed) || 0);
    s.stats = sanitizeStats_(next.stats);
    s.mode = (next.mode === 'learning') ? 'learning' : 'review';
    if (next.seen !== undefined) s.seen = sanitizeSeen_(next.seen);
    s.version++;

    let logRef = null;
    if (isScore) {
      const pushed = Math.max(0, queue.length - prev.qlen);
      const ref = review ? cleanRef_(review.reference) : (prevCard ? prevCard.f : '');
      if (ref) {
        if (!s.tries) s.tries = {};
        s.tries[ref] = (s.tries[ref] || 0) + 1;
        let info;
        if (review && written) {
          const lb = Number(written.before[3]) || 0;
          info = {
            ref: ref,
            result: lb === 0 ? 'Level Up' : (pushed > 0 ? 'Fail' : 'Pass'),
            lb: lb, la: Number(review.group), rb: Number(written.before[2]) || 0, ra: Number(review.rep),
            tries: s.tries[ref], repeat: false
          };
        } else if (prevCard) {
          info = {
            ref: ref,
            result: pushed > 0 ? 'Fail' : 'Pass',
            lb: prevCard.g, la: prevCard.g, rb: prevCard.p, ra: prevCard.p,
            tries: s.tries[ref], repeat: true
          };
        }
        if (info) logRef = logEvent_(info);
      }
    }

    let laterRef = null;
    if (review && review.later === true && Number(review.group) === 1 && Number(review.rep) === 2 && written) {
      const at = laterAt_();
      if (at) {
        s.later = (Array.isArray(s.later) ? s.later : []).filter(function(e) { return e.f !== written.f; });
        s.later.push({ r: written.r, f: written.f, at: at });
        laterRef = written.f;
      }
    }

    let undoId = null;
    if (isScore) {
      undoId = s.version;
      s.undo = {
        id: undoId,
        at: Date.now(),
        idx: prev.idx,
        pushed: Math.max(0, queue.length - prev.qlen),
        stats: prev.stats,
        processed: prev.processed,
        row: written,
        log: logRef,
        tryRef: logRef ? logRef.f : null,
        later: laterRef
      };
    }

    writeState_(s);
    const result = { ok: true, version: s.version, undoId: undoId };
    if (review) result.tomorrow = dueTomorrowCount_(readSheetData_());
    return result;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Undoes the most recent scored card (anywhere): puts its sheet row back as it was,
 * and restores the session's position, retries and stats. Only works for a short time
 * after scoring, and only if nothing has been scored since.
 */
function undoLast(undoId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const today = todayKey_();
    let s = readState_();
    if (!s || s.date !== today) s = createState_(today);

    const u = s.undo;
    if (!u || u.id !== Number(undoId) || (Date.now() - u.at) > UNDO_WINDOW_MS) {
      return { ok: false, reason: 'nothing', session: hydrate_(s) };
    }

    if (u.row) {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Verses');
      if (sheet) {
        let row = u.row.r;
        const rowOk = row <= sheet.getLastRow() && cleanRef_(sheet.getRange(row, 6).getValue()) === u.row.f;
        if (!rowOk) row = findRowByRef_(sheet, u.row.f);
        if (row) {
          const b = u.row.before;
          sheet.getRange(row, 2, 1, 4).setValues([[msToDate_(b[0]), msToDate_(b[1]), b[2], b[3]]]);
        }
      }
    }

    unlogEvent_(u.log);
    if (u.tryRef && s.tries && s.tries[u.tryRef]) s.tries[u.tryRef]--;
    if (u.later && Array.isArray(s.later)) s.later = s.later.filter(function(e) { return e.f !== u.later; });

    for (let i = 0; i < u.pushed; i++) s.queue.pop();
    s.currentIndex = u.idx;
    s.stats = u.stats;
    s.processed = u.processed;
    delete s.undo;
    s.version++;
    writeState_(s);
    return { ok: true, session: hydrate_(s) };
  } finally {
    lock.releaseLock();
  }
}

// Verses with no level yet, in sheet order, for the "Add New Verse" picker.
function listUnassigned() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Verses');
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const data = sheet.getRange(1, 1, lastRow, 8).getValues();
  const out = [];
  const todayMs = today_().getTime();
  for (let i = 1; i < data.length; i++) {
    const g = data[i][4];
    const ref = cleanRef_(data[i][5]);
    if (ref && (g === "" || g === null || g === undefined)) {
      // Days since the date in column C (when it was added), or null if there is none.
      const added = toMidnight_(data[i][2]);
      const waiting = added ? Math.max(0, Math.round((todayMs - added.getTime()) / 86400000)) : null;
      out.push({ r: i + 1, f: ref, v: String(data[i][6] || '').trim(), t: String(data[i][7] || '').trim(), d: waiting });
    }
  }
  return out;
}

// Starts learning the verse you picked (Level 0, Day 1 = today).
function addChosenVerse(rowIndex, reference) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Verses');
    if (!sheet) return { success: false, message: 'Sheet "Verses" not found.' };
    const want = cleanRef_(reference);
    let row = Number(rowIndex);
    const lastRow = sheet.getLastRow();
    if (!(row >= 2 && row <= lastRow) || cleanRef_(sheet.getRange(row, 6).getValue()) !== want) {
      row = findRowByRef_(sheet, want);
    }
    if (!row) return { success: false, message: 'That verse is no longer in the sheet.' };
    const g = sheet.getRange(row, 5).getValue();
    if (!(g === "" || g === null || g === undefined)) {
      return { success: false, message: 'That verse was already added.' };
    }
    sheet.getRange(row, 2, 1, 4).setValues([["", today_(), 1, 0]]);
    return { success: true, reference: want, rowIndex: row };
  } finally {
    lock.releaseLock();
  }
}

// ── Keeping the Verses tab sorted ──
// Order: verses with no level yet first, then Level 0 (New), then Level 1 up to 8.
// Within a level, Bible order (book, chapter, verse). Runs once a day, when the new day's deck is built. Rows are matched by reference
// everywhere else, so moving them never writes to the wrong verse.
// Books in Bible order, with the names and abbreviations a reference might start with.
const BIBLE_BOOKS_ = [
  ["Genesis", "gen|ge|gn"], ["Exodus", "exod|exo|ex"], ["Leviticus", "lev|le|lv"], ["Numbers", "num|nu|nm|nb"],
  ["Deuteronomy", "deut|dt|de"], ["Joshua", "josh|jos|jsh"], ["Judges", "judg|jdg|jg|jdgs"], ["Ruth", "rth|ru"],
  ["1 Samuel", "1 sam|1 sm|1 sa|1sam|first samuel"], ["2 Samuel", "2 sam|2 sm|2 sa|2sam|second samuel"],
  ["1 Kings", "1 kgs|1 ki|1kgs|first kings"], ["2 Kings", "2 kgs|2 ki|2kgs|second kings"],
  ["1 Chronicles", "1 chron|1 chr|1 ch|1chr|first chronicles"], ["2 Chronicles", "2 chron|2 chr|2 ch|2chr|second chronicles"],
  ["Ezra", "ezr"], ["Nehemiah", "neh|ne"], ["Esther", "esth|est|es"], ["Job", "jb"],
  ["Psalms", "psalm|ps|psa|pss|psm|pslm"], ["Proverbs", "prov|pro|prv|pr"], ["Ecclesiastes", "eccles|eccl|ecc|ec|qoh"],
  ["Song of Solomon", "song of songs|song|sos|ss|canticles|cant"], ["Isaiah", "isa|is"], ["Jeremiah", "jer|je|jr"],
  ["Lamentations", "lam|la"], ["Ezekiel", "ezek|eze|ezk"], ["Daniel", "dan|da|dn"], ["Hosea", "hos|ho"],
  ["Joel", "jl"], ["Amos", "am"], ["Obadiah", "obad|ob"], ["Jonah", "jon|jnh"], ["Micah", "mic|mi"],
  ["Nahum", "nah|na"], ["Habakkuk", "hab|hb"], ["Zephaniah", "zeph|zep|zp"], ["Haggai", "hag|hg"],
  ["Zechariah", "zech|zec|zc"], ["Malachi", "mal|ml"],
  ["Matthew", "matt|mat|mt"], ["Mark", "mrk|mk|mr"], ["Luke", "luk|lk"], ["John", "jhn|jn"], ["Acts", "act|ac"],
  ["Romans", "rom|ro|rm"], ["1 Corinthians", "1 cor|1 co|1cor|first corinthians"], ["2 Corinthians", "2 cor|2 co|2cor|second corinthians"],
  ["Galatians", "gal|ga"], ["Ephesians", "eph|ephes"], ["Philippians", "phil|php|pp"], ["Colossians", "col|co"],
  ["1 Thessalonians", "1 thess|1 thes|1 th|1thess|first thessalonians"], ["2 Thessalonians", "2 thess|2 thes|2 th|2thess|second thessalonians"],
  ["1 Timothy", "1 tim|1 ti|1tim|first timothy"], ["2 Timothy", "2 tim|2 ti|2tim|second timothy"],
  ["Titus", "tit|ti"], ["Philemon", "philem|phlm|phm"], ["Hebrews", "heb"], ["James", "jas|jm"],
  ["1 Peter", "1 pet|1 pe|1pet|1 pt|first peter"], ["2 Peter", "2 pet|2 pe|2pet|2 pt|second peter"],
  ["1 John", "1 jn|1 jhn|1john|first john"], ["2 John", "2 jn|2 jhn|2john|second john"], ["3 John", "3 jn|3 jhn|3john|third john"],
  ["Jude", "jud"], ["Revelation", "rev|re|revelations|apocalypse"]
];

const BOOK_INDEX_ = (function() {
  const m = {};
  BIBLE_BOOKS_.forEach(function(b, i) {
    m[normBook_(b[0])] = i;
    b[1].split('|').forEach(function(a) { m[normBook_(a)] = i; });
  });
  return m;
})();

function normBook_(t) {
  return String(t || '').toLowerCase().replace(/\./g, '').replace(/\s+/g, ' ').trim()
    .replace(/^(1st|i)\s+/, '1 ').replace(/^(2nd|ii)\s+/, '2 ').replace(/^(3rd|iii)\s+/, '3 ')
    .replace(/^([123])(?=[a-z])/, '$1 ');
}

// "John 10:27-30" -> { book: index in the Bible (999 if unknown), chapter, verse }
function parseRef_(ref) {
  const m = String(ref || '').match(/^\s*((?:(?:[1-3]|1st|2nd|3rd)\.?\s*|(?:iii|ii|i)\.?\s+)?[A-Za-z][A-Za-z.\s]*?)\s*(\d+)(?:\s*[:.]\s*(\d+))?/i);
  let name = ref, chapter = 0, verse = 0;
  if (m) { name = m[1]; chapter = parseInt(m[2], 10) || 0; verse = parseInt(m[3], 10) || 0; }
  let idx = BOOK_INDEX_[normBook_(name)];
  if (idx === undefined) idx = 999;
  return { book: idx, chapter: chapter, verse: verse };
}

function sortVerses_() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Verses');
    if (!sheet) return;
    const last = sheet.getLastRow();
    const cols = sheet.getLastColumn();
    if (last < 3 || cols < 8) return;
    const range = sheet.getRange(2, 1, last - 1, cols);
    const vals = range.getValues();
    const forms = range.getFormulas();
    const items = vals.map(function(row, i) {
      const ref = cleanRef_(row[5]);
      const str = String(row[4] === null || row[4] === undefined ? '' : row[4]).trim();
      const hasLevel = str !== '' && !isNaN(Number(str));
      const p = parseRef_(ref);
      return {
        i: i,
        empty: !ref,
        level: hasLevel ? Math.max(0, Math.min(MAX_GROUP, Math.floor(Number(str)))) : -1, // unassigned verses sort first
        book: p.book, chapter: p.chapter, verse: p.verse, ref: ref
      };
    });
    const sorted = items.slice().sort(function(a, b) {
      return (a.empty - b.empty) || (a.level - b.level) || (a.book - b.book) ||
        (a.chapter - b.chapter) || (a.verse - b.verse) ||
        (a.ref < b.ref ? -1 : a.ref > b.ref ? 1 : 0) || (a.i - b.i);
    });
    if (sorted.every(function(it, k) { return it.i === k; })) return; // already in order
    const out = sorted.map(function(it) {
      return vals[it.i].map(function(v, c) { return forms[it.i][c] || v; });
    });
    range.setValues(out);
    return true;
  } catch (e) {
    console.error('Sort failed: ' + e);
  }
  return false;
}

// Adds a pasted verse as a new row at the top of the Verses tab: reference, version and text filled in,
// Level left blank (unassigned), and the date it was added in column C.
function addPastedVerse(reference, version, text) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Verses');
    if (!sheet) return { success: false, message: 'Sheet "Verses" not found.' };
    const ref = cleanRef_(reference).replace(/\s+/g, ' ');
    const ver = cleanRef_(version).toUpperCase() || 'ESV';
    const body = String(text === null || text === undefined ? '' : text).replace(/\s+/g, ' ').trim();
    if (!ref || ref.length > 100 || !/^[A-Za-z0-9]/.test(ref) || !/\d/.test(ref)) return { success: false, message: 'That reference does not look right.' };
    if (ver.length > 12 || !/^[A-Z0-9]+$/.test(ver)) return { success: false, message: 'That version does not look right.' };
    if (!body) return { success: false, message: 'No verse text found.' };
    if (body.length > 20000 || /^[=+\-@]/.test(body)) return { success: false, message: 'That verse text cannot be saved.' };
    const last = sheet.getLastRow();
    if (findRowByRef_(sheet, ref)) return { success: false, message: ref + ' is already in your sheet.' };
    const cols = Math.max(8, sheet.getLastColumn());
    if (last >= 2) {
      sheet.insertRowsBefore(2, 1);
      try { // new row takes the look of the row below it, not the header
        sheet.getRange(3, 1, 1, cols).copyTo(sheet.getRange(2, 1, 1, cols), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      } catch (e) {}
    }
    sheet.getRange(2, 1, 1, 8).setValues([['', '', today_(), '', '', ref, ver, body]]);
    return { success: true, reference: ref };
  } finally {
    lock.releaseLock();
  }
}

// For the Verse List screen: every verse you've started (Level 0 and up). Unassigned verses are left out.
function getVerseList() {
  const data = readSheetData_();
  const verses = [];
  const todayMs = today_().getTime();
  for (let i = 1; i < data.length; i++) {
    const ref = cleanRef_(data[i][5]);
    if (!ref) continue;
    const raw = data[i][4];
    const str = String(raw === null || raw === undefined ? '' : raw).trim();
    if (str === '' || isNaN(Number(str))) continue;
    const g = Math.max(0, Math.min(MAX_GROUP, Math.floor(Number(str))));
    // First three words of the verse text (verse-number markers like [3] removed).
    const words = String(data[i][7] || '').replace(/\[\d+\]/g, ' ').replace(/\s+/g, ' ').trim().split(' ').filter(Boolean);
    const first = words.slice(0, 3).join(' ') + (words.length > 3 ? '…' : '');
    // Days from today until the next review (overdue shows as 0); null when there is no review date (Level 0).
    const nd = g > 0 ? toMidnight_(data[i][1]) : null;
    const due = nd ? Math.max(0, Math.round((nd.getTime() - todayMs) / 86400000)) : null;
    verses.push({ f: ref, g: g, v: String(data[i][6] || '').trim(), w: first, x: words.join(' '), d: due });
  }
  return { verses: verses };
}
