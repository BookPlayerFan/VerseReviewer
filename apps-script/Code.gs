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
const BUILD = 'sync-5'; // shown in the app's status readout (tap the "Card x of y" label)

// Intervals (days) are used only here. The client mirrors names and repsNeeded only.
const GROUP_CONFIG = {
  0: { name: "New",               interval: 0,   repsNeeded: 1 },
  1: { name: "Bronze",             interval: 1,   repsNeeded: 4 },
  2: { name: "Silver",             interval: 2,   repsNeeded: 3 },
  3: { name: "Gold",             interval: 5,   repsNeeded: 2 },
  4: { name: "Platinum",            interval: 12,  repsNeeded: 2 },
  5: { name: "Emerald",           interval: 30,  repsNeeded: 2 },
  6: { name: "Ruby",         interval: 75,  repsNeeded: 1 },
  7: { name: "Sapphire",         interval: 180, repsNeeded: 1 },
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
function stampLearningDates_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Verses');
  if (!sheet) return;
  const data = readSheetData_();
  const today = today_();
  for (let i = 1; i < data.length; i++) {
    if (!cleanRef_(data[i][5])) continue;
    const groupStr = String(data[i][4] === null || data[i][4] === undefined ? '' : data[i][4]).trim();
    if (groupStr === '' || Number(groupStr) !== 0) continue;
    if (!toMidnight_(data[i][2])) sheet.getRange(i + 1, 3).setValue(today);
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
//     queue: [ { r: row, f: reference, g: level, p: rep, l: 1 if same-day retry } ] }

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
  let json = '';
  for (let i = 0; i < n; i++) {
    const part = props.getProperty(SESSION_KEY + '_' + i);
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
  props.setProperties(batch);
  for (let i = count; i < oldCount; i++) props.deleteProperty(SESSION_KEY + '_' + i);
}

function compactCard_(c) {
  return { r: c.rowIndex, f: c.reference, g: c.group, p: c.rep, l: c.isLapse ? 1 : 0 };
}

function emptyStats_() {
  return { reviewed: 0, passed: 0, failed: 0, graduated: 0 };
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
function hydrate_(s) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Verses');
  const lastRow = sheet ? sheet.getLastRow() : 0;
  const data = lastRow >= 2 ? sheet.getRange(1, 1, lastRow, 8).getValues() : [];

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
      isLapse: !!e.l
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
    tomorrow: dueTomorrowCount_(data)
  };
}

/**
 * Returns today's session, creating it the first time any device asks today.
 */
function loadSession() {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const today = todayKey_();
    stampLearningDates_();
    let s = readState_();
    if (!s || s.date !== today) s = createState_(today);
    return hydrate_(s);
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
    stampLearningDates_();
    let s = readState_();
    if (!s || s.date !== today) {
      s = createState_(today);
    } else if (s.currentIndex >= s.queue.length) {
      const cards = getDueCards();
      s.queue = cards.map(compactCard_);
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
    out.push({ r: r, f: e.f, g: g, p: p, l: e.l ? 1 : 0 });
  }
  return out;
}

function sanitizeStats_(st) {
  st = st || {};
  const n = function(v) { v = Number(v); return (Number.isInteger(v) && v >= 0) ? v : 0; };
  return { reviewed: n(st.reviewed), passed: n(st.passed), failed: n(st.failed), graduated: n(st.graduated) };
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

    s.queue = queue;
    s.currentIndex = idx;
    s.baseCount = Math.max(0, Math.min(Number(next.baseCount) || 0, queue.length));
    s.processed = Math.max(0, Number(next.processed) || 0);
    s.stats = sanitizeStats_(next.stats);
    s.mode = (next.mode === 'learning') ? 'learning' : 'review';
    s.version++;

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
        row: written
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
  for (let i = 1; i < data.length; i++) {
    const g = data[i][4];
    const ref = cleanRef_(data[i][5]);
    if (ref && (g === "" || g === null || g === undefined)) {
      out.push({ r: i + 1, f: ref, v: String(data[i][6] || '').trim() });
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
