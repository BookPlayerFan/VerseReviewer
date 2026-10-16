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
const BUILD = 'sync-78'; // shown in the app's status readout (hold the Verses tab)

// The Sidebar.html build number the server expects. A page older than this is a stale copy (the phone kept an old one),
// so it offers "Update ready - tap to reload". The number is read from the deployed Sidebar.html itself (its CLIENT_BUILD line),
// so a new Sidebar.html is noticed by pages that are still open without touching this file. CLIENT_EXPECTED is only a floor,
// used if the file can't be read. Looked up at most every two minutes.
const CLIENT_EXPECTED = 152;
function clientExpected_() {
  try {
    const cache = CacheService.getScriptCache();
    const hit = Number(cache.get('VR_CLIENT_BUILD'));
    if (hit) return Math.max(CLIENT_EXPECTED, hit);
    const html = HtmlService.createHtmlOutputFromFile('Sidebar').getContent();
    const m = /const CLIENT_BUILD = 'sync-(\d+)'/.exec(html);
    const n = m ? Number(m[1]) : 0;
    if (n) cache.put('VR_CLIENT_BUILD', String(n), 120);
    return Math.max(CLIENT_EXPECTED, n);
  } catch (e) {
    return CLIENT_EXPECTED;
  }
}

// Intervals (days) are used only here. The client mirrors names and repsNeeded only.
// Load balancing. When a review sets a card's next date, Bronze, Silver and Gold always land on the exact interval (a short
// interval can't be stretched without hurting recall). Platinum and above may move a few days to a quieter day:
//   early / late = the most days the date may move earlier / later than the exact interval (the hard cap)
//   costEarly / costLate = the price, in "reviews", of each day moved. A day is chosen by (reviews already due that day) + price,
//   lowest wins, so a card only moves when another day is clearly quieter. Platinum is pickier (higher price) because 1 day is
//   about 8% of its 12-day interval. Later costs more than earlier because a late review risks forgetting.
const GROUP_CONFIG = {
  0: { name: "New",               interval: 0,   repsNeeded: 1, early: 0,  late: 0 },
  1: { name: "Bronze",             interval: 1,   repsNeeded: 3, early: 0,  late: 0 },
  2: { name: "Silver",             interval: 2,   repsNeeded: 3, early: 0,  late: 0 },
  3: { name: "Gold",             interval: 5,   repsNeeded: 2, early: 0,  late: 0 },
  4: { name: "Platinum",            interval: 12,  repsNeeded: 2, early: 1,  late: 1,  costEarly: 1,    costLate: 2 },
  5: { name: "Sapphire",           interval: 30,  repsNeeded: 2, early: 2,  late: 1,  costEarly: 0.25, costLate: 0.5 },
  6: { name: "Emerald",         interval: 75,  repsNeeded: 1, early: 3,  late: 2,  costEarly: 0.25, costLate: 0.5 },
  7: { name: "Ruby",         interval: 180, repsNeeded: 1, early: 5,  late: 3,  costEarly: 0.25, costLate: 0.5 },
  8: { name: "Diamond",           interval: 360, repsNeeded: 1, early: 7,  late: 4,  costEarly: 0.25, costLate: 0.5 }
};

// How much a verse counts toward a day's review load: its word count / 30, never less than 1 (opening and scoring a card has a
// fixed cost) and never more than 5 (one very long passage cannot dominate). Prices below are in "short verses" of this load.
const LOAD_WORDS = 30;
const LOAD_MAX = 5;
function verseLoad_(text) {
  const n = String(text || '').replace(/\[\d+\]/g, ' ').split(/\s+/).filter(Boolean).length;
  return Math.max(1, Math.min(LOAD_MAX, n / LOAD_WORDS));
}

// The due date for a card moving to `group`. `skipRow` is the card being scheduled, so its own old date is not counted;
// New cards (level 0) and cards with no date are not counted either. A day's load is the sum of verseLoad_ of its verses. On a tie the exact day wins, then the earlier day.
function balancedDueDate_(sheet, group, today, skipRow) {
  const cfg = GROUP_CONFIG[group];
  const base = new Date(today.getTime());
  base.setDate(base.getDate() + cfg.interval);
  const early = cfg.early || 0, late = cfg.late || 0;
  if (!early && !late) return base;
  const last = sheet.getLastRow();
  const counts = {};
  if (last >= 2) {
    const vals = sheet.getRange(2, 2, last - 1, 7).getValues(); // B..H: next date, last date, rep, level, reference, version, text
    for (let i = 0; i < vals.length; i++) {
      if (i + 2 === skipRow) continue;
      const d = toMidnight_(vals[i][0]);
      if (!d || !(Number(vals[i][3]) > 0)) continue;
      const k = d.getTime();
      counts[k] = (counts[k] || 0) + verseLoad_(vals[i][6]);
    }
  }
  let best = base, bestScore = Infinity;
  const reach = Math.max(early, late);
  for (let step = 0; step <= reach * 2; step++) {
    const off = step === 0 ? 0 : (step % 2 === 1 ? -Math.ceil(step / 2) : Math.ceil(step / 2));
    if (off < -early || off > late) continue;
    const c = new Date(base.getTime());
    c.setDate(c.getDate() + off);
    if (c.getTime() <= today.getTime()) continue;
    const score = (counts[c.getTime()] || 0) + (off < 0 ? -off * cfg.costEarly : off * cfg.costLate);
    if (score < bestScore) { best = c; bestScore = score; }
  }
  return best;
}

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

// ───────────── New-verse stages: Learn -> Lock -> Bronze ─────────────
//
// A Level 0 verse is a "New" card. Its Rep (column D) is the rep you are ON, which is also the filled dot:
//   rep 1 .. 7n   Learn   (n = number of verses; verse = ceil(rep / 7), dot = ((rep - 1) mod 7) + 1)
//   rep 7n + 1    Lock AM
//   rep 7n + 2    Lock PM  (its button is Promote: Level 1, Rep 1)
// Column B (Next due) holds the time the button next unlocks (empty = live now). Level 0 rows ignore it everywhere else.
// Learn unlocks at fixed times of the day, one press per time, nothing builds up if one is missed.
const LEARN_SLOTS = [4 * 60, 8 * 60, 10 * 60, 12 * 60, 14 * 60, 16 * 60, 18 * 60]; // minutes after midnight
const REPS_PER_VERSE = 7;
const DAY_START_MIN = 4 * 60;       // a new day of Learn/Lock starts at 4 AM
const LOCK_PM_MIN = 18 * 60;        // Lock PM unlocks at 6 PM
const START_LIVE_BEFORE_MIN = 12 * 60; // a verse added before noon starts today; later it waits for 4 AM tomorrow
// A rep is "late" an hour after it opens (the 4 AM rep: after 7 AM; the 6 PM rep and Lock PM: after 8:30 PM; Lock AM: after noon). The dots turn amber (one late rep
// today) or red (two or more). Late reps are not made up, so the mark stays on a verse until 4 AM the next day. Kept per reference in VR_LATE_MAP.
const LATE_GRACE_MIN = 60;
const LATE_FIRST_SLOT_MIN = 7 * 60;
const LATE_LOCK_AM_MIN = 12 * 60;
const LATE_LOCK_PM_MIN = 20 * 60 + 30;
const LATE_MAP_KEY = 'VR_LATE_MAP';

// The time (ms) of a clock time (minutes after midnight), dayOffset days from the day of `ms`.
function atClock_(ms, dayOffset, mins) {
  const d = new Date(ms);
  d.setDate(d.getDate() + dayOffset);
  d.setHours(Math.floor(mins / 60), mins % 60, 0, 0);
  return d.getTime();
}

function minsOfDay_(ms) {
  const d = new Date(ms);
  return d.getHours() * 60 + d.getMinutes();
}

// The next Learn unlock time after `ms` (today's list, else tomorrow's first).
function nextSlotAfter_(ms) {
  for (let i = 0; i < LEARN_SLOTS.length; i++) {
    const t = atClock_(ms, 0, LEARN_SLOTS[i]);
    if (t > ms) return t;
  }
  return atClock_(ms, 1, LEARN_SLOTS[0]);
}

// The time a rep that opened at `slotMs` counts as late.
function lateDeadline_(slotMs, stage, lockPart) {
  const m = minsOfDay_(slotMs);
  if (stage === 'lock') {
    const lim = lockPart === 'pm' ? LATE_LOCK_PM_MIN : LATE_LOCK_AM_MIN;
    return m < lim ? atClock_(slotMs, 0, lim) : slotMs + LATE_GRACE_MIN * 60000;
  }
  if (m === LEARN_SLOTS[0]) return atClock_(slotMs, 0, LATE_FIRST_SLOT_MIN);
  if (m === LEARN_SLOTS[LEARN_SLOTS.length - 1]) return atClock_(slotMs, 0, LATE_LOCK_PM_MIN); // the 6 PM rep: its reminder goes out at 8 PM, so it is late after 8:30
  return slotMs + LATE_GRACE_MIN * 60000;
}

// For a card that opens (or opened) at readyMs: { late: reps already late, next: the next time a rep will count as late (0 = none) }.
// A Learn card has one rep per slot from readyMs on, and each slot whose deadline has gone by without a press is one late rep.
function lateInfo_(stage, lockPart, readyMs, now) {
  const out = { late: 0, next: 0 };
  if (!readyMs) return out;
  if (stage === 'lock') {
    const dl = lateDeadline_(readyMs, stage, lockPart);
    if (dl <= now) out.late = 1; else out.next = dl;
    return out;
  }
  for (let d = 0; d < 21; d++) {
    for (let i = 0; i < LEARN_SLOTS.length; i++) {
      const t = atClock_(readyMs, d, LEARN_SLOTS[i]);
      if (t < readyMs) continue;
      const dl = lateDeadline_(t, stage, lockPart);
      if (dl <= now) out.late++;
      else if (!out.next) { out.next = dl; return out; }
    }
    if (atClock_(readyMs, d + 1, LEARN_SLOTS[0]) > now + 86400000) break;
  }
  return out;
}

// The day for "missed today" starts at 4 AM, like the Learn day.
function lateDayKey_(ms) {
  return Utilities.formatDate(new Date(ms - DAY_START_MIN * 60000), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}
function lateMap_() {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty(LATE_MAP_KEY) || '{}') || {}; } catch (e) { return {}; }
}
// Reps missed today for one reference (from presses that came late).
function missedToday_(ref, now) {
  const e = lateMap_()[ref];
  return e && e.d === lateDayKey_(now) ? (Number(e.n) || 0) : 0;
}
function addMissed_(ref, n, now) {
  const map = lateMap_(), day = lateDayKey_(now), out = {};
  Object.keys(map).forEach(function(k) { if (map[k] && map[k].d === day) out[k] = map[k]; });
  if (n > 0) out[ref] = { d: day, n: (out[ref] ? Number(out[ref].n) || 0 : 0) + n };
  PropertiesService.getScriptProperties().setProperty(LATE_MAP_KEY, JSON.stringify(out));
}
function clearMissed_(ref, now) {
  const map = lateMap_();
  if (!map[ref]) return;
  delete map[ref];
  PropertiesService.getScriptProperties().setProperty(LATE_MAP_KEY, JSON.stringify(map));
}

// The next 4 AM after `ms`.
function nextDayStart_(ms) {
  const t = atClock_(ms, 0, DAY_START_MIN);
  return t > ms ? t : atClock_(ms, 1, DAY_START_MIN);
}

// When a verse added at `ms` can first be pressed: 0 = now (before noon), otherwise the next 4 AM.
function startUnlock_(ms) {
  const m = minsOfDay_(ms);
  if (m < DAY_START_MIN) return atClock_(ms, 0, DAY_START_MIN);
  if (m < START_LIVE_BEFORE_MIN) return 0;
  return nextDayStart_(ms);
}

// Number of verses in a reference, from the [n] markers in its text (minimum 1).
function verseCount_(text) {
  const t = String(text === null || text === undefined ? '' : text);
  const m = t.match(/\[(\d+)\]/g);
  if (!m) return 1;
  const seen = {};
  let n = 0;
  m.forEach(function(x) { if (!seen[x]) { seen[x] = 1; n++; } });
  // The first verse of a passage has no [n] marker of its own: any text before the first marker is that verse.
  if (t.substring(0, t.search(/\[\d+\]/)).trim()) n++;
  return Math.max(1, n);
}

// Treasury points count verses, not cards: a reference's points are its rank number times its verse count.
// { reference: verse count } for every verse on the sheet (a reference no longer on the sheet counts as 1).
function verseCountMap_() {
  const map = {};
  try {
    const data = readSheetData_();
    for (let i = 1; i < data.length; i++) { const ref = cleanRef_(data[i][5]); if (ref) map[ref] = verseCount_(data[i][7]); }
  } catch (e) { console.error('verseCountMap_ failed: ' + e); }
  return map;
}

// Everything the screen needs to know about a New card's position.
function stageInfo_(rep, n, bRaw, now) {
  const r = Math.max(1, Math.floor(Number(rep) || 1));
  const total = REPS_PER_VERSE * n;
  let unlockAt = 0, readyAt = 0;
  if (bRaw instanceof Date) {
    const t = bRaw.getTime();
    if (!isNaN(t)) { readyAt = t; if (t > now) unlockAt = t; }
  }
  const out = { stage: 'learn', verseIndex: 1, verseCount: n, dot: 1, lockPart: '', label: 'Review', unlockAt: unlockAt, readyAt: readyAt };
  if (r <= total) {
    out.verseIndex = Math.ceil(r / REPS_PER_VERSE);
    out.dot = ((r - 1) % REPS_PER_VERSE) + 1;
    if (out.dot === REPS_PER_VERSE) out.label = out.verseIndex < n ? 'Next Verse' : 'Learned';
  } else {
    out.stage = 'lock';
    out.verseIndex = n;
    out.dot = 0;
    if (r === total + 1) { out.lockPart = 'am'; out.label = 'Review'; }
    else { out.lockPart = 'pm'; out.label = 'Promote'; }
  }
  return out;
}

// What a press does. Returns { rep, unlock } (unlock 0 = live now) or { promote: true }.
function nextStep_(info, rep, now) {
  if (info.stage === 'learn') {
    if (info.dot < REPS_PER_VERSE) return { rep: rep + 1, unlock: nextSlotAfter_(now) };
    return { rep: rep + 1, unlock: nextDayStart_(now) }; // Next Verse or Learned: starts again at 4 AM tomorrow
  }
  if (info.lockPart === 'am') {
    const pm = atClock_(now, 0, LOCK_PM_MIN);
    return { rep: rep + 1, unlock: pm > now ? pm : 0 };
  }
  return { promote: true };
}

// Skip Verse (before the last verse) or Skip to Lock (last verse). Learn cards only.
function skipStep_(info, n, now) {
  if (info.verseIndex < n) return { rep: REPS_PER_VERSE * info.verseIndex + 1, unlock: 0 };
  return { rep: REPS_PER_VERSE * n + 1, unlock: nextDayStart_(now) };
}

// A Level 0 row as a card for the screen.
function learningCard_(row, i, now) {
  const ref = cleanRef_(row[5]);
  const n = verseCount_(row[7]);
  const rep = Number(row[3]) || 1;
  const card = {
    rowIndex: i + 1,
    reference: ref,
    version: row[6] ? String(row[6]).trim() : "",
    text: row[7] ? String(row[7]) : "(No verse text found in Column H)",
    group: 0,
    groupName: GROUP_CONFIG[0].name,
    rep: rep,
    repsNeeded: GROUP_CONFIG[0].repsNeeded,
    isRepeat: false,
    day: learningDay_(row[2])
  };
  const info = stageInfo_(rep, n, row[1], now);
  Object.keys(info).forEach(function(k) { card[k] = info[k]; });
  const li = lateInfo_(info.stage, info.lockPart, info.readyAt, now);
  card.late = li.late;            // reps already late right now (the open rep and any slots that went by)
  card.lateAt = li.next;          // when the next one will count as late (the screen repaints then)
  card.missed = missedToday_(ref, now); // reps that were pressed late earlier today
  return card;
}

function isLevelZero_(raw) {
  const s = String(raw === null || raw === undefined ? '' : raw).trim();
  return s !== '' && Number(s) === 0;
}

// Unlocked New cards right now, split by stage: { learn, lock } counts, plus learnRefs / lockRefs (references, sheet order).
function newCardCounts_(data, now) {
  const out = { learn: 0, lock: 0, learnRefs: [], lockRefs: [] };
  for (let i = 1; i < data.length; i++) {
    if (!cleanRef_(data[i][5]) || !isLevelZero_(data[i][4])) continue;
    const c = learningCard_(data[i], i, now);
    if (c.unlockAt) continue;
    if (c.stage === 'lock') { out.lock++; out.lockRefs.push(c.reference); } else { out.learn++; out.learnRefs.push(c.reference); }
  }
  return out;
}

/**
 * A press on the New card's button.
 *   action 'press'  - the button (refused with 'locked' until its unlock time)
 *   action 'now'    - the same step, ignoring the lock (hold menu)
 *   action 'skip'   - Skip Verse / Skip to Lock (Learn cards only)
 * expectedRep is the rep the screen showed, so a second phone that is behind gets 'stale' and the current card.
 * Lock PM is Promote: the verse becomes Bronze rep 1, due tomorrow.
 */
function learnAction(reference, action, expectedRep) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Verses');
    if (!sheet) return { ok: false, reason: 'nosheet', message: 'Sheet "Verses" not found.' };
    const ref = cleanRef_(reference);
    const row = ref ? findRowByRef_(sheet, ref) : 0;
    if (!row) return { ok: false, reason: 'moved', message: 'That verse is no longer in the sheet.' };
    const vals = sheet.getRange(row, 1, 1, 8).getValues()[0];
    if (!isLevelZero_(vals[4])) return { ok: false, reason: 'notnew', message: 'That verse is not in Learn or Lock.' };
    if (action !== 'press' && action !== 'now' && action !== 'skip') return { ok: false, reason: 'invalid', message: 'Unknown action.' };

    const now = nowMs_();
    const n = verseCount_(vals[7]);
    const rep = Number(vals[3]) || 1;
    if (expectedRep !== undefined && expectedRep !== null && Number(expectedRep) !== rep) {
      return { ok: false, reason: 'stale', card: learningCard_(vals, row - 1, now) };
    }
    const info = stageInfo_(rep, n, vals[1], now);

    let step;
    if (action === 'skip') {
      if (info.stage !== 'learn') return { ok: false, reason: 'noskip', card: learningCard_(vals, row - 1, now) };
      step = skipStep_(info, n, now);
    } else {
      if (action === 'press' && info.unlockAt) {
        return { ok: false, reason: 'locked', unlockAt: info.unlockAt, card: learningCard_(vals, row - 1, now) };
      }
      step = nextStep_(info, rep, now);
    }

    // A rep pressed after its grace time is remembered for the rest of the day (the dots stay amber / red).
    try {
      if (action !== 'skip') {
        const li = lateInfo_(info.stage, info.lockPart, info.readyAt, now);
        if (li.late > 0 && !step.promote) addMissed_(ref, li.late, now);
      }
      if (step.promote) clearMissed_(ref, now);
    } catch (e) { console.error('late record failed: ' + e); }

    let promoted = false;
    let card = null;
    if (step.promote) {
      const w = writeReview_({ rowIndex: row, reference: ref, rep: 1, group: 1 });
      if (!w.ok) return w;
      promoted = true;
      logEvent_({ ref: ref, result: 'Level Up', lb: 0, la: 1, rb: rep, ra: 1, tries: '', repeat: false });
    } else {
      sheet.getRange(row, 2).setValue(step.unlock ? new Date(step.unlock) : '');
      sheet.getRange(row, 4).setValue(step.rep);
      vals[1] = step.unlock ? new Date(step.unlock) : '';
      vals[3] = step.rep;
      logEvent_({ ref: ref, result: action === 'skip' ? 'Skip' : 'Dot', lb: 0, la: 0, rb: rep, ra: step.rep, tries: '', repeat: false });
      card = learningCard_(vals, row - 1, now);
    }

    // Other phones notice the change through the session version, like any other review.
    const s = readState_();
    if (s && s.date === todayKey_()) {
      if (promoted) {  // a new card in Bronze adds one Treasury point for each of its verses
        const pv = verseCount_(vals[7]);
        s.stats = sanitizeStats_(s.stats); s.stats.pts += pv;
        const mine = s.stats.ptsRefs.filter(function(x) { return x.f === ref; })[0];
        if (mine) { mine.d += pv; mine.g = 1; if (mine.d === 0) s.stats.ptsRefs.splice(s.stats.ptsRefs.indexOf(mine), 1); }
        else s.stats.ptsRefs.push({ f: ref, g: 1, d: pv });
      }
      s.version++; writeState_(s);
    }

    // The day's Learn / Lock pushes are kept in line with the new state: the last rep of the day, Skip and Promote drop what is left.
    try { syncSlotPushes_(); } catch (e) { pushLog_('slot sync ERROR', ref, 0, false, String(e)); }

    const result = { ok: true, promoted: promoted, card: card, version: s && s.date === todayKey_() ? s.version : undefined };
    result.tomorrow = dueTomorrowCount_(readSheetData_());
    return result;
  } finally {
    lock.releaseLock();
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
// back and added to the deck later that day: at 6 pm, or an hour after the pass if that is already past 6 pm, but never
// later than 8:30 pm (passed after 7:30 pm -> 8:30 pm; passed after 8:30 pm -> no evening card). A push goes out when the card appears.
// It behaves like a same-day repeat: never written to the sheet, no rep, no demotion.
const LATER_KEY = 'VR_LATER_AT'; // earliest release time (ms), so polling only needs one tiny read
const LATER_HOUR = 18;
const LATER_DELAY_MS = 60 * 60 * 1000;

function nowMs_() { return Date.now(); }

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
  return { r: c.rowIndex, f: c.reference, g: c.group, p: c.rep, l: c.isRepeat ? 1 : 0 };
}

// Where the active deck ends: held repeats (h) sit after it and are not part of the deck until practice starts.
function activeEnd_(s) {
  const i = s.queue.findIndex(function(q) { return q.h; });
  return i < 0 ? s.queue.length : i;
}

// PTS_VER marks a day's Treasury points as counted per verse. A session without it (counted by card, or before points existed) is recounted from the Log.
const PTS_VER = 2;
function emptyStats_() {
  return { reviewed: 0, passed: 0, failed: 0, graduated: 0, pts: 0, ptsV: PTS_VER, ptsRefs: [], failedRefs: [], gradRefs: [], clearedRefs: [] };
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
  const nowMs = nowMs_();

  const byRef = {};
  const learning = [];
  for (let i = 1; i < data.length; i++) {
    const ref = cleanRef_(data[i][5]);
    if (!ref) continue;
    if (!Object.prototype.hasOwnProperty.call(byRef, ref)) byRef[ref] = i;

    if (isLevelZero_(data[i][4])) learning.push(learningCard_(data[i], i, nowMs));
  }
  // Unlocked cards first, Lock before Learn, then sheet order.
  learning.sort(function(a, b) {
    return ((a.unlockAt > 0 ? 1 : 0) - (b.unlockAt > 0 ? 1 : 0)) ||
      ((a.stage === 'lock' ? 0 : 1) - (b.stage === 'lock' ? 0 : 1)) || (a.rowIndex - b.rowIndex);
  });

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
      isRepeat: !!e.l,
      isHeld: !!e.h
    });
  });

  return {
    build: BUILD,
    client: clientExpected_(),
    serverNow: nowMs,
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
// A session whose points were not counted per verse (none yet, or counted by card) is recounted once. Work today's out from the Log tab: the net level change per verse
// (each row's Level After minus Level Before, same-day repeats left out, times the verses in the card), which also counts a promoted new card as +1 per verse.
function backfillPts_(s) {
  if (!s || !s.stats || s.stats.ptsV === PTS_VER) return false;
  const per = {}, vc = verseCountMap_();
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET);
    if (sheet && sheet.getLastRow() >= 2) {
      const last = sheet.getLastRow(), first = Math.max(2, last - 799);
      const vals = sheet.getRange(first, 1, last - first + 1, 9).getValues();
      const tz = Session.getScriptTimeZone(), day = todayKey_();
      vals.forEach(function(v) {
        const t = v[0] instanceof Date ? v[0] : new Date(v[0]);
        if (isNaN(t.getTime()) || Utilities.formatDate(t, tz, 'yyyy-MM-dd') !== day) return;
        if (String(v[8] || '') === 'Yes') return;
        const lb = Number(v[3]), la = Number(v[4]), ref = cleanRef_(v[1]);
        if (!ref || isNaN(lb) || isNaN(la) || lb === la) return;
        const e = per[ref] || (per[ref] = { f: ref, g: la, d: 0 });
        e.d += (la - lb) * (vc[ref] || 1); e.g = la;
      });
    }
  } catch (e) { console.error('backfillPts_ failed: ' + e); }
  const refs = Object.keys(per).map(function(k) { return per[k]; }).filter(function(e) { return e.d !== 0; });
  s.stats.ptsRefs = refs;
  s.stats.pts = refs.reduce(function(a, e) { return a + e.d; }, 0);
  s.stats.ptsV = PTS_VER;
  return true;
}

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
    if (s && s.date === today && backfillPts_(s)) writeState_(s);
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
  return { changed: changed, build: BUILD, client: clientExpected_(), marker: marker || '', today: today };
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
    nextDateVal = balancedDueDate_(sheet, newGroup, today, row);
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
  // ptsRefs: what the Treasury points came from, { f: reference, g: level it now sits at, d: net points from it today (never 0) }.
  const cleanPts = function(list) {
    const out = [], seen = Object.create(null);
    (Array.isArray(list) ? list : []).forEach(function(r) {
      const f = String((r && r.f) || '').slice(0, 120), d = Number(r && r.d), g = Number(r && r.g);
      if (!f || seen[f] || out.length >= 100 || !Number.isInteger(d) || d === 0 || d < -999 || d > 999) return;
      seen[f] = true;
      out.push({ f: f, g: (Number.isInteger(g) && g >= 0 && g <= 8) ? g : 0, d: d });
    });
    return out;
  };
  // pts: Treasury points added today (can be negative after misses): the net change in levels, +1 for each new verse promoted to Bronze.
  const pts = Number.isInteger(Number(st.pts)) ? Math.max(-999, Math.min(999, Number(st.pts))) : 0;
  return { reviewed: n(st.reviewed), passed: n(st.passed), failed: n(st.failed), graduated: n(st.graduated), pts: pts, ptsV: Number(st.ptsV) === PTS_VER ? PTS_VER : 0, ptsRefs: cleanPts(st.ptsRefs), failedRefs: cleanRefs(st.failedRefs), gradRefs: cleanRefs(st.gradRefs), clearedRefs: cleanRefs(st.clearedRefs) };
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
    // Starting a practice round (next.practice) changes the queue but is not a score: no log row, no undo.
    const isScore = !next.practice && (!!review || idx !== prev.idx || queue.length !== prev.qlen);
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
        tryRef: logRef ? logRef.f : null
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

// Starts learning the verse you picked (Level 0). Its first rep is live now if it is before noon, otherwise it unlocks at 4 AM tomorrow.
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
    // One verse in Learn at a time: refuse if another Level 0 verse is still in its Learn reps (a verse in Lock does not block).
    // This also stops a phone with an out-of-date screen from adding a second.
    if (lastRow >= 2) {
      const all = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
      for (let i = 0; i < all.length; i++) {
        if (!cleanRef_(all[i][5]) || !isLevelZero_(all[i][4])) continue;
        if ((Number(all[i][3]) || 1) <= REPS_PER_VERSE * verseCount_(all[i][7])) {
          return { success: false, message: 'You already have a verse in Learn. Finish its reps first, then add the next one.' };
        }
      }
    }
    // Starts today if added before noon, otherwise at 4 AM tomorrow (before 4 AM: at 4 AM).
    const unlock = startUnlock_(nowMs_());
    sheet.getRange(row, 2, 1, 4).setValues([[unlock ? new Date(unlock) : "", today_(), 1, 0]]);
    return { success: true, reference: want, rowIndex: row, unlockAt: unlock };
  } finally {
    lock.releaseLock();
  }
}

// Removes a verse that has no level yet (still in the Choose a verse list) from the Verses tab.
// Started verses (Level 0 and up) are never deleted here. Rows are matched by reference, like everywhere else.
function deleteUnassignedVerse(rowIndex, reference) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Verses');
    if (!sheet) return { success: false, message: 'Sheet "Verses" not found.' };
    const want = cleanRef_(reference);
    if (!want) return { success: false, message: 'No verse given.' };
    let row = Number(rowIndex);
    const lastRow = sheet.getLastRow();
    if (!(row >= 2 && row <= lastRow) || cleanRef_(sheet.getRange(row, 6).getValue()) !== want) {
      row = findRowByRef_(sheet, want);
    }
    if (!row) return { success: false, message: 'That verse is no longer in the sheet.' };
    const g = sheet.getRange(row, 5).getValue();
    if (!(g === '' || g === null || g === undefined)) {
      return { success: false, message: 'That verse has already been started, so it was not deleted.' };
    }
    sheet.deleteRow(row);
    return { success: true, reference: want };
  } finally {
    lock.releaseLock();
  }
}

// ───────────── Verse context (AI) ─────────────
// Hold a verse reference in the app to see a short background note (summary, setting, where it fits).
// It is written by Claude through the Anthropic API, using the key saved in Script Properties as ANTHROPIC_KEY.
// Each verse is asked for once. The answer is kept in column I of that verse's row on the Verses tab and reused after that.
// Clear the cell, or use Regenerate in the app, to get a fresh one.
const CONTEXT_MODEL = 'claude-sonnet-5-5'; // a cheaper, faster choice is 'claude-haiku-5-5'
const CONTEXT_COL = 9;                     // column I on the Verses tab
const CONTEXT_OLD_SHEET = 'Context';       // an earlier build kept answers on this tab; they are reused if it exists
// Claude thinks before it answers and that thinking counts toward this limit, so it needs plenty of room above the ~700 tokens the note itself takes.
// If an answer is still cut off, it is asked once more with double the room (see askClaude_).
const CONTEXT_MAX_TOKENS = 8000;
const CONTEXT_LAST_KEY = 'VR_CTX_LAST'; // last API problem, for troubleshooting

const CONTEXT_SYSTEM = 'You are a careful Bible study helper inside a personal Scripture-memorization app. ' +
  'Give the reader background that makes one passage come alive and easier to remember, the way a knowledgeable friend would explain it. ' +
  'The most useful thing you can give is the story or argument that leads into the passage, and where the passage sits within its chapter and its book. ' +
  'Be concrete and specific: name the people and the situation, refer to chapters and verses, say what the key Hebrew or Greek word meant, and say what the passage\'s own words picture. Prefer particular detail over general statements. ' +
  'Be accurate: use widely agreed information about the author, audience, date, occasion, genre, and the shape of the book. Never invent details. ' +
  'When something is genuinely uncertain, say so briefly where it matters. Do not open with disclaimers about authorship or date, and do not hedge everything. ' +
  'Do not argue denominational or doctrinal positions; where translations or scholars differ in a way that changes how the passage reads, mention only the one or two most important differences, briefly, without taking a side. ' +
  'Write in plain, warm language. No markdown and no bullet points. ' +
  'Keep the whole answer to about 250 to 300 words. Up to 300 is fine, but never over it. Be selective: choose the details that help most and leave the rest out, rather than covering everything.';

// Worked examples set the depth, detail and tone.
const CONTEXT_EXAMPLES = [
  ['Job 19:25-27 (ESV)',
   'SUMMARY: In the middle of his worst suffering, Job declares that he knows his “Redeemer” lives and will stand up for him at last. He expects to see God for himself, as a friend rather than an enemy, and the thought leaves him almost overwhelmed with longing.\n' +
   'LEAD-UP: Job is a righteous man who has lost his children, his wealth and his health. His three friends insist he must have sinned to deserve this. Chapter 19 is Job’s answer to the second speech from Bildad, one of those friends. He has just described being abandoned by his family, his servants and his friends, and God seems to be against him. Then, in a sudden turn, he wishes his words could be carved in rock forever (vv. 23-24) and declares this hope.\n' +
   'KEY WORD: In Hebrew, “Redeemer” is the go’el, a family member who steps in to defend a relative, buy back lost property, or set a wrong right. Job pictures someone who will take his side in a courtroom and clear his name.\n' +
   'WORTH KNOWING: Verse 26 is hard to translate, and Bible versions differ. “In my flesh I shall see God” could mean seeing God in a body after death, or seeing God from beyond death, or being vindicated while still alive. Many Christians read the passage as pointing to resurrection hope and to Christ. Many scholars think Job’s immediate hope is that God himself will vindicate him, even if that comes after he dies.\n' +
   'KEY POINT: Job’s hope doesn’t depend on his circumstances getting better. It depends on who he believes God is.'],
  ['Romans 5:8 (ESV)',
   'SUMMARY: Paul says God\'s love is not just a feeling or a promise; it was shown in a specific act. Christ died for people while they were still sinners, not after they had cleaned up their lives.\n' +
   'LEAD-UP: Paul wrote to the Christians in Rome, a church he had not yet visited, most likely from Corinth around AD 57. In the first four chapters he argued that everyone, Jew and non-Jew alike, stands guilty before God and is put right with him by faith rather than by earning it, using Abraham as his proof. Chapter 5 opens with what that gift produces: peace with God and a hope that holds even in suffering (verses 1-5). Verse 6 then begins to show why that hope is secure: Christ died for the ungodly "at the right time."\n' +
   'WHERE IT SITS: Romans moves from the problem of everyone\'s sin (chapters 1-3), to the gift of being put right by faith (3-5), to life in the Spirit (6-8), and then on to God\'s plan for Israel and everyday Christian living. Verse 8 is the center of this chapter\'s argument. Right after it, in verses 9-10, Paul reasons that if God did this for his enemies, he will surely save those who are now his friends.\n' +
   'KEY WORD: The word "shows" means to demonstrate or put on display, so the cross is presented as the proof of God\'s love. Verse 7 sets up the contrast: hardly anyone would die for a righteous person, though someone might dare to die for a good one.\n' +
   'KEY POINT: God\'s love does not wait for us to deserve it; it came to us at our worst.'],
  ['Psalm 121:1-2 (ESV)',
   'SUMMARY: The psalm opens with a question: looking up at the hills, where will help come from? The answer comes at once. It will not come from the hills but from the LORD, who made heaven and earth.\n' +
   'LEAD-UP: This is one of the "Songs of Ascents" (Psalms 120-134), which pilgrims sang as they climbed toward Jerusalem for the festivals. The road ran through rugged hill country where travelers faced real dangers: heat, rough ground and robbers. The psalm before it, Psalm 120, is the cry of someone far from home in distress. Psalm 121 names no author, and it begins with a traveler looking up at that landscape and asking where help will come from.\n' +
   'WHERE IT SITS: Psalm 121 is an early step on the pilgrims\' journey, and its two opening verses are the question and the answer. The rest of the psalm (verses 3-8) unfolds the answer: the LORD does not sleep, he is a shade at your right hand, and he guards your going out and your coming in, "from this time forth and forevermore."\n' +
   'KEY WORD: The Hebrew word for "keep" or "guard" appears six times in these eight verses, so protection is the psalm\'s main theme.\n' +
   'WORTH KNOWING: "The hills" can be read as the dangers of the road, or as the hilltop shrines where people worshiped other gods. Either way, the point is the same: help comes from the Creator, not from the landscape.\n' +
   'KEY POINT: When the road ahead looks dangerous, the psalm turns your eyes past the obstacle to the one who made it.']
];

function contextPrompt_(ref, version, text) {
  let out = 'Here are worked examples of the depth, detail and tone wanted, for other passages. Match their style and length, but never reuse their wording for a different passage. ' +
    'Note that the WORTH KNOWING label is left out when there is nothing real to flag.\n\n';
  CONTEXT_EXAMPLES.forEach(function(ex) { out += '--- Example: ' + ex[0] + ' ---\n' + ex[1] + '\n\n'; });
  out += '--- Your turn ---\n' +
    'Passage: ' + ref + (version ? ' (' + version + ')' : '') + '\n' +
    'Text (any bracketed numbers are verse numbers): ' + text + '\n\n' +
    'Answer using exactly these labels, each at the start of its own line, in this order:\n' +
    'SUMMARY: two or three sentences, up to 50 words, in plain words on what the passage says and how its thought or feeling moves. Name its main figure or idea directly.\n' +
    'LEAD-UP: the story or argument that leads into these verses: who is speaking or writing, to whom, and what has been happening just before, in this chapter and the book so far. Be specific and refer to chapters and verses. Four to six sentences, up to about 100 words.\n' +
    'WHERE IT SITS: one sentence on what the whole book is and how it moves, then where this chapter and these verses fall in that movement, and what comes right after them. Up to 75 words. Leave this label out if the LEAD-UP already covers where the passage falls.\n' +
    'KEY WORD: what one important word or image means in the original language and why it matters. Two or three sentences, up to 45 words. Leave this label out only if no word stands out.\n' +
    'WORTH KNOWING: only if translations or scholars differ in a way that changes how the passage reads, or there is a common misunderstanding. Give only the one or two most important points, in up to about 70 words. If there is nothing like that, leave this label out completely.\n' +
    'KEY POINT: one or two short sentences, under 25 words in all, giving the takeaway, what to hold on to. Do not just repeat the summary.\n' +
    'All the sections together must stay at or under 300 words. Cut the least useful detail rather than go over.';
  return out;
}

function versesSheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Verses');
}

// The kept answer for a verse (found by reference, like everywhere else), or ''.
function readContext_(ref) {
  const sheet = versesSheet_();
  if (!sheet) return '';
  const row = findRowByRef_(sheet, ref);
  if (!row) return '';
  const v = sheet.getRange(row, CONTEXT_COL).getValue();
  return v === null || v === undefined ? '' : String(v).trim();
}

// Answers saved by an earlier build on a separate "Context" tab.
function oldContext_(ref) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONTEXT_OLD_SHEET);
    if (!sheet || sheet.getLastRow() < 2) return '';
    const want = cleanRef_(ref).toLowerCase();
    const vals = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    for (let i = 0; i < vals.length; i++) {
      if (cleanRef_(vals[i][0]).toLowerCase() === want && String(vals[i][1] || '').trim()) return String(vals[i][1]).trim();
    }
  } catch (e) {}
  return '';
}

function saveContext_(ref, text) {
  const lock = LockService.getScriptLock();
  let got = false;
  try { got = lock.tryLock(10000); } catch (e) {}
  if (!got) return false; // the answer is still shown; it just is not kept this time
  try {
    const sheet = versesSheet_();
    const row = sheet ? findRowByRef_(sheet, ref) : 0; // looked up again: rows may have moved while Claude was thinking
    if (!row) return false;
    sheet.getRange(row, CONTEXT_COL).setValue(text);
    return true;
  } catch (e) {
    console.error('Context save failed: ' + e);
    return false;
  } finally {
    lock.releaseLock();
  }
}

// One call to the Anthropic Messages API. Returns { ok, text, cut } or { ok:false, code, message }.
// cut is true when Claude ran out of room (max_tokens) before it finished, so the text stops part way.
function callClaude_(system, user, maxTokens) {
  const key = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_KEY');
  if (!key) return { ok: false, code: 'nokey', message: 'Add your API key to turn this on.' };
  let res;
  try {
    res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post',
      contentType: 'application/json',
      muteHttpExceptions: true,
      headers: { 'x-api-key': String(key).trim(), 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify({
        model: CONTEXT_MODEL,
        max_tokens: maxTokens,
        system: system,
        messages: [{ role: 'user', content: user }]
      })
    });
  } catch (err) {
    return { ok: false, code: 'net', message: 'Could not reach Claude. Check your connection and try again.' };
  }
  const code = res.getResponseCode();
  const body = res.getContentText();
  let json = null;
  try { json = JSON.parse(body); } catch (e) {}

  if (code >= 200 && code < 300 && json && Array.isArray(json.content)) {
    const text = json.content.filter(function(b) { return b && b.type === 'text' && b.text; })
      .map(function(b) { return b.text; }).join('\n').trim();
    try { if (json.usage) PropertiesService.getScriptProperties().setProperty('VR_CTX_USAGE', new Date().toISOString() + ' in ' + json.usage.input_tokens + ' out ' + json.usage.output_tokens + ' (thinking + answer) limit ' + maxTokens + ' stop ' + json.stop_reason); } catch (e) {}
    if (text) return { ok: true, text: text, cut: json.stop_reason === 'max_tokens' };
    const why = json.stop_details && json.stop_details.type === 'refusal' ? 'Claude declined to answer for this verse.' : 'Claude sent back an empty answer. Try again.';
    return { ok: false, code: 'empty', message: why };
  }

  if (code >= 200 && code < 300) return { ok: false, code: 'empty', message: 'Claude sent back an answer the app could not read. Try again.' };
  const apiMsg = json && json.error && json.error.message ? String(json.error.message) : '';
  const apiType = json && json.error && json.error.type ? String(json.error.type) : '';
  try { PropertiesService.getScriptProperties().setProperty(CONTEXT_LAST_KEY, new Date().toISOString() + ' ' + code + ' ' + body.slice(0, 300)); } catch (e) {}
  if (code === 401) return { ok: false, code: 'badkey', message: 'The API key was not accepted. Check ANTHROPIC_KEY in Script Properties.' };
  if (code === 402 || /credit balance|billing/i.test(apiMsg)) return { ok: false, code: 'nocredit', message: 'The API account is out of credit. Add credit in the Claude Console.' };
  if (code === 403) return { ok: false, code: 'forbidden', message: 'The API key is not allowed to do this. ' + apiMsg.slice(0, 160) };
  if (code === 429 || code === 529 || code >= 500) return { ok: false, code: 'busy', message: 'Claude is busy right now. Try again in a moment.' };
  return { ok: false, code: 'api', message: ('Claude could not answer (' + code + (apiType ? ' ' + apiType : '') + '). ' + apiMsg).slice(0, 260) };
}

// Asks for a note. A note that was cut off is never returned or kept: it is asked again with more room, and if that is still not enough it is an error.
function askClaude_(system, user) {
  let ans = callClaude_(system, user, CONTEXT_MAX_TOKENS);
  if (ans.ok && ans.cut) ans = callClaude_(system, user, CONTEXT_MAX_TOKENS * 2);
  if (ans.ok && ans.cut) {
    try { PropertiesService.getScriptProperties().setProperty(CONTEXT_LAST_KEY, new Date().toISOString() + ' cut off at ' + (CONTEXT_MAX_TOKENS * 2) + ' tokens'); } catch (e) {}
    return { ok: false, code: 'long', message: 'Claude ran out of room before it finished. Tap Try again.' };
  }
  if (ans.ok && /[A-Za-z0-9]$/.test(ans.text)) ans.text += '.'; // finished but no closing full stop: add one so it is never mistaken for a cut-off note later
  return ans;
}

// True for a kept note in this app's format (it starts with SUMMARY:) that stops part way: no KEY POINT, or it ends mid-sentence.
// Notes you type yourself in column I do not start with SUMMARY:, so they are never treated as cut off.
function contextCutOff_(text) {
  const t = String(text || '').trim();
  if (!/^SUMMARY\s*:/i.test(t)) return false;
  if (!/^KEY POINT\s*:/im.test(t)) return true;
  return !/[.!?…"'”’)\]]$/.test(t);
}

// reference/version/text come from the card. force = true asks again even if an answer is already kept.
function getVerseContext(reference, version, text, force) {
  const ref = cleanRef_(reference);
  if (!ref) return { ok: false, code: 'bad', message: 'No verse given.' };
  let cutKept = '';
  if (!force) {
    let kept = readContext_(ref);
    if (!kept) {
      kept = oldContext_(ref);
      if (kept) saveContext_(ref, kept); // move it to column I
    }
    if (kept && !contextCutOff_(kept)) return { ok: true, text: kept, cached: true };
    // a kept note that was cut off by an earlier build is asked for again below; if that fails it is still shown rather than lost
    cutKept = kept;
  }
  let body = String(text || '').trim();
  if (!body) {
    const sheet = versesSheet_();
    const row = sheet ? findRowByRef_(sheet, ref) : 0;
    if (row) body = cleanRef_(sheet.getRange(row, 8).getValue());
  }
  if (!body) return { ok: false, code: 'bad', message: 'That verse has no text to look up.' };
  body = body.slice(0, 4000);

  const ans = askClaude_(CONTEXT_SYSTEM, contextPrompt_(ref, cleanRef_(version), body));
  if (!ans.ok) return cutKept ? { ok: true, text: cutKept, cached: true, partial: true } : ans;
  const kept = saveContext_(ref, ans.text);
  return { ok: true, text: ans.text, cached: false, kept: kept, model: CONTEXT_MODEL };
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
    verses.push({ f: ref, g: g, v: String(data[i][6] || '').trim(), w: first, x: words.join(' '), t: String(data[i][7] || ''), d: due, l: Math.round(verseLoad_(data[i][7]) * 10) / 10 });
  }
  return { verses: verses };
}

// ─────────────────────────── Month line (Perfect Days and points this month) ───────────────────────────
//
// Worked out from the Log tab (one row per scored card), nothing extra is stored. For each day of the month:
//   [reviews (first tries), misses (including repeat tries), level-ups, new verses promoted to Bronze,
//    net change in level points, first-try misses]
// A Perfect Day is a day with at least one review and no miss. Learn / Lock presses (Dot, Skip) are not reviews.
function getMonthStats(year, month) {
  year = Number(year); month = Number(month);
  const tz = Session.getScriptTimeZone();
  const out = { y: year, m: month, days: {}, today: todayKey_() };
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return out;
  const vals = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
  const vc = verseCountMap_();
  vals.forEach(function(v) {
    const t = v[0] instanceof Date ? v[0] : new Date(v[0]);
    if (isNaN(t.getTime())) return;
    const k = Utilities.formatDate(t, tz, 'yyyy-MM-dd');
    if (Number(k.slice(0, 4)) !== year || Number(k.slice(5, 7)) !== month) return;
    const res = String(v[2] || '');
    const lb = Number(v[3]), la = Number(v[4]);
    if (isNaN(lb) || isNaN(la)) return;
    const repeat = String(v[8] || '') === 'Yes';
    const isReview = res === 'Pass' || res === 'Fail' || (res === 'Level Up' && lb >= 1);
    const isPromo = lb === 0 && la >= 1;
    if (!isReview && !isPromo) return;
    const dn = Number(k.slice(8, 10));
    const day = out.days[dn] || (out.days[dn] = [0, 0, 0, 0, 0, 0]);
    if (isReview && !repeat) day[0]++;
    if (res === 'Fail') { day[1]++; if (!repeat) day[5]++; }
    if (isReview && !repeat && la > lb && lb >= 1) day[2]++;
    if (isPromo) day[3]++;
    day[4] += (la - lb) * (vc[cleanRef_(v[1])] || 1);
  });
  return out;
}


// ─────────────────────────── Push reminders (OneSignal) ───────────────────────────
//
// Setup, once:  1) Script Property ONESIGNAL_KEY = your OneSignal REST API key.
//               2) Run installDailyPush() from the editor (creates the ~5:00 AM trigger; approve the permissions).
//               3) Run testPush() to send yourself a test.
// Two kinds of reminder:
//   - Daily: at 6:00 AM (scheduled by a trigger that runs at about 5:00), one push for due reviews plus any Learn / Lock card unlocked at 4 AM (dailyPush). Sent once a day.
//   - Learn / Lock unlocks: the 5 AM job (dailyPush) schedules the whole day's slot pushes up front (syncSlotPushes_): 8, 10, 12, 2, 4 PM,
//     and 8 PM for the 6 PM unlock. There is one push per time, naming every verse waiting then ("Learn A and 1 more", "Learn A, lock B"), so a rep
//     you forgot to press never leaves a gap. After each press the same sync runs again: it cancels what is no longer needed (the last rep of the
//     day, Skip, Promote) and adds anything missing. The 4 AM unlocks are covered by the daily push.

const ONESIGNAL_APP_ID = 'e9bc8efc-620e-49a8-9275-09f9d76cd272';
const PUSH_URL = 'https://bookplayerfan.github.io/VerseReviewer/';
const PUSH_SLOTMAP_KEY = 'VR_PUSH_SLOTMAP'; // { send time in ms: { id: notificationId, learn: [references], lock: [references] } }
const DAILY_PUSH_MIN = 6 * 60;               // the morning push arrives at 6:00 AM
const PUSH_DAILY_KEY = 'VR_PUSH_DAILY';     // date key of the last daily push
const PUSH_LAST_KEY = 'VR_PUSH_LAST';       // last API answer, for troubleshooting

// ───────────── Push log ─────────────
// Every notification the app asks OneSignal for (and every cancel) is added as a row on a "Push Log" tab, newest at the bottom,
// so a missing notification can be traced: was it ever scheduled, for what time, and what did OneSignal answer?
// Run pushReport() from the editor to add what OneSignal itself says happened to the 15 most recent notifications.
// A logging problem never blocks a push.
const PUSH_LOG_SHEET = 'Push Log';
const PUSH_LOG_HEADERS = ['Logged', 'Event', 'Reference / topic', 'Scheduled for', 'Result', 'Detail'];
const PUSH_LOG_KEEP = 400;

function pushLog_(event, ref, whenMs, ok, detail) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(PUSH_LOG_SHEET);
    if (!sheet) {
      sheet = ss.insertSheet(PUSH_LOG_SHEET);
      sheet.appendRow(PUSH_LOG_HEADERS);
      try { sheet.setFrozenRows(1); } catch (e) {}
      try { sheet.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm:ss'); sheet.getRange('D:D').setNumberFormat('yyyy-mm-dd hh:mm:ss'); } catch (e) {}
    }
    sheet.appendRow([new Date(), event, ref || '', whenMs ? new Date(whenMs) : '', ok === null ? '' : (ok ? 'ok' : 'FAILED'), String(detail || '').slice(0, 300)]);
    const extra = sheet.getLastRow() - 1 - PUSH_LOG_KEEP;
    if (extra > 0) sheet.deleteRows(2, extra);
  } catch (e) { console.error('Push log write failed: ' + e); }
}

// Adds OneSignal's own account of the most recent notifications to the Push Log (run it from the editor, e.g. the next morning).
function pushReport() {
  const r = osCall_('get', '/notifications?app_id=' + ONESIGNAL_APP_ID + '&limit=15');
  if (!r.ok || !r.json || !r.json.notifications) { pushLog_('report', '', 0, false, r.code + ' ' + String(r.text || '').slice(0, 200)); return; }
  r.json.notifications.slice().reverse().forEach(function(n) {
    const when = n.send_after ? Number(n.send_after) * 1000 : 0;
    const done = n.completed_at ? new Date(Number(n.completed_at) * 1000).toISOString() : 'not yet';
    pushLog_('report: ' + ((n.headings && n.headings.en) || '?'), n.web_push_topic || '', when, null,
      'delivered ' + (n.successful || 0) + ', failed ' + (n.failed || 0) + ', errored ' + (n.errored || 0) + ', waiting ' + (n.remaining || 0) +
      (n.canceled ? ', CANCELED' : '') + ', finished ' + done + ' [' + n.id + ']');
  });
}

function osCall_(method, path, body) {
  const key = PropertiesService.getScriptProperties().getProperty('ONESIGNAL_KEY');
  if (!key) return { ok: false, text: 'ONESIGNAL_KEY is not set' };
  const opts = {
    method: method,
    muteHttpExceptions: true,
    headers: { 'Authorization': 'Key ' + key, 'accept': 'application/json' }
  };
  if (body) { opts.contentType = 'application/json'; opts.payload = JSON.stringify(body); }
  let res;
  try { res = UrlFetchApp.fetch('https://api.onesignal.com' + path, opts); }
  catch (err) { return { ok: false, text: String(err) }; }
  const code = res.getResponseCode();
  const text = res.getContentText();
  let json = null;
  try { json = JSON.parse(text); } catch (e) {}
  const ok = code >= 200 && code < 300 && !(json && json.errors && !json.id);
  try { PropertiesService.getScriptProperties().setProperty(PUSH_LAST_KEY, new Date().toISOString() + ' ' + code + ' ' + text.slice(0, 300)); } catch (e) {}
  return { ok: ok, code: code, text: text, json: json };
}

function osSend_(heading, message, whenMs, topic, go) {
  const body = {
    app_id: ONESIGNAL_APP_ID,
    target_channel: 'push',
    included_segments: ['Total Subscriptions'],
    headings: { en: heading },
    contents: { en: message },
    url: PUSH_URL + (go ? '?go=' + go : ''),
    web_push_topic: topic || 'verse-review',
    data: go ? { vrgo: go } : {},
    priority: 10, // high priority: delivered right away, even when the phone is idle
    ttl: 4 * 60 * 60 // if the phone is unreachable, drop it after 4 hours instead of delivering a stale reminder
  };
  if (whenMs && whenMs > Date.now() + 30000) {
    body.send_after = Utilities.formatDate(new Date(whenMs), 'UTC', "yyyy-MM-dd'T'HH:mm:ss'Z'");
  }
  const res = osCall_('post', '/notifications?c=push', body);
  pushLog_('send: ' + heading, message, body.send_after ? whenMs : 0, !!res.ok,
    (body.send_after ? 'send_after ' + body.send_after : 'immediate') + ' | ' + (res.ok && res.json && res.json.id ? 'id ' + res.json.id : res.code + ' ' + String(res.text || '').slice(0, 160)));
  return res;
}

function slotMap_() {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty(PUSH_SLOTMAP_KEY) || '{}') || {}; }
  catch (e) { return {}; }
}
function setSlotMap_(map) {
  PropertiesService.getScriptProperties().setProperty(PUSH_SLOTMAP_KEY, JSON.stringify(map));
}

// "Romans 8:35-37", or "Romans 8:35-37 and 1 more" when several verses are waiting.
function refsText_(refs) {
  return refs[0] + (refs.length > 1 ? ' and ' + (refs.length - 1) + ' more' : '');
}

// The title names what is waiting ("Review", "Learn, Lock", "Review, Learn, Lock"); the text gives the specifics.
function pushWords_(due, learnRefs, lockRefs) {
  const titles = [], bits = [];
  if (due > 0) { titles.push('Review'); bits.push('review ' + due + (due === 1 ? ' verse' : ' verses')); }
  if (learnRefs.length) { titles.push('Learn'); bits.push('learn ' + refsText_(learnRefs)); }
  if (lockRefs.length) { titles.push('Lock'); bits.push('lock ' + refsText_(lockRefs)); }
  if (!bits.length) return null;
  const cap = function(t) { return t.charAt(0).toUpperCase() + t.slice(1); };
  let text;
  if (bits.length === 1) text = 'Time to ' + bits[0] + '.';
  else if (due > 0 && bits.length === 2) text = cap(bits[0]) + ', then ' + bits[1] + '.';
  else if (bits.length === 2) text = cap(bits[0]) + ' and ' + bits[1] + '.';
  else text = cap(bits[0]) + ', ' + bits[1] + ', ' + bits[2] + '.';
  return { title: titles.join(', '), text: text };
}

// Runs from the daily trigger (installDailyPush), at about 5:00 AM (Apps Script starts it somewhere in 5:00 to 5:15). It works out what is
// waiting then, and schedules the push for exactly 6:00 AM: due reviews, a Learn card and a Lock AM card (both unlock at 4 AM).
// If it happens to run after 6:00 the push goes out right away. Sent once a day, and only if something is waiting.
function dailyPush() {
  const props = PropertiesService.getScriptProperties();
  const today = todayKey_();
  try { syncSlotPushes_(); } catch (e) { pushLog_('slot schedule ERROR', '', 0, false, String(e)); }
  if (props.getProperty(PUSH_DAILY_KEY) === today) return;
  const due = getDueCards(false).length;
  const nc = newCardCounts_(readSheetData_(), nowMs_());
  const words = pushWords_(due, nc.learnRefs, nc.lockRefs);
  if (!words) return;
  const go = due > 0 ? 'review' : 'new';
  const now = nowMs_();
  const at = atClock_(now, 0, DAILY_PUSH_MIN);
  const r = osSend_(words.title, words.text, at > now + 30000 ? at : 0, 'vr-daily', go);
  if (r.ok) props.setProperty(PUSH_DAILY_KEY, today);
}

// Run from the editor when a morning notification did not arrive. It adds a few "status" rows to the Push Log tab: whether the daily
// trigger exists, the time zone it follows, whether today's push was already sent, and whether anything was waiting to send.
function pushStatus() {
  const props = PropertiesService.getScriptProperties();
  const trig = ScriptApp.getProjectTriggers().filter(function(t) { return t.getHandlerFunction() === 'dailyPush'; });
  pushLog_('status: trigger', 'dailyPush', 0, trig.length === 1, trig.length + ' daily trigger(s)' + (trig.length ? ' (' + trig[0].getEventType() + ')' : ' - run installDailyPush()'));
  pushLog_('status: time', Session.getScriptTimeZone(), 0, null, 'script time zone ' + Session.getScriptTimeZone() + ', now ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') + '. The daily trigger fires between 5:00 and 5:15 in THIS zone and schedules the push for 6:00.');
  pushLog_('status: today', '', 0, null, 'last daily push marked ' + (props.getProperty(PUSH_DAILY_KEY) || 'never') + ', today is ' + todayKey_());
  pushLog_('status: key', 'ONESIGNAL_KEY', 0, !!props.getProperty('ONESIGNAL_KEY'), props.getProperty('ONESIGNAL_KEY') ? 'set' : 'NOT set');
  const slots = slotMap_();
  const times = Object.keys(slots).map(Number).filter(function(k) { return k > Date.now(); }).sort(function(x, y) { return x - y; });
  pushLog_('status: slots', '', 0, null, times.length + ' Learn / Lock push(es) waiting to go out today' + (times.length ? ' (listed below, one row each)' : ''));
  times.forEach(function(k) {
    const e = slots[k] || {};
    const w = pushWords_(0, e.learn || [], e.lock || []);
    pushLog_('status: slot', w ? w.title : '', k, null, Utilities.formatDate(new Date(k), Session.getScriptTimeZone(), 'h:mm a') + '  ' + (w ? w.text : ''));
  });
  const due = getDueCards(false).length;
  const nc = newCardCounts_(readSheetData_(), nowMs_());
  pushLog_('status: waiting', '', 0, null, due + ' due, ' + nc.learnRefs.length + ' learn, ' + nc.lockRefs.length + ' lock' + (pushWords_(due, nc.learnRefs, nc.lockRefs) ? '' : ' (nothing waiting, so no push is sent)'));
}

// Run once from the editor (run it again after updating to this version so the trigger moves to about 6:00 AM).
function installDailyPush() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'dailyPush') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('dailyPush').timeBased().atHour(5).nearMinute(0).everyDays(1).create();
  console.log('Daily push trigger created for about 5:00 AM; it schedules the push for 6:00 AM.');
}

// Run from the editor to check the connection. View -> Logs shows the answer.
function testPush() {
  const r = osSend_('Test Notification', 'Push is working.', 0, 'vr-test');
  console.log('ok=' + r.ok + ' code=' + r.code + ' ' + r.text);
}

// Run from the editor: writes a "Schedule" tab showing, for each of the next 60 days, how many verses are due (and how many of each rank),
// plus their load (word count based, the number the balancing uses). Read-only: it changes nothing in the Verses tab.
function scheduleReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Verses');
  if (!sheet) return;
  const today = today_();
  const last = sheet.getLastRow();
  const vals = last >= 2 ? sheet.getRange(2, 2, last - 1, 7).getValues() : []; // B..H
  const days = {}, later = { n: 0, load: 0 };
  const HORIZON = 60;
  const names = [];
  for (let g = 1; g <= MAX_GROUP; g++) names.push(GROUP_CONFIG[g].name);
  let overdue = 0;
  vals.forEach(function(r) {
    const g = Number(r[3]);
    if (!(g >= 1 && g <= MAX_GROUP)) return;
    const d = toMidnight_(r[0]);
    if (!d) return;
    const k = Math.round((d.getTime() - today.getTime()) / 86400000);
    if (k < 0) { overdue++; return; }
    const ld = verseLoad_(r[6]);
    if (k > HORIZON) { later.n++; later.load += ld; return; }
    const e = days[k] || (days[k] = { n: 0, load: 0, by: {} });
    e.n++; e.load += ld; e.by[g] = (e.by[g] || 0) + 1;
  });
  let out = ss.getSheetByName('Schedule');
  if (!out) out = ss.insertSheet('Schedule'); else out.clear();
  const rows = [['Date', 'Day', 'Due', 'Load'].concat(names)];
  const tz = Session.getScriptTimeZone();
  for (let k = 0; k <= HORIZON; k++) {
    const e = days[k] || { n: 0, load: 0, by: {} };
    const dt = new Date(today.getTime()); dt.setDate(dt.getDate() + k);
    const row = [Utilities.formatDate(dt, tz, 'yyyy-MM-dd'), Utilities.formatDate(dt, tz, 'EEE'), e.n, Math.round(e.load * 10) / 10];
    for (let g = 1; g <= MAX_GROUP; g++) row.push(e.by[g] || '');
    rows.push(row);
  }
  rows.push(['Later than ' + HORIZON + ' days', '', later.n, Math.round(later.load * 10) / 10].concat(names.map(function() { return ''; })));
  if (overdue) rows.push(['Overdue (already past due)', '', overdue, ''].concat(names.map(function() { return ''; })));
  out.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  try { out.setFrozenRows(1); } catch (e) {}
}

// Called by the app (Verse List title -> "Send test notification").
function sendTestPush() {
  const r = osSend_('Test Notification', 'Push is working.', 0, 'vr-test');
  let msg = '';
  if (!r.ok) {
    msg = (r.text || '').slice(0, 120);
    if (/subscribed|recipients/i.test(msg)) msg = 'no subscribed device';
  }
  return { ok: !!r.ok, message: msg };
}

// When the push for a New card's unlock goes out. The 4 AM unlocks ride on the 6 AM daily push (dailyPush), so they get none
// of their own; the 6 PM unlocks get theirs at 8 PM; every other unlock gets a push at its own time. 0 = none.
function learnPushTime_(unlockMs) {
  if (!unlockMs) return 0;
  const m = minsOfDay_(unlockMs);
  if (m === DAY_START_MIN) return 0;
  if (m === LEARN_SLOTS[LEARN_SLOTS.length - 1]) return unlockMs + 2 * 60 * 60 * 1000;
  return unlockMs;
}

// The times of today's remaining slot pushes for a New card, from `now` on. A Learn card gets the slots from its current rep onward
// (8, 10, 12, 2, 4 PM, and 8 PM for the 6 PM one); a Lock card gets the 8 PM one for the 6 PM unlock. A card that is waiting for tomorrow gets none.
function daySlotTimes_(card, now) {
  const endOfDay = atClock_(now, 1, DAY_START_MIN);
  if (card.unlockAt && card.unlockAt >= endOfDay) return [];
  const out = [];
  if (card.stage === 'lock') {
    out.push(atClock_(now, 0, LOCK_PM_MIN) + 2 * 60 * 60 * 1000);
  } else {
    for (let i = 1; i < LEARN_SLOTS.length; i++) {
      if (card.dot > i + 1) continue; // that rep is already done
      out.push(atClock_(now, 0, LEARN_SLOTS[i]) + (i === LEARN_SLOTS.length - 1 ? 2 * 60 * 60 * 1000 : 0));
    }
  }
  return out.filter(function(t) { return t > now + 30000; });
}

// For each remaining slot time today, the Learn and Lock verses that will be waiting then: { time ms: { learn: [refs], lock: [refs] } }.
function wantedSlots_(data, now) {
  const want = {};
  for (let i = 1; i < data.length; i++) {
    if (!cleanRef_(data[i][5]) || !isLevelZero_(data[i][4])) continue;
    const card = learningCard_(data[i], i, now);
    daySlotTimes_(card, now).forEach(function(t) {
      const w = want[t] || (want[t] = { learn: [], lock: [] });
      (card.stage === 'lock' ? w.lock : w.learn).push(card.reference);
    });
  }
  return want;
}

// One-time tidy-up: an older version scheduled a push per verse and kept their ids under VR_PUSH_PENDING. Those would go out as well as the
// new combined ones, so any that are still waiting are cancelled, then the old record is deleted.
function cancelLegacyPushes_() {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty('VR_PUSH_PENDING');
  if (!raw) return 0;
  let n = 0;
  try {
    const map = JSON.parse(raw) || {};
    Object.keys(map).forEach(function(ref) {
      const v = map[ref];
      const list = Array.isArray(v) ? v : [{ id: v, at: 0 }];
      list.forEach(function(e) {
        if (!e || !e.id) return;
        if (e.at && e.at <= Date.now() + 5000) return;
        const c = osCall_('delete', '/notifications/' + encodeURIComponent(e.id) + '?app_id=' + ONESIGNAL_APP_ID);
        pushLog_('cancel (old schedule)', ref, e.at || 0, !!c.ok, 'id ' + e.id + (c.ok ? '' : ' | ' + c.code + ' ' + String(c.text || '').slice(0, 160)));
        n++;
      });
    });
  } catch (err) { console.error('cancelLegacyPushes_ failed: ' + err); }
  props.deleteProperty('VR_PUSH_PENDING');
  return n;
}

function sameRefs_(a, b) {
  return JSON.stringify((a || []).slice().sort()) === JSON.stringify((b || []).slice().sort());
}

// Brings the scheduled Learn / Lock pushes in line with what is true right now: one push per remaining slot time, naming every verse
// waiting then. Pushes that are no longer right (a verse finished its last rep, was skipped or promoted) are cancelled or re-sent with the
// updated names, and missing ones are added. Safe to repeat; a push that is about to go out (within 2 minutes) is left alone.
function syncSlotPushes_() {
  const now = nowMs_();
  const want = wantedSlots_(readSheetData_(), now);
  const map = slotMap_();
  const out = { sent: 0, cancelled: cancelLegacyPushes_() };
  Object.keys(map).forEach(function(k) {
    const e = map[k];
    if (Number(k) <= now + 120000) { if (Number(k) <= now - 60000) delete map[k]; return; } // gone out (or about to)
    const w = want[k];
    if (w && sameRefs_(e.learn, w.learn) && sameRefs_(e.lock, w.lock)) return;
    delete map[k];
    const c = osCall_('delete', '/notifications/' + encodeURIComponent(e.id) + '?app_id=' + ONESIGNAL_APP_ID);
    pushLog_('cancel', (e.learn || []).concat(e.lock || []).join(', '), Number(k), !!c.ok, 'id ' + e.id + (c.ok ? '' : ' | ' + c.code + ' ' + String(c.text || '').slice(0, 160)));
    out.cancelled++;
  });
  Object.keys(want).sort().forEach(function(k) {
    if (map[k]) return;
    const w = want[k];
    const words = pushWords_(0, w.learn, w.lock);
    if (!words) return;
    const r = osSend_(words.title, words.text, Number(k), 'vr-new', 'new');
    if (r.ok && r.json && r.json.id) { map[k] = { id: r.json.id, learn: w.learn, lock: w.lock }; out.sent++; }
  });
  setSlotMap_(map);
  return out;
}

// Run from the editor to schedule the rest of today's slot pushes right now (for example after updating mid-day). Results show on the Push Log tab.
function scheduleTodayPushes() {
  const r = syncSlotPushes_();
  pushLog_('status: slots', '', 0, null, r.sent + ' slot push(es) scheduled, ' + r.cancelled + ' cancelled, for the rest of today');
}

// ───────────── One-time migration: Bronze goes from 4 reps to 3 ─────────────
// 1) Run migrationPreview() from the editor and read View -> Logs. It changes nothing.
// 2) Run migrationRun(). It first copies the Verses tab to a backup tab ("Verses backup <date time>"), then lowers every
//    Bronze verse at rep 2 or higher by one, so a verse already part-way through keeps its Silver date. Bronze verses at rep 1 stay.
//    Level 0 verses are already at rep 1 and simply start the new Learn flow.
// It refuses to run a second time (script property VR_MIGRATED_BRONZE3).
const MIGRATION_KEY = 'VR_MIGRATED_BRONZE3';

function migrationPreview() { return migrateBronze_(false); }
function migrationRun() { return migrateBronze_(true); }

function migrateBronze_(apply) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const props = PropertiesService.getScriptProperties();
    if (apply && props.getProperty(MIGRATION_KEY)) {
      console.log('The migration already ran (' + props.getProperty(MIGRATION_KEY) + '). Nothing changed.');
      return { ran: false, reason: 'already' };
    }
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Verses');
    if (!sheet) { console.log('Sheet "Verses" not found.'); return { ran: false, reason: 'nosheet' }; }
    const data = readSheetData_();
    const changes = [];
    for (let i = 1; i < data.length; i++) {
      const ref = cleanRef_(data[i][5]);
      if (!ref) continue;
      const g = String(data[i][4] === null || data[i][4] === undefined ? '' : data[i][4]).trim();
      if (g === '' || Number(g) !== 1) continue;
      const rep = Number(data[i][3]) || 1;
      if (rep >= 2) changes.push({ row: i + 1, ref: ref, from: rep, to: rep - 1 });
    }
    changes.forEach(function(c) { console.log((apply ? 'Changed ' : 'Would change ') + c.ref + ' (row ' + c.row + '): Bronze rep ' + c.from + ' -> ' + c.to); });
    console.log((apply ? 'Changed ' : 'Would change ') + changes.length + ' verse(s).');
    if (!apply) return { ran: false, preview: true, changes: changes };

    const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HHmm');
    let backupName = '';
    try {
      const copy = sheet.copyTo(ss);
      backupName = 'Verses backup ' + stamp;
      copy.setName(backupName);
    } catch (e) {
      console.log('Could not make the backup tab, so nothing was changed: ' + e);
      return { ran: false, reason: 'backup' };
    }
    changes.forEach(function(c) { sheet.getRange(c.row, 4).setValue(c.to); });
    props.setProperty(MIGRATION_KEY, stamp);
    console.log('Done. Backup tab: ' + backupName);
    return { ran: true, changes: changes, backup: backupName };
  } finally {
    lock.releaseLock();
  }
}
