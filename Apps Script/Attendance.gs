/**
 * ============================================================
 *  SLACK ATTENDANCE TRACKER  (Google Apps Script) — Attendance.gs
 *  #daily-check-in  ->  Leave Portal Google Sheet
 *  Isi Apps Script project mein Code.gs ke sath ek nayi file banayein.
 *  SHEET_ID Code.gs se aata hai.
 * ============================================================
 *  Kya karta hai:
 *   - Slack channel ke saare check in / check out messages parhta hai
 *     (thread replies bhi, edited messages bhi).
 *   - Spelling mistakes khud samajh leta hai (chek in, checkout namz, brek...).
 *   - Message mein likha time (e.g. "Check in (12:14)") use karta hai.
 *   - Har employee ka roz ka hisaab: Work, Break, Namaz, Overtime, Short.
 *   - Jo message samajh na aaye ya hisaab mein gadbad ho (check out bhool gaya,
 *     do baar check in, etc.) -> "Issues" sheet + admin ko Slack DM alert.
 *   - Admin "Raw Log" mein Manual Type / Manual Time likh kar ghalti theek
 *     kar sakta hai; agle sync par hisaab khud update ho jata hai.
 *
 *  Setup steps neeche SETUP section mein, aur chat mein bhi diye gaye hain.
 * ============================================================
 */

/************************ SETTINGS (yahan change karein) ************************/
const CONFIG = {
  CHANNEL_ID: 'C0AGSCB14JK',                 // #daily-check-in
  SLACK_WORKSPACE_URL: 'https://othtech.slack.com',
  ADMIN_SLACK_IDS: ['U0A63SDPKNE'],          // Shoaib Mansoor — alerts isi DM mein aate hain
  ADMIN_EMAILS: [],                          // optional: ['admin@company.com']

  TZ_OFFSET_MINUTES: 300,                    // Pakistan = UTC+5 (DST nahi hota)
  DAY_START_HOUR: 5,                         // subah 5 baje se pehle ki entry pichle din mein jayegi

  TARGET_WORK_MINUTES: 8 * 60,               // roz ka target (saaf kaam)
  WEEKLY_OFF_DAYS: [0],                      // 0 = Sunday. Off day par saara kaam overtime
  ALLOWED_BREAK_MINUTES: 60,                 // allowed break
  NAMAZ_COUNTS_AS_WORK: true,                // namaz ka time kaam mein ginta hai (9h office = 8h kaam + 1h break)
  EXTRA_BREAK_COUNTS_AS_SHORT: false,        // true = allowed se zyada break short time mein bhi add ho

  // ---- Overtime -> extra leave (website sript.js ke ATT mein bhi yahi rakhein) ----
  OVERTIME_START_DATE: '2026-10-24',         // overtime / short ISI DIN se ginna shuru (pehle ke din sirf dikhte hain)
  OVERTIME_PER_LEAVE_MINUTES: 8 * 60,        // har 8h overtime = 1 extra leave, baqi aage
  SHORT_REDUCES_OVERTIME: true,              // true = short time overtime balance se minus hota hai
  OVERTIME_HISTORY_WEEKS: 2,                 // "Overtime History" tab mein sirf itne haftay (is + pichla)
  SETTLE_GRACE_DAYS: 3,                      // mahina khatam hone ke itne din baad "Final" (lock) hota hai

  // ---- Reminders / warnings (channel mein employee ko tag, aur DM) ----
  LATE_AFTER: '12:10',                       // is ke baad check in = late (Mon–Sat; Sunday/chutti par nahi)
  BREAK_REMIND_MINUTES: 60,                  // break / namaz ke baad itni der check in nahi -> tag
  NIGHT_CHECK: '23:00',                      // is waqt tak check out nahi -> "overtime kar rahe hain?" tag
  HR_NAMES: 'HR ya Haseeb bhai',             // messages mein sirf naam (tag nahi)
  SEND_EMPLOYEE_MESSAGES: true,              // false = employees ko koi channel message / DM nahi

  TYPED_TIME_MAX_BACK_MINUTES: 180,          // message mein likha time max kitna purana ho sakta hai
  TYPED_TIME_MAX_BACK_CHECKOUT_MINUTES: 12 * 60, // check out par likha time (raat ko "Checkout (9:30)") 12 ghante tak
  ALERT_ON_TYPED_TIME: true,                 // likha hua time use ho to admin ko batao
  SHORT_SESSION_MINUTES: 2,                  // is se chhota session = shayad ghalti
  LONG_SESSION_MINUTES: 10 * 60,             // is se lamba ek session = shayad check out bhool gaya

  RESYNC_DAYS: 3,                            // har sync par pichle itne din dobara check (edits/deletes)
  QUICK_SYNC_DAYS: 1,                        // Slack message aate hi "quick sync" sirf itne din ka data laati hai
  INITIAL_SYNC_DAYS: 30,                     // pehli dafa kitne din ka data laye
};

const SHEETS = { RAW: 'Raw Log', DAILY: 'Daily Summary', MONTHLY: 'Monthly Summary', ISSUES: 'Issues',
  OT_HISTORY: 'Overtime History', EARNED: 'Earned Leaves', OT_MONTHS: 'Overtime Months', NOTICES: 'Notices' };
const RAW_HEADERS = ['Slack TS', 'Date', 'Time', 'Employee', 'User ID', 'Message', 'Detected',
  'Typed Time Used', 'Manual Type', 'Manual Time', 'Note', 'Posted At', 'Thread TS'];
const R = { TS: 0, DAY: 1, TIME: 2, NAME: 3, UID: 4, TEXT: 5, DETECTED: 6, TYPED: 7,
  MAN_TYPE: 8, MAN_TIME: 9, NOTE: 10, POSTED: 11, THREAD: 12 };
const TYPES = ['IN', 'OUT_BREAK', 'OUT_NAMAZ', 'OUT_END', 'OUT_HALFDAY', 'LATE', 'IGNORE'];
const TYPE_LABEL = { IN: 'Check in', OUT_BREAK: 'Out - Break', OUT_NAMAZ: 'Out - Namaz',
  OUT_END: 'Out - Day end', OUT_HALFDAY: 'Out - Half day', LATE: 'Late (informed)', IGNORE: 'Ignore', UNREADABLE: '❓ UNREADABLE' };

function ss_() { return SpreadsheetApp.openById(SHEET_ID); }

/************************ MENU & SETUP ************************/
function onOpen() {
  SpreadsheetApp.getUi().createMenu('⏱ Attendance')
    .addItem('Sync now', 'syncNow')
    .addItem('Recalculate only (no Slack fetch)', 'recalculateOnly')
    .addSeparator()
    .addItem('1) Set Slack bot token', 'setSlackToken')
    .addItem('2) Setup sheets + auto-sync (15 min)', 'setup')
    .addItem('Full resync (last 30 days)', 'fullResync')
    .addToUi();
}

function setSlackToken() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('Slack Bot Token', 'xoxb-... wala token paste karein:', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const t = res.getResponseText().trim();
  if (!/^xoxb-/.test(t)) { ui.alert('Token "xoxb-" se shuru hona chahiye.'); return; }
  PropertiesService.getScriptProperties().setProperty('SLACK_BOT_TOKEN', t);
  ui.alert('Token save ho gaya ✅');
}

function setup() {
  if (!PropertiesService.getScriptProperties().getProperty('SLACK_BOT_TOKEN')) {
    throw new Error('Pehle Project Settings > Script Properties mein SLACK_BOT_TOKEN add karein.');
  }
  const ss = ss_();
  ensureSheet_(ss, SHEETS.RAW, RAW_HEADERS);
  const raw = ss.getSheetByName(SHEETS.RAW);
  raw.getRange('A:B').setNumberFormat('@');
  raw.getRange('J:J').setNumberFormat('@');
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(TYPES, true).setAllowInvalid(false).build();
  raw.getRange(2, R.MAN_TYPE + 1, raw.getMaxRows() - 1, 1).setDataValidation(rule);
  ensureSheet_(ss, SHEETS.DAILY, []);
  ensureSheet_(ss, SHEETS.MONTHLY, []);
  ensureSheet_(ss, SHEETS.ISSUES, []);

  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'syncNow') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncNow').timeBased().everyMinutes(15).create();
  syncNow();
  Logger.log('Setup complete ✅  Har 15 minute auto-sync hoga.');
}

/** Saare attendance tabs maujood hon (setup na chala ho tab bhi sync na toote) */
function ensureAllSheets_(ss) {
  const isNew = !ss.getSheetByName(SHEETS.RAW);
  ensureSheet_(ss, SHEETS.RAW, RAW_HEADERS);
  if (isNew) {
    const raw = ss.getSheetByName(SHEETS.RAW);
    raw.getRange('A:B').setNumberFormat('@');
    raw.getRange('J:J').setNumberFormat('@');
    const rule = SpreadsheetApp.newDataValidation().requireValueInList(TYPES, true).setAllowInvalid(false).build();
    raw.getRange(2, R.MAN_TYPE + 1, raw.getMaxRows() - 1, 1).setDataValidation(rule);
  }
  [SHEETS.DAILY, SHEETS.MONTHLY, SHEETS.ISSUES].forEach(n => ensureSheet_(ss, n, []));
}

function fullResync() {
  PropertiesService.getScriptProperties().deleteProperty('LAST_SYNC_MS');
  syncNow();
}

/************************ MAIN SYNC ************************/
/** Poori sync (har 15 minute + menu). Backup: koi message miss na ho, edits/deletes bhi pakre. */
function syncNow() { syncWindow_(0); }

/** days = 0: normal sync (pichle RESYNC_DAYS din) · days > 0: quick sync (sirf itne din) */
function syncWindow_(days) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) { if (days) scheduleQuickSync_(); return; }
  try {
    const props = PropertiesService.getScriptProperties();
    const last = Number(props.getProperty('LAST_SYNC_MS') || 0);
    const now = Date.now();
    const oldestMs = days ? now - days * 86400000
      : Math.min(last || now, now) - (last ? CONFIG.RESYNC_DAYS : CONFIG.INITIAL_SYNC_DAYS) * 86400000;

    ensureAllSheets_(ss_());
    const messages = fetchSlackMessages_(oldestMs / 1000);
    upsertRawLog_(messages, oldestMs);
    recalculateOnly();
    if (!days) props.setProperty('LAST_SYNC_MS', String(now));
  } finally {
    lock.releaseLock();
  }
}

/************************ SLACK EVENTS (message aate hi update) ************************
 *  Slack App > Event Subscriptions > Request URL = is web app ka /exec link
 *  Subscribe to bot events: message.channels (+ message.groups agar channel private ho)
 *  Slack ko 3 second mein jawab dena hota hai, is liye yahan sirf ~1 minute baad ki
 *  "quickSync" schedule hoti hai; asal kaam woh karti hai.
 *  (optional) Script Properties: SLACK_VERIFICATION_TOKEN = Slack App > Basic Information > Verification Token
 ***************************************************************************************/
function handleSlackEvent_(body) {
  const out = t => ContentService.createTextOutput(t);
  const vt = PropertiesService.getScriptProperties().getProperty('SLACK_VERIFICATION_TOKEN');
  if (vt && body.token !== vt) return out('forbidden');
  if (body.type === 'url_verification') return out(String(body.challenge || ''));
  if (body.type !== 'event_callback') return out('ok');
  const ev = body.event || {};
  if (ev.channel !== CONFIG.CHANNEL_ID) return out('ok');
  if (ev.bot_id || ev.subtype === 'bot_message') return out('ok');   // bot ke apne messages par sync nahi
  const cache = CacheService.getScriptCache();
  if (body.event_id) {                                   // Slack retry -> dobara kaam nahi
    if (cache.get('ev_' + body.event_id)) return out('ok');
    cache.put('ev_' + body.event_id, '1', 21600);
  }
  scheduleQuickSync_();
  return out('ok');
}

function scheduleQuickSync_() {
  const cache = CacheService.getScriptCache();
  if (cache.get('QUICK_SCHEDULED')) return;              // pehle se schedule hai
  cache.put('QUICK_SCHEDULED', '1', 180);
  ScriptApp.newTrigger('quickSync').timeBased().after(1000).create();
}

/** Slack event ke ~1 minute baad chalti hai */
function quickSync() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'quickSync') ScriptApp.deleteTrigger(t);
  });
  CacheService.getScriptCache().remove('QUICK_SCHEDULED');
  syncWindow_(CONFIG.QUICK_SYNC_DAYS);
}

function recalculateOnly() {
  const ss = ss_();
  ensureAllSheets_(ss);
  const raw = ss.getSheetByName(SHEETS.RAW);
  const data = raw.getLastRow() > 1
    ? raw.getRange(2, 1, raw.getLastRow() - 1, RAW_HEADERS.length).getDisplayValues() : [];
  const empSheet = ss.getSheetByName('Employees');
  const empRows = empSheet && empSheet.getLastRow() > 1
    ? empSheet.getRange(2, 1, empSheet.getLastRow() - 1, Math.max(3, empSheet.getLastColumn())).getDisplayValues() : [];
  const result = buildReport(data, Date.now(), empRows);

  // Raw log: Date / Time / Note columns refresh (manual changes ke baad)
  if (data.length) {
    raw.getRange(2, R.DAY + 1, data.length, 2).setValues(result.rawDayTime);
    raw.getRange(2, R.NOTE + 1, data.length, 1).setValues(result.rawNotes);
    // manual entries ko TS de do
    raw.getRange(2, 1, data.length, 1).setValues(result.rawTs);
  }
  writeTable_(ss.getSheetByName(SHEETS.DAILY), result.dailyHeaders, result.daily);
  writeTable_(ss.getSheetByName(SHEETS.MONTHLY), result.monthlyHeaders, result.monthly);
  writeIssues_(ss.getSheetByName(SHEETS.ISSUES), result.issues);
  const now = Date.now();
  const plan = updateOvertime_(ss, result.records, now);
  try { runNotices_(ss, result, plan, now); } catch (e) { Logger.log('Notices error: ' + e.message); }
  try { scheduleReminders_(result.records, now); } catch (e) { Logger.log('Reminder schedule error: ' + e.message); }
}

/************************ SLACK API ************************/
function slackGet_(method, params) {
  const token = PropertiesService.getScriptProperties().getProperty('SLACK_BOT_TOKEN');
  if (!token) throw new Error('Slack token set nahi hai. Menu > Attendance > Set Slack bot token');
  const qs = Object.keys(params).map(k => k + '=' + encodeURIComponent(params[k])).join('&');
  for (let attempt = 0; attempt < 5; attempt++) {
    const resp = UrlFetchApp.fetch('https://slack.com/api/' + method + '?' + qs, {
      headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true });
    if (resp.getResponseCode() === 429) {
      const wait = Number(resp.getHeaders()['Retry-After'] || 3);
      Utilities.sleep(wait * 1000);
      continue;
    }
    const json = JSON.parse(resp.getContentText());
    if (!json.ok) throw new Error('Slack ' + method + ' error: ' + json.error);
    return json;
  }
  throw new Error('Slack rate limit: ' + method);
}

function slackPost_(method, payload) {
  const token = PropertiesService.getScriptProperties().getProperty('SLACK_BOT_TOKEN');
  const resp = UrlFetchApp.fetch('https://slack.com/api/' + method, {
    method: 'post', contentType: 'application/json; charset=utf-8',
    headers: { Authorization: 'Bearer ' + token }, payload: JSON.stringify(payload), muteHttpExceptions: true });
  return JSON.parse(resp.getContentText());
}

function fetchSlackMessages_(oldestSec) {
  const out = [];
  let cursor = '';
  do {
    const params = { channel: CONFIG.CHANNEL_ID, oldest: String(oldestSec), limit: 200 };
    if (cursor) params.cursor = cursor;
    const res = slackGet_('conversations.history', params);
    (res.messages || []).forEach(m => {
      if (isHumanMessage_(m)) out.push(m);
      if (m.reply_count) {                       // thread replies bhi parho
        let c2 = '';
        do {
          const p2 = { channel: CONFIG.CHANNEL_ID, ts: m.ts, limit: 200 };
          if (c2) p2.cursor = c2;
          const r2 = slackGet_('conversations.replies', p2);
          (r2.messages || []).forEach(rm => {
            if (rm.ts !== m.ts && isHumanMessage_(rm)) { rm._thread = m.ts; out.push(rm); }
          });
          c2 = (r2.response_metadata && r2.response_metadata.next_cursor) || '';
        } while (c2);
      }
    });
    cursor = (res.response_metadata && res.response_metadata.next_cursor) || '';
  } while (cursor);
  return out;
}

function isHumanMessage_(m) {
  if (m.bot_id || m.subtype === 'bot_message') return false;
  if (m.subtype && m.subtype !== 'thread_broadcast' && m.subtype !== 'file_share') return false;
  return !!m.user;
}

function getUserName_(uid) {
  const cache = CacheService.getScriptCache();
  const key = 'u_' + uid;
  const hit = cache.get(key);
  if (hit) return hit;
  let name = uid;
  try {
    const u = slackGet_('users.info', { user: uid }).user;
    name = (u.profile && (u.profile.real_name || u.profile.display_name)) || u.name || uid;
  } catch (e) {}
  cache.put(key, name, 21600);
  return name;
}

function upsertRawLog_(messages, windowStartMs) {
  const sh = ss_().getSheetByName(SHEETS.RAW);
  const lastRow = sh.getLastRow();
  const existing = lastRow > 1 ? sh.getRange(2, 1, lastRow - 1, RAW_HEADERS.length).getValues() : [];
  const index = {};
  existing.forEach((r, i) => { index[String(r[R.TS])] = i; });
  const seen = {};
  const newRows = [];

  messages.forEach(m => {
    seen[m.ts] = true;
    const posted = Math.round(Number(m.ts) * 1000);
    const name = getUserName_(m.user);
    const text = (m.text || '').replace(/\n/g, ' ');
    if (index[m.ts] !== undefined) {
      const row = existing[index[m.ts]];
      row[R.NAME] = name; row[R.UID] = m.user; row[R.TEXT] = text;
      row[R.POSTED] = fmtDateTime(posted); row[R.THREAD] = m._thread || '';
    } else {
      const row = new Array(RAW_HEADERS.length).fill('');
      row[R.TS] = m.ts; row[R.NAME] = name; row[R.UID] = m.user; row[R.TEXT] = text;
      row[R.POSTED] = fmtDateTime(posted); row[R.THREAD] = m._thread || '';
      newRows.push(row);
    }
  });

  // Slack se delete ho gaye messages ko IGNORE mark karo (sirf sync window ke andar)
  existing.forEach(r => {
    const ts = String(r[R.TS]);
    if (!ts || ts.indexOf('manual') === 0) return;
    if (Number(ts) * 1000 >= windowStartMs && !seen[ts] && !r[R.MAN_TYPE]) {
      r[R.MAN_TYPE] = 'IGNORE';
      r[R.NOTE] = 'Deleted in Slack';
    }
  });

  if (existing.length) sh.getRange(2, 1, existing.length, RAW_HEADERS.length).setValues(existing);
  if (newRows.length) {
    newRows.sort((a, b) => Number(a[R.TS]) - Number(b[R.TS]));
    sh.getRange(sh.getLastRow() + 1, 1, newRows.length, RAW_HEADERS.length).setValues(newRows);
  }
}

/************************ SHEET WRITERS ************************/
function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (headers.length && sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function writeTable_(sh, headers, rows) {
  sh.clearContents();
  sh.getRange(1, 1, sh.getMaxRows(), 2).setNumberFormat('@');   // dates text hi rahein
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sh.setFrozenRows(1);
  if (rows.length) sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
}

function writeIssues_(sh, issues) {
  const headers = ['Key', 'Date', 'Employee', 'Severity', 'Problem', 'Message', 'Slack Link', 'Alerted', 'Resolved'];
  const old = {};
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, headers.length).getValues().forEach(r => {
      old[r[0]] = { alerted: r[7], resolved: r[8] === true };
    });
  }
  const toAlert = [];
  const rows = issues.map(i => {
    const prev = old[i.key] || {};
    let alerted = prev.alerted || '';
    if (!alerted && i.severity !== 'INFO' && !prev.resolved) { toAlert.push(i); alerted = 'Yes'; }
    return [i.key, i.day, i.name, i.severity, i.problem, i.text || '', i.link || '', alerted, !!prev.resolved];
  });
  rows.sort((a, b) => (a[8] === b[8] ? (a[1] < b[1] ? 1 : -1) : (a[8] ? 1 : -1)));
  writeTable_(sh, headers, rows);
  if (rows.length) sh.getRange(2, 9, rows.length, 1).insertCheckboxes();
  if (toAlert.length) sendAlerts_(toAlert);
}

function sendAlerts_(all) {
  const notes = all.filter(i => i.severity === 'NOTICE');
  const issues = all.filter(i => i.severity !== 'NOTICE');
  if (notes.length) {
    const t = ':information_source: *Half day:*\n' + notes.slice(0, 25).map(i =>
      `• *${i.name}* (${i.day}) — ${i.problem}` + (i.link ? `  <${i.link}|open>` : '')).join('\n');
    CONFIG.ADMIN_SLACK_IDS.forEach(id => slackPost_('chat.postMessage', { channel: id, text: t }));
  }
  if (!issues.length) return;
  const lines = issues.slice(0, 25).map(i =>
    `• *${i.name}* (${i.day}) — ${i.problem}` + (i.text ? `\n    _"${i.text}"_` : '') + (i.link ? `  <${i.link}|open>` : ''));
  const more = issues.length > 25 ? `\n…aur ${issues.length - 25} issues. Sheet mein "Issues" tab dekhein.` : '';
  const text = `:warning: *Attendance script ko ${issues.length} cheez(on) mein masla aaya:*\n` +
    lines.join('\n') + more +
    `\n\nTheek karne ke liye: Sheet > "Raw Log" mein us row ka *Manual Type* / *Manual Time* bharein. ` +
    `<${ss_().getUrl()}|Sheet kholein>`;
  CONFIG.ADMIN_SLACK_IDS.forEach(id => slackPost_('chat.postMessage', { channel: id, text: text }));
  if (CONFIG.ADMIN_EMAILS.length) {
    MailApp.sendEmail(CONFIG.ADMIN_EMAILS.join(','), 'Attendance issues (' + issues.length + ')',
      text.replace(/[*_]/g, '').replace(/<([^|>]+)\|([^>]+)>/g, '$2: $1'));
  }
}

/* =====================================================================
 *  CORE LOGIC  (pure functions — Slack/Sheets ke baghair test ho sakti hain)
 * ===================================================================== */

/** ---------- time helpers ---------- */
function pad2(n) { return (n < 10 ? '0' : '') + n; }
function localParts(ms) {
  const d = new Date(ms + CONFIG.TZ_OFFSET_MINUTES * 60000);
  return { y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, d: d.getUTCDate(),
    h: d.getUTCHours(), mi: d.getUTCMinutes(), s: d.getUTCSeconds() };
}
function localToMs(y, m, d, h, mi) { return Date.UTC(y, m - 1, d, h, mi) - CONFIG.TZ_OFFSET_MINUTES * 60000; }
function workdayOf(ms) {
  const p = localParts(ms - CONFIG.DAY_START_HOUR * 3600000);
  return p.y + '-' + pad2(p.m) + '-' + pad2(p.d);
}
function fmtTime12(ms) {
  const p = localParts(ms);
  const h = p.h % 12 === 0 ? 12 : p.h % 12;
  return h + ':' + pad2(p.mi) + (p.h < 12 ? ' AM' : ' PM');
}
function fmtDateTime(ms) {
  const p = localParts(ms);
  return p.y + '-' + pad2(p.m) + '-' + pad2(p.d) + ' ' + fmtTime12(ms);
}
function parsePostedAt(s) {
  // "2026-10-23 12:27 PM"  (seconds Slack TS se aate hain; yeh sirf fallback hai)
  const m = String(s).match(/(\d{4})-(\d{2})-(\d{2})\s+(\d{1,2}):(\d{2})\s*(AM|PM)/i);
  if (!m) return null;
  let h = Number(m[4]) % 12; if (/pm/i.test(m[6])) h += 12;
  return localToMs(+m[1], +m[2], +m[3], h, +m[5]);
}
function fmtDur(min) {
  min = Math.round(min);
  const sign = min < 0 ? '-' : '';
  min = Math.abs(min);
  return sign + Math.floor(min / 60) + 'h ' + pad2(min % 60) + 'm';
}

/** ---------- fuzzy matching ---------- */
function osaDistance(a, b) {
  const d = [];
  for (let i = 0; i <= a.length; i++) { d[i] = [i]; }
  for (let j = 0; j <= b.length; j++) { d[0][j] = j; }
  for (let i = 1; i <= a.length; i++) {
    for (let j = 1; j <= b.length; j++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      d[i][j] = Math.min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + cost);
      if (i > 1 && j > 1 && a[i - 1] === b[j - 2] && a[i - 2] === b[j - 1]) {
        d[i][j] = Math.min(d[i][j], d[i - 2][j - 2] + 1);
      }
    }
  }
  return d[a.length][b.length];
}

const WORDS = {
  check: ['check', 'checked', 'chek', 'chk', 'chck', 'cheak', 'chack', 'chec', 'cheek', 'chcek'],
  namaz: ['namaz', 'nimaz', 'namaaz', 'nmaz', 'namz', 'nemaz', 'prayer', 'prayers', 'salah', 'salat',
    'jumma', 'juma', 'jummah', 'jumah', 'zuhr', 'zohr', 'zohar', 'zuhar', 'asr', 'asar', 'maghrib',
    'magrib', 'isha', 'esha', 'fajr', 'taraweeh'],
  brk: ['break', 'brk', 'brek', 'lunch', 'dinner', 'tea', 'chai', 'khana', 'lunchbreak', 'teabreak', 'rest'],
  outTypos: ['ot', 'otu', 'uot', 'outt', 'aut', 'oit', 'iut'],
  inTypos: ['inn', 'ib', 'un'],
  backIn: ['back', 'wapis', 'wapas', 'returned', 'resume', 'resumed'],
  half: ['halfday', 'half', 'halfdays', 'emergency', 'emergancy', 'urgent'],   // "Checkout (Halfday)"
  late: ['late', 'der', 'dair', 'dyr', 'lete'],                                  // "Aaj late aaunga" = pehle se bata diya
  filler: ['day', 'aaunga', 'aunga', 'aaungi', 'aungi', 'aaenge', 'aayenge', 'ayenge', 'jaunga', 'jaungi', 'se',
    'thora', 'thoda', 'thori', 'thodi', 'minute', 'minutes', 'min', 'mins', 'baje', 'bje', 'tak', 'gi', 'ga', 'sorry', 'kal',
    'sir', 'ji', 'g', 'for', 'the', 'and', 'to', 'from', 'after', 'am', 'pm', 'time', 'at', 'now',
    'done', 'ho', 'hon', 'hn', 'raha', 'rahi', 'hu', 'hoon', 'gaya', 'gya', 'kr', 'ker', 'kar', 'de', 'diya',
    'please', 'pls', 'today', 'aaj', 'aj', 'ok', 'jee', 'mein', 'me', 'my', 'of', 'a', 'is'],
};

function matchList(token, list) {
  if (list.indexOf(token) >= 0) return 'exact';
  if (token.length < 3) return null;
  for (let i = 0; i < list.length; i++) {
    if (list[i].length >= 4 && osaDistance(token, list[i]) <= 1) return 'fuzzy';
  }
  return null;
}

/** Message mein likha hua time nikaalo: "12:14", "12.14", "(12:14 pm)" */
function extractTypedTime(text) {
  const m = String(text).match(/(^|[^\d])(\d{1,2})\s*[:.]\s*(\d{2})\s*(a\.?\s?m\.?|p\.?\s?m\.?)?(?!\d)/i);
  if (!m) return null;
  const h = Number(m[2]), mi = Number(m[3]);
  if (h > 23 || mi > 59) return { invalid: true, raw: m[0].slice(m[1].length).trim() };
  const ap = m[4] ? (m[4].toLowerCase()[0] === 'a' ? 'am' : 'pm') : '';
  return { h: h, m: mi, ampm: ap, raw: m[0].slice(m[1].length).trim() };
}

/** Likha hua time message ke time ke hisaab se asli date-time banao */
function resolveTypedTime(tt, refMs, maxBackMin, allowForward) {
  const p = localParts(refMs);
  let hours;
  if (tt.ampm === 'am') hours = [tt.h % 12];
  else if (tt.ampm === 'pm') hours = [(tt.h % 12) + 12];
  else if (tt.h > 12 || tt.h === 0) hours = [tt.h];
  else hours = [tt.h % 12, (tt.h % 12) + 12];
  let best = null;
  [0, -1, 1].forEach(shift => {
    hours.forEach(h => {
      const ms = localToMs(p.y, p.m, p.d, h, tt.m) + shift * 86400000;
      const diff = (refMs - ms) / 60000;          // + matlab peeche ka time
      const ok = allowForward ? Math.abs(diff) <= maxBackMin : (diff >= -2 && diff <= maxBackMin);
      if (ok && (!best || Math.abs(diff) < Math.abs(best.diff))) best = { ms: ms, diff: diff };
    });
  });
  return best;
}

/**
 * Ek message parse karo.
 * return { type: IN|OUT_BREAK|OUT_NAMAZ|OUT_END|UNREADABLE, fuzzy, warn, typed }
 */
function parseMessage(text) {
  const typed = extractTypedTime(text);
  let s = String(text || '').toLowerCase()
    .replace(/<[^>]*>/g, ' ')                       // mentions / links
    .replace(/:[a-z_+\-][a-z0-9_+\-]*:/g, ' ');     // :emoji:
  if (typed) s = s.replace(typed.raw.toLowerCase(), ' ');
  s = s.replace(/check\s*-\s*/g, 'check').replace(/[^a-z]+/g, ' ').trim();
  const tokens = s ? s.split(/\s+/) : [];

  let hasCheck = false, hasIn = false, hasOut = false, weakIn = false, weakOutTypo = false, weakInTypo = false;
  let namaz = false, brk = false, half = false, late = false, fuzzy = false, unknown = [];

  tokens.forEach(t => {
    let used = false;
    if (t === 'in') { hasIn = true; used = true; }
    else if (t === 'out') { hasOut = true; used = true; }
    else if (t.length > 3 && /out$/.test(t) && matchList(t.slice(0, -3), WORDS.check)) {
      hasOut = true; hasCheck = true; used = true;
      if (matchList(t.slice(0, -3), WORDS.check) === 'fuzzy') fuzzy = true;
    } else if (t.length > 2 && /in$/.test(t) && matchList(t.slice(0, -2), WORDS.check)) {
      hasIn = true; hasCheck = true; used = true;
      if (matchList(t.slice(0, -2), WORDS.check) === 'fuzzy') fuzzy = true;
    }
    if (!used) {
      const c = matchList(t, WORDS.check);
      if (c) { hasCheck = true; used = true; if (c === 'fuzzy') fuzzy = true; }
    }
    if (!used && WORDS.late.indexOf(t) >= 0) { late = true; used = true; }
    if (!used) {
      const hd = matchList(t, WORDS.half);
      if (hd) { half = true; used = true; if (hd === 'fuzzy') fuzzy = true; }
    }
    if (!used) {
      const n = matchList(t, WORDS.namaz);
      if (n) { namaz = true; used = true; if (n === 'fuzzy') fuzzy = true; }
    }
    if (!used) {
      const b = matchList(t, WORDS.brk);
      if (b) { brk = true; used = true; if (b === 'fuzzy') fuzzy = true; }
    }
    if (!used && WORDS.outTypos.indexOf(t) >= 0) { weakOutTypo = true; used = true; }
    if (!used && WORDS.inTypos.indexOf(t) >= 0) { weakInTypo = true; used = true; }
    if (!used && WORDS.backIn.indexOf(t) >= 0) { weakIn = true; used = true; }
    if (!used && WORDS.filler.indexOf(t) >= 0) used = true;
    if (!used) unknown.push(t);
  });

  if (hasCheck && !hasIn && !hasOut) {
    if (weakOutTypo && !weakInTypo) { hasOut = true; fuzzy = true; }
    else if (weakInTypo && !weakOutTypo) { hasIn = true; fuzzy = true; }
  }

  const res = { type: 'UNREADABLE', fuzzy: fuzzy, warn: '', typed: typed };
  if (hasIn && hasOut) { res.warn = 'Message mein "in" aur "out" dono hain'; return res; }
  if (hasOut) {
    res.type = half ? 'OUT_HALFDAY' : namaz && brk ? 'OUT_BREAK' : namaz ? 'OUT_NAMAZ' : brk ? 'OUT_BREAK' : 'OUT_END';
    if (namaz && brk) res.warn = 'Namaz aur break dono likha hai — Break maana';
  } else if (hasIn) {
    res.type = 'IN';
  } else if (weakIn) {
    res.type = 'IN';
    if (!namaz && !brk) res.warn = '"back/wapis" ko check in maana';   // "back from namaz" = saaf check in
  } else if (late) {
    res.type = 'LATE';                        // in/out nahi, sirf "late aaunga" -> warning tag nahi jayega
    return res;
  } else {
    res.warn = 'Check in / check out samajh nahi aaya';
    return res;
  }
  if (unknown.length > 3) res.warn = (res.warn ? res.warn + '; ' : '') + 'Message mein extra text hai: "' + unknown.join(' ') + '"';
  return res;
}

/** Ek employee ke ek din ke events se hisaab */
function computeDay(events, dayIsOver, target) {
  if (target === undefined) target = CONFIG.TARGET_WORK_MINUTES;
  events.sort((a, b) => a.ms - b.ms);
  let state = 'NONE', start = 0, lastOut = 0, firstIn = 0, lastEnd = 0;
  let work = 0, brk = 0, namaz = 0, sessions = 0, halfOpen = false;
  const issues = [];
  events.forEach(e => {
    if (e.type === 'IN') {
      if (state === 'IN') {
        issues.push({ sev: 'WARNING', problem: 'Do baar check in (beech mein check out nahi) — doosra ignore kiya', ev: e });
        return;
      }
      if (state === 'BREAK') brk += (e.ms - lastOut) / 60000;
      if (state === 'NAMAZ') namaz += (e.ms - lastOut) / 60000;
      if (state === 'ENDED') {                // check out ke baad usi din wapas = beech ka waqt BREAK
        brk += (e.ms - lastOut) / 60000;
        halfOpen = false;                     // half day ke baad wapas aa gaya -> half day nahi
        issues.push({ sev: 'INFO', problem: 'Check out ke baad dobara check in — beech ka waqt break mein gina', ev: e });
      }
      if (!firstIn) firstIn = e.ms;
      start = e.ms; state = 'IN';
    } else {
      if (state !== 'IN') {
        issues.push({ sev: 'WARNING', problem: 'Check out bina check in ke — ignore kiya (shayad check in bhool gaya)', ev: e });
        return;
      }
      const dur = (e.ms - start) / 60000;
      if (dur < CONFIG.SHORT_SESSION_MINUTES) issues.push({ sev: 'INFO', problem: 'Bohat chhota session (' + Math.round(dur * 60) + ' sec) — shayad ghalti', ev: e });
      if (dur > CONFIG.LONG_SESSION_MINUTES) issues.push({ sev: 'WARNING', problem: 'Ek session ' + fmtDur(dur) + ' ka — check karein', ev: e });
      work += dur; sessions++; lastOut = e.ms; lastEnd = e.ms;
      state = e.type === 'OUT_BREAK' ? 'BREAK' : e.type === 'OUT_NAMAZ' ? 'NAMAZ' : 'ENDED';
      if (e.type === 'OUT_HALFDAY') {
        halfOpen = true;
        const sofar = work + (CONFIG.NAMAZ_COUNTS_AS_WORK ? namaz : 0);
        issues.push({ sev: 'NOTICE', problem: 'Half day check out ' + fmtTime12(e.ms) + ' — ' + fmtDur(sofar) + ' kaam kiya', ev: e });
      } else if (state === 'ENDED') halfOpen = false;
    }
  });
  const halfDay = halfOpen && state === 'ENDED';

  let status = 'Complete';
  if (state === 'IN') {
    if (dayIsOver) {
      status = 'Incomplete';
      issues.push({ sev: 'ERROR', problem: 'Check out nahi kiya — aakhri session gina nahi gaya', ev: null });
    } else status = 'Working now';
  } else if (state === 'BREAK' || state === 'NAMAZ') {
    if (dayIsOver) {
      status = 'No final check out';
      issues.push({ sev: 'WARNING', problem: (state === 'BREAK' ? 'Break' : 'Namaz') + ' ke liye check out kiya, wapis check in nahi — din yahin khatam maana', ev: null });
    } else status = state === 'BREAK' ? 'On break' : 'At namaz';
  } else if (state === 'NONE') {
    status = 'No valid check in';
  }
  if (halfDay) status = 'Half day';

  const counted = work + (CONFIG.NAMAZ_COUNTS_AS_WORK ? namaz : 0);
  const extraBreak = Math.max(0, brk - CONFIG.ALLOWED_BREAK_MINUTES);
  let diff = counted - target - (CONFIG.EXTRA_BREAK_COUNTS_AS_SHORT ? extraBreak : 0);
  let relief = 0;
  if (halfDay && diff < 0) { relief = -diff; diff = 0; }   // Half day: short time maaf (relief)
  return {
    halfDay: halfDay, relief: relief,
    firstIn: firstIn, lastOut: lastEnd, work: counted, brk: brk, namaz: namaz, extraBreak: extraBreak,
    sessions: sessions, overtime: diff > 0 ? diff : 0, short: diff < 0 ? -diff : 0,
    required: target, openSince: state === 'IN' && !dayIsOver ? start : 0,
    final: dayIsOver && status !== 'Incomplete' && status !== 'No valid check in', status: status, issues: issues,
  };
}

/** Raw log rows (display values) -> poori report */
function buildReport(rows, nowMs, empRows) {
  const nameToUid = {};
  rows.forEach(r => { if (r[R.UID]) nameToUid[r[R.NAME]] = r[R.UID]; });
  const todayWorkday = workdayOf(nowMs);
  const issues = [];
  const matchEmp = buildEmployeeMatcher(empRows || []);
  const uidToPortal = {}, portalToUids = {}, unmatched = {};
  const rawDayTime = [], rawNotes = [], rawTs = [];
  const groups = {};   // key uid|day -> {name, day, events}
  const lateInformed = {};   // key portal|day -> employee ne pehle se late bata diya
  let manualCounter = 0;

  rows.forEach((r, i) => {
    let ts = String(r[R.TS] || '').trim();
    const name = r[R.NAME], uid = r[R.UID] || nameToUid[name] || name;
    const manType = String(r[R.MAN_TYPE] || '').trim().toUpperCase();
    const manTime = String(r[R.MAN_TIME] || '').trim();
    const notes = [];
    let posted = ts && ts.indexOf('manual') !== 0 ? Math.round(Number(ts) * 1000) : null;
    if (!posted && r[R.POSTED]) posted = parsePostedAt(r[R.POSTED]);

    // Manual row (admin ne khud add ki) — TS khali ho to
    if (!ts) {
      if (!name || (!manType && !manTime)) { rawDayTime.push([r[R.DAY], r[R.TIME]]); rawNotes.push([r[R.NOTE]]); rawTs.push([ts]); return; }
      ts = 'manual-' + nowMs + '-' + (manualCounter++);
    }
    rawTs.push([ts]);

    const link = ts.indexOf('manual') === 0 ? '' :
      CONFIG.SLACK_WORKSPACE_URL + '/archives/' + CONFIG.CHANNEL_ID + '/p' + ts.replace('.', '') +
      (r[R.THREAD] ? '?thread_ts=' + r[R.THREAD] + '&cid=' + CONFIG.CHANNEL_ID : '');
    const parsed = parseMessage(r[R.TEXT]);

    // ----- type -----
    let type = parsed.type;
    if (manType) {
      if (TYPES.indexOf(manType) < 0) { notes.push('Manual Type ghalat'); }
      else { type = manType; notes.push('Manual type'); }
    }
    if (type === 'IGNORE') {
      rawDayTime.push([r[R.DAY], r[R.TIME]]);
      rawNotes.push([r[R.NOTE] === 'Deleted in Slack' ? 'Deleted in Slack' : 'Ignored']);
      return;
    }

    // ----- time -----
    let ms = posted;
    let issueDay = posted ? workdayOf(posted) : String(r[R.DAY] || '');
    if (manTime) {
      const tt = extractTypedTime(manTime);
      let ref = posted;
      if (!ref && /^\d{4}-\d{2}-\d{2}$/.test(String(r[R.DAY]))) {
        const d = String(r[R.DAY]).split('-').map(Number);
        ref = localToMs(d[0], d[1], d[2], 12, 0) + 0;     // din ka beech (12 PM) reference
      }
      const res = tt && !tt.invalid && ref ? resolveTypedTime(tt, ref, 18 * 60, true) : null;
      if (res) { ms = res.ms; notes.push('Manual time'); }
      else {
        issues.push({ key: 'MANTIME|' + ts, day: issueDay, name: name, severity: 'ERROR',
          problem: 'Manual Time samajh nahi aaya (format: 12:14 PM, aur manual row mein Date bhi likhein)', text: manTime, link: link });
      }
    } else if (parsed.typed && posted && !manType && parsed.type !== 'LATE') {
      if (parsed.typed.invalid) {
        notes.push('Likha time ghalat: ' + parsed.typed.raw);
        issues.push({ key: 'TYPEDBAD|' + ts, day: issueDay, name: name, severity: 'WARNING',
          problem: 'Message mein likha time ghalat hai — Slack ka time use kiya', text: r[R.TEXT], link: link });
      } else {
        const maxBack = /^OUT/.test(parsed.type) ? CONFIG.TYPED_TIME_MAX_BACK_CHECKOUT_MINUTES : CONFIG.TYPED_TIME_MAX_BACK_MINUTES;
        const res = resolveTypedTime(parsed.typed, posted, maxBack, false);
        if (res && Math.abs(res.diff) >= 2) {
          ms = res.ms; notes.push('Likha time use kiya (' + fmtTime12(res.ms) + ')');
          if (CONFIG.ALERT_ON_TYPED_TIME) issues.push({ key: 'TYPED|' + ts, day: workdayOf(ms), name: name, severity: 'WARNING',
            problem: 'Message mein likha time use kiya: ' + fmtTime12(res.ms) + ' (Slack par ' + fmtTime12(posted) + ' bheja)', text: r[R.TEXT], link: link });
        } else if (!res) {
          notes.push('Likha time match nahi hua — Slack time use kiya');
          issues.push({ key: 'TYPEDODD|' + ts, day: issueDay, name: name, severity: 'WARNING',
            problem: 'Likha time (' + parsed.typed.raw + ') message ke time se match nahi karta — Slack ka time use kiya', text: r[R.TEXT], link: link });
        }
      }
    }

    if (!ms && type === 'LATE' && /^\d{4}-\d{2}-\d{2}$/.test(String(r[R.DAY]))) ms = dayMs_(String(r[R.DAY]), '09:00');
    if (!ms) {
      rawDayTime.push([r[R.DAY], r[R.TIME]]); rawNotes.push([notes.concat('Time missing').join('; ')]);
      return;
    }
    const day = workdayOf(ms);
    rawDayTime.push([day, fmtTime12(ms)]);

    if (type === 'UNREADABLE') {
      notes.unshift('❓ ' + (parsed.warn || 'Samajh nahi aaya'));
      issues.push({ key: 'UNREAD|' + ts, day: day, name: name, severity: 'ERROR',
        problem: 'Message samajh nahi aaya — Manual Type bharein (IN / OUT_BREAK / OUT_NAMAZ / OUT_END / IGNORE)', text: r[R.TEXT], link: link });
      rawNotes.push([notes.join('; ')]);
      return;
    }
    if (!manType) {
      if (parsed.warn) {
        notes.push('⚠ ' + parsed.warn);
        issues.push({ key: 'WARN|' + ts, day: day, name: name, severity: 'WARNING',
          problem: 'Parha gaya "' + TYPE_LABEL[type] + '" lekin: ' + parsed.warn, text: r[R.TEXT], link: link });
      }
      if (parsed.fuzzy) notes.push('Spelling auto-correct');
    }
    notes.unshift(TYPE_LABEL[type]);
    rawNotes.push([notes.join('; ')]);

    // Slack bande ko portal Employee ID se milao
    let portal = uidToPortal[uid];
    if (portal === undefined) {
      portal = matchEmp(name);
      uidToPortal[uid] = portal;
      if (portal) (portalToUids[portal] = portalToUids[portal] || {})[uid] = name;
      else unmatched[uid] = name;
    }
    const gk = (portal || 'slack:' + uid) + '|' + day;
    if (type === 'LATE') { lateInformed[gk] = true; return; }      // sirf "bata diya" — hisaab mein koi event nahi
    if (!groups[gk]) groups[gk] = { uid: uid, portal: portal, name: name, day: day, events: [] };
    groups[gk].events.push({ ms: ms, type: type, ts: ts, text: r[R.TEXT], link: link });
  });

  Object.keys(unmatched).forEach(uid => {
    issues.push({ key: 'NOMATCH|' + uid, day: todayWorkday, name: unmatched[uid], severity: 'WARNING',
      problem: 'Yeh Slack banda portal ke "Employees" tab mein kisi ID se match nahi hua — Employees tab ke column C (Slack Name) mein "' + unmatched[uid] + '" likhein', text: '', link: '' });
  });
  Object.keys(portalToUids).forEach(pid => {
    const names = Object.keys(portalToUids[pid]).map(u => portalToUids[pid][u]);
    if (names.length > 1) issues.push({ key: 'MULTI|' + pid, day: todayWorkday, name: pid, severity: 'WARNING',
      problem: 'Ek hi Employee ID se kai Slack bande match hue: ' + names.join(', ') + ' — Employees tab ke column C mein sahi Slack Name likhein', text: '', link: '' });
  });

  // ----- daily -----
  const daily = [];
  const monthly = {};
  const records = [];   // overtime ke liye: sirf portal se match hue employees
  Object.keys(groups).forEach(gk => {
    const g = groups[gk];
    const over = g.day < todayWorkday;
    const dow = new Date(g.day + 'T12:00:00Z').getUTCDay();
    const offDay = CONFIG.WEEKLY_OFF_DAYS.indexOf(dow) >= 0;
    const d = computeDay(g.events, over, offDay ? 0 : CONFIG.TARGET_WORK_MINUTES);
    if (offDay && d.final) d.status = 'Weekly Off — Worked';
    const who = g.portal || g.name;
    d.issues.forEach(iss => {
      issues.push({
        key: 'DAY|' + gk + '|' + iss.problem.split(' ')[0] + '|' + (iss.ev ? iss.ev.ts : 'end'),
        day: g.day, name: who, severity: iss.sev, problem: iss.problem,
        text: iss.ev ? iss.ev.text : '', link: iss.ev ? iss.ev.link : '' });
    });
    daily.push([g.day, g.portal || '', g.name, d.firstIn ? fmtTime12(d.firstIn) : '', d.lastOut ? fmtTime12(d.lastOut) : '',
      fmtDur(d.work), fmtDur(d.brk), fmtDur(d.namaz), d.extraBreak ? fmtDur(d.extraBreak) : '',
      d.final ? (d.overtime ? fmtDur(d.overtime) : '') : '', d.final ? (d.short ? fmtDur(d.short) : '') : '',
      d.sessions, d.status,
      Math.round(d.work), Math.round(d.brk), Math.round(d.namaz),
      d.final ? Math.round(d.overtime) : 0, d.final ? Math.round(d.short) : 0, d.final ? 'Y' : 'N',
      Math.round(d.required), d.openSince ? String(d.openSince) : '',
      d.halfDay ? 'Y' : '', d.final ? Math.round(d.relief) : 0,
      lateInformed[gk] ? 'Informed' : (!offDay && d.firstIn && d.firstIn > lateCutoffMs_(g.day)) ? 'Y' : '']);

    if (g.portal) records.push({ id: g.portal, uid: g.uid, day: g.day, final: d.final, work: Math.round(d.work),
      required: Math.round(d.required), ot: d.final ? Math.round(d.overtime) : 0, short: d.final ? Math.round(d.short) : 0,
      status: d.status, firstInMs: d.firstIn || 0, lastOutMs: d.lastOut || 0, halfDay: !!d.halfDay });

    const mk = g.day.slice(0, 7) + '|' + (g.portal || g.uid);
    if (!monthly[mk]) monthly[mk] = { month: g.day.slice(0, 7), name: who, days: 0, work: 0, ot: 0, sh: 0, brk: 0, nmz: 0, inc: 0, half: 0 };
    const M = monthly[mk];
    M.days++; M.work += d.work; M.brk += d.brk; M.nmz += d.namaz;
    if (d.final) { M.ot += d.overtime; M.sh += d.short; }
    if (d.status === 'Incomplete' || d.status === 'No final check out') M.inc++;
    if (d.final && d.halfDay) M.half++;
  });
  daily.sort((a, b) => (a[0] === b[0] ? (a[2] < b[2] ? -1 : 1) : (a[0] < b[0] ? 1 : -1)));

  const monthlyRows = Object.keys(monthly).map(k => {
    const M = monthly[k];
    return [M.month, M.name, M.days, fmtDur(M.work), fmtDur(M.work / M.days), fmtDur(M.brk), fmtDur(M.nmz),
      fmtDur(M.ot), fmtDur(M.sh), fmtDur(M.ot - M.sh), M.inc, M.half];
  }).sort((a, b) => (a[0] === b[0] ? (a[1] < b[1] ? -1 : 1) : (a[0] < b[0] ? 1 : -1)));

  return {
    rawDayTime: rawDayTime, rawNotes: rawNotes, rawTs: rawTs, issues: issues, records: records, lateInformed: lateInformed,
    dailyHeaders: DAILY_HEADERS,
    daily: daily,
    monthlyHeaders: ['Month', 'Employee', 'Days', 'Total Work', 'Avg / Day', 'Total Break', 'Total Namaz',
      'Overtime', 'Short', 'Net (OT − Short)', 'Incomplete Days', 'Half Days'],
    monthly: monthlyRows,
  };
}

const DAILY_HEADERS = ['Date', 'Employee ID', 'Slack Name', 'First In', 'Last Out', 'Work', 'Break', 'Namaz',
  'Extra Break', 'Overtime', 'Short', 'Sessions', 'Status',
  'Work (min)', 'Break (min)', 'Namaz (min)', 'Overtime (min)', 'Short (min)', 'Final',
  'Required (min)', 'Open Since (ms)', 'Half Day', 'Relief (min)', 'Late'];

/** Employees tab rows [ID, Password, Slack Name?] -> function(slackName) => ID ya '' */
function buildEmployeeMatcher(empRows) {
  const list = empRows.map(r => ({
    id: String(r[0] || '').trim(),
    slack: String(r[2] || '').trim().toLowerCase().replace(/\s+/g, ' '),
  })).filter(e => e.id);
  const one = arr => (arr.length === 1 ? arr[0].id : '');
  return function (slackName) {
    const n = String(slackName || '').trim().toLowerCase().replace(/\s+/g, ' ');
    if (!n) return '';
    const words = n.split(' ');
    const free = list.filter(e => !e.slack);
    return one(list.filter(e => e.slack && e.slack === n))
      || one(free.filter(e => e.id.toLowerCase() === n))
      || one(free.filter(e => words.indexOf(e.id.toLowerCase()) >= 0))
      || one(free.filter(e => e.id.toLowerCase().split(/\s+/)[0] === words[0]));
  };
}

/* =====================================================================
 *  PORTAL KE LIYE DATA (Code.gs isay use karta hai)
 * ===================================================================== */

/** Daily Summary se ek employee (ya sab) ki attendance rows */
function getAttendanceData_(ss, onlyId) {
  const sh = ss.getSheetByName(SHEETS.DAILY);
  const out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  const vals = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getDisplayValues();
  const h = vals[0];
  const c = name => h.indexOf(name);
  const iDate = c('Date'), iId = c('Employee ID');
  if (iId < 0) return out;
  for (let i = 1; i < vals.length; i++) {
    const r = vals[i];
    const id = r[iId];
    if (!id || (onlyId && id !== onlyId)) continue;
    (out[id] = out[id] || []).push({
      date: r[iDate], firstIn: r[c('First In')], lastOut: r[c('Last Out')],
      work: Number(r[c('Work (min)')]) || 0, brk: Number(r[c('Break (min)')]) || 0,
      namaz: Number(r[c('Namaz (min)')]) || 0, ot: Number(r[c('Overtime (min)')]) || 0,
      short: Number(r[c('Short (min)')]) || 0, status: r[c('Status')], final: r[c('Final')] === 'Y',
      required: c('Required (min)') >= 0 && r[c('Required (min)')] !== '' ? Number(r[c('Required (min)')]) : null,
      openSince: Number(r[c('Open Since (ms)')]) || 0,
      halfDay: c('Half Day') >= 0 && r[c('Half Day')] === 'Y', relief: Number(r[c('Relief (min)')]) || 0,
      late: c('Late') >= 0 ? String(r[c('Late')] || '') : '',
    });
  }
  return out;
}

/** Issues tab se khule (unresolved) masle — admin panel ke liye */
function getOpenIssues_(ss, limit) {
  const sh = ss.getSheetByName(SHEETS.ISSUES);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 9).getDisplayValues()
    .filter(r => String(r[8]).toUpperCase() !== 'TRUE' && r[3] !== 'INFO' && r[3] !== 'NOTICE')
    .slice(0, limit || 50)
    .map(r => ({ date: String(r[1]), name: String(r[2]), severity: r[3], problem: r[4], text: r[5], link: r[6] }));
}

/* =====================================================================
 *  OVERTIME  ->  EXTRA LEAVES   (3 tabs)
 *
 *  "Overtime Months"   PERMANENT. Har employee ka har mahina:
 *                      Carried In + Overtime − Short = Total
 *                      Har baar balance 8h tak pohanche = 1 extra leave (−8h)
 *                      Baqi = Carried Forward (agle mahine ka Carried In)
 *                      Status: "Running" (chal raha) -> "Final" (lock)
 *                      Mahina khatam hone ke SETTLE_GRACE_DAYS din baad Final.
 *                      Final row ko script kabhi dobara nahi badalti / mitati.
 *  "Earned Leaves"     PERMANENT. Har extra leave ki alag row (kis din, kis mahine,
 *                      kitna overtime laga). Final rows kabhi nahi mitti.
 *  "Overtime History"  Sirf aakhri OVERTIME_HISTORY_WEEKS haftay (din-ba-din + week total).
 *                      Har sync par dobara likha jata hai — purana hafta khud hat jata hai.
 *                      Cleanup SIRF isi tab par hoti hai.
 *
 *  Balance kabhi reset nahi hota: agla mahina pichle FINAL mahine ke
 *  "Carried Forward" se shuru hota hai (daily data ki zaroorat nahi).
 *
 *  Admin ko kisi Final mahine ko dobara calculate karwana ho (bohat kam):
 *  "Overtime Months" mein us mahine AUR baad ke saare mahinon ki rows, aur
 *  "Earned Leaves" mein un mahinon ki rows delete karein, phir "Recalculate only".
 * ===================================================================== */
const OTM_HEADERS = ['Key', 'Employee ID', 'Month', 'Status', 'Carried In', 'Overtime', 'Short', 'Total',
  'Extra Leaves', 'Carried Forward', 'Days Counted', 'Updated On',
  'Carried In (min)', 'Overtime (min)', 'Short (min)', 'Total (min)', 'Carried Forward (min)', 'Paid Leaves Cut'];
const OTM = { KEY: 0, ID: 1, MONTH: 2, STATUS: 3, LEAVES: 8, DAYS: 10, CIN: 12, OT: 13, SH: 14, TOT: 15, COUT: 16, CUT: 17, TEXT_COLS: 12 };
const EARNED_HEADERS = ['Key', 'Employee ID', 'Date Generated', 'Month', 'Overtime Used', 'Extra Leaves',
  'Balance After', 'Status', 'Saved On', 'Balance After (min)', 'Type'];
const EAR = { KEY: 0, ID: 1, DATE: 2, MONTH: 3, LEAVES: 5, STATUS: 7, AFTER: 9, TYPE: 10, TEXT_COLS: 9 };
const OTH_HEADERS = ['Employee ID', 'Week Starting', 'Date', 'Day', 'Worked', 'Required', 'Overtime', 'Short',
  'Worked (min)', 'Required (min)', 'Overtime (min)', 'Short (min)'];

/** Ek mahine ka hisaab, din-ba-din (pure function) */
function runMonth(carryIn, days) {
  const per = CONFIG.OVERTIME_PER_LEAVE_MINUTES, red = CONFIG.SHORT_REDUCES_OVERTIME;
  let bal = carryIn, ot = 0, short = 0, counted = 0;
  const events = [], warnings = [];
  days.forEach(d => {
    if (!d.final) return;
    counted++; ot += d.ot; short += d.short;
    const before = bal;
    bal += d.ot - (red ? d.short : 0);
    while (bal >= per) { bal -= per; events.push({ day: d.day, after: bal, type: 'Earned' }); }     // +1 extra leave
    while (bal <= -per) { bal += per; events.push({ day: d.day, after: bal, type: 'Cut' }); }       // −1 paid leave
    if (before >= 0 && bal < 0) warnings.push({ day: d.day, after: bal });                          // minus shuru
  });
  const earned = events.filter(e => e.type === 'Earned').length;
  return { carryIn: carryIn, ot: ot, short: short, total: carryIn + ot - (red ? short : 0),
    leaves: earned, cut: events.length - earned, carryOut: bal, counted: counted, events: events, warnings: warnings };
}

function nextMonth_(m) {
  const y = Number(m.slice(0, 4)), mo = Number(m.slice(5, 7));
  return mo === 12 ? (y + 1) + '-01' : y + '-' + pad2(mo + 1);
}
function ymdUTC_(d) { return d.getUTCFullYear() + '-' + pad2(d.getUTCMonth() + 1) + '-' + pad2(d.getUTCDate()); }
/** Mahina khatam + grace din guzar gaye? */
function monthClosed_(m, today) {
  const y = Number(m.slice(0, 4)), mo = Number(m.slice(5, 7));
  return today >= ymdUTC_(new Date(Date.UTC(y, mo, CONFIG.SETTLE_GRACE_DAYS)));   // (y, mo, 0) = mahine ka aakhri din
}
function weekStartOf_(day) {
  const d = new Date(day + 'T12:00:00Z');
  d.setUTCDate(d.getUTCDate() - (d.getUTCDay() + 6) % 7);   // Monday
  return ymdUTC_(d);
}
function addDaysStr_(day, n) {
  const d = new Date(day + 'T12:00:00Z'); d.setUTCDate(d.getUTCDate() + n); return ymdUTC_(d);
}

function readRows_(sh, n) {
  return sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, n).getValues() : [];
}
/** Bina clear kiye likho (pehle text format, phir values, phir sirf faltu neeche wali rows saaf) */
function writeRowsSafe_(sh, headers, rows, textCols) {
  sh.getRange(1, 1, Math.max(sh.getMaxRows(), rows.length + 1), textCols).setNumberFormat('@');
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sh.setFrozenRows(1);
  if (rows.length) sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  const extra = sh.getLastRow() - (rows.length + 1);
  if (extra > 0) sh.getRange(rows.length + 2, 1, extra, headers.length).clearContent();
}

/** Har sync ke baad: Overtime Months + Earned Leaves (permanent) aur Overtime History (2 haftay) */
function updateOvertime_(ss, records, nowMs) {
  const today = workdayOf(nowMs), curMonth = today.slice(0, 7), stamp = fmtDateTime(nowMs);
  const mSh = ensureSheet_(ss, SHEETS.OT_MONTHS, OTM_HEADERS);
  const eSh = ensureSheet_(ss, SHEETS.EARNED, EARNED_HEADERS);
  const mRows = readRows_(mSh, OTM_HEADERS.length);
  const eRows = readRows_(eSh, EARNED_HEADERS.length);
  const plan = planOvertime(records, mRows, eRows, today, stamp);

  // ---- SAFETY: koi Final row badli ya gum hui to kuch mat likho ----
  const outM = {}; plan.months.forEach(r => outM[r[OTM.KEY]] = JSON.stringify(r));
  const outE = {}; plan.events.forEach(r => outE[r[EAR.KEY]] = JSON.stringify(r));
  mRows.forEach(r => {
    if (r[OTM.STATUS] === 'Final' && outM[r[OTM.KEY]] !== JSON.stringify(r)) throw new Error('Safety stop: Final month ' + r[OTM.KEY] + ' change ho raha tha. Kuch nahi likha gaya.');
  });
  eRows.forEach(r => {
    if (r[EAR.STATUS] === 'Final' && plan.finalMonths[String(r[EAR.ID]) + '|' + String(r[EAR.MONTH])] && outE[r[EAR.KEY]] !== JSON.stringify(r))
      throw new Error('Safety stop: Final earned leave ' + r[EAR.KEY] + ' change ho rahi thi. Kuch nahi likha gaya.');
  });

  writeRowsSafe_(mSh, OTM_HEADERS, plan.months, OTM.TEXT_COLS);
  writeRowsSafe_(eSh, EARNED_HEADERS, plan.events, EAR.TEXT_COLS);

  // ---- Overtime History: sirf aakhri N haftay (yahi cleanup hai) ----
  const hSh = ensureSheet_(ss, SHEETS.OT_HISTORY, []);
  hSh.clearContents();
  hSh.getRange(1, 1, hSh.getMaxRows(), 4).setNumberFormat('@');
  hSh.getRange(1, 1, 1, OTH_HEADERS.length).setValues([OTH_HEADERS]).setFontWeight('bold');
  hSh.setFrozenRows(1);
  const hist = buildOvertimeHistory(records, today);
  if (hist.length) hSh.getRange(2, 1, hist.length, OTH_HEADERS.length).setValues(hist);
  return plan;
}

/** Pure: nayi Overtime Months + Earned Leaves rows banao. Final rows jaisi ki taisi. */
function planOvertime(records, mRows, eRows, today, stamp) {
  const curMonth = today.slice(0, 7);
  const start = CONFIG.OVERTIME_START_DATE, startMonth = start.slice(0, 7);
  records = records.filter(r => r.day >= start);          // start date se pehle ka overtime/short nahi ginta
  const months = mRows.map(r => r.slice());
  const idx = {}; months.forEach((r, i) => idx[r[OTM.KEY]] = i);
  const finalMonths = {}; months.forEach(r => { if (r[OTM.STATUS] === 'Final') finalMonths[r[OTM.KEY]] = true; });

  const byEmp = {};
  records.forEach(r => {
    const m = r.day.slice(0, 7), e = byEmp[r.id] = byEmp[r.id] || {};
    (e[m] = e[m] || []).push(r);
  });
  months.forEach(r => { byEmp[r[OTM.ID]] = byEmp[r[OTM.ID]] || {}; });

  const touched = {}, newEvents = [], warnings = [];
  Object.keys(byEmp).forEach(id => {
    // start month se pehle ki purani rows sheet mein rehti hain lekin hisaab mein nahi aatin
    const list = Object.keys(byEmp[id]).concat(months.filter(r => r[OTM.ID] === id).map(r => String(r[OTM.MONTH])))
      .filter(m => m >= startMonth).sort();
    if (!list.length) return;
    let carry = 0;
    for (let m = list[0]; m <= curMonth; m = nextMonth_(m)) {
      const key = id + '|' + m, i = idx[key], row = i === undefined ? null : months[i];
      if (row && row[OTM.STATUS] === 'Final') { carry = Number(row[OTM.COUT]) || 0; continue; }   // LOCKED
      const days = (byEmp[id][m] || []).slice().sort((a, b) => (a.day < b.day ? -1 : 1));
      if (!days.length && !carry && !row) continue;
      const res = runMonth(carry, days);
      const status = monthClosed_(m, today) ? 'Final' : 'Running';
      const out = [key, id, m, status, fmtDur(res.carryIn), fmtDur(res.ot), fmtDur(res.short), fmtDur(res.total),
        String(res.leaves), fmtDur(res.carryOut), String(res.counted), stamp,
        res.carryIn, res.ot, res.short, res.total, res.carryOut, res.cut];
      if (row) months[i] = out; else { idx[key] = months.length; months.push(out); }
      touched[key] = true;
      res.events.forEach((ev, n) => newEvents.push([key + '|' + (n + 1), id, ev.day, m,
        ev.type === 'Cut' ? fmtDur(CONFIG.OVERTIME_PER_LEAVE_MINUTES) + ' short' : fmtDur(CONFIG.OVERTIME_PER_LEAVE_MINUTES),
        ev.type === 'Cut' ? '-1' : '1', fmtDur(ev.after), status === 'Final' ? 'Final' : 'This month', stamp, ev.after, ev.type]));
      res.warnings.forEach(w => warnings.push({ id: id, day: w.day, after: w.after }));
      carry = res.carryOut;
    }
  });

  // Earned Leaves: jo mahine dobara calculate nahi hue un ki rows jaisi ki taisi; baqi nayi
  const events = eRows.filter(r => !touched[String(r[EAR.ID]) + '|' + String(r[EAR.MONTH])]).map(r => r.slice()).concat(newEvents);
  events.sort((a, b) => (String(a[EAR.DATE]) === String(b[EAR.DATE]) ? (a[EAR.KEY] < b[EAR.KEY] ? -1 : 1) : (String(a[EAR.DATE]) < String(b[EAR.DATE]) ? -1 : 1)));
  return { months: months, events: events, finalMonths: finalMonths, warnings: warnings };
}

/** Pure: aakhri N haftay, din-ba-din + har haftay ka total */
function buildOvertimeHistory(records, today) {
  const n = CONFIG.OVERTIME_HISTORY_WEEKS, first = addDaysStr_(weekStartOf_(today), -7 * (n - 1));
  records = records.filter(r => r.day >= CONFIG.OVERTIME_START_DATE);
  const rows = [];
  const ids = {}; records.forEach(r => { if (r.final && r.day >= first) ids[r.id] = true; });
  Object.keys(ids).sort().forEach(id => {
    for (let w = 0; w < n; w++) {
      const ws = addDaysStr_(first, 7 * w), we = addDaysStr_(ws, 6);
      const days = records.filter(r => r.id === id && r.final && r.day >= ws && r.day <= we).sort((a, b) => (a.day < b.day ? -1 : 1));
      if (!days.length) continue;
      const t = { work: 0, req: 0, ot: 0, sh: 0 };
      days.forEach(d => {
        const dow = new Date(d.day + 'T12:00:00Z').toUTCString().slice(0, 3);
        rows.push([id, ws, d.day, dow, fmtDur(d.work), fmtDur(d.required), fmtDur(d.ot), fmtDur(d.short), d.work, d.required, d.ot, d.short]);
        t.work += d.work; t.req += d.required; t.ot += d.ot; t.sh += d.short;
      });
      rows.push([id, ws, '', 'Week total', fmtDur(t.work), fmtDur(t.req), fmtDur(t.ot), fmtDur(t.sh), t.work, t.req, t.ot, t.sh]);
    }
  });
  return rows;
}

/** Website ke liye: ek employee (ya sab) ka overtime data */
function getOvertimeData_(ss, onlyId) {
  const out = {};
  const get = id => (out[id] = out[id] || { months: [], earned: [], history: [] });
  const keep = id => id && (!onlyId || id === onlyId);
  const start = CONFIG.OVERTIME_START_DATE, startMonth = start.slice(0, 7);
  const mSh = ss.getSheetByName(SHEETS.OT_MONTHS), eSh = ss.getSheetByName(SHEETS.EARNED), hSh = ss.getSheetByName(SHEETS.OT_HISTORY);
  if (mSh) readRows_(mSh, OTM_HEADERS.length).forEach(r => {
    const id = String(r[OTM.ID]); if (!keep(id) || String(r[OTM.MONTH]) < startMonth) return;
    get(id).months.push({ month: String(r[OTM.MONTH]), status: String(r[OTM.STATUS]), carryIn: Number(r[OTM.CIN]) || 0,
      ot: Number(r[OTM.OT]) || 0, short: Number(r[OTM.SH]) || 0, total: Number(r[OTM.TOT]) || 0,
      leaves: Number(r[OTM.LEAVES]) || 0, carryOut: Number(r[OTM.COUT]) || 0, days: Number(r[OTM.DAYS]) || 0,
      cut: Number(r[OTM.CUT]) || 0 });
  });
  if (eSh) readRows_(eSh, EARNED_HEADERS.length).forEach(r => {
    const id = String(r[EAR.ID]); if (!keep(id) || String(r[EAR.DATE]) < start) return;
    get(id).earned.push({ date: String(r[EAR.DATE]), month: String(r[EAR.MONTH]), used: CONFIG.OVERTIME_PER_LEAVE_MINUTES,
      leaves: Number(r[EAR.LEAVES]) || 1, after: Number(r[EAR.AFTER]) || 0, status: String(r[EAR.STATUS]),
      type: String(r[EAR.TYPE] || '') === 'Cut' || Number(r[EAR.LEAVES]) < 0 ? 'Cut' : 'Earned' });
  });
  if (hSh) readRows_(hSh, OTH_HEADERS.length).forEach(r => {
    const id = String(r[0]); if (!keep(id) || !r[2] || String(r[2]) < start) return;      // "Week total" rows website khud jorti hai
    get(id).history.push({ week: String(r[1]), date: String(r[2]), work: Number(r[8]) || 0, required: Number(r[9]) || 0,
      ot: Number(r[10]) || 0, short: Number(r[11]) || 0 });
  });
  Object.keys(out).forEach(id => out[id].months.sort((a, b) => (a.month < b.month ? -1 : 1)));
  return out;
}


/* =====================================================================
 *  REMINDERS + WARNINGS  (channel mein employee ko tag, aur DM)
 *
 *  Channel (#daily-check-in), employee @tag:
 *   - LATE        12:10 tak check in nahi (Mon–Sat, chutti/Sunday nahi, "late" bata diya ho to nahi)
 *   - LATE_REPEAT pichla kaam ka din bhi late tha (bina bataye) -> "routine sahi karein"
 *   - BREAK       break / namaz ke baad 1 ghanta, check in nahi
 *   - NIGHT       raat 11 baje tak check out nahi
 *  DM (employee ko):
 *   - NOT_COUNTED pichle din check out nahi -> time gina nahi gaya
 *   - EARNED      +1 extra leave 🎉
 *   - NEGATIVE    time balance minus mein gaya -> paid leave bacha lo
 *   - CUT         −1 paid leave (8h short)
 *  Har message sirf EK dafa: "Notices" tab mein record (yahi history + duplicate se bachao).
 *  Script theek waqt par khud jaagti hai (reminderTick): 12:10, break + 1 ghanta, 11 PM, 5 AM.
 * ===================================================================== */
const NOTICE_HEADERS = ['Key', 'Date', 'Employee ID', 'Type', 'Sent To', 'Message', 'Sent At', 'Result'];

function hm_(s) { const p = String(s).split(':'); return { h: Number(p[0]), m: Number(p[1] || 0) }; }
function dayMs_(day, hhmm) {
  const t = hm_(hhmm), d = day.split('-').map(Number);
  return localToMs(d[0], d[1], d[2], t.h, t.m);
}
function lateCutoffMs_(day) { return dayMs_(day, CONFIG.LATE_AFTER); }
function isOffDay_(day) { return CONFIG.WEEKLY_OFF_DAYS.indexOf(new Date(day + 'T12:00:00Z').getUTCDay()) >= 0; }
function niceDate_(day) {
  const M = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const d = day.split('-').map(Number); return d[2] + ' ' + M[d[1] - 1];
}
/** "14-Feb-2026", "01-June-2026", "2026-02-14", "14/02/2026", Date -> "2026-02-14" */
function normDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const s = String(v || '').trim();
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/); if (m) return m[1] + '-' + pad2(+m[2]) + '-' + pad2(+m[3]);
  m = s.match(/^(\d{1,2})[-\s/]([A-Za-z]+)[-\s/,]*(\d{4})$/);
  if (m) { const i = ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'].indexOf(m[2].slice(0, 3).toLowerCase()); if (i >= 0) return m[3] + '-' + pad2(i + 1) + '-' + pad2(+m[1]); }
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/); if (m) return m[3] + '-' + pad2(+m[2]) + '-' + pad2(+m[1]);
  return '';
}

/** Pure: abhi kaun se messages banne chahiyein (duplicate check baad mein) */
function planNotices(records, plan, lateInformed, leaveDays, nowMs) {
  const out = [];
  const today = workdayOf(nowMs);
  const recent = addDaysStr_(today, -1);
  const byKey = {}, uidOf = {}, lastSeen = {};
  records.forEach(r => {
    byKey[r.id + '|' + r.day] = r;
    if (r.uid && String(r.uid).indexOf('U') === 0) uidOf[r.id] = r.uid;
    if (!lastSeen[r.id] || r.day > lastSeen[r.id]) lastSeen[r.id] = r.day;
  });
  const tag = id => '<@' + uidOf[id] + '>';
  const onLeave = (id, day) => !!(leaveDays[id] && leaveDays[id][day]);
  const working = (id, day) => !isOffDay_(day) && !onLeave(id, day);
  const HR = CONFIG.HR_NAMES;
  const active = Object.keys(uidOf).filter(id => lastSeen[id] >= addDaysStr_(today, -14));   // pichle 14 din mein aaya

  active.forEach(id => {
    const rec = byKey[id + '|' + today];
    // ---- LATE (12:10 ke baad, Mon–Sat, chutti nahi, bataya nahi) ----
    const cut = lateCutoffMs_(today), night = dayMs_(today, CONFIG.NIGHT_CHECK);
    if (working(id, today) && nowMs >= cut && nowMs < cut + 2 * 3600000 && !lateInformed[id + '|' + today]) {
      const lateToday = !rec || !rec.firstInMs || rec.firstInMs > cut;
      if (lateToday) {
        let prev = addDaysStr_(today, -1), guard = 0;
        while (guard++ < 10 && !working(id, prev)) prev = addDaysStr_(prev, -1);
        const pr = byKey[id + '|' + prev];
        const prevLate = pr && pr.firstInMs > lateCutoffMs_(prev) && !lateInformed[id + '|' + prev];
        const text = prevLate
          ? `⚠️ ${tag(id)} aap daily late aa rahe hain, yeh routine sahi nahi hai. Meharbani kar ke is routine ko sahi karein. Shukriya.`
          : rec && rec.firstInMs
            ? `⚠️ ${tag(id)} aap aaj late aaye (${fmtTime12(rec.firstInMs)}). Kindly ${HR} ko inform karein.`
            : `⚠️ ${tag(id)} aap ne abhi tak check in nahi kiya (late arrival). Kindly check in karein ya ${HR} ko inform karein.`;
        out.push({ key: 'LATE|' + id + '|' + today, day: today, id: id, type: prevLate ? 'LATE_REPEAT' : 'LATE', to: 'channel', text: text });
      }
    }
    // ---- BREAK / NAMAZ 1 ghanta ----
    if (rec && !rec.final && (rec.status === 'On break' || rec.status === 'At namaz') && rec.lastOutMs) {
      const due = rec.lastOutMs + CONFIG.BREAK_REMIND_MINUTES * 60000;
      if (nowMs >= due && nowMs < due + 2 * 3600000) {
        out.push({ key: 'BREAK|' + id + '|' + rec.lastOutMs, day: today, id: id, type: 'BREAK', to: 'channel',
          text: `⚠️ ${tag(id)} aap ne ${rec.status === 'At namaz' ? 'namaz' : 'break'} ke baad check in nahi kiya, ya aap ka break time over ho gaya hai. Kindly check karein ya ${HR} ko inform karein.` });
      }
    }
    // ---- RAAT 11 baje tak check out nahi ----
    if (rec && !rec.final && rec.status === 'Working now' && nowMs >= dayMs_(today, CONFIG.NIGHT_CHECK)) {
      out.push({ key: 'NIGHT|' + id + '|' + today, day: today, id: id, type: 'NIGHT', to: 'channel',
        text: `${tag(id)} aap ne check out nahi kiya. Kya aap overtime kar rahe hain? Agar haan, to kaam khatam kar ke check out karein. Agar nahi, to jis time nikle the woh time likh kar check out karein, jaise: Checkout (9:30)` });
    }
    // ---- Kal ka din gina nahi gaya (check out nahi) -> DM ----
    const yday = addDaysStr_(today, -1), yr = byKey[id + '|' + yday];
    if (yr && yr.status === 'Incomplete') {
      out.push({ key: 'NOTCOUNTED|' + id + '|' + yday, day: yday, id: id, type: 'NOT_COUNTED', to: 'dm',
        text: `Aap ka ${niceDate_(yday)} ka time gina nahi gaya kyun ke check out nahi tha. Meharbani kar ke ${HR} se raabta kar ke isay sahi karwayein, aur next time khayal rakhein.` });
    }
  });

  // ---- Leaves: +1 mili / −1 kati, aur balance minus mein gaya -> DM ----
  const seen = {};
  (plan && plan.events || []).forEach(r => {
    const id = String(r[EAR.ID]), day = String(r[EAR.DATE]);
    if (!uidOf[id] || day < recent || day < CONFIG.OVERTIME_START_DATE) return;
    const type = String(r[EAR.TYPE] || '') === 'Cut' || Number(r[EAR.LEAVES]) < 0 ? 'CUT' : 'EARNED';
    const k = id + '|' + day + '|' + type; seen[k] = (seen[k] || 0) + 1;
    const after = fmtDur(Number(r[EAR.AFTER]) || 0);
    out.push({ key: type + '|' + k + '|' + seen[k], day: day, id: id, type: type, to: 'dm',
      text: type === 'EARNED'
        ? `🎉 Mubarak ho! Aap ka overtime 8 ghante poora ho gaya, aap ko 1 extra leave mil gayi (${niceDate_(day)}). Baqi overtime balance: ${after}. Shabash, aise hi lage rahein!`
        : `⚠️ Aap ka short time 8 ghante ho gaya, is liye 1 paid leave kat gayi (${niceDate_(day)}). Balance ab: ${after}. Overtime kar ke time wapas cover kar sakte hain.` });
  });
  (plan && plan.warnings || []).forEach(w => {
    if (!uidOf[w.id] || w.day < recent) return;
    out.push({ key: 'NEGATIVE|' + w.id + '|' + w.day, day: w.day, id: w.id, type: 'NEGATIVE', to: 'dm',
      text: `⚠️ Aap ka time balance minus mein chala gaya (${fmtDur(w.after)}). Minus 8 ghante hone par 1 paid leave kat jayegi, apni paid leave bacha lein bhai 🙂 Thora overtime kar ke yeh time wapas pura ho jata hai.` });
  });
  out.forEach(n => { n.uid = uidOf[n.id]; });
  return out;
}

/** Notices bhejo (sirf naye), aur "Notices" tab mein likho */
function runNotices_(ss, result, plan, nowMs) {
  const firstRun = !ss.getSheetByName(SHEETS.NOTICES);      // pehli dafa: purani cheezon ke messages NAHI bhejte
  const sh = ensureSheet_(ss, SHEETS.NOTICES, NOTICE_HEADERS);
  const sent = {};
  readRows_(sh, 1).forEach(r => { sent[String(r[0])] = true; });
  const leaveDays = {};
  if (typeof buildEmployees_ === 'function') {
    const emps = buildEmployees_(ss);
    Object.keys(emps).forEach(id => (emps[id].leaves || []).forEach(l => {
      const d = normDate_(l.date); if (d) (leaveDays[id] = leaveDays[id] || {})[d] = true;
    }));
  }
  const list = planNotices(result.records, plan, result.lateInformed || {}, leaveDays, nowMs).filter(n => !sent[n.key]);
  if (!list.length) return;
  const stamp = fmtDateTime(nowMs), rows = [];
  list.forEach(n => {
    let res = 'Not sent (SEND_EMPLOYEE_MESSAGES = false)';
    if (firstRun) res = 'Skipped (first run)';
    else if (CONFIG.SEND_EMPLOYEE_MESSAGES) {
      const r = slackPost_('chat.postMessage', { channel: n.to === 'dm' ? n.uid : CONFIG.CHANNEL_ID, text: n.text });
      res = r && r.ok ? 'Sent' : 'Failed: ' + (r && r.error);
    }
    rows.push([n.key, n.day, n.id, n.type, n.to === 'dm' ? 'DM' : '#daily-check-in', n.text, stamp, res]);
  });
  sh.getRange(1, 1, sh.getMaxRows(), 2).setNumberFormat('@');
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, NOTICE_HEADERS.length).setValues(rows);
}

/** Agla reminder kab? Usi waqt ke liye ek trigger (12:10, break + 1h, 11 PM, 5 AM) */
function nextReminderMs(records, nowMs) {
  const today = workdayOf(nowMs), times = [];
  if (!isOffDay_(today)) times.push(lateCutoffMs_(today) + 30000);
  times.push(dayMs_(today, CONFIG.NIGHT_CHECK) + 30000);
  times.push(dayMs_(addDaysStr_(today, 1), '05:05'));                   // din badla -> "gina nahi gaya" DM
  records.forEach(r => {
    if (r.day === today && !r.final && (r.status === 'On break' || r.status === 'At namaz') && r.lastOutMs)
      times.push(r.lastOutMs + CONFIG.BREAK_REMIND_MINUTES * 60000 + 30000);
  });
  const future = times.filter(t => t > nowMs + 30000).sort((a, b) => a - b);
  return future.length ? future[0] : 0;
}

function scheduleReminders_(records, nowMs) {
  const next = nextReminderMs(records, nowMs);
  const props = PropertiesService.getScriptProperties();
  const have = ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'reminderTick');
  if (next && have.length === 1 && props.getProperty('NEXT_TICK') === String(next)) return;   // pehle se set
  have.forEach(t => ScriptApp.deleteTrigger(t));
  if (!next) return;
  ScriptApp.newTrigger('reminderTick').timeBased().at(new Date(next)).create();
  props.setProperty('NEXT_TICK', String(next));
}

/** Reminder ke waqt chalti hai: taaza Slack data laao, phir check (recalculate khud agla reminder set karta hai) */
function reminderTick() {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'reminderTick') ScriptApp.deleteTrigger(t); });
  PropertiesService.getScriptProperties().deleteProperty('NEXT_TICK');
  syncWindow_(CONFIG.QUICK_SYNC_DAYS);
}
