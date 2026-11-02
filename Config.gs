/**
 * Config.gs
 * Single place for sheet names, days, periods, column positions and shared helpers.
 *
 * Only reference CONFIG inside functions (never at file top level): Apps Script
 * loads files in an unspecified order, so a top-level reference from another
 * file could run before this file is loaded.
 */

const CONFIG = {
  SHEETS: {
    TEACHERS: 'Teachers',
    SUBJECTS: 'Subjects',
    CLASSES: 'Classes',
    SCHEDULE: 'Master_Schedule',
    CLASS_VIEW: 'Class_View',
    TEACHER_VIEW: 'Teacher_View',
    TEACHER_DAY_VIEW: 'Teacher_Day_View',
    MASTER_GRID: 'Master_Grid_View',
    // Retired: the old free/busy grid (Teacher_Day_View shows the same, plus DAY OFF).
    // Setup / Repair renames and hides the tab if an older version created it.
    LEGACY_AVAILABILITY: 'Teacher_Availability_Grid',
    LEGACY_AVAILABILITY_RETIRED: 'Teacher_Availability_Grid (old)',
    COVER: 'Cover_Manager',
    CHANGE_LOG: 'Change_Log',
    PREFILL_BACKUP: '_Prefill_Data_Backup'
  },

  // One row per period changed from a view or Master_Schedule (written by logChanges_).
  CHANGE_LOG_HEADERS: ['Time', 'Edited By', 'Edited In', 'Day', 'Period', 'Class', 'Before', 'After'],

  DAYS: ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],
  PERIODS: [1, 2, 3, 4, 5, 6, 7, 8],

  // Bell schedule shown under each period header. The headers read RAW_PERIOD_ROWS in
  // PrefillData.gs (taken from the Excel header rows by `npm run convert`); this table
  // is only the fallback when that file is missing. A test keeps the two in step.
  PERIOD_TIMES: {
    1: '10:00–10:50', 2: '10:50–11:30', 3: '11:30–12:10', 4: '12:10–12:50',
    5: '1:20–2:00', 6: '2:00–2:40', 7: '2:40–3:20', 8: '3:20–4:00'
  },
  // Breaks shown as their own column in every view; `after` = the period they follow
  // (0 = before period 1). Read from RAW_BREAK_ROWS in PrefillData.gs; this is the fallback.
  BREAKS: [
    { name: 'Prayer', time: '9:10–10:00', after: 0 },
    { name: 'Lunch', time: '12:50–1:20', after: 4 }
  ],

  // Master_Schedule columns (1-based)
  COL: { DAY: 1, PERIOD: 2, CLASS: 3, TIER: 4, SUBJECT: 5, TEACHER: 6, ROOM: 7, CLASH: 8 },
  SCHEDULE_HEADERS: ['Day', 'Period', 'Class', 'Academic Tier', 'Subject', 'Teacher', 'Room', 'Clash Status'],

  // A teacher teaching the same subject to several sections at once is a combined
  // lecture, not a clash — but only within one grade (9th C + 9th D). These whole-
  // school activities may also combine sections of different grades. Matched as a
  // case-insensitive prefix of the subject.
  CROSS_GRADE_SUBJECTS: ['Sports', 'Yoga', 'Library', 'Music', 'ICT Lab', 'SUPW', 'Bal Sabha', 'UHV', 'Life Skills', 'Prayer', 'Assembly'],

  ALL_CLASSES: 'All Classes',
  FREE: 'FREE',
  // Shown in Teacher_Day_View for a free period on a day listed in Teachers → Days Unavailable
  DAY_OFF: 'DAY OFF',
  // Room values that mean "no room recorded" — never treated as a shared room.
  ROOM_PLACEHOLDERS: ['', 'n/a', 'na', '-', '—', 'none'],

  // Edits that touch more cells than this (e.g. a huge paste) are not synced cell by cell.
  MAX_SYNCED_CELLS: 400,
  // Unique id of the installable edit trigger recorded by installEditTrigger_().
  PROP_EDIT_TRIGGER: 'EDIT_TRIGGER_UID',
  // Unique id of the installable change trigger (rows deleted / inserted in Master_Schedule).
  PROP_CHANGE_TRIGGER: 'CHANGE_TRIGGER_UID',
  // Time (ms) the installable edit trigger last started; the simple trigger checks it to warn
  // an editor when sync has stopped (trigger deleted, owner lost access, daily quota used up).
  PROP_LAST_SYNC: 'LAST_SYNC_AT',
  SYNC_CHECK_WAIT_MS: 10000,
  // A sync mark newer than this counts as "sync is running" (wide on purpose: trigger start
  // times vary, and a false "not saved" warning makes people retype edits that went through)
  SYNC_ALIVE_WINDOW_MS: 60000,
  // Max wait for another edit to finish (stays well under the 30 s simple-trigger limit).
  LOCK_WAIT_MS: 10000,

  // Data tabs are styled this many rows past the last filled row, so new rows look the same.
  TABLE_SPARE_ROWS: 100,

  FONT: 'Montserrat'
};

/**
 * "9:30–10:20" for a period, from the bell schedule in PrefillData.gs (RAW_PERIOD_ROWS,
 * 24-hour "13:30" values shown as "1:30"), else CONFIG.PERIOD_TIMES.
 */
function periodTime_(period) {
  const rows = typeof RAW_PERIOD_ROWS !== 'undefined' ? RAW_PERIOD_ROWS : [];
  const row = rows.find(r => Number(r[0]) === Number(period));
  if (!row) return CONFIG.PERIOD_TIMES[period] || '';
  return clockTime_(row[1]) + '–' + clockTime_(row[2]);
}

/** "13:30" → "1:30" (12-hour clock without am/pm, as the school writes it). */
function clockTime_(t) {
  const m = String(t).match(/^(\d{1,2}):(\d{2})$/);
  return m ? (Number(m[1]) % 12 || 12) + ':' + m[2] : String(t);
}

/** Prayer / Lunch breaks: [{name, time, after}] from RAW_BREAK_ROWS, else CONFIG.BREAKS. */
function viewBreaks_() {
  const rows = typeof RAW_BREAK_ROWS !== 'undefined' ? RAW_BREAK_ROWS : null;
  if (!rows) return CONFIG.BREAKS;
  return rows.map(r => ({ name: String(r[0]), time: clockTime_(r[1]) + '–' + clockTime_(r[2]), after: Number(r[3]) }));
}

/**
 * The columns of every view after its first (label) column, in order: a period
 * ({period: 3}) or a break ({period: null, name: 'Lunch', time: '12:50–1:20'}).
 * With the default bell schedule: Prayer, P1–P4, Lunch, P5–P8 (sheet columns B–K).
 */
function viewSlots_() {
  const breaks = viewBreaks_();
  const out = [];
  const addBreaks = after => breaks.filter(b => b.after === after)
    .forEach(b => out.push({ period: null, name: b.name, time: b.time }));
  addBreaks(0);
  CONFIG.PERIODS.forEach(p => { out.push({ period: p }); addBreaks(p); });
  return out;
}

/** Header text of a view column: "Period 3\n11:30–12:10" or "Lunch\n12:50–1:20". */
function viewSlotHeader_(slot) {
  return slot.period ? periodHeader_(slot.period) : slot.name + (slot.time ? '\n' + slot.time : '');
}

/** Number of columns of a view grid, including its label column. */
function viewWidth_() {
  return viewSlots_().length + 1;
}

/** The period shown in a view's sheet column (1-based), or null (label or break column). */
function periodOfViewColumn_(col) {
  const slot = viewSlots_()[col - 2];
  return slot && slot.period ? slot.period : null;
}

/** The view sheet column (1-based) that shows a period. */
function viewColumnOfPeriod_(period) {
  return viewSlots_().findIndex(s => s.period === Number(period)) + 2;
}

/**
 * Styles the break columns of a rendered view: one grey cell down the whole body with the
 * break's name, no dropdowns or notes. The grid's values must already be written.
 * @param {number} headerRow row of the period headers
 * @param {number} numRows   header row + body rows
 */
function styleBreakColumns_(sheet, headerRow, numRows) {
  viewSlots_().forEach((slot, i) => {
    if (slot.period) return;
    const col = i + 2;
    sheet.getRange(headerRow, col).setBackground('#566573');
    sheet.setColumnWidth(col, 70);
    if (numRows < 2) return;
    const body = sheet.getRange(headerRow + 1, col, numRows - 1, 1);
    body.clearDataValidations()
        .setNotes(Array.from({ length: numRows - 1 }, () => ['']))
        .setBackground('#EAECEE')
        .setFontColor('#566573')
        .setFontWeight('bold')
        .setFontSize(11);
    if (numRows > 2) body.merge();
    body.setVerticalText(true);
  });
}

/**
 * Splits edited view cells into period cells (each given `period`) and break-column cells.
 * @returns {{cells: Object[], breakHit: boolean}}
 */
function periodCells_(cells) {
  const out = [];
  let breakHit = false;
  cells.forEach(c => {
    const p = periodOfViewColumn_(c.col);
    if (p) { c.period = p; out.push(c); } else breakHit = true;
  });
  return { cells: out, breakHit: breakHit };
}

/** Told to the editor when they type in a Prayer / Lunch column (the view is redrawn). */
function breakColumnNotice_() {
  return viewBreaks_().map(b => b.name).join(' and ') + ' columns are not part of the timetable — nothing was saved there.';
}

/** "Period 3\n11:00–11:40" */
function periodHeader_(period) {
  const time = periodTime_(period);
  return 'Period ' + period + (time ? '\n' + time : '');
}

/** Returns the day name ("Monday"…) for a "YYYY-MM-DD" string, or '' for Sunday / invalid. */
function dayNameForDate_(isoDate) {
  const m = String(isoDate || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return '';
  const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  if (d.getMonth() !== Number(m[2]) - 1 || d.getDate() !== Number(m[3])) return ''; // e.g. 2026-02-31
  const idx = d.getDay(); // 0 = Sunday
  return idx === 0 ? '' : CONFIG.DAYS[idx - 1];
}

/** Matches "monday", " Mon ", "MONDAY" → "Monday"; returns '' when not a school day. */
function normalizeDay_(value) {
  const v = String(value || '').trim().toLowerCase();
  if (!v) return '';
  return CONFIG.DAYS.find(d => d.toLowerCase() === v || (v.length >= 3 && d.toLowerCase().indexOf(v) === 0)) || '';
}

/**
 * Reads a "Days Unavailable" cell the way people write it: "Monday, Friday", "Mon/Fri",
 * "Monday & Friday", "Monday and Friday", "Mondays", "Mon-Wed", "Monday to Wednesday".
 * @returns {{days: string[], unreadable: string[]}} canonical day names, plus the parts
 *   that are not a school day (reported by Check For Conflicts instead of being ignored)
 */
function parseDaysList_(text) {
  const out = { days: [], unreadable: [] };
  const raw = String(text == null ? '' : text).trim();
  // "-", "N/A", "None", "nil" mean no days off
  if (!raw || CONFIG.ROOM_PLACEHOLDERS.indexOf(raw.toLowerCase()) !== -1 || /^(none|nil|no)$/i.test(raw)) return out;
  const days = [];
  const add = d => { if (days.indexOf(d) === -1) days.push(d); };
  const dayOf = w => normalizeDay_(String(w).replace(/s$/i, '')) || normalizeDay_(w);
  const bad = [];
  raw
    .replace(/\s+(?:to|till|until)\s+|\s*[–—-]\s*/gi, '-')
    .split(/\s*(?:[,\/;&+]|\band\b)\s*|\s+/i)
    .map(s => s.trim()).filter(Boolean)
    .forEach(part => {
      const range = part.split('-');
      if (range.length === 2 && dayOf(range[0]) && dayOf(range[1])) {
        // "Thursday-Monday" wraps round the week: Thu, Fri, Sat, Mon
        const n = CONFIG.DAYS.length;
        const a = CONFIG.DAYS.indexOf(dayOf(range[0])), b = CONFIG.DAYS.indexOf(dayOf(range[1]));
        for (let i = a; ; i = (i + 1) % n) { add(CONFIG.DAYS[i]); if (i === b) break; }
      } else if (dayOf(part)) {
        add(dayOf(part));
      } else {
        bad.push(part);
      }
    });
  // A note such as "Monday (morning)", "Except Saturday" or "2nd & 4th Saturday" is not a
  // list of whole days off: use none of it and report the whole value instead of guessing
  if (bad.length) { out.unreadable.push(raw); return out; }
  out.days = days;
  return out;
}

/** True when a room value is a real room (not blank / "N/A"). */
function isRealRoom_(room) {
  return CONFIG.ROOM_PLACEHOLDERS.indexOf(String(room == null ? '' : room).trim().toLowerCase()) === -1;
}

/** Class names in the order of the Classes sheet (keeps 6th … 12th in school order). */
function getClassNames_() {
  return DataAccess.getSheetDataAsObjects(CONFIG.SHEETS.CLASSES)
    .map(c => String(c['Class Name'] || '').trim())
    .filter(Boolean);
}

/** Teacher names in the order of the Teachers sheet, de-duplicated. */
function getTeacherNames_() {
  const seen = {};
  return DataAccess.getSheetDataAsObjects(CONFIG.SHEETS.TEACHERS)
    .map(t => String(t['Teacher Name'] || '').trim())
    .filter(t => {
      const k = t.toLowerCase();
      if (!t || seen[k]) return false;
      seen[k] = true;
      return true;
    });
}

/**
 * Days each teacher is off, from Teachers → "Days Unavailable" ("Monday, Friday").
 * @returns {Object<string, string[]>} normalized teacher name → canonical day names
 */
function teacherDaysOff_() {
  const off = {};
  DataAccess.getSheetDataAsObjects(CONFIG.SHEETS.TEACHERS).forEach(t => {
    const key = ScheduleParser.normalizeName(t['Teacher Name']);
    if (!key) return;
    off[key] = parseDaysList_(t['Days Unavailable']).days;
  });
  return off;
}

/** Teacher names A→Z (case-insensitive) for dropdowns and pickers, so a name is easy to find. */
function getTeacherNamesSorted_() {
  return getTeacherNames_().sort((a, b) => a.localeCompare(b, 'en', { sensitivity: 'base' }));
}

/**
 * Runs fn while holding the document lock so two people editing at once
 * cannot interleave read-modify-write cycles on Master_Schedule.
 */
function withDocumentLock_(fn) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(CONFIG.LOCK_WAIT_MS)) {
    throw new Error('Not saved — another timetable edit is still being processed. What you typed is still on screen but was NOT saved: re-enter it in a few seconds.');
  }
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}
