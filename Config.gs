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
    AVAILABILITY: 'Teacher_Availability_Grid',
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
    1: '9:30–10:20', 2: '10:20–11:00', 3: '11:00–11:40', 4: '11:40–12:20',
    5: '12:50–1:30', 6: '1:30–2:10', 7: '2:10–2:50', 8: '2:50–3:30'
  },

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
  // Room values that mean "no room recorded" — never treated as a shared room.
  ROOM_PLACEHOLDERS: ['', 'n/a', 'na', '-', '—', 'none'],

  // Edits that touch more cells than this (e.g. a huge paste) are not synced cell by cell.
  MAX_SYNCED_CELLS: 400,
  // Unique id of the installable edit trigger recorded by installEditTrigger_().
  PROP_EDIT_TRIGGER: 'EDIT_TRIGGER_UID',
  // Unique id of the installable change trigger (rows deleted / inserted in Master_Schedule).
  PROP_CHANGE_TRIGGER: 'CHANGE_TRIGGER_UID',
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
  const clock = t => {
    const m = String(t).match(/^(\d{1,2}):(\d{2})$/);
    return m ? (Number(m[1]) % 12 || 12) + ':' + m[2] : String(t);
  };
  return clock(row[1]) + '–' + clock(row[2]);
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
