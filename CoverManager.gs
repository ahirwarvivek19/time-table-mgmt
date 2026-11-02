/**
 * CoverManager.gs
 * Backend API for the Daily Arrangement / Substitution system (CoverManagerUI.html).
 * All api* functions return plain strings/numbers/objects so google.script.run
 * can serialise them.
 */

function openCoverManagerUI() {
  const html = HtmlService.createTemplateFromFile('CoverManagerUI')
    .evaluate()
    .setWidth(760)
    .setHeight(640)
    .setTitle('Daily Arrangement Manager');
  SpreadsheetApp.getUi().showModalDialog(html, 'Daily Arrangement Manager');
}

/** Returns all teachers A→Z (de-duplicated) for the Cover Manager pickers. */
function apiGetTeachers() {
  return getTeacherNamesSorted_();
}

/** "2026-10-12" → "Monday" ('' for Sunday). Used by the dialog to pick the timetable day. */
function apiGetDayForDate(isoDate) {
  return dayNameForDate_(isoDate);
}

/**
 * Returns the absent teacher's periods for a day, sorted by period.
 * @returns {Array<{Period: number, Class: string, Subject: string}>}
 */
/**
 * True when the absent teacher's lessons in this period are one combined lecture (same
 * subject, same grade), so one cover teacher can take them all. Two different lessons at
 * once are a clash: each needs its own cover.
 */
function absentLessonIsCombined_(absentTeacher, day, period) {
  const lessons = apiGetTeacherSchedule(absentTeacher, day).filter(r => r.Period === period)
    .map(r => ({ cls: r.Class, subject: r.Subject }));
  return lessons.length < 2 || !ScheduleParser.classifyTeacherSlot(lessons).clash;
}

function apiGetTeacherSchedule(teacherName, day) {
  day = normalizeDay_(day);
  return DataAccess.getSheetDataAsObjects(CONFIG.SHEETS.SCHEDULE)
    .filter(row => normalizeDay_(row.Day) === day && ScheduleParser.rowIncludesTeacher(row, teacherName))
    .map(row => ({
      Period: parseInt(row.Period),
      Class: String(row.Class || ''),
      Subject: ScheduleParser.getSubjectForTeacher(row, teacherName) || String(row.Subject || '')
    }))
    .sort((a, b) => a.Period - b.Period);
}

/**
 * Teachers who can take a cover for this day/period: not teaching then, not
 * marked unavailable that day, not the absent teacher, and not already given a
 * cover for the same date + period. Sorted by fewest covers already that date.
 *
 * @param {string} day
 * @param {number|string} period
 * @param {string} [date]           "YYYY-MM-DD" — enables the already-covering check
 * @param {string} [absentTeacher]
 * @returns {string[]}
 */
function apiGetAvailableTeachers(day, period, date, absentTeacher) {
  day = normalizeDay_(day);
  const p = parseInt(period);
  const norm = s => ScheduleParser.normalizeName(s);

  const busy = {};
  DataAccess.getSheetDataAsObjects(CONFIG.SHEETS.SCHEDULE).forEach(row => {
    if (normalizeDay_(row.Day) === day && parseInt(row.Period) === p) {
      ScheduleParser.parseRowAssignments(row).forEach(a => { if (a.teacher) busy[norm(a.teacher)] = true; });
    }
  });

  const offToday = {};
  DataAccess.getSheetDataAsObjects(CONFIG.SHEETS.TEACHERS).forEach(t => {
    const days = parseDaysList_(t['Days Unavailable']).days;
    if (days.indexOf(day) !== -1) offToday[norm(t['Teacher Name'])] = true;
  });

  const coversToday = {};
  const coveringNow = {};
  const coveringSameLesson = {}; // already covering this absent teacher's combined lecture now
  const combined = absentTeacher ? absentLessonIsCombined_(absentTeacher, day, p) : false;
  if (date) {
    getCoverRows_().forEach(c => {
      if (c.date !== date) return;
      const key = norm(c.coverTeacher);
      coversToday[key] = (coversToday[key] || 0) + 1;
      if (parseInt(c.period) === p) {
        // A combined lecture of the absent teacher (several sections at once) is one class
        // to cover, so the teacher covering one section may take the others too
        if (combined && ScheduleParser.sameName(c.absentTeacher, absentTeacher)) coveringSameLesson[key] = true;
        else coveringNow[key] = true;
      }
      // The absent teacher of another cover is absent all day
      offToday[norm(c.absentTeacher)] = true;
    });
  }
  if (absentTeacher) offToday[norm(absentTeacher)] = true;

  return apiGetTeachers()
    .filter(t => { const k = norm(t); return !busy[k] && !offToday[k] && !coveringNow[k]; })
    .sort((a, b) => (coveringSameLesson[norm(b)] ? 1 : 0) - (coveringSameLesson[norm(a)] ? 1 : 0) ||
      (coversToday[norm(a)] || 0) - (coversToday[norm(b)] || 0) || a.localeCompare(b));
}

/**
 * Cover assignments already recorded for a date.
 * @returns {Array<{period: number, classToCover: string, absentTeacher: string, coverTeacher: string}>}
 */
function apiGetCoversForDate(date) {
  return getCoverRows_().filter(c => c.date === date).map(c => ({
    period: parseInt(c.period), classToCover: c.classToCover, absentTeacher: c.absentTeacher, coverTeacher: c.coverTeacher
  }));
}

/**
 * Appends the cover assignment to the Cover_Manager sheet.
 * Refuses a second cover for the same date/period/class/absent teacher, a lesson the
 * absent teacher does not teach, or a cover teacher already covering another lesson at
 * that time. Two absent teachers of one split period each get their own cover; the
 * sections of one absent teacher's combined lecture may share one cover teacher.
 */
function apiAssignCover(date, absentTeacher, period, classToCover, coverTeacher) {
  if (!date || !absentTeacher || !period || !classToCover || !coverTeacher) {
    throw new Error('Date, absent teacher, period, class and cover teacher are all required.');
  }
  // Server-side checks: the dialog's choices are not trusted
  const day = dayNameForDate_(date);
  if (!day) throw new Error(date + ' is not a school day (Monday–Saturday).');
  const p = parseInt(period);
  if (CONFIG.PERIODS.indexOf(p) === -1) throw new Error('Invalid period: ' + period);
  const teachers = getTeacherNames_();
  absentTeacher = teachers.find(t => ScheduleParser.sameName(t, absentTeacher));
  coverTeacher = teachers.find(t => ScheduleParser.sameName(t, coverTeacher)) || coverTeacher;
  if (!absentTeacher) throw new Error('Unknown absent teacher.');
  const cls = getClassNames_().find(c => ScheduleParser.sameName(c, classToCover)) ||
    (ScheduleParser.splitClasses(classToCover).every(c => getClassNames_().some(k => ScheduleParser.sameName(k, c))) ? classToCover : '');
  if (!cls) throw new Error('Unknown class: ' + classToCover);
  classToCover = cls;
  const teaches = apiGetTeacherSchedule(absentTeacher, day).some(r => r.Period === p &&
    (ScheduleParser.sameName(r.Class, classToCover) || ScheduleParser.rowIncludesClass({ Class: r.Class }, classToCover)));
  if (!teaches) throw new Error(absentTeacher + ' does not teach ' + classToCover + ' in period ' + p + ' on ' + day + '.');

  return withDocumentLock_(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const coverSheet = ss.getSheetByName(CONFIG.SHEETS.COVER);
    if (!coverSheet) throw new Error('Cover_Manager sheet not found. Run Setup first.');

    const same = (a, b) => ScheduleParser.sameName(a, b);
    const combined = absentLessonIsCombined_(absentTeacher, day, p);
    getCoverRows_().forEach(c => {
      if (c.date !== date || parseInt(c.period) !== p) return;
      const overlap = same(c.classToCover, classToCover) || ScheduleParser.rowIncludesClass({ Class: c.classToCover }, classToCover) ||
        ScheduleParser.rowIncludesClass({ Class: classToCover }, c.classToCover);
      if (overlap && same(c.absentTeacher, absentTeacher)) {
        throw new Error(classToCover + ' period ' + p + ' is already covered by ' + c.coverTeacher + ' on ' + date + '.');
      }
      if (same(c.coverTeacher, coverTeacher) && !(combined && same(c.absentTeacher, absentTeacher))) {
        throw new Error(coverTeacher + ' is already covering ' + c.classToCover + ' in period ' + p + ' on ' + date + '.');
      }
    });

    const suggestions = apiGetAvailableTeachers(day, p, date, absentTeacher);
    if (!suggestions.some(t => ScheduleParser.sameName(t, coverTeacher))) {
      throw new Error(coverTeacher + ' is not free in period ' + p + ' on ' + date + ' (teaching, unavailable, absent or already covering).');
    }
    coverSheet.appendRow([
      date,
      absentTeacher,
      p,
      classToCover,
      suggestions.slice(0, 10).join(', '), // what the "Suggested Available Teachers" column says
      coverTeacher
    ]);

    const lastRow = coverSheet.getLastRow();
    const rowRange = coverSheet.getRange(lastRow, 1, 1, 6);
    rowRange.setFontFamily(CONFIG.FONT);
    rowRange.setBackground(tableRowColor_(lastRow));
    rowRange.setBorder(true, true, true, true, true, true, '#D5D8DC', SpreadsheetApp.BorderStyle.SOLID);
    return true;
  });
}

/** Reads Cover_Manager rows; dates are normalised to "YYYY-MM-DD". */
function getCoverRows_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.COVER);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone
    ? SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() : 'Asia/Kolkata';
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues()
    .filter(r => r[0] !== '' && r[1] !== '')
    .map(r => ({
      date: Object.prototype.toString.call(r[0]) === '[object Date]' ? Utilities.formatDate(r[0], tz, 'yyyy-MM-dd') : String(r[0]).trim(),
      absentTeacher: String(r[1]),
      period: r[2],
      classToCover: String(r[3]),
      coverTeacher: String(r[5])
    }));
}
