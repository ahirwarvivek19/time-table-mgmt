/**
 * TeacherViewManager.gs
 * Renders the weekly schedule for one teacher on the 'Teacher_View' tab and
 * writes cell edits back to Master_Schedule.
 * Shows FREE vs BUSY slots with class + subject information.
 */

const TeacherViewManager = {

  /**
   * Renders the full weekly grid for a teacher.
   * @param {string} teacherName
   */
  renderTeacherView: function(teacherName) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(CONFIG.SHEETS.TEACHER_VIEW);
    if (!sheet) sheet = ss.insertSheet(CONFIG.SHEETS.TEACHER_VIEW);

    sheet.setFrozenRows(0);    // reset frozen state before any unmerges/clears
    sheet.setFrozenColumns(0);
    sheet.clear();
    sheet.getDataRange().breakApart();
    sheet.setHiddenGridlines(true);

    const days    = CONFIG.DAYS;
    const periods = CONFIG.PERIODS;
    const numCols = periods.length + 1;
    const startRow = 5;

    // ── 1. TEACHER SELECTION ──────────────────────────────────────────────
    const teacherNames = getTeacherNames_();
    const match = teacherNames.find(t => ScheduleParser.sameName(t, teacherName));
    // Unknown / removed teacher → fall back to the first one so the grid and B3 agree
    teacherName = match || teacherNames[0] || '';

    // ── 2. TITLE BANNER (Rows 1–2) ────────────────────────────────────────
    // Col 1 stays unmerged so setFrozenColumns(1) works without conflict.
    sheet.getRange(1, 1, 2, numCols)
         .setBackground('#1A3A4A')
         .setFontFamily(CONFIG.FONT);
    sheet.getRange(1, 2, 2, numCols - 1)
         .merge()
         .setValue('📅  Teacher Schedule  ·  ' + (teacherName || 'Select a Teacher'))
         .setFontFamily(CONFIG.FONT)
         .setFontSize(16)
         .setFontWeight('bold')
         .setFontColor('#FFFFFF')
         .setHorizontalAlignment('center')
         .setVerticalAlignment('middle');
    sheet.setRowHeight(1, 28);
    sheet.setRowHeight(2, 28);

    // ── 3. TEACHER SELECTOR (Row 3) ───────────────────────────────────────
    sheet.getRange('A3')
         .setValue('Select Teacher:')
         .setFontWeight('bold')
         .setFontFamily(CONFIG.FONT)
         .setFontColor('#5D6D7E')
         .setHorizontalAlignment('right');

    const b3Cell = sheet.getRange('B3');
    if (teacherNames.length > 0) {
      b3Cell.setDataValidation(
        SpreadsheetApp.newDataValidation().requireValueInList(getTeacherNamesSorted_(), true).build()
      );
    }
    b3Cell.setValue(teacherName);

    sheet.getRange(3, 1, 1, numCols).setBackground('#F2F3F4').setFontFamily(CONFIG.FONT);
    sheet.setRowHeight(3, 36);

    // Spacer row 4
    sheet.getRange(4, 1, 1, numCols).setBackground('#E8EAED');
    sheet.setRowHeight(4, 6);

    // ── 4. FETCH SCHEDULE DATA ────────────────────────────────────────────
    const scheduleData = DataAccess.getSheetDataAsObjects(CONFIG.SHEETS.SCHEDULE);
    const lookup = {};
    days.forEach(d => {
      lookup[d] = {};
      periods.forEach(p => lookup[d][p] = []);
    });

    scheduleData.forEach(row => {
      const rowDay = normalizeDay_(row.Day); // "monday " still counts as Monday
      if (!rowDay || !row.Period || !lookup[rowDay]) return;
      const periodInt = parseInt(row.Period);
      if (!lookup[rowDay][periodInt]) return;
      if (!ScheduleParser.rowIncludesTeacher(row, teacherName)) return;
      lookup[rowDay][periodInt].push({
        cls: String(row.Class || ''),
        subject: ScheduleParser.getSubjectForTeacher(row, teacherName) || String(row.Subject || '')
      });
    });

    // ── 5. BUILD GRID DATA ────────────────────────────────────────────────
    const headers = ['Day / Period'].concat(periods.map(periodHeader_));
    const gridOutput = [headers];
    const bg = [];
    const fontColors = [];
    const fontWeights = [];

    days.forEach((day, r) => {
      const row = [day];
      const rowBg = ['#2C3E50'], rowFont = ['#ECF0F1'], rowWeight = ['bold'];
      periods.forEach(period => {
        const slots = lookup[day][period];
        if (slots.length === 0) {
          row.push(CONFIG.FREE);
          rowBg.push('#EAFAF1'); rowFont.push('#1E8449'); rowWeight.push('bold');
          return;
        }
        // Same subject in several sections → one combined line; different subjects → clash (red)
        const grouped = ScheduleParser.groupTeacherSlots(slots);
        row.push(grouped.isClash
          ? slots.map(s => s.cls + '\n' + s.subject).join('\n---\n')
          : (grouped.isCombined
              ? [...new Set(slots.flatMap(s => ScheduleParser.splitClasses(s.cls)))].join(', ') + '\n' + slots[0].subject
              : slots[0].cls + '\n' + slots[0].subject));
        if (grouped.isClash) {
          rowBg.push('#FDEDEC'); rowFont.push('#C0392B'); rowWeight.push('bold');
        } else {
          rowBg.push(r % 2 === 1 ? '#EBF5FB' : '#FDFEFE'); rowFont.push('#1A5276'); rowWeight.push('normal');
        }
      });
      gridOutput.push(row);
      bg.push(rowBg); fontColors.push(rowFont); fontWeights.push(rowWeight);
    });

    // ── 6. WRITE DATA ─────────────────────────────────────────────────────
    const numRows  = gridOutput.length;
    const dataRange = sheet.getRange(startRow, 1, numRows, numCols);
    dataRange.setValues(gridOutput);

    // ── 7. PREMIUM STYLING (batched) ──────────────────────────────────────
    dataRange.setFontFamily(CONFIG.FONT)
             .setHorizontalAlignment('center')
             .setVerticalAlignment('middle')
             .setWrap(true)
             .setBorder(true, true, true, true, true, true,
                        '#CCD1D9', SpreadsheetApp.BorderStyle.SOLID);

    sheet.getRange(startRow, 1, 1, numCols)
         .setBackground('#1A3A4A')
         .setFontColor('#FFFFFF')
         .setFontWeight('bold')
         .setFontSize(11);
    sheet.setRowHeight(startRow, 40);

    const body = sheet.getRange(startRow + 1, 1, numRows - 1, numCols);
    body.setBackgrounds(bg);
    body.setFontColors(fontColors);
    body.setFontWeights(fontWeights);
    sheet.getRange(startRow + 1, 1, numRows - 1, 1).setFontSize(10);
    sheet.setRowHeights(startRow + 1, numRows - 1, 70);
    for (let r = startRow + 1; r < startRow + numRows; r++) {
      sheet.getRange(r, 1, 1, numCols)
           .setBorder(null, null, true, null, null, null,
                      '#7F8C8D', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    }

    sheet.setColumnWidth(1, 120);
    for (let i = 2; i <= numCols; i++) sheet.setColumnWidth(i, 160);

    sheet.setFrozenRows(startRow);
    sheet.setFrozenColumns(1);
  },

  /**
   * Updates Master_Schedule from an inline cell edit made on Teacher_View.
   *
   * Cell value formats accepted:
   *   "ClassName - Subject"   e.g. "6th A - Mathematics"
   *   "ClassName\nSubject"    e.g. "6th A\nMathematics"  (the rendered format)
   *   "ClassName"             keep / share the slot's existing subject
   *   "FREE" or ""            → remove this teacher from the slot
   *
   * Multiple classes are separated by "," (e.g. "11th A, 11th B - IP").
   * The teacher is removed from any other class in the same slot, so moving a
   * period never leaves the teacher booked twice.
   *
   * @param {string} teacherName  Teacher shown in B3
   * @param {string} day          e.g. 'Monday'
   * @param {number} period       1-indexed
   * @param {string} newValue     Raw cell value typed by user
   * @returns {string} '' on success, otherwise a message explaining why nothing was saved
   */
  updateMasterFromTeacherView: function(teacherName, day, period, newValue, changes, notices) {
    return updateTeacherSlot_(teacherName, day, period, newValue, changes, notices);
  }
};

/**
 * Shared write-back for Teacher_View and Teacher_Day_View.
 * @param {Array} [changes]  receives {cls, before, after} per changed row (for the change log)
 * @param {Array} [notices]  receives messages about other teachers this edit affected
 * @returns {string} '' on success, otherwise a reason the edit was rejected
 */
function updateTeacherSlot_(teacherName, day, period, newValue, changes, notices) {
  if (!teacherName || !day || !period) return '';
  changes = changes || [];
  notices = notices || [];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const msSheet = ss.getSheetByName(CONFIG.SHEETS.SCHEDULE);
  if (!msSheet) return '';

  const C = CONFIG.COL;
  const values = msSheet.getDataRange().getValues();
  const cleanValue = String(newValue == null ? '' : newValue).trim();
  // Day typed as "monday " in Master_Schedule still matches
  const inSlot = i => normalizeDay_(values[i][C.DAY - 1]) === day && values[i][C.PERIOD - 1] == period;

  const removeFromRow = i => {
    const res = ScheduleParser.removeTeacher(values[i][C.TEACHER - 1], values[i][C.SUBJECT - 1], teacherName, values[i][C.ROOM - 1]);
    if (!res.removed) return;
    writeLaneCells_(msSheet, values, i, res, changes);
    if (String(res.subject).trim() && !String(res.teacher).trim()) {
      notices.push(day + ' P' + period + ': ' + values[i][C.CLASS - 1] + ' ' + res.subject + ' now has no teacher.');
    }
  };

  // ── FREE / empty: remove teacher from this slot ─────────────────────────
  if (!cleanValue || cleanValue.toUpperCase() === CONFIG.FREE) {
    for (let i = 1; i < values.length; i++) if (inSlot(i)) removeFromRow(i);
    return '';
  }

  // A clash cell lists several class/subject blocks ("6th A\nMaths\n---\n7th A\nScience");
  // re-saving it cannot be interpreted safely.
  if (cleanValue.indexOf('---') !== -1) {
    return day + ' P' + period + ': this period has a clash. Fix it in Master_Schedule, or type one "Class - Subject".';
  }

  // ── Parse "ClassName - Subject" or "ClassName\nSubject" ─────────────────
  let targetClass = cleanValue;
  let targetSubject = '';
  if (cleanValue.includes('\n')) {
    const parts = cleanValue.split('\n').map(s => s.trim()).filter(s => s && s !== '---');
    targetClass = parts[0] || '';
    targetSubject = parts.slice(1).join(' ').trim();
  } else if (cleanValue.includes(' - ')) {
    const parts = cleanValue.split(' - ');
    targetClass = parts[0].trim();
    targetSubject = parts.slice(1).join(' - ').trim();
  }

  // Validate the class names so a typo does not create a phantom class row
  const classes = DataAccess.getSheetDataAsObjects(CONFIG.SHEETS.CLASSES);
  const targetClasses = ScheduleParser.splitClasses(targetClass);
  const resolved = targetClasses.map(tc => classes.find(c => ScheduleParser.sameName(c['Class Name'], tc)));
  if (targetClasses.length === 0 || resolved.some(c => !c)) {
    const bad = targetClasses.filter((tc, i) => !resolved[i]);
    return day + ' P' + period + ': unknown class "' + (bad.join(', ') || targetClass) + '". Use "Class - Subject", e.g. "6th A - Mathematics".';
  }

  // A row shared by several classes ("11th A, 11th B") can only be edited as a whole here
  for (let i = 1; i < values.length; i++) {
    if (!inSlot(i)) continue;
    const rowClasses = ScheduleParser.splitClasses(values[i][C.CLASS - 1]);
    const covered = rowClasses.filter(rc => targetClasses.some(tc => ScheduleParser.sameName(rc, tc)));
    if (rowClasses.length > 1 && covered.length > 0 && covered.length < rowClasses.length) {
      return day + ' P' + period + ': "' + values[i][C.CLASS - 1] + '" is one combined row — type all of its classes, or edit it in Master_Schedule.';
    }
  }

  // Check every target row can take the teacher BEFORE touching anything, so a
  // rejected edit never leaves the teacher removed from their old class.
  for (let i = 1; i < values.length; i++) {
    if (!inSlot(i) || !ScheduleParser.rowIncludesClass({ Class: values[i][C.CLASS - 1] }, targetClass)) continue;
    const trial = ScheduleParser.assignTeacher(values[i][C.TEACHER - 1], values[i][C.SUBJECT - 1], teacherName, targetSubject, values[i][C.ROOM - 1]);
    if (trial.error) return day + ' P' + period + ': ' + values[i][C.CLASS - 1] + ': ' + trial.error + '.';
  }

  // Remove the teacher from every other class at this slot (a move, not a copy)
  for (let i = 1; i < values.length; i++) {
    if (inSlot(i) && !ScheduleParser.rowIncludesClass({ Class: values[i][C.CLASS - 1] }, targetClass)) removeFromRow(i);
  }

  // Assign to each target class (one Master_Schedule row per class)
  const errors = [];
  resolved.forEach(classData => {
    const className = String(classData['Class Name']);
    let rowIndex = -1;
    for (let i = 1; i < values.length; i++) {
      if (inSlot(i) && ScheduleParser.rowIncludesClass({ Class: values[i][C.CLASS - 1] }, className)) { rowIndex = i; break; }
    }
    if (rowIndex > 0) {
      const res = ScheduleParser.assignTeacher(values[rowIndex][C.TEACHER - 1], values[rowIndex][C.SUBJECT - 1], teacherName, targetSubject, values[rowIndex][C.ROOM - 1]);
      if (res.error) { errors.push(className + ': ' + res.error); return; }
      writeLaneCells_(msSheet, values, rowIndex, res, changes);
      (res.replaced || []).forEach(r => notices.push(day + ' P' + period + ': ' + className + ' ' + (r.subject || '') + ' — ' + r.teacher + ' replaced by ' + teacherName + '.'));
    } else {
      const row = [day, period, className, classData['Academic Tier'] || '', targetSubject, teacherName, classData['Room Assigned'] || '', ''];
      msSheet.appendRow(row);
      values.push(row);
      changes.push({ cls: className, before: '', after: describeLesson_(targetSubject, teacherName) });
    }
    if (targetSubject) addMissingSubjects_([targetSubject]);
  });
  return errors.length ? day + ' P' + period + ': ' + errors.join('; ') : '';
}

/** Writes Subject, Teacher and Room (columns E:G) of Master_Schedule row index i (0-based in values). */
function writeLaneCells_(msSheet, values, i, res, changes) {
  const C = CONFIG.COL;
  if (changes) changes.push({ cls: String(values[i][C.CLASS - 1]), before: describeRow_(values[i]), after: describeLesson_(res.subject, res.teacher) });
  msSheet.getRange(i + 1, C.SUBJECT, 1, 3).setValues([[res.subject, res.teacher, res.room]]);
  values[i][C.SUBJECT - 1] = res.subject;
  values[i][C.TEACHER - 1] = res.teacher;
  values[i][C.ROOM - 1] = res.room;
}
