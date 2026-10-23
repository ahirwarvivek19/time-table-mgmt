/**
 * ClassViewManager.gs
 * Handles the 2D editable Class View grid with premium styling.
 * Supports multi-class / combined class selection (e.g. 11th A, 11th B) and optional subjects.
 */

const ClassViewManager = {

  /**
   * Renders the full 2D timetable grid for a specific class (or combined classes).
   * @param {string} className
   */
  renderClassView: function(className) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(CONFIG.SHEETS.CLASS_VIEW);
    if (!sheet) sheet = ss.insertSheet(CONFIG.SHEETS.CLASS_VIEW);

    sheet.setFrozenRows(0);    // reset frozen state before any unmerges/clears
    sheet.setFrozenColumns(0);
    sheet.clear();
    sheet.getDataRange().breakApart();
    sheet.setHiddenGridlines(true);

    const days    = CONFIG.DAYS;
    const periods = CONFIG.PERIODS;
    const numCols = periods.length + 1; // Day col + period cols
    const startRow = 5; // Grid header at row 5, data from row 6

    // ── 1. CLASS SELECTOR (Row 3) ──────────────────────────────────────────
    // Keep the Classes sheet order (alphabetical sorting put "10th A" before "6th A")
    const classNames  = getClassNames_();

    if (!className && classNames.length > 0) {
      className = classNames[0];
    }

    sheet.getRange('A3')
         .setValue('Select Class:')
         .setFontWeight('bold')
         .setFontFamily('Montserrat')
         .setFontColor('#5D6D7E')
         .setHorizontalAlignment('right');

    const classCell = sheet.getRange('B3');
    if (classNames.length > 0) {
      classCell.setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireValueInList(classNames, true)
          .setAllowInvalid(true) // Allows combined class entries e.g. "11th A, 11th B"
          .build()
      );
    }
    classCell.setValue(className);
    sheet.getRange(3, 1, 1, numCols)
         .setBackground('#F2F3F4')
         .setFontFamily('Montserrat');
    sheet.setRowHeight(3, 36);

    // ── 2. TITLE BANNER (Rows 1–2) ────────────────────────────────────────
    sheet.getRange(1, 1, 2, numCols)
         .setBackground('#1A252F')
         .setFontFamily('Montserrat');
    sheet.getRange(1, 2, 2, numCols - 1)
         .merge()
         .setValue('📚  Class Timetable  ·  ' + className)
         .setFontFamily('Montserrat')
         .setFontSize(16)
         .setFontWeight('bold')
         .setFontColor('#FFFFFF')
         .setHorizontalAlignment('center')
         .setVerticalAlignment('middle');
    sheet.setRowHeight(1, 28);
    sheet.setRowHeight(2, 28);

    // Spacer row 4
    sheet.getRange(4, 1, 1, numCols).setBackground('#E8EAED');
    sheet.setRowHeight(4, 6);

    // ── 3. FETCH SCHEDULE DATA ────────────────────────────────────────────
    const scheduleData = DataAccess.getSheetDataAsObjects(CONFIG.SHEETS.SCHEDULE);
    // Rows of every class, by slot: needed to tell combined lectures and clashes apart
    const bySlot = {};
    scheduleData.forEach(r => {
      const d = normalizeDay_(r.Day);
      const p = parseInt(r.Period);
      if (!d || isNaN(p)) return;
      (bySlot[d + '|' + p] = bySlot[d + '|' + p] || []).push(r);
    });

    const lookup = {};
    days.forEach(d => {
      lookup[d] = {};
      periods.forEach(p => {
        const info = ScheduleParser.describeSlot(bySlot[d + '|' + p] || [], className);
        // Lanes keep subject i ↔ teacher i aligned even when several rows are merged
        const text = ScheduleParser.formatLanes(info.lanes);
        lookup[d][p] = { subject: text.subject, teacher: text.teacher, info: info };
      });
    });

    // ── 4. DROPDOWN VALIDATION RULES ─────────────────────────────────────
    let subjectsList = DataAccess.getSheetDataAsObjects(CONFIG.SHEETS.SUBJECTS)
                                 .map(r => r['Subject Name']).filter(s => s);
    if (subjectsList.length === 0) {
      subjectsList = [...new Set(scheduleData.map(r => r['Subject']).filter(s => s))].sort();
    }
    const subjectRule = subjectsList.length > 0
      ? SpreadsheetApp.newDataValidation()
          .requireValueInList(subjectsList, true)
          .setAllowInvalid(true)
          .build()
      : null;

    const teachersList = getTeacherNamesSorted_(); // A→Z, easier to find a name
    const teacherRule = teachersList.length > 0
      ? SpreadsheetApp.newDataValidation()
          .requireValueInList(teachersList, true)
          .setAllowInvalid(true)
          .build()
      : null;

    // ── 5. BUILD GRID DATA ────────────────────────────────────────────────
    const periodLabels = ['Day / Row'].concat(periods.map(periodHeader_));
    const gridOutput = [periodLabels];

    days.forEach(day => {
      const subRow  = [day];
      const tchrRow = [''];
      periods.forEach(p => {
        subRow.push(lookup[day][p].subject);
        tchrRow.push(lookup[day][p].teacher);
      });
      gridOutput.push(subRow);
      gridOutput.push(tchrRow);
    });

    // ── 6. WRITE DATA ─────────────────────────────────────────────────────
    const numRows  = gridOutput.length;
    const dataRange = sheet.getRange(startRow, 1, numRows, numCols);
    dataRange.setValues(gridOutput);

    // ── 7. PREMIUM STYLING ────────────────────────────────────────────────
    dataRange.setFontFamily('Montserrat')
             .setHorizontalAlignment('center')
             .setVerticalAlignment('middle')
             .setWrap(true)
             .setBorder(true, true, true, true, true, true,
                        '#CCD1D9', SpreadsheetApp.BorderStyle.SOLID);

    // Period header row (startRow)
    sheet.getRange(startRow, 1, 1, numCols)
         .setBackground('#2C3E50')
         .setFontColor('#FFFFFF')
         .setFontWeight('bold')
         .setFontSize(11);
    sheet.setRowHeight(startRow, 40);

    // Day rows: 2 rows each (subject + teacher)
    let currentRow = startRow + 1;
    days.forEach(day => {
      const subjectRowNum = currentRow;
      const teacherRowNum = currentRow + 1;

      // Merge Day label cell vertically across both rows
      sheet.getRange(subjectRowNum, 1, 2, 1)
           .merge()
           .setBackground('#2C3E50')
           .setFontColor('#ECF0F1')
           .setFontWeight('bold')
           .setFontSize(10)
           .setHorizontalAlignment('center')
           .setVerticalAlignment('middle');

      // Subject row cells (cols 2–9)
      const subRange = sheet.getRange(subjectRowNum, 2, 1, periods.length);
      subRange.setBackground('#EBF5FB')   // light blue
              .setFontColor('#1A5276')     // deep blue text
              .setFontWeight('bold');
      if (subjectRule) subRange.setDataValidation(subjectRule);
      else             subRange.clearDataValidations();
      sheet.setRowHeight(subjectRowNum, 36);

      // Teacher row cells (cols 2–9)
      const tchrRange = sheet.getRange(teacherRowNum, 2, 1, periods.length);
      tchrRange.setBackground('#FDFEFE')  // near-white
               .setFontColor('#717D7E')   // muted grey
               .setFontSize(9);
      if (teacherRule) tchrRange.setDataValidation(teacherRule);
      else             tchrRange.clearDataValidations();
      sheet.setRowHeight(teacherRowNum, 28);

      // Thick bottom border between day groups
      sheet.getRange(teacherRowNum, 1, 1, numCols)
           .setBorder(null, null, true, null, null, null,
                      '#7F8C8D', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

      currentRow += 2;
    });

    // ── 8. SPLIT / COMBINED / CLASH MARKERS ───────────────────────────────
    // Colour the subject cell and explain it in a note (hover to read).
    const mark = { split: '#FEF5E7', combined: '#F5EEF8', clash: '#FADBD8' };
    const bgs = [], notes = [];
    days.forEach(day => {
      const bgRow = [], noteRow = [];
      periods.forEach(p => {
        const info = lookup[day][p].info;
        bgRow.push(info.clashes.length ? mark.clash : info.split ? mark.split : info.combinedWith.length ? mark.combined : '#EBF5FB');
        noteRow.push(info.note);
      });
      bgs.push(bgRow);
      notes.push(noteRow);
    });
    days.forEach((day, i) => {
      const range = sheet.getRange(startRow + 1 + i * 2, 2, 1, periods.length);
      range.setBackgrounds([bgs[i]]);
      range.setNotes([notes[i]]);
    });
    sheet.getRange(3, 4, 1, numCols - 3)
         .merge()
         .setValue('🟧 split class   🟪 taught with another section   🟥 teacher clash   — hover a cell for details')
         .setFontSize(9)
         .setFontColor('#5D6D7E')
         .setHorizontalAlignment('left');

    // Column widths
    sheet.setColumnWidth(1, 110);
    for (let i = 2; i <= numCols; i++) {
      sheet.setColumnWidth(i, 150);
    }

    // Freeze header rows and day column
    sheet.setFrozenRows(startRow);
    sheet.setFrozenColumns(1);
  },

  /**
   * Saves a Class_View edit of one period into Master_Schedule, after checking it.
   * The Subject and Teacher cells of the period are applied together, so pasting or
   * clearing both at once is one edit.
   *
   * Rejected (nothing saved, the view is redrawn so the typed text disappears):
   *   - a split period whose subjects and teacher groups do not line up
   *     ("Hindi / IT" with one teacher, or 3 subjects with 2 teacher groups)
   *   - clearing the subject while the teacher stays
   *   - a teacher name that is not on the Teachers tab
   *   - an edit that books a teacher in another class at the same time
   *     (a combined lecture — same subject, same grade — is allowed)
   *   - a class that is not on the Classes tab, or only part of a combined row
   * Clearing both cells removes the period.
   *
   * @param {string} className  the class shown in B3 ("6th A" or "6th A, 6th B")
   * @param {string} day
   * @param {number} period
   * @param {{subject: (string|undefined), teacher: (string|undefined)}} edit  undefined = cell not edited
   * @param {{keepEmptyRow: boolean, onlyRow: number}} [opts]  keepEmptyRow: clearing both cells blanks the row
   *   instead of deleting it; onlyRow: change only this sheet row (Master_Schedule edits)
   * @returns {{error: string, changes: Array<{cls: string, before: string, after: string}>}}
   */
  applyClassViewEdit: function(className, day, period, edit, opts) {
    const fail = msg => ({ error: day + ' P' + period + ': ' + msg + '.', changes: [] });
    if (!className || !day || !period) return { error: '', changes: [] };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const msSheet = ss.getSheetByName(CONFIG.SHEETS.SCHEDULE);
    if (!msSheet) return { error: '', changes: [] };
    const C = CONFIG.COL;
    const values = msSheet.getDataRange().getValues();
    const clean = v => String(v == null ? '' : v).trim();
    const inSlot = i => normalizeDay_(values[i][C.DAY - 1]) === day && values[i][C.PERIOD - 1] == period;

    // ── Classes being edited ──────────────────────────────────────────────
    const classes = DataAccess.getSheetDataAsObjects(CONFIG.SHEETS.CLASSES);
    const targets = ScheduleParser.splitClasses(className);
    const resolved = targets.map(t => classes.find(c => ScheduleParser.sameName(c['Class Name'], t)));
    if (!targets.length || resolved.some(c => !c)) {
      return fail('unknown class "' + targets.filter((t, i) => !resolved[i]).join(', ') + '"');
    }
    const isTarget = c => targets.some(t => ScheduleParser.sameName(t, c));

    // Rows of this slot that belong to the edited classes
    const rows = [];
    for (let i = 1; i < values.length; i++) {
      if (!inSlot(i)) continue;
      const rowClasses = ScheduleParser.splitClasses(values[i][C.CLASS - 1]);
      if (!rowClasses.some(isTarget)) continue;
      if (rowClasses.some(c => !isTarget(c))) {
        return fail('"' + values[i][C.CLASS - 1] + '" is one combined row — select all of its classes, or edit it in Master_Schedule');
      }
      rows.push(i);
    }
    // Master_Schedule edits change the one row being edited, never its duplicates
    if (opts && opts.onlyRow) {
      const only = rows.filter(i => i + 1 === opts.onlyRow);
      rows.length = 0;
      only.forEach(i => rows.push(i));
    }
    // Several classes shown together can only be edited when they hold the same period
    if (targets.length > 1 && rows.length > 1) {
      const sig = i => clean(values[i][C.SUBJECT - 1]) + '|' + clean(values[i][C.TEACHER - 1]);
      if (rows.some(i => sig(i) !== sig(rows[0]))) {
        return fail(targets.join(', ') + ' have different lessons in this period — select one class to edit it');
      }
    }

    // ── New Subject / Teacher text ────────────────────────────────────────
    const base = rows.length ? values[rows[0]] : null;
    // One class with two rows in this period (a hand-added duplicate): the view shows them
    // merged, so the edit starts from the merged text and is saved into a single row.
    const merged = targets.length === 1 && rows.length > 1;
    const baseText = merged
      ? ScheduleParser.formatLanes(ScheduleParser.mergeLanes(rows.map(i => ({ Subject: values[i][C.SUBJECT - 1], Teacher: values[i][C.TEACHER - 1], Room: values[i][C.ROOM - 1] }))))
      : { subject: base ? clean(base[C.SUBJECT - 1]) : '', teacher: base ? clean(base[C.TEACHER - 1]) : '' };
    const subject = edit.subject !== undefined ? clean(edit.subject) : clean(baseText.subject);
    let teacher = edit.teacher !== undefined ? clean(edit.teacher) : clean(baseText.teacher);
    // Picking one name from the dropdown on a split period adds that teacher to the next
    // group without one, so several picks build "Asha / Bina / Chetan" in subject order.
    const groupCount = ScheduleParser.splitList(subject).length;
    const pickedOne = edit.subject === undefined && edit.teacher !== undefined && teacher &&
      !/[\/+,]/.test(teacher) && getTeacherNames_().some(k => ScheduleParser.sameName(k, teacher));
    if (pickedOne && groupCount > 1) {
      const filled = ScheduleParser.fillNextGroup(baseText.teacher, teacher, groupCount);
      if (filled.error) return fail(filled.error);
      teacher = filled.teacher;
    }
    if (!subject && !teacher) {
      // Both cleared: the period is now free
      const changes = [];
      rows.slice().sort((a, b) => b - a).forEach(i => {
        changes.push({ cls: String(values[i][C.CLASS - 1]), before: describeRow_(values[i]), after: '' });
        // In Master_Schedule the row stays (it is the row the editor is standing on)
        if (opts && opts.keepEmptyRow) msSheet.getRange(i + 1, C.SUBJECT, 1, 2).setValues([['', '']]);
        else msSheet.deleteRow(i + 1);
      });
      return { error: '', changes: changes };
    }
    // Clearing only the subject would leave the teacher booked for nothing
    if (!subject && teacher && edit.subject !== undefined) {
      return fail('the subject was cleared but the teacher was not — clear both cells to free the period, or type a subject');
    }
    const laneError = ScheduleParser.checkLaneText(subject, teacher);
    if (laneError) return fail(laneError);

    // Teacher names typed into the subject row (the rows alternate, so this is an easy slip)
    const known = getTeacherNames_();
    if (edit.subject !== undefined && known.length) {
      const typed = ScheduleParser.splitList(subject).filter(s => !ScheduleParser._isEmptyMark(s));
      if (typed.length && typed.every(s => known.some(k => ScheduleParser.sameName(k, s)))) {
        return fail('"' + subject + '" ' + (typed.length === 1 ? 'is a teacher name' : 'are teacher names') + ' — type teachers in the teacher row (the row below the subject)');
      }
    }

    const lanes = ScheduleParser.parseLanes({ Subject: subject, Teacher: teacher }).lanes;
    const newTeachers = [].concat.apply([], lanes.map(l => l.teachers));

    // Teacher names must be on the Teachers tab (a typo would vanish from every teacher view)
    if (known.length) {
      const unknown = newTeachers.filter(t => !known.some(k => ScheduleParser.sameName(k, t)));
      if (unknown.length) return fail('unknown teacher "' + unknown.join('", "') + '" — add them to the Teachers tab first, or pick a name from the list');
    }

    // Clash check: only clashes this edit creates are refused, so an edit that fixes
    // part of an existing clash is never blocked by it.
    const others = [];
    for (let i = 1; i < values.length; i++) {
      if (inSlot(i) && rows.indexOf(i) === -1) others.push(i);
    }
    const entriesFor = (t, ownLanes) => {
      const out = [];
      others.forEach(i => ScheduleParser.parseLanes({ Subject: values[i][C.SUBJECT - 1], Teacher: values[i][C.TEACHER - 1] }).lanes
        .forEach(l => { if (l.teachers.some(x => ScheduleParser.sameName(x, t))) ScheduleParser.splitClasses(values[i][C.CLASS - 1]).forEach(c => out.push({ cls: c, subject: l.subject })); }));
      ownLanes.forEach(l => { if (l.teachers.some(x => ScheduleParser.sameName(x, t))) targets.forEach(c => out.push({ cls: c, subject: l.subject })); });
      return out;
    };
    const oldLanes = rows.length ? ScheduleParser.parseLanes({ Subject: baseText.subject, Teacher: baseText.teacher }).lanes : [];
    for (let k = 0; k < newTeachers.length; k++) {
      const t = newTeachers[k];
      const after = entriesFor(t, lanes);
      if (!ScheduleParser.classifyTeacherSlot(after).clash) continue;
      if (ScheduleParser.classifyTeacherSlot(entriesFor(t, oldLanes)).clash && oldLanes.some(l => l.teachers.some(x => ScheduleParser.sameName(x, t)))) continue;
      const busy = after.filter(e => !isTarget(e.cls)).map(e => e.cls + ' (' + e.subject + ')');
      return fail(t + ' already teaches ' + (busy.join(', ') || 'another group') + ' at this time — remove them there first, or use Teacher_View to move them');
    }

    // ── Write ─────────────────────────────────────────────────────────────
    const changes = [];
    const writeRows = merged ? rows.slice(0, 1) : rows;
    writeRows.forEach(i => {
      const before = merged ? describeLesson_(baseText.subject, baseText.teacher) : describeRow_(values[i]);
      msSheet.getRange(i + 1, C.SUBJECT, 1, 2).setValues([[subject, teacher]]);
      changes.push({ cls: String(values[i][C.CLASS - 1]), before: before, after: describeLesson_(subject, teacher) });
    });
    // The duplicate rows are now part of the first one: remove them (bottom-up keeps indexes valid)
    if (merged) rows.slice(1).sort((a, b) => b - a).forEach(i => msSheet.deleteRow(i + 1));
    // Classes without a row yet get one each (one row per class, like the rest of the data)
    resolved.forEach(c => {
      const name = String(c['Class Name']);
      const has = rows.some(i => ScheduleParser.splitClasses(values[i][C.CLASS - 1]).some(x => ScheduleParser.sameName(x, name)));
      if (has) return;
      msSheet.appendRow([day, period, name, c['Academic Tier'] || '', subject, teacher, c['Room Assigned'] || '', '']);
      changes.push({ cls: name, before: '', after: describeLesson_(subject, teacher) });
    });
    addMissingSubjects_(lanes.map(l => l.subject));
    return { error: '', changes: changes };
  }
};

/** "Maths — A Teacher" for a Master_Schedule row (used in the change log). */
function describeRow_(row) {
  const C = CONFIG.COL;
  return describeLesson_(row[C.SUBJECT - 1], row[C.TEACHER - 1]);
}

/** "Maths — A Teacher", "Maths" (no teacher), "" (free). */
function describeLesson_(subject, teacher) {
  const s = String(subject == null ? '' : subject).trim();
  const t = String(teacher == null ? '' : teacher).trim();
  return t ? (s + ' — ' + t).trim() : s;
}

/** Adds subjects typed in a view to the Subjects tab, so they appear in the dropdowns. */
function addMissingSubjects_(subjects) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.SUBJECTS);
  if (!sheet) return;
  const existing = DataAccess.getSheetDataAsObjects(CONFIG.SHEETS.SUBJECTS).map(r => r['Subject Name']);
  const added = [];
  subjects.forEach(s => {
    s = String(s == null ? '' : s).trim();
    if (!s || s === ScheduleParser.EMPTY_LANE) return;
    if (existing.concat(added).some(e => ScheduleParser.sameName(e, s))) return;
    sheet.appendRow([s]);
    added.push(s);
  });
}
