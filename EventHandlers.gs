/**
 * EventHandlers.gs
 * Centralizes all edit events.
 *
 * Source of truth: Master_Schedule.
 * Every editable view writes back to Master_Schedule, then re-renders the other
 * views (programmatic writes never fire onEdit, so the cascade is explicit).
 *
 * Editable sheets & write-back targets:
 *   Class_View        — B3 dropdown → re-render
 *                       Grid cell   → write-back Subject/Teacher to Master_Schedule
 *   Teacher_View      — B3 dropdown → re-render
 *                       Grid cell   → write-back assignment to Master_Schedule
 *   Teacher_Day_View  — B3/E3 selectors → re-render
 *                       Grid cell        → write-back assignment to Master_Schedule
 *   Master_Schedule   — any cell edit → refresh all views
 *
 * Multi-cell edits (paste, drag-fill, deleting a block) are synced cell by cell.
 *
 * Triggers: the simple onEdit trigger has a 30-second limit, which a full
 * cascade can exceed on a large timetable. "⚙️ Setup / Repair" adds an
 * installable trigger and records its unique id. From then on
 * only that trigger processes edits: the simple trigger and any other installable
 * onEdit trigger (hand-made, or created by another editor) return immediately,
 * so each edit is processed exactly once.
 *
 * The views are shared sheets with one selector (B3) for everybody; two people
 * using the same view at the same time will switch each other's selection.
 *
 * Delimiter convention (enforced in ScheduleParser):
 *   Teachers / Subjects : "/" separator  (e.g. "Mrs. X / Mr. Y")
 *   Classes             : "," separator  (e.g. "11th A, 11th B")
 */

/** Email of the person whose edit is being processed (set per edit by handleEdit_). */
var CURRENT_EDITOR_ = '';

function onEdit(e) {
  if (!e || !e.range) return;
  const installedUid = getInstalledEditTriggerUid_();
  if (e.triggerUid) {
    // Installable trigger: only the one recorded by installEditTrigger_ does the work.
    // Older / duplicate triggers (e.g. one added by hand, or by another editor) are ignored.
    if (installedUid && e.triggerUid === installedUid) {
      markSyncAlive_();
      handleEdit_(e);
    }
    return;
  }
  // Simple trigger: steps aside once the installable trigger exists — but checks that it
  // really ran, so a trigger that stopped (its owner lost access, it was deleted, or the
  // daily trigger quota ran out) does not drop every edit without a word.
  if (installedUid) {
    if (isSyncedSheet_(e.range.getSheet().getName())) warnIfSyncStopped_(e);
    return;
  }
  handleEdit_(e);
}

/** Tabs whose edits the installable trigger writes back or reacts to. */
function isSyncedSheet_(name) {
  const S = CONFIG.SHEETS;
  return [S.CLASS_VIEW, S.TEACHER_VIEW, S.TEACHER_DAY_VIEW, S.SCHEDULE, S.TEACHERS, S.CLASSES, S.SUBJECTS].indexOf(name) !== -1;
}

/** Called by the installable trigger as it starts: proof for the simple trigger that sync runs. */
function markSyncAlive_() {
  try { PropertiesService.getDocumentProperties().setProperty(CONFIG.PROP_LAST_SYNC, String(Date.now())); } catch (err) { /* best effort */ }
}

/**
 * Simple-trigger side of the check: waits a few seconds for the installable trigger to
 * record that it started after this edit, and warns the editor when it did not.
 */
function warnIfSyncStopped_(e) {
  const started = Date.now();
  const lastSync = () => {
    try { return Number(PropertiesService.getDocumentProperties().getProperty(CONFIG.PROP_LAST_SYNC)) || 0; } catch (err) { return -1; }
  };
  // Sync ran within the last minute: it is alive, no need to wait (keeps simple-trigger time low)
  const before = lastSync();
  if (before < 0 || before >= started - CONFIG.SYNC_ALIVE_WINDOW_MS) return;
  Utilities.sleep(CONFIG.SYNC_CHECK_WAIT_MS);
  const after = lastSync();
  if (after < 0 || after >= started - CONFIG.SYNC_ALIVE_WINDOW_MS) return;
  e.source.toast('This edit may NOT have been saved: timetable sync did not respond. If this keeps happening, ' +
    'run Timetable System → ⚙️ Setup / Repair (it reinstalls the sync).', '⚠️ Sync not running', 20);
}

function handleEdit_(e) {
  const sheetName = e.range.getSheet().getName();
  const S = CONFIG.SHEETS;
  // Who made the edit, for Change_Log (the installable trigger runs as whoever installed it)
  CURRENT_EDITOR_ = '';
  try { CURRENT_EDITOR_ = (e.user && e.user.getEmail && e.user.getEmail()) || ''; } catch (err) { CURRENT_EDITOR_ = ''; }
  try {
    if (sheetName === S.CLASS_VIEW) {
      handleClassViewEdit(e);
    } else if (sheetName === S.TEACHER_VIEW) {
      handleTeacherViewEdit(e);
    } else if (sheetName === S.TEACHER_DAY_VIEW) {
      handleTeacherDayViewEdit(e);
    } else if (sheetName === S.SCHEDULE) {
      handleMasterScheduleEdit(e);
    } else if (sheetName === S.TEACHERS || sheetName === S.CLASSES || sheetName === S.SUBJECTS) {
      handleRegistryEdit(e);
    }
  } catch (err) {
    e.source.toast(String(err && err.message || err), '⚠️ Timetable sync failed', 10);
    throw err;
  }
}

function getInstalledEditTriggerUid_() {
  try {
    return PropertiesService.getDocumentProperties().getProperty(CONFIG.PROP_EDIT_TRIGGER) || '';
  } catch (err) {
    return '';
  }
}

function installEditTrigger_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Only editors may (re)point the sync trigger: a viewer's trigger could not write.
  // This harmless write throws for anyone without edit access.
  const probe = ss.getSheetByName(CONFIG.SHEETS.SCHEDULE) || ss.getSheets()[0];
  probe.getRange(1, 1).setValue(probe.getRange(1, 1).getValue());

  // Remove this user's own duplicates (other users' triggers cannot be listed;
  // they are made harmless by the uid check in onEdit).
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'onEdit' && t.getEventType() === ScriptApp.EventType.ON_EDIT)
    .forEach(t => ScriptApp.deleteTrigger(t));
  const trigger = ScriptApp.newTrigger('onEdit').forSpreadsheet(ss).onEdit().create();
  PropertiesService.getDocumentProperties().setProperty(CONFIG.PROP_EDIT_TRIGGER, trigger.getUniqueId());

  // Change trigger: refreshes the views after Master_Schedule rows are deleted or inserted
  // (those never fire onEdit). Same uid scheme, so a duplicate does nothing.
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'onScheduleChange')
    .forEach(t => ScriptApp.deleteTrigger(t));
  const change = ScriptApp.newTrigger('onScheduleChange').forSpreadsheet(ss).onChange().create();
  PropertiesService.getDocumentProperties().setProperty(CONFIG.PROP_CHANGE_TRIGGER, change.getUniqueId());
}

/**
 * Reads the edited cells NOW (before waiting for the lock), so a re-render by
 * another editor cannot change what this edit means.
 * @returns {Array<{row: number, col: number, value: *}>|null} null = too many cells
 */
function collectEditedCells_(e, minRow, minCol, maxCol) {
  const range = e.range;
  const rows = range.getNumRows();
  const cols = range.getNumColumns();
  if (rows * cols > CONFIG.MAX_SYNCED_CELLS) {
    e.source.toast('Edit too large to sync cell by cell (' + rows * cols + ' cells). Use Refresh All Views after editing Master_Schedule instead.', 'Timetable', 10);
    return null;
  }
  const values = rows * cols === 1
    ? [[e.value !== undefined ? e.value : range.getValue()]]
    : range.getValues();
  const cells = [];
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const row = range.getRow() + r;
      const col = range.getColumn() + c;
      if (row < minRow || col < minCol || col > maxCol) continue;
      cells.push({ row: row, col: col, value: values[r][c] });
    }
  }
  return cells;
}

/** Re-renders every view except the one named in `skip` (it re-renders itself). */
function refreshViewsAfterWrite_(ss, skip) {
  const S = CONFIG.SHEETS;
  if (skip !== S.CLASS_VIEW) {
    const cv = ss.getSheetByName(S.CLASS_VIEW);
    const activeClass = cv ? cv.getRange('B3').getValue() : '';
    if (activeClass) ClassViewManager.renderClassView(activeClass);
  }
  if (skip !== S.TEACHER_VIEW) {
    const tv = ss.getSheetByName(S.TEACHER_VIEW);
    const activeTeacher = tv ? tv.getRange('B3').getValue() : '';
    if (activeTeacher) TeacherViewManager.renderTeacherView(activeTeacher);
  }
  if (skip !== S.TEACHER_DAY_VIEW) {
    const tdv = ss.getSheetByName(S.TEACHER_DAY_VIEW);
    if (tdv) {
      const activeDay = tdv.getRange('B3').getValue() || CONFIG.DAYS[0];
      const activeClassFilter = tdv.getRange('E3').getValue() || CONFIG.ALL_CLASSES;
      TeacherDayViewManager.renderTeacherDayView(activeDay, activeClassFilter);
    }
  }
  refreshMasterGridData_();
}

// ─────────────────────────────────────────────────────────
// TEACHERS / CLASSES / SUBJECTS TABS
// ─────────────────────────────────────────────────────────

/**
 * A new teacher, a changed "Days Unavailable", or a new class or subject changes what the
 * views, dropdowns, DAY OFF marks and Clash Status show, so they are brought up to date.
 * A renamed teacher or class is not renamed in Master_Schedule; the editor is told how
 * many periods still use the old name.
 */
function handleRegistryEdit(e) {
  const range = e.range;
  if (range.getLastRow() < 2) return; // header row
  const sheet = range.getSheet();
  const S = CONFIG.SHEETS;
  // The owner's own columns (phone numbers, notes) and the computed Total Hours change
  // nothing the views show: Teachers A-D, Classes A-C, Subjects A-B do
  const width = sheet.getName() === S.TEACHERS ? 4 : sheet.getName() === S.CLASSES ? 3 : 2;
  if (range.getColumn() > width) return;
  let notice = '';
  withDocumentLock_(() => {
    const single = range.getNumRows() === 1 && range.getNumColumns() === 1;
    const nameCol = sheet.getName() === S.SUBJECTS ? 0 : 1; // Teacher Name / Class Name are column A
    if (single && nameCol && range.getColumn() === nameCol && e.oldValue !== undefined && typeof e.oldValue !== 'object') {
      const oldName = String(e.oldValue).trim();
      const newName = String(range.getValue() == null ? '' : range.getValue()).trim();
      if (oldName && !ScheduleParser.sameName(oldName, newName)) {
        const isTeacher = sheet.getName() === S.TEACHERS;
        const rows = DataAccess.getSheetDataAsObjects(S.SCHEDULE).filter(r => isTeacher
          ? ScheduleParser.rowIncludesTeacher(r, oldName) : ScheduleParser.rowIncludesClass(r, oldName)).length;
        if (rows) {
          notice = rows + ' period(s) in Master_Schedule still use the old name "' + oldName + '". Change them there ' +
            '(Edit → Find and replace), or they show as unknown in Check For Conflicts.';
        }
      }
    }
    applyMasterScheduleDropdowns_(/* silent= */ true);
    afterScheduleChange_(e.source, null);
  });
  if (notice) e.source.toast(notice, 'Name changed', 20);
}

// ─────────────────────────────────────────────────────────
// CLASS VIEW
// ─────────────────────────────────────────────────────────

function handleClassViewEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();

  // B3 dropdown → re-render for new class
  if (range.getRow() === 3 && range.getColumn() === 2) {
    const className = range.getValue();
    if (className) withDocumentLock_(() => ClassViewManager.renderClassView(className));
    return;
  }

  // Grid: row 5 = header, rows 6+ alternate Subject/Teacher per day; columns B+ = Prayer,
  // periods 1–4, Lunch, periods 5–8 (viewSlots_). Capture the selector and values before
  // waiting for the lock.
  const width = viewWidth_();
  const className = sheet.getRange('B3').getValue();
  if (!className) {
    // Only an edit inside the timetable grid needs a class; a title or note edit does not
    if (range.getLastRow() >= 6 && range.getLastColumn() >= 2 && range.getColumn() <= width) {
      e.source.toast('Pick a class in B3 first — this edit was not saved.', 'Not saved', 10);
    }
    return;
  }
  const collected = collectEditedCells_(e, 6, 2, width);
  if (!collected || !collected.length) return;
  const split = periodCells_(collected);
  const cells = split.cells;

  // A period's Subject and Teacher cells are applied together (one edit per period),
  // so pasting or clearing both at once is checked as a whole.
  const edits = {};
  const order = [];
  cells.forEach(cell => {
    const dayIndex = Math.floor((cell.row - 6) / 2);
    if (dayIndex < 0 || dayIndex >= CONFIG.DAYS.length) return;
    const key = dayIndex + '|' + cell.period;
    if (!edits[key]) { edits[key] = { day: CONFIG.DAYS[dayIndex], period: cell.period, edit: {} }; order.push(key); }
    edits[key].edit[((cell.row - 6) % 2 === 0) ? 'subject' : 'teacher'] = cell.value;
  });
  if (!order.length) {
    // Only a Prayer / Lunch cell was typed in: put the view back
    if (split.breakHit) {
      withDocumentLock_(() => ClassViewManager.renderClassView(className));
      toastEditResult_(e.source, [breakColumnNotice_()], []);
    }
    return;
  }

  const problems = split.breakHit ? [breakColumnNotice_()] : [];
  const notices = [];
  withDocumentLock_(() => {
    order.forEach(key => {
      const x = edits[key];
      const res = ClassViewManager.applyClassViewEdit(className, x.day, x.period, x.edit);
      if (res.error) problems.push(res.error);
      (res.notices || []).forEach(n => notices.push(n));
      logChanges_(CONFIG.SHEETS.CLASS_VIEW, x.day, x.period, res.changes);
    });
    // Redraw this view too: reverts rejected edits and updates the split/combined/clash colours
    ClassViewManager.renderClassView(className);
    afterScheduleChange_(e.source, CONFIG.SHEETS.CLASS_VIEW);
  });
  toastEditResult_(e.source, problems, notices);
}

// ─────────────────────────────────────────────────────────
// TEACHER VIEW
// ─────────────────────────────────────────────────────────

function handleTeacherViewEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();

  // B3 dropdown → re-render for new teacher
  if (range.getRow() === 3 && range.getColumn() === 2) {
    const teacherName = range.getValue();
    if (teacherName) withDocumentLock_(() => TeacherViewManager.renderTeacherView(teacherName));
    return;
  }

  // Grid: rows 6–11 = Mon–Sat; columns B+ = Prayer, periods 1–4, Lunch, periods 5–8
  const teacherName = sheet.getRange('B3').getValue();
  if (!teacherName) {
    e.source.toast('Pick a teacher in B3 first — this edit was not saved.', 'Not saved', 10);
    return;
  }
  const collected = collectEditedCells_(e, 6, 2, viewWidth_());
  if (!collected || !collected.length) return;
  const split = periodCells_(collected);

  const problems = split.breakHit ? [breakColumnNotice_()] : [];
  const notices = [];
  withDocumentLock_(() => {
    split.cells.forEach(cell => {
      const dayIndex = cell.row - 6;
      if (dayIndex < 0 || dayIndex >= CONFIG.DAYS.length) return;
      const day = CONFIG.DAYS[dayIndex];
      const changes = [];
      const err = updateTeacherSlot_(teacherName, day, cell.period, cell.value, changes, notices);
      if (err) problems.push(err);
      logChanges_(CONFIG.SHEETS.TEACHER_VIEW, day, cell.period, changes);
    });
    // Re-render this view too: normalises the typed text and reverts rejected edits
    TeacherViewManager.renderTeacherView(teacherName);
    if (split.cells.length) afterScheduleChange_(e.source, CONFIG.SHEETS.TEACHER_VIEW);
  });
  toastEditResult_(e.source, problems, notices);
}

// ─────────────────────────────────────────────────────────
// TEACHER DAY VIEW
// ─────────────────────────────────────────────────────────

function handleTeacherDayViewEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();

  // Row 3 edits (B3 = Day selector, E3 = Class filter)
  if (range.getRow() === 3) {
    const day = sheet.getRange('B3').getValue() || CONFIG.DAYS[0];
    const filterClass = sheet.getRange('E3').getValue() || CONFIG.ALL_CLASSES;
    withDocumentLock_(() => TeacherDayViewManager.renderTeacherDayView(day, filterClass));
    return;
  }

  // Grid: rows 6+ = one teacher per row; columns B+ = Prayer, periods 1–4, Lunch, periods 5–8.
  // Capture day, filter, teacher names and values before waiting for the lock.
  const selectedDay = sheet.getRange('B3').getValue() || CONFIG.DAYS[0];
  const activeClassFilter = sheet.getRange('E3').getValue() || CONFIG.ALL_CLASSES;
  const collected = collectEditedCells_(e, 6, 2, viewWidth_());
  if (!collected || !collected.length) return;
  const split = periodCells_(collected);
  const cells = split.cells;
  if (cells.length) {
    const firstRow = Math.min.apply(null, cells.map(c => c.row));
    const lastRow = Math.max.apply(null, cells.map(c => c.row));
    const names = sheet.getRange(firstRow, 1, lastRow - firstRow + 1, 1).getValues();
    cells.forEach(c => { c.teacher = names[c.row - firstRow][0]; });
  }

  const problems = split.breakHit ? [breakColumnNotice_()] : [];
  const notices = [];
  const day = normalizeDay_(selectedDay) || CONFIG.DAYS[0];
  withDocumentLock_(() => {
    cells.forEach(cell => {
      if (!cell.teacher) return;
      const changes = [];
      const err = updateTeacherSlot_(cell.teacher, day, cell.period, cell.value, changes, notices);
      if (err) problems.push(cell.teacher + ', ' + err);
      logChanges_(CONFIG.SHEETS.TEACHER_DAY_VIEW, day, cell.period, changes);
    });
    TeacherDayViewManager.renderTeacherDayView(selectedDay, activeClassFilter);
    if (cells.length) afterScheduleChange_(e.source, CONFIG.SHEETS.TEACHER_DAY_VIEW);
  });
  toastEditResult_(e.source, problems, notices);
}

// ─────────────────────────────────────────────────────────
// MASTER SCHEDULE → cascade refresh
// ─────────────────────────────────────────────────────────

function handleMasterScheduleEdit(e) {
  if (e.range.getLastRow() < 2) return; // header-only edit
  const C = CONFIG.COL;
  const range = e.range;
  // Notes in the owner's own columns (or a hand edit of Clash Status) cannot change the
  // timetable, so they do not pay for a full refresh
  if (range.getColumn() > C.ROOM) return;
  // Single-cell edits are logged with the old value Sheets passes in (e.oldValue)
  const single = range.getNumRows() === 1 && range.getNumColumns() === 1;
  withDocumentLock_(() => {
    const sheet = range.getSheet();
    // Day names typed as "monday" / "Mon" are written back as "Monday", so every view
    // (and the clash check) reads them the same way.
    if (range.getColumn() <= C.DAY && range.getLastColumn() >= C.DAY) {
      const first = Math.max(2, range.getRow());
      const n = range.getLastRow() - first + 1;
      if (n > 0) {
        const days = sheet.getRange(first, C.DAY, n, 1).getValues();
        let changed = false;
        const fixed = days.map(r => {
          const d = normalizeDay_(r[0]);
          if (d && d !== r[0]) { changed = true; return [d]; }
          return [r[0]];
        });
        if (changed) sheet.getRange(first, C.DAY, n, 1).setValues(fixed);
      }
    }
    const problems = [];
    const unchecked = [];
    const notes = [];
    let logged = false;
    // A single Subject / Teacher cell goes through the same checks as Class_View
    // (pick-to-fill on split periods, names, new clashes, groups lining up).
    if (single && range.getRow() >= 2 && (range.getColumn() === C.SUBJECT || range.getColumn() === C.TEACHER)) {
      const res = checkMasterScheduleCell_(sheet, range, e.oldValue);
      if (res) {
        logged = true;
        if (res.error) (res.kept ? unchecked : problems).push(res.error);
        if (res.notice) notes.push(res.notice);
        if (!res.error || res.kept) logChanges_(CONFIG.SHEETS.SCHEDULE, res.day, res.period, res.changes);
      }
    }
    if (!logged && single && range.getRow() >= 2 && range.getColumn() <= C.ROOM) {
      const row = sheet.getRange(range.getRow(), 1, 1, CONFIG.SCHEDULE_HEADERS.length).getValues()[0];
      logChanges_(CONFIG.SHEETS.SCHEDULE, normalizeDay_(row[C.DAY - 1]) || row[C.DAY - 1], row[C.PERIOD - 1], [{
        cls: String(row[C.CLASS - 1]),
        before: CONFIG.SCHEDULE_HEADERS[range.getColumn() - 1] + ': ' + (e.oldValue == null ? '' : e.oldValue),
        after: CONFIG.SCHEDULE_HEADERS[range.getColumn() - 1] + ': ' + (e.value == null ? range.getValue() : e.value)
      }]);
    }
    // Pastes, fill-downs and block deletes: one log line per row (Sheets does not report
    // the old values of a multi-cell edit)
    if (!single) {
      const first = Math.max(2, range.getRow());
      const n = range.getLastRow() - first + 1;
      if (n > 0) {
        const lastCol = Math.min(range.getLastColumn(), C.ROOM);
        const rows = sheet.getRange(first, 1, n, CONFIG.SCHEDULE_HEADERS.length).getValues();
        // Built first and written to Change_Log in one call: a 1,000-row paste holds the lock briefly
        const batch = rows.map(r => ({
          day: normalizeDay_(r[C.DAY - 1]) || r[C.DAY - 1], period: r[C.PERIOD - 1], cls: String(r[C.CLASS - 1]),
          before: '(multi-cell edit: previous values not reported)',
          after: CONFIG.SCHEDULE_HEADERS.slice(range.getColumn() - 1, lastCol)
            .map((h, k) => h + ': ' + r[range.getColumn() - 1 + k]).join(' · ')
        }));
        logChanges_(CONFIG.SHEETS.SCHEDULE, '', '', batch);
      }
    }
    // Teacher/class names may have been added, removed or replaced (including in
    // multi-teacher cells), so every view showing schedule data is refreshed.
    afterScheduleChange_(e.source, null);
    if (problems.length) e.source.toast(problems.join('\n'), 'Not saved', 15);
    if (notes.length) e.source.toast(notes.join('\n'), 'Saved', 10);
    if (unchecked.length) {
      e.source.toast(unchecked.join('\n') + '\nThe cell was left as it is: Sheets did not report what it held before ' +
        '(a paste, or a cell that was empty), so nothing was undone. On a split period, pick the next teacher or type the full list.',
        'Check this cell', 15);
    }
  });
}

/**
 * Runs a typed Subject / Teacher cell of Master_Schedule through applyClassViewEdit.
 * The old value is put back first, so the checks see the row as it was; on success the
 * checked value is written, on failure the old value stays.
 *
 * Sheets does not report the previous value for a paste, or for a cell that was empty
 * (e.oldValue is undefined). Then the previous text is unknown, so the cell is never
 * rewritten: it is only checked as it now stands, and a problem is reported without
 * undoing anything (`kept: true`).
 * Rows still being filled in (no valid day, period or known class yet) are left alone.
 * @returns {{error: string, kept: boolean, notice: string, changes: Array, day: string, period: number}|null}  null = not checked
 */
function checkMasterScheduleCell_(sheet, range, oldValue) {
  const C = CONFIG.COL;
  const row = sheet.getRange(range.getRow(), 1, 1, CONFIG.SCHEDULE_HEADERS.length).getValues()[0];
  const day = normalizeDay_(row[C.DAY - 1]);
  const period = Number(row[C.PERIOD - 1]);
  const cls = String(row[C.CLASS - 1] || '').trim();
  if (!day || CONFIG.PERIODS.indexOf(period) === -1 || !cls) return null;
  const known = getClassNames_();
  if (!ScheduleParser.splitClasses(cls).every(c => known.some(k => ScheduleParser.sameName(k, c)))) return null;

  const newValue = range.getValue();
  const edit = range.getColumn() === C.SUBJECT ? { subject: newValue } : { teacher: newValue };
  const oldKnown = oldValue !== undefined && oldValue !== null && typeof oldValue !== 'object';
  if (!oldKnown) {
    // A single name picked into a split row's Teacher cell is the first pick-to-fill step
    // ("Asha" on "Hindi / IT" → "Asha / -"). Writing the filled form only adds empty-group
    // marks to what is already in the cell, so it is safe even when the old text is unknown.
    let checked = edit;
    const subjectText = String(row[C.SUBJECT - 1] || '');
    const groups = ScheduleParser.splitList(subjectText).length;
    const picked = String(newValue == null ? '' : newValue).trim();
    const fill = range.getColumn() === C.TEACHER && groups > 1 && picked && !/[\/+,]/.test(picked) &&
      getTeacherNames_().some(k => ScheduleParser.sameName(k, picked));
    if (fill) checked = { teacher: ScheduleParser.fillNextGroup('', picked, groups).teacher };
    const res = ClassViewManager.applyClassViewEdit(cls, day, period, checked, { keepEmptyRow: true, onlyRow: range.getRow(), dryRun: true });
    // One subject pasted over a split row: the old groups are not reported, but the Teacher
    // cell still lists several teacher groups, so the other groups' subjects were lost
    if (!res.error && range.getColumn() === C.SUBJECT && groups === 1) {
      const staffed = ScheduleParser.splitList(String(row[C.TEACHER - 1] || '')).filter(t => t && !/^[-–—]$/.test(t));
      if (staffed.length > 1) {
        res.error = 'one subject with ' + staffed.length + ' teacher groups (' + row[C.TEACHER - 1] +
          '). Fine if they share the class; if this was a split period, type every group\'s subject in order (e.g. "Hindi / IT"), or press Ctrl+Z';
      }
    }
    if (fill && !res.error) range.setValue(checked.teacher);
    const before = '(not reported by Sheets: the cell was empty, or a paste replaced it)';
    const changes = [{ cls: cls, before: before, after: CONFIG.SCHEDULE_HEADERS[range.getColumn() - 1] + ': ' + (checked.teacher !== undefined ? checked.teacher : newValue) }];
    const notice = fill && !res.error
      ? day + ' P' + period + ': ' + picked + ' is in the first empty group (' + checked.teacher + '). Pick the next teacher, or type the full list.'
      : '';
    return { error: res.error, kept: true, notice: notice, changes: changes, day: day, period: period };
  }
  range.setValue(oldValue);
  const res = ClassViewManager.applyClassViewEdit(cls, day, period, edit, { keepEmptyRow: true, onlyRow: range.getRow() });
  return { error: res.error, kept: false, changes: res.changes, day: day, period: period };
}

// ─────────────────────────────────────────────────────────
// ROWS DELETED / INSERTED (onChange)
// ─────────────────────────────────────────────────────────

/**
 * Deleting a Master_Schedule row removes a period but fires no edit trigger, so the
 * views would keep showing it. The installable change trigger (added by Setup / Repair)
 * refreshes everything after rows are removed.
 */
function onScheduleChange(e) {
  if (!e) return;
  const installedUid = getInstalledChangeTriggerUid_();
  if (!installedUid || e.triggerUid !== installedUid) return;
  if (['REMOVE_ROW', 'INSERT_ROW', 'REMOVE_GRID'].indexOf(String(e.changeType)) === -1) return;
  const ss = e.source || SpreadsheetApp.getActiveSpreadsheet();
  withDocumentLock_(() => afterScheduleChange_(ss, null));
}

function getInstalledChangeTriggerUid_() {
  try {
    return PropertiesService.getDocumentProperties().getProperty(CONFIG.PROP_CHANGE_TRIGGER) || '';
  } catch (err) {
    return '';
  }
}

// ─────────────────────────────────────────────────────────
// AFTER EVERY SAVED EDIT
// ─────────────────────────────────────────────────────────

/**
 * Brings everything that depends on Master_Schedule up to date after an edit:
 * the other views, the Master Grid, Clash Status and each teacher's period count. Without this the Clash Status column and the load totals
 * stayed stale until someone ran Check For Conflicts.
 */
function afterScheduleChange_(ss, skip) {
  refreshViewsAfterWrite_(ss, skip);
  validateSchedule_();
}

/** Shows what was refused and what changed for other teachers, in one toast. */
function toastEditResult_(ss, problems, notices) {
  if (problems.length) ss.toast(problems.concat(notices).join('\n'), 'Not saved', 15);
  else if (notices.length) ss.toast(notices.join('\n'), 'Saved', 10);
}

/**
 * Appends one line per changed period to the Change_Log tab (created on first use):
 * time, editor, where it was edited, day, period, class, before, after.
 */
/** Text that Sheets would read as a formula ("=…", "+…", "@…") is stored as plain text. */
function asText_(v) {
  const s = String(v == null ? '' : v);
  return /^[=+@-]/.test(s) ? "'" + s : s;
}

function logChanges_(where, day, period, changes) {
  if (!changes || !changes.length) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEETS.CHANGE_LOG);
  if (!sheet) {
    const active = ss.getActiveSheet(); // insertSheet switches tabs; keep the editor where they were
    sheet = ss.insertSheet(CONFIG.SHEETS.CHANGE_LOG);
    sheet.getRange(1, 1, 1, CONFIG.CHANGE_LOG_HEADERS.length).setValues([CONFIG.CHANGE_LOG_HEADERS]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    if (active) ss.setActiveSheet(active);
  }
  let editor = '';
  editor = CURRENT_EDITOR_;
  if (!editor) {
    try { editor = typeof Session !== 'undefined' ? Session.getActiveUser().getEmail() : ''; } catch (err) { editor = ''; }
  }
  const now = new Date();
  const rows = changes.filter(c => c.before !== c.after)
    .map(c => [now, editor || '(unknown)', where, c.day !== undefined ? c.day : day, c.period !== undefined ? c.period : period,
      c.cls, asText_(c.before), asText_(c.after)]);
  if (!rows.length) return;
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}
