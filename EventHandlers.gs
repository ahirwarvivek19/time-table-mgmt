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

function onEdit(e) {
  if (!e || !e.range) return;
  const installedUid = getInstalledEditTriggerUid_();
  if (e.triggerUid) {
    // Installable trigger: only the one recorded by installEditTrigger_ does the work.
    // Older / duplicate triggers (e.g. one added by hand, or by another editor) are ignored.
    if (installedUid && e.triggerUid === installedUid) handleEdit_(e);
    return;
  }
  // Simple trigger: steps aside once the installable trigger exists.
  if (installedUid) return;
  handleEdit_(e);
}

function handleEdit_(e) {
  const sheetName = e.range.getSheet().getName();
  const S = CONFIG.SHEETS;
  try {
    if (sheetName === S.CLASS_VIEW) {
      handleClassViewEdit(e);
    } else if (sheetName === S.TEACHER_VIEW) {
      handleTeacherViewEdit(e);
    } else if (sheetName === S.TEACHER_DAY_VIEW) {
      handleTeacherDayViewEdit(e);
    } else if (sheetName === S.SCHEDULE) {
      handleMasterScheduleEdit(e);
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

  // Grid: row 5 = header, rows 6+ alternate Subject/Teacher per day, cols 2–9 = periods.
  // Capture the selector and values before waiting for the lock.
  const className = sheet.getRange('B3').getValue();
  if (!className) return;
  const cells = collectEditedCells_(e, 6, 2, 9);
  if (!cells || !cells.length) return;

  // A period's Subject and Teacher cells are applied together (one edit per period),
  // so pasting or clearing both at once is checked as a whole.
  const edits = {};
  const order = [];
  cells.forEach(cell => {
    const dayIndex = Math.floor((cell.row - 6) / 2);
    if (dayIndex < 0 || dayIndex >= CONFIG.DAYS.length) return;
    const key = dayIndex + '|' + (cell.col - 1);
    if (!edits[key]) { edits[key] = { day: CONFIG.DAYS[dayIndex], period: cell.col - 1, edit: {} }; order.push(key); }
    edits[key].edit[((cell.row - 6) % 2 === 0) ? 'subject' : 'teacher'] = cell.value;
  });
  if (!order.length) return;

  const problems = [];
  withDocumentLock_(() => {
    order.forEach(key => {
      const x = edits[key];
      const res = ClassViewManager.applyClassViewEdit(className, x.day, x.period, x.edit);
      if (res.error) problems.push(res.error);
      logChanges_(CONFIG.SHEETS.CLASS_VIEW, x.day, x.period, res.changes);
    });
    // Redraw this view too: reverts rejected edits and updates the split/combined/clash colours
    ClassViewManager.renderClassView(className);
    afterScheduleChange_(e.source, CONFIG.SHEETS.CLASS_VIEW);
  });
  if (problems.length) e.source.toast(problems.join('\n'), 'Not saved', 15);
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

  // Grid: rows 6–11 = Mon–Sat, cols 2–9 = Period 1–8
  const teacherName = sheet.getRange('B3').getValue();
  if (!teacherName) return;
  const cells = collectEditedCells_(e, 6, 2, 9);
  if (!cells || !cells.length) return;

  const problems = [];
  const notices = [];
  withDocumentLock_(() => {
    cells.forEach(cell => {
      const dayIndex = cell.row - 6;
      if (dayIndex < 0 || dayIndex >= CONFIG.DAYS.length) return;
      const day = CONFIG.DAYS[dayIndex];
      const changes = [];
      const err = updateTeacherSlot_(teacherName, day, cell.col - 1, cell.value, changes, notices);
      if (err) problems.push(err);
      logChanges_(CONFIG.SHEETS.TEACHER_VIEW, day, cell.col - 1, changes);
    });
    // Re-render this view too: normalises the typed text and reverts rejected edits
    TeacherViewManager.renderTeacherView(teacherName);
    afterScheduleChange_(e.source, CONFIG.SHEETS.TEACHER_VIEW);
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

  // Grid: rows 6+ = one teacher per row, cols 2–9 = Period 1–8.
  // Capture day, filter, teacher names and values before waiting for the lock.
  const selectedDay = sheet.getRange('B3').getValue() || CONFIG.DAYS[0];
  const activeClassFilter = sheet.getRange('E3').getValue() || CONFIG.ALL_CLASSES;
  const cells = collectEditedCells_(e, 6, 2, 9);
  if (!cells || !cells.length) return;
  const firstRow = Math.min.apply(null, cells.map(c => c.row));
  const lastRow = Math.max.apply(null, cells.map(c => c.row));
  const names = sheet.getRange(firstRow, 1, lastRow - firstRow + 1, 1).getValues();
  cells.forEach(c => { c.teacher = names[c.row - firstRow][0]; });

  const problems = [];
  const notices = [];
  const day = normalizeDay_(selectedDay) || CONFIG.DAYS[0];
  withDocumentLock_(() => {
    cells.forEach(cell => {
      if (!cell.teacher) return;
      const changes = [];
      const err = updateTeacherSlot_(cell.teacher, day, cell.col - 1, cell.value, changes, notices);
      if (err) problems.push(cell.teacher + ', ' + err);
      logChanges_(CONFIG.SHEETS.TEACHER_DAY_VIEW, day, cell.col - 1, changes);
    });
    TeacherDayViewManager.renderTeacherDayView(selectedDay, activeClassFilter);
    afterScheduleChange_(e.source, CONFIG.SHEETS.TEACHER_DAY_VIEW);
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
    let logged = false;
    // A single Subject / Teacher cell goes through the same checks as Class_View
    // (pick-to-fill on split periods, names, new clashes, groups lining up).
    if (single && range.getRow() >= 2 && (range.getColumn() === C.SUBJECT || range.getColumn() === C.TEACHER)) {
      const res = checkMasterScheduleCell_(sheet, range, e.oldValue);
      if (res) {
        logged = true;
        if (res.error) problems.push(res.error);
        else logChanges_(CONFIG.SHEETS.SCHEDULE, res.day, res.period, res.changes);
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
    // Teacher/class names may have been added, removed or replaced (including in
    // multi-teacher cells), so every view showing schedule data is refreshed.
    afterScheduleChange_(e.source, null);
    if (problems.length) e.source.toast(problems.join('\n'), 'Not saved', 15);
  });
}

/**
 * Runs a typed Subject / Teacher cell of Master_Schedule through applyClassViewEdit.
 * The old value is put back first, so the checks see the row as it was; on success the
 * checked value is written, on failure the old value stays.
 * Rows still being filled in (no valid day, period or known class yet) are left alone.
 * @returns {{error: string, changes: Array, day: string, period: number}|null}  null = not checked
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
  range.setValue(oldValue == null ? '' : oldValue);
  const edit = range.getColumn() === C.SUBJECT ? { subject: newValue } : { teacher: newValue };
  const res = ClassViewManager.applyClassViewEdit(cls, day, period, edit, { keepEmptyRow: true, onlyRow: range.getRow() });
  return { error: res.error, changes: res.changes, day: day, period: period };
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
 * the other views, the Master Grid, Clash Status, each teacher's period count and
 * the availability grid. Without this the Clash Status column and the load totals
 * stayed stale until someone ran Check For Conflicts.
 */
function afterScheduleChange_(ss, skip) {
  refreshViewsAfterWrite_(ss, skip);
  validateSchedule_();
  const grid = ss.getSheetByName(CONFIG.SHEETS.AVAILABILITY);
  const day = grid ? availabilityGridDay_(grid) : '';
  if (day) renderTeacherAvailabilityGrid_(day, true);
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
  try { editor = typeof Session !== 'undefined' ? Session.getActiveUser().getEmail() : ''; } catch (err) { editor = ''; }
  const now = new Date();
  const rows = changes.filter(c => c.before !== c.after)
    .map(c => [now, editor || '(unknown)', where, day, period, c.cls, c.before, c.after]);
  if (!rows.length) return;
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}
