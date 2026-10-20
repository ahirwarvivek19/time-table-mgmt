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
 * cascade can exceed on a large timetable. "Install Edit Sync Trigger" (also run
 * by Setup) adds an installable trigger and records its unique id. From then on
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

/** Menu: installs (or repairs) the installable onEdit trigger. */
function installEditTrigger() {
  installEditTrigger_();
  SpreadsheetApp.getUi().alert(
    '✅ Edit sync trigger installed.\n\n' +
    'Edits in the views now sync to Master_Schedule. The trigger runs under your account — ' +
    'if you ever lose edit access, another editor should run this menu item again.'
  );
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

  withDocumentLock_(() => {
    cells.forEach(cell => {
      const dayIndex = Math.floor((cell.row - 6) / 2);
      if (dayIndex < 0 || dayIndex >= CONFIG.DAYS.length) return;
      const editType = ((cell.row - 6) % 2 === 0) ? 'Subject' : 'Teacher';
      ClassViewManager.updateMasterFromClassView(className, CONFIG.DAYS[dayIndex], cell.col - 1, editType, cell.value);
    });
    refreshViewsAfterWrite_(e.source, CONFIG.SHEETS.CLASS_VIEW);
  });
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
  withDocumentLock_(() => {
    cells.forEach(cell => {
      const dayIndex = cell.row - 6;
      if (dayIndex < 0 || dayIndex >= CONFIG.DAYS.length) return;
      const err = TeacherViewManager.updateMasterFromTeacherView(teacherName, CONFIG.DAYS[dayIndex], cell.col - 1, cell.value);
      if (err) problems.push(err);
    });
    // Re-render this view too: normalises the typed text and reverts rejected edits
    TeacherViewManager.renderTeacherView(teacherName);
    refreshViewsAfterWrite_(e.source, CONFIG.SHEETS.TEACHER_VIEW);
  });
  if (problems.length) e.source.toast(problems.join('\n'), 'Not saved', 10);
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
  withDocumentLock_(() => {
    cells.forEach(cell => {
      if (!cell.teacher) return;
      const err = TeacherDayViewManager.updateMasterFromTeacherDayView(selectedDay, cell.teacher, cell.col - 1, cell.value);
      if (err) problems.push(err);
    });
    TeacherDayViewManager.renderTeacherDayView(selectedDay, activeClassFilter);
    refreshViewsAfterWrite_(e.source, CONFIG.SHEETS.TEACHER_DAY_VIEW);
  });
  if (problems.length) e.source.toast(problems.join('\n'), 'Not saved', 10);
}

// ─────────────────────────────────────────────────────────
// MASTER SCHEDULE → cascade refresh
// ─────────────────────────────────────────────────────────

function handleMasterScheduleEdit(e) {
  if (e.range.getLastRow() < 2) return; // header-only edit
  // Teacher/class names may have been added, removed or replaced (including in
  // multi-teacher cells), so every view showing schedule data is refreshed.
  withDocumentLock_(() => refreshViewsAfterWrite_(e.source, null));
}
