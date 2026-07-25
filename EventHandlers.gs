/**
 * EventHandlers.gs
 * Centralizes all trigger events (onEdit).
 *
 * Source of truth: Master_Schedule.
 * Every editable view writes back to Master_Schedule, then cascades
 * re-renders to all other views so they always reference Master_Schedule.
 *
 * Editable sheets & write-back targets:
 *   Class_View        — B3 dropdown → re-render
 *                       Grid cell   → write-back Subject/Teacher to Master_Schedule
 *   Teacher_View      — B3 dropdown → re-render
 *                       Grid cell   → write-back assignment to Master_Schedule
 *   Teacher_Day_View  — B3/E3 selectors → re-render
 *                       Grid cell        → write-back assignment to Master_Schedule
 *   Master_Schedule   — any cell edit → cascade refresh to all views
 *
 * Delimiter convention (enforced in ScheduleParser):
 *   Teachers / Subjects : "/" separator  (e.g. "Mrs. X / Mr. Y")
 *   Classes             : "," separator  (e.g. "11th A, 11th B")
 */

function onEdit(e) {
  if (!e) return;
  const sheetName = e.source.getActiveSheet().getName();

  if (sheetName === 'Class_View') {
    handleClassViewEdit(e);
  } else if (sheetName === 'Teacher_View') {
    handleTeacherViewEdit(e);
  } else if (sheetName === 'Teacher_Day_View') {
    handleTeacherDayViewEdit(e);
  } else if (sheetName === 'Master_Schedule') {
    handleMasterScheduleEdit(e);
  }
}

// ─────────────────────────────────────────────────────────
// CLASS VIEW
// ─────────────────────────────────────────────────────────

function handleClassViewEdit(e) {
  const range = e.range;
  const sheet = e.source.getActiveSheet();

  // B3 dropdown → re-render for new class
  if (range.getRow() === 3 && range.getColumn() === 2) {
    const className = range.getValue();
    if (className) ClassViewManager.renderClassView(className);
    return;
  }

  // Grid area: rows 6+ (row 5 is header), cols 2–9
  // Layout: row 5 = header, rows 6+ alternate Subject/Teacher per day
  if (range.getRow() >= 6 && range.getColumn() >= 2 && range.getColumn() <= 9) {
    const className = sheet.getRange('B3').getValue();
    if (!className) return;

    const row    = range.getRow();
    const col    = range.getColumn();
    const period = col - 1; // Col 2 = Period 1

    const days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
    // Rows 6–7 = Monday (subject/teacher), 8–9 = Tuesday, etc.
    const dayIndex  = Math.floor((row - 6) / 2);
    if (dayIndex < 0 || dayIndex >= days.length) return;

    const day       = days[dayIndex];
    const isSubject = ((row - 6) % 2 === 0); // even offset = subject row
    const editType  = isSubject ? 'Subject' : 'Teacher';
    const newValue  = range.getValue();

    // Write-back to Master_Schedule
    ClassViewManager.updateMasterFromClassView(className, day, period, editType, newValue);

    // Cascade: programmatic writes to Master_Schedule do NOT re-fire onEdit,
    // so we must manually refresh every other affected view here.
    const tvSheet = e.source.getSheetByName('Teacher_View');
    if (tvSheet) {
      const activeTeacher = tvSheet.getRange('B3').getValue();
      if (activeTeacher) TeacherViewManager.renderTeacherView(activeTeacher);
    }
    const tdvSheet = e.source.getSheetByName('Teacher_Day_View');
    if (tdvSheet) {
      const activeDay = tdvSheet.getRange('B3').getValue() || 'Monday';
      const activeClassFilter = tdvSheet.getRange('E3').getValue() || 'All Classes';
      TeacherDayViewManager.renderTeacherDayView(activeDay, activeClassFilter);
    }
    refreshMasterGridData_();
  }
}

// ─────────────────────────────────────────────────────────
// TEACHER VIEW
// ─────────────────────────────────────────────────────────

function handleTeacherViewEdit(e) {
  const range = e.range;
  const sheet = e.source.getActiveSheet();
  const row   = range.getRow();
  const col   = range.getColumn();

  // B3 dropdown → re-render for new teacher
  if (row === 3 && col === 2) {
    const teacherName = range.getValue();
    if (teacherName) TeacherViewManager.renderTeacherView(teacherName);
    return;
  }

  // Grid area: rows 6–11 (row 5 is header), cols 2–9 (Period 1–8)
  // Layout: 1 row per day — Mon=6, Tue=7, Wed=8, Thu=9, Fri=10, Sat=11
  if (row >= 6 && col >= 2 && col <= 9) {
    const teacherName = sheet.getRange('B3').getValue();
    if (!teacherName) return;

    const days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
    const dayIndex = row - 6;
    if (dayIndex < 0 || dayIndex >= days.length) return;

    const day    = days[dayIndex];
    const period = col - 1; // Col 2 = Period 1
    const newValue = range.getValue();

    // Write-back to Master_Schedule
    TeacherViewManager.updateMasterFromTeacherView(teacherName, day, period, newValue);

    // Cascade: programmatic writes to Master_Schedule do NOT re-fire onEdit
    TeacherViewManager.renderTeacherView(teacherName);

    const cvSheet = e.source.getSheetByName('Class_View');
    if (cvSheet) {
      const activeClass = cvSheet.getRange('B3').getValue();
      if (activeClass) ClassViewManager.renderClassView(activeClass);
    }
    const tdvSheet = e.source.getSheetByName('Teacher_Day_View');
    if (tdvSheet) {
      const activeDay         = tdvSheet.getRange('B3').getValue() || 'Monday';
      const activeClassFilter = tdvSheet.getRange('E3').getValue() || 'All Classes';
      TeacherDayViewManager.renderTeacherDayView(activeDay, activeClassFilter);
    }
    refreshMasterGridData_();
  }
}

// ─────────────────────────────────────────────────────────
// TEACHER DAY VIEW
// ─────────────────────────────────────────────────────────

function handleTeacherDayViewEdit(e) {
  const range = e.range;
  const sheet = e.source.getActiveSheet();
  const row = range.getRow();
  const col = range.getColumn();

  // Row 3 edits (B3 = Day selector, E3 = Class filter)
  if (row === 3) {
    const day = sheet.getRange('B3').getValue() || 'Monday';
    const filterClass = sheet.getRange('E3').getValue() || 'All Classes';
    TeacherDayViewManager.renderTeacherDayView(day, filterClass);
    return;
  }

  // Grid edits: Row 6+ (Row 5 is header), Cols 2–9 (Period 1 to 8)
  if (row >= 6 && col >= 2 && col <= 9) {
    const selectedDay = sheet.getRange('B3').getValue() || 'Monday';
    const teacherName = sheet.getRange(row, 1).getValue();
    if (!teacherName) return;

    const period = col - 1; // Col 2 = Period 1
    const newValue = range.getValue();

    // Write-back to Master_Schedule
    TeacherDayViewManager.updateMasterFromTeacherDayView(selectedDay, teacherName, period, newValue);

    // Cascade: programmatic writes to Master_Schedule do NOT re-fire onEdit,
    // so we must manually refresh every other affected view here.
    const activeClassFilter = sheet.getRange('E3').getValue() || 'All Classes';
    TeacherDayViewManager.renderTeacherDayView(selectedDay, activeClassFilter);

    const cvSheet = e.source.getSheetByName('Class_View');
    if (cvSheet) {
      const activeClass = cvSheet.getRange('B3').getValue();
      if (activeClass) ClassViewManager.renderClassView(activeClass);
    }
    const tvSheet = e.source.getSheetByName('Teacher_View');
    if (tvSheet) {
      const activeTeacher = tvSheet.getRange('B3').getValue();
      if (activeTeacher) TeacherViewManager.renderTeacherView(activeTeacher);
    }
    refreshMasterGridData_();
  }
}

// ─────────────────────────────────────────────────────────
// MASTER SCHEDULE → cascade refresh
// ─────────────────────────────────────────────────────────

function handleMasterScheduleEdit(e) {
  const sheet = e.source;
  const range = e.range;

  if (range.getRow() < 2) return; // ignore header edits

  const classEdited   = sheet.getSheetByName('Master_Schedule')
                             .getRange(range.getRow(), 3).getValue();
  const teacherEdited = sheet.getSheetByName('Master_Schedule')
                             .getRange(range.getRow(), 6).getValue();

  // 1. Refresh Class_View if it's showing the edited class
  const classViewSheet = sheet.getSheetByName('Class_View');
  if (classViewSheet) {
    const activeClass = classViewSheet.getRange('B3').getValue();
    if (activeClass && activeClass === classEdited) {
      ClassViewManager.renderClassView(activeClass);
    }
  }

  // 2. Refresh Teacher_View if it's showing the edited teacher
  const teacherViewSheet = sheet.getSheetByName('Teacher_View');
  if (teacherViewSheet) {
    const activeTeacher = teacherViewSheet.getRange('B3').getValue();
    if (activeTeacher && (activeTeacher === teacherEdited || !teacherEdited)) {
      TeacherViewManager.renderTeacherView(activeTeacher);
    }
  }

  // 3. Refresh Teacher_Day_View if active
  const teacherDayViewSheet = sheet.getSheetByName('Teacher_Day_View');
  if (teacherDayViewSheet) {
    const activeDay = teacherDayViewSheet.getRange('B3').getValue() || 'Monday';
    const activeClassFilter = teacherDayViewSheet.getRange('E3').getValue() || 'All Classes';
    TeacherDayViewManager.renderTeacherDayView(activeDay, activeClassFilter);
  }

  // 4. Refresh Master_Grid_View data (lightweight — no re-styling)
  refreshMasterGridData_();

  // NOTE: Teacher_Day_View re-render above already covers the Master_Schedule edit path.
  // No additional cascade needed here since this handler is only triggered by direct
  // user edits on Master_Schedule (not by programmatic writes from other views).
}
