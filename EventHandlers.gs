/**
 * EventHandlers.gs
 * Centralizes all trigger events (onEdit, onOpen callbacks).
 *
 * Handled sheets:
 *   Class_View   — B3 dropdown change → re-render class view
 *                  Grid cell edit     → write-back to Master_Schedule
 *   Teacher_View — B3 dropdown change → re-render teacher view
 *   Master_Schedule — any data edit   → refresh Class_View, Teacher_View, Master_Grid_View
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

    ClassViewManager.updateMasterFromClassView(className, day, period, editType, newValue);
  }
}

// ─────────────────────────────────────────────────────────
// TEACHER VIEW
// ─────────────────────────────────────────────────────────

function handleTeacherViewEdit(e) {
  const range = e.range;

  // B3 dropdown → re-render for new teacher
  if (range.getRow() === 3 && range.getColumn() === 2) {
    const teacherName = range.getValue();
    if (teacherName) TeacherViewManager.renderTeacherView(teacherName);
  }
}

// ─────────────────────────────────────────────────────────
// TEACHER DAY VIEW
// ─────────────────────────────────────────────────────────

function handleTeacherDayViewEdit(e) {
  const range = e.range;

  // B3 dropdown → re-render for new day
  if (range.getRow() === 3 && range.getColumn() === 2) {
    const day = range.getValue();
    if (day) TeacherDayViewManager.renderTeacherDayView(day);
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
    TeacherDayViewManager.renderTeacherDayView(activeDay);
  }

  // 4. Refresh Master_Grid_View data (lightweight — no re-styling)
  refreshMasterGridData_();
}
