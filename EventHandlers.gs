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

  // 1. Multi-Select dropdown handling for Subject & Teacher fields
  handleMultiSelectDropdown(e);

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

/**
 * Enables multi-select dropdown functionality for Subject and Teacher cells.
 * Selecting a new dropdown item appends it with ' / '.
 * Selecting an existing item toggles it off.
 *
 * Supported sheets:
 * - Master_Schedule: Col 5 (Subject), Col 6 (Teacher)
 * - Class_View: Rows 6+, Cols 2-9 (Subject and Teacher rows)
 *
 * @param {Object} e Event object from onEdit
 */
function handleMultiSelectDropdown(e) {
  if (!e || !e.range || e.value === undefined || e.oldValue === undefined) return;

  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();
  const row = e.range.getRow();
  const col = e.range.getColumn();

  let isTargetCell = false;

  if (sheetName === 'Master_Schedule') {
    // Header is row 1. Data starts row 2+. Col 5 = Subject, Col 6 = Teacher.
    if (row >= 2 && (col === 5 || col === 6)) {
      isTargetCell = true;
    }
  } else if (sheetName === 'Class_View') {
    // Header is row 5. Data starts row 6+. Cols 2–9 = Periods 1–8.
    if (row >= 6 && col >= 2 && col <= 9) {
      isTargetCell = true;
    }
  }

  if (!isTargetCell) return;

  const newValue = String(e.value).trim();
  const oldValue = String(e.oldValue).trim();

  if (!newValue || !oldValue) return;

  // Split existing values by / or comma
  const existingItems = oldValue.split(/[\/\,]/).map(s => s.trim()).filter(Boolean);

  let updatedItems = [];
  if (existingItems.includes(newValue)) {
    // Toggle OFF: remove the selected item
    updatedItems = existingItems.filter(item => item !== newValue);
  } else {
    // Toggle ON: append the selected item
    updatedItems = [...existingItems, newValue];
  }

  const finalValue = updatedItems.join(' / ');
  e.range.setValue(finalValue);
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

  // Row 3 edits (B3 = Day selector, E3 = Class filter)
  if (range.getRow() === 3) {
    const sheet = e.source.getActiveSheet();
    const day = sheet.getRange('B3').getValue() || 'Monday';
    const filterClass = sheet.getRange('E3').getValue() || 'All Classes';
    TeacherDayViewManager.renderTeacherDayView(day, filterClass);
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
}
