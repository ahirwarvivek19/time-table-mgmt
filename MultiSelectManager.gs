/**
 * MultiSelectManager.gs
 * Backend controller for the Multi-Select Sidebar Picker UI.
 * Allows admins to check multiple subjects or teachers using checkboxes and apply to active cell.
 */

function openMultiSelectSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('MultiSelectUI')
    .setTitle('Multi-Select Picker')
    .setWidth(320);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Returns options and current selections for the active cell.
 * @returns {Object}
 */
function apiGetPickerOptions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  const sheetName = activeSheet.getName();
  const activeRange = activeSheet.getActiveRange();

  if (!activeRange) return null;

  const row = activeRange.getRow();
  const col = activeRange.getColumn();
  const rawValue = String(activeRange.getValue() || '').trim();

  let fieldType = 'Subject'; // Default

  if (sheetName === 'Master_Schedule') {
    if (col === 6) fieldType = 'Teacher';
    else if (col === 5) fieldType = 'Subject';
  } else if (sheetName === 'Class_View') {
    if (row >= 6) {
      const isSubject = ((row - 6) % 2 === 0);
      fieldType = isSubject ? 'Subject' : 'Teacher';
    }
  } else if (sheetName === 'Teacher_Day_View') {
    if (col === 5 && row === 3) {
      // Class Filter
      const classesData = DataAccess.getSheetDataAsObjects('Classes');
      const classList = classesData.map(c => c['Class Name']).filter(c => c).sort();
      const currentSelections = ScheduleParser.splitClasses(rawValue);
      return {
        fieldType: 'Class Filter',
        sheetName: sheetName,
        cellAddress: activeRange.getA1Notation(),
        currentSelections: currentSelections,
        allOptions: classList
      };
    }
  }

  // Parse current selections
  const currentSelections = ScheduleParser.splitList(rawValue);

  // Fetch options list based on fieldType
  let allOptions = [];
  if (fieldType === 'Teacher') {
    const teachersData = DataAccess.getSheetDataAsObjects('Teachers');
    allOptions = teachersData.map(t => t['Teacher Name']).filter(t => t).sort();
  } else {
    // Subject options
    const subjectsData = DataAccess.getSheetDataAsObjects('Subjects');
    allOptions = subjectsData.map(s => s['Subject Name']).filter(s => s).sort();
    if (allOptions.length === 0) {
      const scheduleData = DataAccess.getSheetDataAsObjects('Master_Schedule');
      allOptions = [...new Set(scheduleData.map(r => r.Subject).filter(s => s))].sort();
    }
  }

  return {
    fieldType: fieldType,
    sheetName: sheetName,
    cellAddress: activeRange.getA1Notation(),
    currentSelections: currentSelections,
    allOptions: allOptions
  };
}

/**
 * Applies multi-selected items to the active spreadsheet cell and refreshes dashboards.
 * @param {Array<string>} selectedItems
 */
function apiApplyPickerSelections(selectedItems) {
  if (!Array.isArray(selectedItems)) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  const activeRange = activeSheet.getActiveRange();

  if (!activeRange) return;

  const sheetName = activeSheet.getName();
  const row = activeRange.getRow();
  const col = activeRange.getColumn();

  const formattedValue = selectedItems.join(' / ');
  activeRange.setValue(formattedValue);

  // Trigger cascade edits to Master_Schedule and views
  if (sheetName === 'Class_View') {
    if (row >= 6 && col >= 2 && col <= 9) {
      const className = activeSheet.getRange('B3').getValue();
      const period = col - 1;
      const days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
      const dayIndex = Math.floor((row - 6) / 2);
      const isSubject = ((row - 6) % 2 === 0);
      if (dayIndex >= 0 && dayIndex < days.length) {
        const day = days[dayIndex];
        const editType = isSubject ? 'Subject' : 'Teacher';
        ClassViewManager.updateMasterFromClassView(className, day, period, editType, formattedValue);
      }
    } else if (row === 3 && col === 2) {
      ClassViewManager.renderClassView(formattedValue);
    }
  } else if (sheetName === 'Teacher_Day_View') {
    const day = activeSheet.getRange('B3').getValue() || 'Monday';
    const filterClass = activeSheet.getRange('E3').getValue() || 'All Classes';
    TeacherDayViewManager.renderTeacherDayView(day, filterClass);
  } else if (sheetName === 'Teacher_View') {
    if (row === 3 && col === 2) {
      TeacherViewManager.renderTeacherView(formattedValue);
    }
  } else if (sheetName === 'Master_Schedule') {
    // Refresh associated views
    const classEdited = activeSheet.getRange(row, 3).getValue();
    const teacherEdited = activeSheet.getRange(row, 6).getValue();

    const classViewSheet = ss.getSheetByName('Class_View');
    if (classViewSheet && classViewSheet.getRange('B3').getValue() === classEdited) {
      ClassViewManager.renderClassView(classEdited);
    }

    const teacherViewSheet = ss.getSheetByName('Teacher_View');
    if (teacherViewSheet && teacherViewSheet.getRange('B3').getValue() === teacherEdited) {
      TeacherViewManager.renderTeacherView(teacherEdited);
    }

    const teacherDayViewSheet = ss.getSheetByName('Teacher_Day_View');
    if (teacherDayViewSheet) {
      const activeDay = teacherDayViewSheet.getRange('B3').getValue() || 'Monday';
      const activeClassFilter = teacherDayViewSheet.getRange('E3').getValue() || 'All Classes';
      TeacherDayViewManager.renderTeacherDayView(activeDay, activeClassFilter);
    }

    refreshMasterGridData_();
  }
}
