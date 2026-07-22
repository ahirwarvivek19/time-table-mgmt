/**
 * @OnlyCurrentDoc
 * School Timetable Management System
 * Entry Point and UI
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Timetable System')
    .addItem('📋 Open Multi-Select Picker Sidebar', 'openMultiSelectSidebar')
    .addItem('🔄 Refresh All Timetable Views & Tabs', 'refreshAllViews')
    .addSeparator()
    .addItem('Check For Conflicts', 'runValidation')
    .addItem('Generate Master Grid View', 'generateMasterGrid')
    .addSeparator()
    .addItem('Initialize / Refresh Class View', 'initClassView')
    .addItem('Initialize / Refresh Teacher View', 'initTeacherView')
    .addItem('Initialize / Refresh Day-wise Teacher View', 'initTeacherDayView')
    .addSeparator()
    .addItem('View Teacher Free Slots', 'generateTeacherAvailabilityGrid')
    .addItem('Launch Cover Manager UI', 'openCoverManagerUI')
    .addSeparator()
    .addItem('Setup Spreadsheets (Run Once)', 'setupInitialSpreadsheet')
    .addItem('Apply Master Schedule Dropdowns', 'applyMasterScheduleDropdowns')
    .addItem('Apply Global Styling', 'styleEntireSheet')
    .addItem('Import Timetable Data', 'importExcelData')
    .addToUi();
}

/**
 * Refreshes all timetable views and dashboards (Class View, Teacher View, Teacher Day View, Master Grid View).
 * Re-applies validation dropdowns and styling across the entire workbook.
 */
function refreshAllViews() {
  // 1. Re-apply Master Schedule Dropdowns
  applyMasterScheduleDropdowns_(/* silent= */ true);

  // 2. Refresh Class View
  const classViewSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Class_View');
  let activeClass = '';
  if (classViewSheet && classViewSheet.getLastRow() >= 3) {
    activeClass = classViewSheet.getRange('B3').getValue();
  }
  const classesData = DataAccess.getSheetDataAsObjects('Classes');
  if (!activeClass && classesData.length > 0) activeClass = classesData[0]['Class Name'];
  if (activeClass) ClassViewManager.renderClassView(activeClass);

  // 3. Refresh Teacher View
  const teacherViewSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Teacher_View');
  let activeTeacher = '';
  if (teacherViewSheet && teacherViewSheet.getLastRow() >= 3) {
    activeTeacher = teacherViewSheet.getRange('B3').getValue();
  }
  const teachersData = DataAccess.getSheetDataAsObjects('Teachers');
  if (!activeTeacher && teachersData.length > 0) activeTeacher = teachersData[0]['Teacher Name'];
  if (activeTeacher) TeacherViewManager.renderTeacherView(activeTeacher);

  // 4. Refresh Day-wise Teacher View
  const teacherDayViewSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Teacher_Day_View');
  let activeDay = 'Monday';
  let activeClassFilter = 'All Classes';
  if (teacherDayViewSheet && teacherDayViewSheet.getLastRow() >= 3) {
    activeDay = teacherDayViewSheet.getRange('B3').getValue() || 'Monday';
    activeClassFilter = teacherDayViewSheet.getRange('E3').getValue() || 'All Classes';
  }
  TeacherDayViewManager.renderTeacherDayView(activeDay, activeClassFilter);

  // 5. Refresh Master Grid View Data
  refreshMasterGridData_();

  // 6. Style entire sheet
  styleEntireSheet();

  SpreadsheetApp.getUi().alert('All Timetable Views & Tabs refreshed successfully!');
}

/**
 * Initializes or Refreshes the Class View with the first available class.
 */
function initClassView() {
  const classesData = DataAccess.getSheetDataAsObjects('Classes');
  if (classesData.length > 0 && classesData[0]['Class Name']) {
    ClassViewManager.renderClassView(classesData[0]['Class Name']);
    SpreadsheetApp.getUi().alert('Class View initialized successfully!');
  } else {
    SpreadsheetApp.getUi().alert('No classes found in the Classes tab.');
  }
}

/**
 * Initializes or Refreshes the Teacher View with the first available teacher.
 */
function initTeacherView() {
  const teachersData = DataAccess.getSheetDataAsObjects('Teachers');
  if (teachersData.length > 0 && teachersData[0]['Teacher Name']) {
    TeacherViewManager.renderTeacherView(teachersData[0]['Teacher Name']);
    SpreadsheetApp.getUi().alert('Teacher View initialized successfully!');
  } else {
    SpreadsheetApp.getUi().alert('No teachers found in the Teachers tab.');
  }
}

/**
 * Initializes or Refreshes the Day-wise Teacher Schedule View for Monday.
 */
function initTeacherDayView() {
  TeacherDayViewManager.renderTeacherDayView('Monday');
  SpreadsheetApp.getUi().alert('Day-wise Teacher View initialized successfully!');
}

/**
 * Applies global frontend styling to every single tab in the spreadsheet.
 * Uses batch setBackgrounds() calls instead of per-row loops to stay within
 * Apps Script quota limits.
 */
function styleEntireSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  // Sheets that have their own custom styling & layouts — skip default table styling for these
  const customSheets = ['Master_Grid_View', 'Class_View', 'Teacher_View', 'Teacher_Day_View', 'Teacher_Availability_Grid'];

  sheets.forEach(sheet => {
    sheet.setHiddenGridlines(true);

    // Skip custom view dashboards so their custom banners, frozen rows & merges are preserved
    if (customSheets.includes(sheet.getName())) {
      return;
    }

    const lastCol = sheet.getLastColumn();
    const lastRow = sheet.getLastRow();

    if (lastCol > 0) {
      // Header Row
      sheet.getRange(1, 1, 1, lastCol)
        .setFontWeight('bold')
        .setFontColor('#FFFFFF')
        .setBackground('#2C3E50')
        .setFontFamily('Montserrat');
      sheet.setFrozenRows(1);
    }

    if (lastRow > 1 && lastCol > 0) {
      const bodyRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
      bodyRange.setFontFamily('Montserrat');

      // Build 2D colour array — one batch call instead of per-row API calls
      const bgColors = [];
      for (let i = 0; i < lastRow - 1; i++) {
        bgColors.push(Array(lastCol).fill(i % 2 === 0 ? '#F8F9FA' : '#FFFFFF'));
      }
      bodyRange.setBackgrounds(bgColors);
    }
  });

  SpreadsheetApp.getUi().alert('Global styling applied to all standard data tabs!');
}

/**
 * Creates all required tabs and headers if they don't already exist.
 * Safe to run multiple times — existing sheets are left untouched.
 */
function setupInitialSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const requiredSheets = [
    {
      name: 'Teachers',
      headers: ['Teacher Name', 'Subject Specialization', 'Max Hours / Week', 'Days Unavailable', 'Total Hours Scheduled']
    },
    {
      name: 'Subjects',
      headers: ['Subject Name']
    },
    {
      name: 'Classes',
      headers: ['Class Name', 'Academic Tier', 'Room Assigned']
    },
    {
      name: 'Rooms',
      headers: ['Room Name', 'Capacity', 'Specialized Type']
    },
    {
      name: 'Master_Schedule',
      headers: ['Day', 'Period', 'Class', 'Academic Tier', 'Subject', 'Teacher', 'Room', 'Clash Status']
    },
    {
      name: 'Class_View',
      headers: []   // rendered programmatically by ClassViewManager
    },
    {
      name: 'Teacher_View',
      headers: []   // rendered programmatically by TeacherViewManager
    },
    {
      name: 'Teacher_Day_View',
      headers: []   // rendered programmatically by TeacherDayViewManager
    },
    {
      name: 'Cover_Manager',
      headers: ['Date', 'Absent Teacher', 'Period', 'Class to Cover', 'Suggested Available Teachers', 'Assigned Cover Teacher']
    }
  ];

  requiredSheets.forEach(sheetInfo => {
    let sheet = ss.getSheetByName(sheetInfo.name);
    if (!sheet) {
      sheet = ss.insertSheet(sheetInfo.name);
    }
    if (sheetInfo.headers.length > 0) {
      sheet.getRange(1, 1, 1, sheetInfo.headers.length).setValues([sheetInfo.headers]);
      sheet.autoResizeColumns(1, sheetInfo.headers.length);
      for (let i = 1; i <= sheetInfo.headers.length; i++) {
        sheet.setColumnWidth(i, sheet.getColumnWidth(i) + 50);
      }
    }
  });

  styleEntireSheet();

  // Reorder sheets to the correct tab order
  reorderSheets_();

  // Apply dropdowns to Master_Schedule
  applyMasterScheduleDropdowns_(/* silent= */ true);

  // Initialize views with first available data
  const classesData = DataAccess.getSheetDataAsObjects('Classes');
  if (classesData.length > 0 && classesData[0]['Class Name']) {
    ClassViewManager.renderClassView(classesData[0]['Class Name']);
  }
  const teachersData = DataAccess.getSheetDataAsObjects('Teachers');
  if (teachersData.length > 0 && teachersData[0]['Teacher Name']) {
    TeacherViewManager.renderTeacherView(teachersData[0]['Teacher Name']);
  }
  TeacherDayViewManager.renderTeacherDayView('Monday');

  SpreadsheetApp.getUi().alert(
    'Setup Complete!\n' +
    'Tabs created: Teachers, Subjects, Classes, Rooms, Master_Schedule, Class_View, Teacher_View, Cover_Manager.\n\n' +
    'Next step: run \'Import Timetable Data\' from the menu to load all data.'
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// MASTER SCHEDULE DROPDOWNS
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Public menu entry-point — applies dropdowns and shows a confirmation alert.
 */
function applyMasterScheduleDropdowns() {
  applyMasterScheduleDropdowns_(/* silent= */ false);
}

/**
 * Applies data-validation dropdowns to the Subject (col E) and Teacher (col F)
 * columns of the Master_Schedule sheet, mirroring the Class_View behaviour.
 *
 * - Subjects sourced from the Subjects sheet (falls back to unique values already
 *   present in Master_Schedule).
 * - Teachers sourced from the Teachers sheet.
 * - Both rules use setAllowInvalid(true) so existing free-text values are kept.
 *
 * @param {boolean} silent  When true, no UI alert is shown (used during setup).
 */
function applyMasterScheduleDropdowns_(silent) {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const msSheet  = ss.getSheetByName('Master_Schedule');
  if (!msSheet) return;

  const lastRow = msSheet.getLastRow();
  if (lastRow < 2) {
    if (!silent) SpreadsheetApp.getUi().alert('Master Schedule has no data rows yet.');
    return;
  }

  const dataRowCount = lastRow - 1; // exclude header

  // ── Subjects list ─────────────────────────────────────────────────────────
  let subjectsList = DataAccess.getSheetDataAsObjects('Subjects')
                               .map(r => r['Subject Name']).filter(s => s);
  if (subjectsList.length === 0) {
    // Fallback: derive from existing values in col E (Subject)
    const colE = msSheet.getRange(2, 5, dataRowCount, 1).getValues();
    subjectsList = [...new Set(colE.flat().filter(v => v))].sort();
  }

  // ── Teachers list ─────────────────────────────────────────────────────────
  const teachersList = DataAccess.getSheetDataAsObjects('Teachers')
                                 .map(t => t['Teacher Name']).filter(t => t);

  // ── Build validation rules ────────────────────────────────────────────────
  const subjectRule = subjectsList.length > 0
    ? SpreadsheetApp.newDataValidation()
        .requireValueInList(subjectsList, true)
        .setAllowInvalid(true)
        .build()
    : null;

  const teacherRule = teachersList.length > 0
    ? SpreadsheetApp.newDataValidation()
        .requireValueInList(teachersList, true)
        .setAllowInvalid(true)
        .build()
    : null;

  // ── Apply to columns E (Subject) and F (Teacher) ─────────────────────────
  const subjectRange = msSheet.getRange(2, 5, dataRowCount, 1); // col E
  const teacherRange = msSheet.getRange(2, 6, dataRowCount, 1); // col F

  if (subjectRule) {
    subjectRange.setDataValidation(subjectRule);
  } else {
    subjectRange.clearDataValidations();
  }

  if (teacherRule) {
    teacherRange.setDataValidation(teacherRule);
  } else {
    teacherRange.clearDataValidations();
  }

  if (!silent) {
    SpreadsheetApp.getUi().alert(
      'Dropdowns applied!\n' +
      'Subject column: ' + (subjectsList.length > 0 ? subjectsList.length + ' options' : 'no list found') + '\n' +
      'Teacher column: ' + (teachersList.length > 0 ? teachersList.length + ' options' : 'no list found')
    );
  }
}

/**
 * Reorders the spreadsheet tabs to the canonical display order.
 * Called automatically by setupInitialSpreadsheet.
 */
function reorderSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const desiredOrder = [
    'Teachers', 'Subjects', 'Classes', 'Rooms',
    'Master_Schedule', 'Master_Grid_View', 'Class_View', 'Teacher_View', 'Teacher_Day_View', 'Cover_Manager',
    'Teacher_Availability_Grid'
  ];
  desiredOrder.forEach((name, idx) => {
    const sheet = ss.getSheetByName(name);
    if (sheet) {
      ss.setActiveSheet(sheet);
      ss.moveActiveSheet(idx + 1);
    }
  });
}
