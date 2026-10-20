/**
 * @OnlyCurrentDoc
 * School Timetable Management System
 * Entry Point and UI
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  const setupMenu = ui.createMenu('⚙️ Setup & Utilities')
    .addItem('Setup Spreadsheets (Run Once)', 'setupInitialSpreadsheet')
    .addItem('Generate Master Grid (Full Restyle)', 'generateMasterGrid')
    .addSeparator()
    .addItem('Apply Master Schedule Dropdowns', 'applyMasterScheduleDropdowns')
    .addItem('Apply Global Styling', 'styleEntireSheet')
    .addItem('Reorder Spreadsheet Tabs', 'reorderSheets')
    .addItem('Format Teacher Names (Title Case)', 'formatAllTeacherNames')
    .addSeparator()
    .addItem('Install Edit Sync Trigger', 'installEditTrigger');

  const importMenu = ui.createMenu('📥 Data Import')
    .addItem('Import All Pre-fill Data', 'importExcelData')
    .addItem('Import Static Teacher Registry', 'importTeachersData')
    .addItem('Import Dynamic Schedule Data', 'importScheduleData')
    .addSeparator()
    .addItem('Save Current Sheet as Pre-fill Data', 'saveCurrentSheetAsPrefillData')
    .addItem('Reset Pre-fill Data to Default', 'resetPrefillDataToDefault');

  ui.createMenu('Timetable System')
    .addItem('🔄 Refresh All Views', 'refreshAllViews')
    .addItem('✅ Check For Conflicts', 'runValidation')
    .addItem('📊 Teacher Availability', 'generateTeacherAvailabilityGrid')
    .addItem('👥 Cover Manager', 'openCoverManagerUI')
    .addSeparator()
    .addSubMenu(setupMenu)
    .addSubMenu(importMenu)
    .addToUi();
}

/**
 * Refreshes all timetable views and dashboards (Class View, Teacher View, Teacher Day View, Master Grid View).
 * Re-applies validation dropdowns and styling across the entire workbook.
 * @param {boolean} [silent=false] When true, suppresses the completion UI alert.
 */
function refreshAllViews(silent) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const S = CONFIG.SHEETS;

  // 1. Re-apply Master Schedule Dropdowns
  applyMasterScheduleDropdowns_(/* silent= */ true);

  // 2. Refresh Class View (keep the selected class if it still exists)
  const classViewSheet = ss.getSheetByName(S.CLASS_VIEW);
  const classNames = getClassNames_();
  let activeClass = classViewSheet && classViewSheet.getLastRow() >= 3 ? String(classViewSheet.getRange('B3').getValue() || '') : '';
  if (!activeClass && classNames.length > 0) activeClass = classNames[0];
  if (activeClass) ClassViewManager.renderClassView(activeClass);

  // 3. Refresh Teacher View (falls back to the first teacher if the selected one was removed)
  const teacherViewSheet = ss.getSheetByName(S.TEACHER_VIEW);
  const activeTeacher = teacherViewSheet && teacherViewSheet.getLastRow() >= 3 ? teacherViewSheet.getRange('B3').getValue() : '';
  if (getTeacherNames_().length > 0) TeacherViewManager.renderTeacherView(activeTeacher);

  // 4. Refresh Day-wise Teacher View
  const teacherDayViewSheet = ss.getSheetByName(S.TEACHER_DAY_VIEW);
  let activeDay = CONFIG.DAYS[0];
  let activeClassFilter = CONFIG.ALL_CLASSES;
  if (teacherDayViewSheet && teacherDayViewSheet.getLastRow() >= 3) {
    activeDay = teacherDayViewSheet.getRange('B3').getValue() || CONFIG.DAYS[0];
    activeClassFilter = teacherDayViewSheet.getRange('E3').getValue() || CONFIG.ALL_CLASSES;
  }
  TeacherDayViewManager.renderTeacherDayView(activeDay, activeClassFilter);

  // 4b. Teachers → "Total Hours Scheduled" (periods per week)
  DataAccess.updateTeacherLoad();

  // 4c. Lets the optional read-only web page (WebApp.gs) find this spreadsheet
  rememberSpreadsheetId_();

  // 5. Refresh Master Grid View Data
  refreshMasterGridData_();

  // 6. Reorder tabs (Teachers, Subjects, Classes, Master_Schedule, then all Views)
  reorderSheets();

  // 7. Style entire sheet silently
  styleEntireSheet(/* silent= */ true);

  if (!silent) {
    SpreadsheetApp.getUi().alert('All Timetable Views & Tabs refreshed successfully!');
  }
}

/**
 * Initializes or Refreshes the Class View with the first available class.
 */
function initClassView() {
  const classNames = getClassNames_();
  if (classNames.length > 0) {
    ClassViewManager.renderClassView(classNames[0]);
    SpreadsheetApp.getUi().alert('Class View initialized successfully!');
  } else {
    SpreadsheetApp.getUi().alert('No classes found in the Classes tab.');
  }
}

/**
 * Initializes or Refreshes the Teacher View with the first available teacher.
 */
function initTeacherView() {
  const teacherNames = getTeacherNames_();
  if (teacherNames.length > 0) {
    TeacherViewManager.renderTeacherView(teacherNames[0]);
    SpreadsheetApp.getUi().alert('Teacher View initialized successfully!');
  } else {
    SpreadsheetApp.getUi().alert('No teachers found in the Teachers tab.');
  }
}

/**
 * Initializes or Refreshes the Day-wise Teacher Schedule View for Monday.
 */
function initTeacherDayView() {
  TeacherDayViewManager.renderTeacherDayView(CONFIG.DAYS[0]);
  SpreadsheetApp.getUi().alert('Day-wise Teacher View initialized successfully!');
}

/**
 * Applies global frontend styling to every single tab in the spreadsheet.
 * Uses batch setBackgrounds() calls instead of per-row loops to stay within
 * Apps Script quota limits.
 * @param {boolean} [silent=false] When true, suppresses the completion UI alert.
 */
function styleEntireSheet(silent) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  // Sheets that have their own custom styling & layouts — skip default table styling for these
  const S = CONFIG.SHEETS;
  const customSheets = [S.MASTER_GRID, S.CLASS_VIEW, S.TEACHER_VIEW, S.TEACHER_DAY_VIEW, S.AVAILABILITY];

  sheets.forEach(sheet => {
    sheet.setHiddenGridlines(true);

    // Skip custom view dashboards & internal hidden sheets so their custom banners, frozen rows & merges are preserved
    if (customSheets.includes(sheet.getName()) || sheet.getName().startsWith('_')) {
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
        .setFontFamily(CONFIG.FONT);
      sheet.setFrozenRows(1);
    }

    if (lastRow > 1 && lastCol > 0) {
      const bodyRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
      bodyRange.setFontFamily(CONFIG.FONT);

      // Build 2D colour array — one batch call instead of per-row API calls
      const bgColors = [];
      for (let i = 0; i < lastRow - 1; i++) {
        bgColors.push(Array(lastCol).fill(i % 2 === 0 ? '#F8F9FA' : '#FFFFFF'));
      }
      bodyRange.setBackgrounds(bgColors);
    }
  });

  if (!silent) {
    SpreadsheetApp.getUi().alert('Global styling applied to all standard data tabs!');
  }
}

/**
 * Creates all required tabs and headers if they don't already exist.
 * Safe to run multiple times — existing sheets are left untouched.
 */
function setupInitialSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const S = CONFIG.SHEETS;
  const requiredSheets = [
    {
      name: S.TEACHERS,
      headers: ['Teacher Name', 'Subject Specialization', 'Max Hours / Week', 'Days Unavailable', 'Total Hours Scheduled']
    },
    {
      name: S.SUBJECTS,
      headers: ['Subject Name']
    },
    {
      name: S.CLASSES,
      headers: ['Class Name', 'Academic Tier', 'Room Assigned']
    },
    {
      name: S.SCHEDULE,
      headers: CONFIG.SCHEDULE_HEADERS
    },
    {
      name: S.CLASS_VIEW,
      headers: []   // rendered programmatically by ClassViewManager
    },
    {
      name: S.TEACHER_VIEW,
      headers: []   // rendered programmatically by TeacherViewManager
    },
    {
      name: S.TEACHER_DAY_VIEW,
      headers: []   // rendered programmatically by TeacherDayViewManager
    },
    {
      name: S.COVER,
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

  styleEntireSheet(/* silent= */ true);

  // Reorder sheets to the correct tab order
  reorderSheets();

  // Apply dropdowns to Master_Schedule
  applyMasterScheduleDropdowns_(/* silent= */ true);

  // Initialize views with first available data
  const classNames = getClassNames_();
  if (classNames.length > 0) ClassViewManager.renderClassView(classNames[0]);
  const teacherNames = getTeacherNames_();
  if (teacherNames.length > 0) TeacherViewManager.renderTeacherView(teacherNames[0]);
  TeacherDayViewManager.renderTeacherDayView(CONFIG.DAYS[0]);

  rememberSpreadsheetId_();

  // Installable edit trigger (no 30-second limit, processes each edit once)
  let triggerNote = 'Edit sync trigger installed.';
  try {
    installEditTrigger_();
  } catch (err) {
    triggerNote = 'Could not install the edit sync trigger (' + err.message + '). Run Setup & Utilities → Install Edit Sync Trigger.';
  }

  SpreadsheetApp.getUi().alert(
    'Setup Complete!\n' +
    'Tabs created: Teachers, Subjects, Classes, Master_Schedule, Class_View, Teacher_View, Teacher_Day_View, Cover_Manager.\n' +
    triggerNote + '\n\n' +
    'Next step: run Data Import → Import All Pre-fill Data to load the timetable.'
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
  const msSheet  = ss.getSheetByName(CONFIG.SHEETS.SCHEDULE);
  if (!msSheet) return;

  const lastRow = msSheet.getLastRow();
  if (lastRow < 2) {
    if (!silent) SpreadsheetApp.getUi().alert('Master Schedule has no data rows yet.');
    return;
  }

  const dataRowCount = lastRow - 1; // exclude header

  // ── Subjects list ─────────────────────────────────────────────────────────
  let subjectsList = DataAccess.getSheetDataAsObjects(CONFIG.SHEETS.SUBJECTS)
                               .map(r => r['Subject Name']).filter(s => s);
  if (subjectsList.length === 0) {
    // Fallback: derive from existing values in col E (Subject)
    const colE = msSheet.getRange(2, 5, dataRowCount, 1).getValues();
    subjectsList = [...new Set(colE.flat().filter(v => v))].sort();
  }

  // ── Teachers list ─────────────────────────────────────────────────────────
  const teachersList = getTeacherNames_();

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
 * Reorders the spreadsheet tabs: Teachers, Subjects, Classes, Master_Schedule, followed by all Views.
 * Public menu entry point.
 */
function reorderSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const S = CONFIG.SHEETS;
  const desiredOrder = [
    S.TEACHERS, S.SUBJECTS, S.CLASSES, S.SCHEDULE, S.CLASS_VIEW, S.TEACHER_VIEW,
    S.TEACHER_DAY_VIEW, S.MASTER_GRID, S.AVAILABILITY, S.COVER
  ];

  // moveActiveSheet() needs each sheet activated; remember where the user was
  // and return them there (previously every refresh left them on the last tab).
  const original = ss.getActiveSheet();
  let currentPos = 1;
  desiredOrder.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    const atPos = ss.getSheets()[currentPos - 1];
    if (!atPos || atPos.getName() !== name) {
      ss.setActiveSheet(sheet);
      ss.moveActiveSheet(currentPos);
    }
    currentPos++;
  });
  if (original) ss.setActiveSheet(original);
}
