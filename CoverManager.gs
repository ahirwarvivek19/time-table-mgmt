/**
 * CoverManager.gs
 * Backend API for the Daily Arrangement / Substitution system.
 */

function openCoverManagerUI() {
  const html = HtmlService.createHtmlOutputFromFile('CoverManagerUI')
    .setWidth(700)
    .setHeight(600)
    .setTitle('Daily Arrangement Manager');
  SpreadsheetApp.getUi().showModalDialog(html, 'Daily Arrangement Manager');
}

/**
 * Returns a list of all active teachers.
 */
function apiGetTeachers() {
  const teachersData = DataAccess.getSheetDataAsObjects('Teachers');
  return teachersData.map(t => t['Teacher Name']).filter(t => t);
}

/**
 * Returns the schedule for an absent teacher on a specific day.
 */
function apiGetTeacherSchedule(teacherName, day) {
  const scheduleData = DataAccess.getSheetDataAsObjects('Master_Schedule');
  return scheduleData.filter(row => row.Day === day && ScheduleParser.rowIncludesTeacher(row, teacherName))
                     .sort((a, b) => parseInt(a.Period) - parseInt(b.Period));
}

/**
 * Returns a list of teachers who are NOT teaching during the specified day and period.
 */
function apiGetAvailableTeachers(day, period) {
  const allTeachers = apiGetTeachers();
  const scheduleData = DataAccess.getSheetDataAsObjects('Master_Schedule');
  
  // Find who IS teaching
  const busyTeachers = new Set();
  scheduleData.forEach(row => {
    if (row.Day === day && parseInt(row.Period) === parseInt(period)) {
      const assignments = ScheduleParser.parseRowAssignments(row);
      assignments.forEach(assign => {
        if (assign.teacher) busyTeachers.add(assign.teacher);
      });
    }
  });
  
  // Subtract busy from all
  const available = allTeachers.filter(t => !busyTeachers.has(t));
  return available.sort();
}

/**
 * Writes the cover assignment to the Cover_Manager sheet.
 */
function apiAssignCover(date, absentTeacher, period, classToCover, coverTeacher) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let coverSheet = ss.getSheetByName('Cover_Manager');
  
  if (!coverSheet) {
    throw new Error("Cover_Manager sheet not found. Run Setup first.");
  }
  
  coverSheet.appendRow([
    date,
    absentTeacher,
    period,
    classToCover,
    'System Assigned', // Replaced the suggestion list logging for cleaner UX
    coverTeacher
  ]);
  
  // Style the new row
  const lastRow = coverSheet.getLastRow();
  const rowRange = coverSheet.getRange(lastRow, 1, 1, 6);
  rowRange.setFontFamily('Montserrat');
  if (lastRow % 2 === 0) {
    rowRange.setBackground('#F8F9FA');
  } else {
    rowRange.setBackground('#FFFFFF');
  }
  rowRange.setBorder(true, true, true, true, true, true, '#BDC3C7', SpreadsheetApp.BorderStyle.SOLID);
  
  return true;
}
