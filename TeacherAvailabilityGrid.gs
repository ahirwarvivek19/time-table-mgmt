/**
 * TeacherAvailabilityGrid.gs
 * Generates a 2D grid showing free vs busy slots for all teachers on a specific day.
 */

function generateTeacherAvailabilityGrid() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Teacher Availability Grid',
    'Which day do you want to check? (e.g., Monday, Tuesday)',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const day = response.getResponseText().trim();
  if (!day) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let gridSheet = ss.getSheetByName('Teacher_Availability_Grid');
  
  if (!gridSheet) {
    gridSheet = ss.insertSheet('Teacher_Availability_Grid');
  }
  
  gridSheet.clear();
  
  // 1. Get Data
  const teachers = DataAccess.getSheetDataAsObjects('Teachers').map(t => t['Teacher Name']).filter(t => t).sort();
  const scheduleData = DataAccess.getSheetDataAsObjects('Master_Schedule').filter(row => row.Day === day);
  const periods = [1, 2, 3, 4, 5, 6, 7, 8];
  
  if (teachers.length === 0) {
    ui.alert('No teachers found in the Teachers tab.');
    return;
  }

  // 2. Build Lookup
  // lookup[teacher][period] = "Busy (Class Name)" or "FREE"
  const lookup = {};
  teachers.forEach(t => {
    lookup[t] = {};
    periods.forEach(p => lookup[t][p] = 'FREE'); // Default FREE
  });
  
  scheduleData.forEach(row => {
    if (row.Teacher && lookup[row.Teacher]) {
      lookup[row.Teacher][row.Period] = `BUSY (${row.Class})`;
    }
  });

  // 3. Build Output Array
  const gridOutput = [];
  const headers = [`Teacher (${day})`, 'Period 1', 'Period 2', 'Period 3', 'Period 4', 'Period 5', 'Period 6', 'Period 7', 'Period 8'];
  gridOutput.push(headers);
  
  teachers.forEach(t => {
    const rowOutput = [t];
    periods.forEach(p => {
      rowOutput.push(lookup[t][p]);
    });
    gridOutput.push(rowOutput);
  });

  // 4. Write Data
  const numRows = gridOutput.length;
  const numCols = headers.length;
  const dataRange = gridSheet.getRange(1, 1, numRows, numCols);
  dataRange.setValues(gridOutput);

  // 5. Styling
  gridSheet.setHiddenGridlines(true);
  gridSheet.setFrozenRows(1);
  gridSheet.setFrozenColumns(1);
  
  dataRange.setFontFamily('Montserrat').setVerticalAlignment('middle').setHorizontalAlignment('center');
  
  // Header
  gridSheet.getRange(1, 1, 1, numCols)
           .setBackground('#2C3E50')
           .setFontColor('#FFFFFF')
           .setFontWeight('bold');

  // Body Colors (Green for FREE, Light Gray for BUSY)
  for (let r = 2; r <= numRows; r++) {
    for (let c = 2; c <= numCols; c++) {
      const cellValue = gridOutput[r-1][c-1];
      const cell = gridSheet.getRange(r, c);
      if (cellValue === 'FREE') {
        cell.setBackground('#E8F8F5').setFontColor('#27AE60').setFontWeight('bold'); // Soft Green
      } else {
        cell.setBackground('#F8F9FA').setFontColor('#7F8C8D'); // Gray out busy slots
      }
    }
  }

  // Teacher Column
  gridSheet.getRange(2, 1, numRows - 1, 1).setBackground('#34495E').setFontColor('#FFFFFF').setFontWeight('bold');

  // Borders & Resize
  dataRange.setBorder(true, true, true, true, true, true, '#BDC3C7', SpreadsheetApp.BorderStyle.SOLID);
  gridSheet.setColumnWidth(1, 180);
  for(let i=2; i<=numCols; i++) {
    gridSheet.setColumnWidth(i, 120);
  }
  
  ss.setActiveSheet(gridSheet);
}
