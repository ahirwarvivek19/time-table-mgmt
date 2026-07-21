/**
 * TeacherViewManager.gs
 * Renders a read-only weekly schedule for a specific teacher on the 'Teacher_View' tab.
 * Shows FREE vs BUSY slots with class + subject information.
 */

const TeacherViewManager = {

  /**
   * Renders the full weekly grid for a teacher.
   * @param {string} teacherName
   */
  renderTeacherView: function(teacherName) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Teacher_View');
    if (!sheet) sheet = ss.insertSheet('Teacher_View');

    sheet.clear();
    sheet.getDataRange().breakApart();
    sheet.setFrozenRows(0);    // reset frozen state before any merges
    sheet.setFrozenColumns(0);
    sheet.setHiddenGridlines(true);

    const days    = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
    const periods = [1, 2, 3, 4, 5, 6, 7, 8];
    const numCols = 9;
    const startRow = 5;

    // ── 1. TITLE BANNER (Rows 1–2) ────────────────────────────────────────
    // Col 1 stays unmerged so setFrozenColumns(1) works without conflict.
    // Cols 2–9 are merged for the title text.
    sheet.getRange(1, 1, 2, numCols)
         .setBackground('#1A3A4A')
         .setFontFamily('Montserrat');
    sheet.getRange(1, 2, 2, numCols - 1)
         .merge()
         .setValue('📅  Teacher Schedule  ·  ' + (teacherName || 'Select a Teacher'))
         .setFontFamily('Montserrat')
         .setFontSize(16)
         .setFontWeight('bold')
         .setFontColor('#FFFFFF')
         .setHorizontalAlignment('center')
         .setVerticalAlignment('middle');
    sheet.setRowHeight(1, 28);
    sheet.setRowHeight(2, 28);

    // ── 2. TEACHER SELECTOR (Row 3) ───────────────────────────────────────
    const teachersData = DataAccess.getSheetDataAsObjects('Teachers');
    const teacherNames = teachersData.map(t => t['Teacher Name']).filter(t => t);

    sheet.getRange('A3')
         .setValue('Select Teacher:')
         .setFontWeight('bold')
         .setFontFamily('Montserrat')
         .setFontColor('#5D6D7E')
         .setHorizontalAlignment('right');

    const b3Cell = sheet.getRange('B3');
    if (teacherNames.length > 0) {
      b3Cell.setDataValidation(
        SpreadsheetApp.newDataValidation().requireValueInList(teacherNames, true).build()
      );
    }

    if (teacherName && teacherNames.includes(teacherName)) {
      b3Cell.setValue(teacherName);
    } else if (!teacherName && teacherNames.length > 0) {
      teacherName = teacherNames[0];
      b3Cell.setValue(teacherName);
    }

    sheet.getRange(3, 1, 1, numCols).setBackground('#F2F3F4').setFontFamily('Montserrat');
    sheet.setRowHeight(3, 36);

    // Spacer row 4
    sheet.getRange(4, 1, 1, numCols).setBackground('#E8EAED');
    sheet.setRowHeight(4, 6);

    // ── 3. FETCH SCHEDULE DATA ────────────────────────────────────────────
    const scheduleData = DataAccess.getSheetDataAsObjects('Master_Schedule');

    const lookup = {};
    days.forEach(d => {
      lookup[d] = {};
      periods.forEach(p => lookup[d][p] = null);
    });

    scheduleData.forEach(row => {
      if (row.Day && row.Period && lookup[row.Day] !== undefined) {
        if (ScheduleParser.rowIncludesTeacher(row, teacherName)) {
          const specificSubject = ScheduleParser.getSubjectForTeacher(row, teacherName);
          const periodInt = parseInt(row.Period);
          const existingSlot = lookup[row.Day][periodInt];
          const entryText = (row.Class || '') + '\n' + (specificSubject || row.Subject || '');
          
          if (!existingSlot) {
            lookup[row.Day][periodInt] = {
              cls: row.Class || '',
              subject: specificSubject || row.Subject || '',
              text: entryText
            };
          } else {
            lookup[row.Day][periodInt].text += '\n---\n' + entryText;
          }
        }
      }
    });

    // ── 4. BUILD GRID DATA ────────────────────────────────────────────────
    const headers = ['Day / Period', 'Period 1', 'Period 2', 'Period 3',
                     'Period 4', 'Period 5', 'Period 6', 'Period 7', 'Period 8'];
    const gridOutput = [headers];

    days.forEach(day => {
      const row = [day];
      periods.forEach(period => {
        const slot = lookup[day][period];
        row.push(slot && slot.text ? slot.text : 'FREE');
      });
      gridOutput.push(row);
    });

    // ── 5. WRITE DATA ─────────────────────────────────────────────────────
    const numRows  = gridOutput.length;
    const dataRange = sheet.getRange(startRow, 1, numRows, numCols);
    dataRange.setValues(gridOutput);

    // ── 6. PREMIUM STYLING ────────────────────────────────────────────────
    dataRange.setFontFamily('Montserrat')
             .setHorizontalAlignment('center')
             .setVerticalAlignment('middle')
             .setWrap(true)
             .setBorder(true, true, true, true, true, true,
                        '#CCD1D9', SpreadsheetApp.BorderStyle.SOLID);

    // Period header row
    sheet.getRange(startRow, 1, 1, numCols)
         .setBackground('#1A3A4A')
         .setFontColor('#FFFFFF')
         .setFontWeight('bold')
         .setFontSize(11);
    sheet.setRowHeight(startRow, 40);

    // Day column + individual cell coloring
    for (let r = 1; r < numRows; r++) {
      const day = days[r - 1];
      const sheetRow = startRow + r;

      // Day label cell
      sheet.getRange(sheetRow, 1)
           .setBackground('#2C3E50')
           .setFontColor('#ECF0F1')
           .setFontWeight('bold')
           .setFontSize(10);

      sheet.setRowHeight(sheetRow, 70);

      // Period cells
      for (let c = 1; c < numCols; c++) {
        const period = periods[c - 1];
        const slot   = lookup[day][period];
        const cell   = sheet.getRange(sheetRow, c + 1);

        if (!slot || !slot.cls) {
          // FREE slot — green tint
          cell.setBackground('#EAFAF1').setFontColor('#1E8449').setFontWeight('bold');
        } else {
          // BUSY slot — alternating blues
          cell.setBackground(r % 2 === 0 ? '#EBF5FB' : '#FDFEFE')
              .setFontColor('#1A5276');
        }
      }

      // Separator border at bottom of each day row
      sheet.getRange(sheetRow, 1, 1, numCols)
           .setBorder(null, null, true, null, null, null,
                      '#7F8C8D', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    }

    // Column widths
    sheet.setColumnWidth(1, 120);
    for (let i = 2; i <= numCols; i++) sheet.setColumnWidth(i, 160);

    sheet.setFrozenRows(startRow);
    sheet.setFrozenColumns(1);
  }
};
