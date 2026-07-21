/**
 * TeacherDayViewManager.gs
 * Renders a read-only matrix view for ALL teachers for a specific selected Day.
 * Helps admins plan and review school-wide daily coverage and schedules.
 */

const TeacherDayViewManager = {

  /**
   * Renders the day-wise teacher matrix for a specific day.
   * @param {string} selectedDay  e.g. 'Monday', 'Tuesday', ...
   */
  renderTeacherDayView: function(selectedDay) {
    const validDays = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
    if (!selectedDay || !validDays.includes(selectedDay)) {
      selectedDay = 'Monday';
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Teacher_Day_View');
    if (!sheet) sheet = ss.insertSheet('Teacher_Day_View');

    sheet.clear();
    sheet.getDataRange().breakApart();
    sheet.setFrozenRows(0);
    sheet.setFrozenColumns(0);
    sheet.setHiddenGridlines(true);

    const periods = [1, 2, 3, 4, 5, 6, 7, 8];
    const numCols = 9; // Teacher col + 8 periods
    const startRow = 5;

    // ── 1. TITLE BANNER (Rows 1–2) ────────────────────────────────────────
    sheet.getRange(1, 1, 2, numCols)
         .setBackground('#1A3A4A')
         .setFontFamily('Montserrat');
    sheet.getRange(1, 2, 2, numCols - 1)
         .merge()
         .setValue('📅  Day-Wise Teacher Schedule  ·  ' + selectedDay)
         .setFontFamily('Montserrat')
         .setFontSize(16)
         .setFontWeight('bold')
         .setFontColor('#FFFFFF')
         .setHorizontalAlignment('center')
         .setVerticalAlignment('middle');
    sheet.setRowHeight(1, 28);
    sheet.setRowHeight(2, 28);

    // ── 2. DAY SELECTOR (Row 3) ───────────────────────────────────────────
    sheet.getRange('A3')
         .setValue('Select Day:')
         .setFontWeight('bold')
         .setFontFamily('Montserrat')
         .setFontColor('#5D6D7E')
         .setHorizontalAlignment('right');

    const b3Cell = sheet.getRange('B3');
    b3Cell.setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(validDays, true).build()
    );
    b3Cell.setValue(selectedDay);

    sheet.getRange(3, 1, 1, numCols).setBackground('#F2F3F4').setFontFamily('Montserrat');
    sheet.setRowHeight(3, 36);

    // Spacer row 4
    sheet.getRange(4, 1, 1, numCols).setBackground('#E8EAED');
    sheet.setRowHeight(4, 6);

    // ── 3. FETCH TEACHERS & SCHEDULE ──────────────────────────────────────
    const teachersData = DataAccess.getSheetDataAsObjects('Teachers');
    const teacherNames = teachersData.map(t => t['Teacher Name']).filter(t => t).sort();

    if (teacherNames.length === 0) {
      sheet.getRange(startRow, 1).setValue('No teachers found in Teachers tab.');
      return;
    }

    const scheduleData = DataAccess.getSheetDataAsObjects('Master_Schedule')
      .filter(r => r.Day === selectedDay);

    // Build lookup: lookup[teacher][period] = Array of { cls, subject }
    const lookup = {};
    teacherNames.forEach(t => {
      lookup[t] = {};
      periods.forEach(p => lookup[t][p] = []);
    });

    scheduleData.forEach(row => {
      if (!row.Period) return;
      const period = parseInt(row.Period);
      if (isNaN(period) || period < 1 || period > 8) return;

      const assignments = ScheduleParser.parseRowAssignments(row);
      assignments.forEach(assign => {
        if (assign.teacher && lookup[assign.teacher] && lookup[assign.teacher][period]) {
          lookup[assign.teacher][period].push({
            cls: row.Class || '',
            subject: assign.subject || row.Subject || ''
          });
        }
      });
    });

    // ── 4. BUILD GRID OUTPUT ──────────────────────────────────────────────
    const headers = ['Teacher / Period', 'Period 1', 'Period 2', 'Period 3',
                     'Period 4', 'Period 5', 'Period 6', 'Period 7', 'Period 8'];
    const gridOutput = [headers];

    teacherNames.forEach(teacher => {
      const row = [teacher];
      periods.forEach(period => {
        const slots = lookup[teacher][period];
        if (slots.length === 0) {
          row.push('FREE');
        } else if (slots.length === 1) {
          const s = slots[0];
          row.push(s.cls + (s.subject ? ' - ' + s.subject : ''));
        } else {
          // Multiple classes double-booked (Clash)
          const clashDesc = slots.map(s => s.cls + ' (' + s.subject + ')').join(' / ');
          row.push('⚠️ CLASH: ' + clashDesc);
        }
      });
      gridOutput.push(row);
    });

    // ── 5. WRITE DATA & STYLING ───────────────────────────────────────────
    const numRows = gridOutput.length;
    const dataRange = sheet.getRange(startRow, 1, numRows, numCols);
    dataRange.setValues(gridOutput);

    dataRange.setFontFamily('Montserrat')
             .setHorizontalAlignment('center')
             .setVerticalAlignment('middle')
             .setWrap(true)
             .setBorder(true, true, true, true, true, true,
                        '#CCD1D9', SpreadsheetApp.BorderStyle.SOLID);

    // Period header row (Row 5)
    sheet.getRange(startRow, 1, 1, numCols)
         .setBackground('#1A3A4A')
         .setFontColor('#FFFFFF')
         .setFontWeight('bold')
         .setFontSize(11);
    sheet.setRowHeight(startRow, 40);

    // Build background color matrix for body cells
    const bgColors = [];
    const fontColors = [];
    const fontWeights = [];

    for (let r = 1; r < numRows; r++) {
      const sheetRow = startRow + r;
      sheet.setRowHeight(sheetRow, 38);

      const rowBg = [];
      const rowFont = [];
      const rowWeight = [];

      // Teacher label cell (Col 1)
      rowBg.push('#2C3E50');
      rowFont.push('#ECF0F1');
      rowWeight.push('bold');

      // Periods 1-8 (Cols 2-9)
      for (let c = 1; c < numCols; c++) {
        const val = gridOutput[r][c];
        if (val === 'FREE') {
          rowBg.push('#EAFAF1');    // soft green
          rowFont.push('#1E8449');  // dark green
          rowWeight.push('bold');
        } else if (val.startsWith('⚠️ CLASH')) {
          rowBg.push('#FDEDEC');    // soft red
          rowFont.push('#C0392B');  // dark red
          rowWeight.push('bold');
        } else {
          rowBg.push(r % 2 === 0 ? '#EBF5FB' : '#FDFEFE'); // light blue / white
          rowFont.push('#1A5276');  // deep blue
          rowWeight.push('normal');
        }
      }

      bgColors.push(rowBg);
      fontColors.push(rowFont);
      fontWeights.push(rowWeight);
    }

    if (numRows > 1) {
      const bodyRange = sheet.getRange(startRow + 1, 1, numRows - 1, numCols);
      bodyRange.setBackgrounds(bgColors);
      bodyRange.setFontColors(fontColors);
      bodyRange.setFontWeights(fontWeights);
    }

    // Column widths
    sheet.setColumnWidth(1, 190);
    for (let i = 2; i <= numCols; i++) {
      sheet.setColumnWidth(i, 150);
    }

    sheet.setFrozenRows(startRow);
    sheet.setFrozenColumns(1);
  }
};
