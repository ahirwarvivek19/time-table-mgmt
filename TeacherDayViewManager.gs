/**
 * TeacherDayViewManager.gs
 * Renders a read-only matrix view for ALL teachers for a specific selected Day and Class filter.
 * Supports combined lectures across multiple classes for optional subjects (e.g. 11th A, 11th B - IP).
 */

const TeacherDayViewManager = {

  /**
   * Renders the day-wise teacher matrix.
   * @param {string} [selectedDay]  e.g. 'Monday', 'Tuesday', ...
   * @param {string} [selectedClass] e.g. 'All Classes' or '11th A, 11th B'
   */
  renderTeacherDayView: function(selectedDay, selectedClass) {
    const validDays = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
    if (!selectedDay || !validDays.includes(selectedDay)) {
      selectedDay = 'Monday';
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Teacher_Day_View');
    if (!sheet) sheet = ss.insertSheet('Teacher_Day_View');

    // Preserve previous selections if not passed explicitly
    if (!selectedClass && sheet.getLastRow() >= 3) {
      selectedClass = sheet.getRange('E3').getValue();
    }
    if (!selectedClass) selectedClass = 'All Classes';

    sheet.clear();
    sheet.getDataRange().breakApart();
    sheet.setFrozenRows(0);
    sheet.setFrozenColumns(0);
    sheet.setHiddenGridlines(true);

    const periods = [1, 2, 3, 4, 5, 6, 7, 8];
    const numCols = 9; // Teacher col + 8 periods
    const startRow = 5;

    // ── 1. FETCH CLASSES & TEACHERS DATA ──────────────────────────────────
    const classesData = DataAccess.getSheetDataAsObjects('Classes');
    const classNames = classesData.map(c => c['Class Name']).filter(c => c).sort();
    const filterOptions = ['All Classes', ...classNames];

    const teachersData = DataAccess.getSheetDataAsObjects('Teachers');
    const teacherNames = teachersData.map(t => t['Teacher Name']).filter(t => t).sort();

    // ── 2. TITLE BANNER (Rows 1–2) ────────────────────────────────────────
    sheet.getRange(1, 1, 2, numCols)
         .setBackground('#1A3A4A')
         .setFontFamily('Montserrat');
    sheet.getRange(1, 2, 2, numCols - 1)
         .merge()
         .setValue('📅  Day-Wise Teacher Schedule  ·  ' + selectedDay + (selectedClass !== 'All Classes' ? ' (' + selectedClass + ')' : ''))
         .setFontFamily('Montserrat')
         .setFontSize(16)
         .setFontWeight('bold')
         .setFontColor('#FFFFFF')
         .setHorizontalAlignment('center')
         .setVerticalAlignment('middle');
    sheet.setRowHeight(1, 28);
    sheet.setRowHeight(2, 28);

    // ── 3. CONTROLS BAR (Row 3): Day & Class Selectors ────────────────────
    // A3 & B3: Day Dropdown
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

    // D3 & E3: Class Filter Dropdown (Multi-select / combined class support)
    sheet.getRange('D3')
         .setValue('Filter Class:')
         .setFontWeight('bold')
         .setFontFamily('Montserrat')
         .setFontColor('#5D6D7E')
         .setHorizontalAlignment('right');

    const e3Cell = sheet.getRange('E3');
    if (filterOptions.length > 0) {
      e3Cell.setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireValueInList(filterOptions, true)
          .setAllowInvalid(true) // Allows free-text combined classes e.g. "11th A, 11th B"
          .build()
      );
    }
    e3Cell.setValue(selectedClass);

    sheet.getRange(3, 1, 1, numCols).setBackground('#F2F3F4').setFontFamily('Montserrat');
    sheet.setRowHeight(3, 36);

    // Spacer row 4
    sheet.getRange(4, 1, 1, numCols).setBackground('#E8EAED');
    sheet.setRowHeight(4, 6);

    if (teacherNames.length === 0) {
      sheet.getRange(startRow, 1).setValue('No teachers found in Teachers tab.');
      return;
    }

    // ── 4. FETCH SCHEDULE DATA ────────────────────────────────────────────
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

      // Filter by Class selection if specific class requested
      if (selectedClass !== 'All Classes' && !ScheduleParser.rowIncludesClass(row, selectedClass)) {
        return;
      }

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

    // ── 5. BUILD GRID OUTPUT ──────────────────────────────────────────────
    const headers = ['Teacher / Period', 'Period 1', 'Period 2', 'Period 3',
                     'Period 4', 'Period 5', 'Period 6', 'Period 7', 'Period 8'];
    const gridOutput = [headers];
    const metadataGrid = []; // Stores { isFree, isCombined, isClash } per cell

    teacherNames.forEach(teacher => {
      const row = [teacher];
      const metaRow = [{ isFree: false, isCombined: false, isClash: false }];

      periods.forEach(period => {
        const slots = lookup[teacher][period];
        const grouped = ScheduleParser.groupTeacherSlots(slots);
        row.push(grouped.display);
        metaRow.push(grouped);
      });

      gridOutput.push(row);
      metadataGrid.push(metaRow);
    });

    // ── 6. WRITE DATA & APPLY PREMIUM STYLING ────────────────────────────
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

    // Build background color & typography matrices for body cells
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
        const meta = metadataGrid[r - 1][c];
        if (meta.isFree) {
          rowBg.push('#EAFAF1');    // soft green for FREE
          rowFont.push('#1E8449');
          rowWeight.push('bold');
        } else if (meta.isClash) {
          rowBg.push('#FDEDEC');    // soft red for true double-booking clash
          rowFont.push('#C0392B');
          rowWeight.push('bold');
        } else if (meta.isCombined) {
          rowBg.push('#F5EEF8');    // soft lavender/purple for combined optional lectures
          rowFont.push('#5B2C6F');
          rowWeight.push('bold');
        } else {
          rowBg.push(r % 2 === 0 ? '#EBF5FB' : '#FDFEFE'); // light blue / white
          rowFont.push('#1A5276');
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
      sheet.setColumnWidth(i, 155);
    }

    sheet.setFrozenRows(startRow);
    sheet.setFrozenColumns(1);
  }
};
