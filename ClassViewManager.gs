/**
 * ClassViewManager.gs
 * Handles the 2D editable Class View grid with premium styling.
 * Supports multi-class / combined class selection (e.g. 11th A, 11th B) and optional subjects.
 */

const ClassViewManager = {

  /**
   * Renders the full 2D timetable grid for a specific class (or combined classes).
   * @param {string} className
   */
  renderClassView: function(className) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(CONFIG.SHEETS.CLASS_VIEW);
    if (!sheet) sheet = ss.insertSheet(CONFIG.SHEETS.CLASS_VIEW);

    sheet.setFrozenRows(0);    // reset frozen state before any unmerges/clears
    sheet.setFrozenColumns(0);
    sheet.clear();
    sheet.getDataRange().breakApart();
    sheet.setHiddenGridlines(true);

    const days    = CONFIG.DAYS;
    const periods = CONFIG.PERIODS;
    const numCols = periods.length + 1; // Day col + period cols
    const startRow = 5; // Grid header at row 5, data from row 6

    // ── 1. CLASS SELECTOR (Row 3) ──────────────────────────────────────────
    // Keep the Classes sheet order (alphabetical sorting put "10th A" before "6th A")
    const classNames  = getClassNames_();

    if (!className && classNames.length > 0) {
      className = classNames[0];
    }

    sheet.getRange('A3')
         .setValue('Select Class:')
         .setFontWeight('bold')
         .setFontFamily('Montserrat')
         .setFontColor('#5D6D7E')
         .setHorizontalAlignment('right');

    const classCell = sheet.getRange('B3');
    if (classNames.length > 0) {
      classCell.setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireValueInList(classNames, true)
          .setAllowInvalid(true) // Allows combined class entries e.g. "11th A, 11th B"
          .build()
      );
    }
    classCell.setValue(className);
    sheet.getRange(3, 1, 1, numCols)
         .setBackground('#F2F3F4')
         .setFontFamily('Montserrat');
    sheet.setRowHeight(3, 36);

    // ── 2. TITLE BANNER (Rows 1–2) ────────────────────────────────────────
    sheet.getRange(1, 1, 2, numCols)
         .setBackground('#1A252F')
         .setFontFamily('Montserrat');
    sheet.getRange(1, 2, 2, numCols - 1)
         .merge()
         .setValue('📚  Class Timetable  ·  ' + className)
         .setFontFamily('Montserrat')
         .setFontSize(16)
         .setFontWeight('bold')
         .setFontColor('#FFFFFF')
         .setHorizontalAlignment('center')
         .setVerticalAlignment('middle');
    sheet.setRowHeight(1, 28);
    sheet.setRowHeight(2, 28);

    // Spacer row 4
    sheet.getRange(4, 1, 1, numCols).setBackground('#E8EAED');
    sheet.setRowHeight(4, 6);

    // ── 3. FETCH SCHEDULE DATA ────────────────────────────────────────────
    const scheduleData = DataAccess.getSheetDataAsObjects(CONFIG.SHEETS.SCHEDULE);
    // Rows of every class, by slot: needed to tell combined lectures and clashes apart
    const bySlot = {};
    scheduleData.forEach(r => {
      const d = normalizeDay_(r.Day);
      const p = parseInt(r.Period);
      if (!d || isNaN(p)) return;
      (bySlot[d + '|' + p] = bySlot[d + '|' + p] || []).push(r);
    });

    const lookup = {};
    days.forEach(d => {
      lookup[d] = {};
      periods.forEach(p => {
        const info = ScheduleParser.describeSlot(bySlot[d + '|' + p] || [], className);
        // Lanes keep subject i ↔ teacher i aligned even when several rows are merged
        const text = ScheduleParser.formatLanes(info.lanes);
        lookup[d][p] = { subject: text.subject, teacher: text.teacher, info: info };
      });
    });

    // ── 4. DROPDOWN VALIDATION RULES ─────────────────────────────────────
    let subjectsList = DataAccess.getSheetDataAsObjects(CONFIG.SHEETS.SUBJECTS)
                                 .map(r => r['Subject Name']).filter(s => s);
    if (subjectsList.length === 0) {
      subjectsList = [...new Set(scheduleData.map(r => r['Subject']).filter(s => s))].sort();
    }
    const subjectRule = subjectsList.length > 0
      ? SpreadsheetApp.newDataValidation()
          .requireValueInList(subjectsList, true)
          .setAllowInvalid(true)
          .build()
      : null;

    const teachersList = getTeacherNames_();
    const teacherRule = teachersList.length > 0
      ? SpreadsheetApp.newDataValidation()
          .requireValueInList(teachersList, true)
          .setAllowInvalid(true)
          .build()
      : null;

    // ── 5. BUILD GRID DATA ────────────────────────────────────────────────
    const periodLabels = ['Day / Row'].concat(periods.map(periodHeader_));
    const gridOutput = [periodLabels];

    days.forEach(day => {
      const subRow  = [day];
      const tchrRow = [''];
      periods.forEach(p => {
        subRow.push(lookup[day][p].subject);
        tchrRow.push(lookup[day][p].teacher);
      });
      gridOutput.push(subRow);
      gridOutput.push(tchrRow);
    });

    // ── 6. WRITE DATA ─────────────────────────────────────────────────────
    const numRows  = gridOutput.length;
    const dataRange = sheet.getRange(startRow, 1, numRows, numCols);
    dataRange.setValues(gridOutput);

    // ── 7. PREMIUM STYLING ────────────────────────────────────────────────
    dataRange.setFontFamily('Montserrat')
             .setHorizontalAlignment('center')
             .setVerticalAlignment('middle')
             .setWrap(true)
             .setBorder(true, true, true, true, true, true,
                        '#CCD1D9', SpreadsheetApp.BorderStyle.SOLID);

    // Period header row (startRow)
    sheet.getRange(startRow, 1, 1, numCols)
         .setBackground('#2C3E50')
         .setFontColor('#FFFFFF')
         .setFontWeight('bold')
         .setFontSize(11);
    sheet.setRowHeight(startRow, 40);

    // Day rows: 2 rows each (subject + teacher)
    let currentRow = startRow + 1;
    days.forEach(day => {
      const subjectRowNum = currentRow;
      const teacherRowNum = currentRow + 1;

      // Merge Day label cell vertically across both rows
      sheet.getRange(subjectRowNum, 1, 2, 1)
           .merge()
           .setBackground('#2C3E50')
           .setFontColor('#ECF0F1')
           .setFontWeight('bold')
           .setFontSize(10)
           .setHorizontalAlignment('center')
           .setVerticalAlignment('middle');

      // Subject row cells (cols 2–9)
      const subRange = sheet.getRange(subjectRowNum, 2, 1, periods.length);
      subRange.setBackground('#EBF5FB')   // light blue
              .setFontColor('#1A5276')     // deep blue text
              .setFontWeight('bold');
      if (subjectRule) subRange.setDataValidation(subjectRule);
      else             subRange.clearDataValidations();
      sheet.setRowHeight(subjectRowNum, 36);

      // Teacher row cells (cols 2–9)
      const tchrRange = sheet.getRange(teacherRowNum, 2, 1, periods.length);
      tchrRange.setBackground('#FDFEFE')  // near-white
               .setFontColor('#717D7E')   // muted grey
               .setFontSize(9);
      if (teacherRule) tchrRange.setDataValidation(teacherRule);
      else             tchrRange.clearDataValidations();
      sheet.setRowHeight(teacherRowNum, 28);

      // Thick bottom border between day groups
      sheet.getRange(teacherRowNum, 1, 1, numCols)
           .setBorder(null, null, true, null, null, null,
                      '#7F8C8D', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

      currentRow += 2;
    });

    // ── 8. SPLIT / COMBINED / CLASH MARKERS ───────────────────────────────
    // Colour the subject cell and explain it in a note (hover to read).
    const mark = { split: '#FEF5E7', combined: '#F5EEF8', clash: '#FADBD8' };
    const bgs = [], notes = [];
    days.forEach(day => {
      const bgRow = [], noteRow = [];
      periods.forEach(p => {
        const info = lookup[day][p].info;
        bgRow.push(info.clashes.length ? mark.clash : info.split ? mark.split : info.combinedWith.length ? mark.combined : '#EBF5FB');
        noteRow.push(info.note);
      });
      bgs.push(bgRow);
      notes.push(noteRow);
    });
    days.forEach((day, i) => {
      const range = sheet.getRange(startRow + 1 + i * 2, 2, 1, periods.length);
      range.setBackgrounds([bgs[i]]);
      range.setNotes([notes[i]]);
    });
    sheet.getRange(3, 4, 1, numCols - 3)
         .merge()
         .setValue('🟧 split class   🟪 taught with another section   🟥 teacher clash   — hover a cell for details')
         .setFontSize(9)
         .setFontColor('#5D6D7E')
         .setHorizontalAlignment('left');

    // Column widths
    sheet.setColumnWidth(1, 110);
    for (let i = 2; i <= numCols; i++) {
      sheet.setColumnWidth(i, 150);
    }

    // Freeze header rows and day column
    sheet.setFrozenRows(startRow);
    sheet.setFrozenColumns(1);
  },

  /**
   * Updates the Master_Schedule from an inline edit made on Class_View.
   * @param {string} className
   * @param {string} day
   * @param {number} period
   * @param {string} editedType  'Subject' | 'Teacher'
   * @param {string} newValue
   * @returns {null}
   */
  updateMasterFromClassView: function(className, day, period, editedType, newValue) {
    if (!className || !day || !period) return null;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const msSheet = ss.getSheetByName(CONFIG.SHEETS.SCHEDULE);
    if (!msSheet) return null;

    const C = CONFIG.COL;
    const values = msSheet.getDataRange().getValues();
    let rowIndexToUpdate = -1;

    for (let i = 1; i < values.length; i++) {
      if (normalizeDay_(values[i][C.DAY - 1]) === day && values[i][C.PERIOD - 1] == period && ScheduleParser.rowIncludesClass({ Class: values[i][C.CLASS - 1] }, className)) {
        rowIndexToUpdate = i + 1; // 1-indexed
        break;
      }
    }

    const col = editedType === 'Subject' ? C.SUBJECT : editedType === 'Teacher' ? C.TEACHER : 0;
    if (!col) return null;

    if (rowIndexToUpdate > 0) {
      msSheet.getRange(rowIndexToUpdate, col).setValue(newValue);
    } else if (String(newValue || '').trim()) {
      // Row doesn't exist yet — create it (previously a Teacher edit on an empty slot was lost)
      const classes   = DataAccess.getSheetDataAsObjects(CONFIG.SHEETS.CLASSES);
      const classData = classes.find(c => ScheduleParser.sameName(c['Class Name'], className));
      const tier = classData ? classData['Academic Tier'] : '';
      const room = classData ? classData['Room Assigned']  : '';
      msSheet.appendRow([day, period, className, tier,
                         editedType === 'Subject' ? newValue : '',
                         editedType === 'Teacher' ? newValue : '',
                         room, '']);
    }
    return null;
  }
};
