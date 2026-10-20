/**
 * MasterGrid.gs
 * Generates a unified 2D grid view with premium styling.
 *
 * Two modes:
 *  generateMasterGrid()      — full generation + styling (menu / first run)
 *  refreshMasterGridData_()  — data-only refresh, no re-styling (called from onEdit)
 */

// ─────────────────────────────────────────────────────────────────────────────
// PUBLIC: Full generation with styling (called from menu)
// ─────────────────────────────────────────────────────────────────────────────

/**
 * @param {boolean} [silent=false] When true, no alert is shown (used after imports).
 */
function generateMasterGrid(silent) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let gridSheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_GRID);
  if (!gridSheet) gridSheet = ss.insertSheet(CONFIG.SHEETS.MASTER_GRID);
  gridSheet.clear();

  const result = _buildMasterGridOutput_();
  if (!result) {
    if (!silent) SpreadsheetApp.getUi().alert('Master Schedule is empty!');
    return;
  }

  const { gridOutput, classes } = result;
  const numRows = gridOutput.length;
  const numCols = gridOutput[0].length;

  // Write data
  gridSheet.getRange(1, 1, numRows, numCols).setValues(gridOutput);

  // Apply full premium styling
  _applyMasterGridStyling_(gridSheet, numRows, numCols, classes.length);

  if (!silent) SpreadsheetApp.getUi().alert('Master Grid View generated with premium styling!');
}

// ─────────────────────────────────────────────────────────────────────────────
// INTERNAL: Lightweight data refresh (called from onEdit — no styling)
// ─────────────────────────────────────────────────────────────────────────────

function refreshMasterGridData_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gridSheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_GRID);
  if (!gridSheet || gridSheet.getLastRow() === 0) return; // don't create from scratch

  const result = _buildMasterGridOutput_();
  if (!result) return;

  const { gridOutput } = result;
  const numRows = gridOutput.length;
  const numCols = gridOutput[0].length;

  // Clear only data area and rewrite — preserves all formatting
  const existingRows = gridSheet.getLastRow();
  if (existingRows > numRows) {
    gridSheet.getRange(numRows + 1, 1, existingRows - numRows, numCols).clearContent();
  }
  gridSheet.getRange(1, 1, numRows, numCols).setValues(gridOutput);
  // Rows added since the last full restyle (e.g. new classes) get the basic body style
  if (numRows > existingRows && existingRows > 0) {
    gridSheet.getRange(existingRows + 1, 1, numRows - existingRows, numCols)
             .setFontFamily(CONFIG.FONT).setWrap(true)
             .setVerticalAlignment('middle').setHorizontalAlignment('center');
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// INTERNAL: Build the grid output array (shared by both modes)
// ─────────────────────────────────────────────────────────────────────────────

function _buildMasterGridOutput_() {
  const scheduleData = DataAccess.getSheetDataAsObjects(CONFIG.SHEETS.SCHEDULE);
  if (scheduleData.length === 0) return null;

  // Classes sheet order first (6th … 12th), then any class only found in the schedule
  const known = getClassNames_();
  const extra = [...new Set(scheduleData.map(r => String(r.Class || '').trim()).filter(Boolean))]
    .filter(c => known.indexOf(c) === -1);
  const inSchedule = {};
  scheduleData.forEach(r => { inSchedule[String(r.Class || '').trim()] = true; });
  const classes = known.filter(c => inSchedule[c]).concat(extra);
  const periods = CONFIG.PERIODS;
  const numDays = CONFIG.DAYS.length;

  // Build lookup: lookup[class][period] → array of { day, entry }
  const lookup = {};
  classes.forEach(c => {
    lookup[c] = {};
    periods.forEach(p => lookup[c][p] = []);
  });

  scheduleData.forEach(row => {
    if (row.Class && row.Period && row.Day) {
      const p = parseInt(row.Period);
      const cls = String(row.Class).trim();
      if (lookup[cls] && lookup[cls][p] && (row.Subject || row.Teacher)) {
        let entry = row.Subject || '';
        if (row.Teacher) entry += ' (' + row.Teacher + ')';
        lookup[cls][p].push({ day: String(row.Day), entry: entry });
      }
    }
  });

  // Build output array
  const headers = ['Class'].concat(periods.map(periodHeader_));
  const gridOutput = [headers];

  classes.forEach(c => {
    const rowOutput = [c];
    periods.forEach(p => {
      const entries = lookup[c][p];
      if (entries.length === 0) {
        rowOutput.push('—');
      } else {
        // Group by entry text, collect day abbreviations
        const grouped = {};
        entries.forEach(e => {
          if (!grouped[e.entry]) grouped[e.entry] = [];
          grouped[e.entry].push(e.day.substring(0, 3));
        });
        const cellLines = [];
        for (const [entryText, daysArr] of Object.entries(grouped)) {
          // Only an entry that runs every school day is shown without its day list
          // (">= 5" used to hide the one day an entry did NOT run)
          cellLines.push(daysArr.length >= numDays ? entryText
                                                   : entryText + ' [' + daysArr.join(',') + ']');
        }
        rowOutput.push(cellLines.join('\n'));
      }
    });
    gridOutput.push(rowOutput);
  });

  return { gridOutput, classes };
}

// ─────────────────────────────────────────────────────────────────────────────
// INTERNAL: Apply premium styling (called only on full generation)
// ─────────────────────────────────────────────────────────────────────────────

function _applyMasterGridStyling_(gridSheet, numRows, numCols, numClasses) {
  const dataRange = gridSheet.getRange(1, 1, numRows, numCols);

  // Base styles
  gridSheet.setHiddenGridlines(true);
  gridSheet.setFrozenRows(1);
  gridSheet.setFrozenColumns(1);

  dataRange.setFontFamily(CONFIG.FONT)
           .setVerticalAlignment('middle')
           .setHorizontalAlignment('center')
           .setWrap(true)
           .setBorder(true, true, true, true, true, true,
                      '#CCD1D9', SpreadsheetApp.BorderStyle.SOLID);

  // Header row
  gridSheet.getRange(1, 1, 1, numCols)
           .setBackground('#1A252F')
           .setFontColor('#FFFFFF')
           .setFontWeight('bold')
           .setFontSize(11);
  gridSheet.setRowHeight(1, 44);

  // Class column (Y-axis)
  gridSheet.getRange(2, 1, numRows - 1, 1)
           .setBackground('#2C3E50')
           .setFontColor('#FFFFFF')
           .setFontWeight('bold');

  // Body: alternating row colors with batch setBackgrounds
  if (numRows > 1) {
    const bodyRange = gridSheet.getRange(2, 2, numRows - 1, numCols - 1);
    const bgColors = [];
    for (let i = 0; i < numRows - 1; i++) {
      bgColors.push(Array(numCols - 1).fill(i % 2 === 0 ? '#EBF5FB' : '#FDFEFE'));
    }
    bodyRange.setBackgrounds(bgColors);
    bodyRange.setFontColor('#1A2530');
  }

  // Column & row sizing
  gridSheet.setColumnWidth(1, 140);
  for (let i = 2; i <= numCols; i++) gridSheet.setColumnWidth(i, 210);
  gridSheet.setRowHeights(2, numRows - 1, 85);
}
