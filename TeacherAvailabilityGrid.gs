/**
 * TeacherAvailabilityGrid.gs
 * Generates a 2D grid showing free vs busy slots for all teachers on a specific day.
 */

function generateTeacherAvailabilityGrid() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Teacher Availability Grid',
    'Which day do you want to check? (' + CONFIG.DAYS.join(', ') + ')',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  // "monday", "Mon", " MONDAY " all work; anything else is rejected instead of
  // silently showing every teacher as FREE.
  const day = normalizeDay_(response.getResponseText());
  if (!day) {
    ui.alert('Please enter a school day: ' + CONFIG.DAYS.join(', ') + '.');
    return;
  }
  renderTeacherAvailabilityGrid_(day);
}

/** The day the availability grid currently shows (from its "Teacher (Monday)" header), or ''. */
function availabilityGridDay_(gridSheet) {
  const m = String(gridSheet.getRange(1, 1).getValue() || '').match(/\(([^)]+)\)/);
  return m ? normalizeDay_(m[1]) : '';
}

/**
 * Renders Teacher_Availability_Grid for one day.
 * @param {string} day  Canonical day name, e.g. "Monday"
 * @param {boolean} [stayOnSheet]  true = refresh in the background (after an edit) without switching tabs
 */
function renderTeacherAvailabilityGrid_(day, stayOnSheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let gridSheet = ss.getSheetByName(CONFIG.SHEETS.AVAILABILITY);
  if (!gridSheet) gridSheet = ss.insertSheet(CONFIG.SHEETS.AVAILABILITY);
  gridSheet.clear();

  const teachersData = DataAccess.getSheetDataAsObjects(CONFIG.SHEETS.TEACHERS);
  const teachers = getTeacherNames_();
  if (teachers.length === 0) {
    SpreadsheetApp.getUi().alert('No teachers found in the Teachers tab.');
    return;
  }
  const scheduleData = DataAccess.getSheetDataAsObjects(CONFIG.SHEETS.SCHEDULE).filter(row => normalizeDay_(row.Day) === day);
  const periods = CONFIG.PERIODS;
  const UNAVAILABLE = 'UNAVAILABLE';

  // lookup[normalized teacher][period] = "FREE" | "BUSY (classes)" | "UNAVAILABLE"
  const lookup = {};
  teachers.forEach(t => {
    const key = ScheduleParser.normalizeName(t);
    const record = teachersData.find(r => ScheduleParser.sameName(r['Teacher Name'], t)) || {};
    const offDays = String(record['Days Unavailable'] || '').split(/[,\/;]/).map(normalizeDay_);
    lookup[key] = {};
    periods.forEach(p => lookup[key][p] = offDays.indexOf(day) !== -1 ? UNAVAILABLE : CONFIG.FREE);
  });

  const busy = {}; // key|period → [classes]
  scheduleData.forEach(row => {
    const p = parseInt(row.Period);
    if (!row.Period || periods.indexOf(p) === -1) return;
    ScheduleParser.parseRowAssignments(row).forEach(assign => {
      const key = ScheduleParser.normalizeName(assign.teacher);
      if (!key || !lookup[key]) return;
      const k = key + '|' + p;
      (busy[k] = busy[k] || []).push(String(row.Class));
    });
  });
  Object.keys(busy).forEach(k => {
    const [key, p] = k.split('|');
    lookup[key][p] = 'BUSY (' + [...new Set(busy[k])].join(', ') + ')';
  });

  // Build output
  const headers = ['Teacher (' + day + ')'].concat(periods.map(periodHeader_));
  const gridOutput = [headers];
  teachers.forEach(t => {
    const key = ScheduleParser.normalizeName(t);
    gridOutput.push([t].concat(periods.map(p => lookup[key][p])));
  });

  const numRows = gridOutput.length;
  const numCols = headers.length;
  const dataRange = gridSheet.getRange(1, 1, numRows, numCols);
  dataRange.setValues(gridOutput);

  // Styling
  gridSheet.setHiddenGridlines(true);
  gridSheet.setFrozenRows(1);
  gridSheet.setFrozenColumns(1);
  dataRange.setFontFamily(CONFIG.FONT).setVerticalAlignment('middle').setHorizontalAlignment('center').setWrap(true)
           .setBorder(true, true, true, true, true, true, '#D5D8DC', SpreadsheetApp.BorderStyle.SOLID);

  gridSheet.getRange(1, 1, 1, numCols)
           .setBackground('#2C3E50')
           .setFontColor('#FFFFFF')
           .setFontWeight('bold');

  // Body colours: GREEN = free, grey = busy, amber = unavailable (batched)
  const bg = [], fc = [], fw = [];
  for (let r = 1; r < numRows; r++) {
    const rowBg = [], rowFont = [], rowWeight = [];
    for (let c = 1; c < numCols; c++) {
      const v = gridOutput[r][c];
      if (v === CONFIG.FREE) { rowBg.push('#E8F8F5'); rowFont.push('#27AE60'); rowWeight.push('bold'); }
      else if (v === UNAVAILABLE) { rowBg.push('#FEF5E7'); rowFont.push('#B9770E'); rowWeight.push('bold'); }
      else { rowBg.push('#F8F9FA'); rowFont.push('#7F8C8D'); rowWeight.push('normal'); }
    }
    bg.push(rowBg); fc.push(rowFont); fw.push(rowWeight);
  }
  const periodBody = gridSheet.getRange(2, 2, numRows - 1, numCols - 1);
  periodBody.setBackgrounds(bg);
  periodBody.setFontColors(fc);
  periodBody.setFontWeights(fw);

  gridSheet.getRange(2, 1, numRows - 1, 1)
           .setBackground('#34495E')
           .setFontColor('#FFFFFF')
           .setFontWeight('bold');

  dataRange.setBorder(true, true, true, true, true, true, '#BDC3C7', SpreadsheetApp.BorderStyle.SOLID);
  gridSheet.setColumnWidth(1, 180);
  for (let i = 2; i <= numCols; i++) gridSheet.setColumnWidth(i, 120);

  if (!stayOnSheet) ss.setActiveSheet(gridSheet);
}
