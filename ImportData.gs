/**
 * ImportData.gs
 * Bulk import of the pre-fill timetable into the data tabs, plus the
 * "saved snapshot" (_Prefill_Data_Backup) workflow.
 *
 * The pre-fill data itself (RAW_CLASS_ROWS, RAW_TEACHER_ROWS, RAW_SUBJECT_ROWS,
 * RAW_SCHEDULE) lives in PrefillData.gs, generated from the Excel timetables by
 * `npm run convert` (tools/convert/build.js).
 *
 * Snapshot layout (_Prefill_Data_Backup, hidden):
 *   A:C  Classes   ·   E:I  Teachers   ·   K  Subjects   ·   M:T  Master_Schedule
 */

/**
 * Decides whether to import the saved snapshot or the built-in pre-fill data.
 * Asks only when a snapshot exists.
 * @returns {'backup'|'default'|null} null = user cancelled
 */
function choosePrefillSource_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const backupSheet = ss.getSheetByName(CONFIG.SHEETS.PREFILL_BACKUP);
  if (!backupSheet || backupSheet.getLastRow() < 2) return 'default';

  const ui = SpreadsheetApp.getUi();
  const savedAt = PropertiesService.getDocumentProperties().getProperty('PREFILL_SAVED_AT');
  const answer = ui.alert(
    'Which timetable should be imported?',
    'A saved snapshot of this sheet exists' + (savedAt ? ' (saved ' + savedAt + ')' : '') + '.\n\n' +
    'YES  → import the saved snapshot\n' +
    'NO   → import the built-in timetable (' + RAW_CLASS_ROWS.length + ' classes, ' + RAW_SCHEDULE.length + ' periods)\n' +
    'CANCEL → do nothing',
    ui.ButtonSet.YES_NO_CANCEL
  );
  if (answer === ui.Button.YES) return 'backup';
  if (answer === ui.Button.NO) return 'default';
  return null;
}

/** Reads the four data blocks from the snapshot sheet. */
function readPrefillBackup_() {
  const backupSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.PREFILL_BACKUP);
  const n = backupSheet.getLastRow() - 1;
  const nonBlank = r => String(r[0]).trim() !== '';
  return {
    classRows: backupSheet.getRange(2, 1, n, 3).getValues().filter(nonBlank),
    teacherRows: backupSheet.getRange(2, 5, n, 5).getValues().filter(nonBlank),
    subjectRows: backupSheet.getRange(2, 11, n, 1).getValues().filter(nonBlank),
    scheduleRows: backupSheet.getRange(2, 13, n, 8).getValues()
      .filter(r => String(r[0]).trim() !== '' || String(r[2]).trim() !== '')
  };
}

/** The built-in data from PrefillData.gs, in sheet-row form. */
function readDefaultPrefill_() {
  return {
    classRows: RAW_CLASS_ROWS.map(r => r.slice(0, 3)),
    teacherRows: RAW_TEACHER_ROWS.map(r => r.slice(0, 5)),
    subjectRows: RAW_SUBJECT_ROWS.map(r => r.slice(0, 1)),
    scheduleRows: RAW_SCHEDULE.map(r => [r.Day, r.Period, r.Class, r.Tier, r.Subject, r.Teacher, r.Room || '', r.ClashStatus || ''])
  };
}

/** Replaces all data rows (below the header) of a sheet. */
function replaceSheetRows_(sheetName, rows, width) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return 0;
  // Always clear, so an empty block in a snapshot does not leave old rows behind
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, Math.max(width, sheet.getLastColumn())).clearContent();
  if (rows.length) sheet.getRange(2, 1, rows.length, width).setValues(rows);
  return rows.length;
}

function importTeachers_(data) {
  return replaceSheetRows_(CONFIG.SHEETS.TEACHERS, data.teacherRows, 5);
}

function importSchedule_(data) {
  return {
    classes: replaceSheetRows_(CONFIG.SHEETS.CLASSES, data.classRows, 3),
    subjects: replaceSheetRows_(CONFIG.SHEETS.SUBJECTS, data.subjectRows, 1),
    schedule: replaceSheetRows_(CONFIG.SHEETS.SCHEDULE, data.scheduleRows, CONFIG.SCHEDULE_HEADERS.length)
  };
}

/** Rebuilds dropdowns, views, styling and the Master Grid after an import. */
function afterImport_() {
  refreshAllViews(/* silent= */ true);
}

function loadPrefill_(source) {
  return source === 'backup' ? readPrefillBackup_() : readDefaultPrefill_();
}

/**
 * Imports everything (teachers, classes, subjects, schedule), then refreshes
 * dropdowns and all views once, and shows a single summary.
 * (This function used to be defined twice; the second copy skipped the refresh.)
 */
function importExcelData() {
  const source = choosePrefillSource_();
  if (!source) return;
  const data = loadPrefill_(source);
  if (!data.classRows.length || !data.scheduleRows.length) {
    SpreadsheetApp.getUi().alert('⚠️ The selected data has no classes or no schedule rows — nothing was imported.');
    return;
  }
  // Write phase under the lock so a concurrent view edit cannot interleave
  const result = withDocumentLock_(() => ({ teachers: importTeachers_(data), counts: importSchedule_(data) }));
  const teachers = result.teachers;
  const counts = result.counts;
  afterImport_();
  const check = validateSchedule_();

  SpreadsheetApp.getUi().alert(
    'Import Complete! Loaded ' + (source === 'backup' ? 'the saved snapshot' : 'the built-in timetable') + ':\n' +
    '• Classes  : ' + counts.classes + '\n' +
    '• Teachers : ' + teachers + '\n' +
    '• Subjects : ' + counts.subjects + '\n' +
    '• Schedule : ' + counts.schedule + ' entries\n' +
    (check && check.rowsFlagged ? '\n⚠️ ' + check.rowsFlagged + ' row(s) have conflicts — see the Clash Status column.' : '')
  );
}

/**
 * Saves all current data (Classes, Teachers, Subjects, Master_Schedule) as the
 * pre-fill snapshot.
 */
function saveCurrentSheetAsPrefillData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const S = CONFIG.SHEETS;

  const read = (name, width, keep) => {
    const sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() < 2) return [];
    return sheet.getRange(2, 1, sheet.getLastRow() - 1, width).getValues().filter(keep);
  };
  const nonBlank = r => String(r[0]).trim() !== '';
  const classRows = read(S.CLASSES, 3, nonBlank);
  const teacherRows = read(S.TEACHERS, 5, nonBlank);
  const subjectRows = read(S.SUBJECTS, 1, nonBlank);
  const scheduleRows = read(S.SCHEDULE, 8, r => String(r[0]).trim() !== '' || String(r[2]).trim() !== '');

  if (classRows.length === 0 && teacherRows.length === 0 && scheduleRows.length === 0) {
    ui.alert('⚠️ No valid data found in the current sheets to save as pre-fill import data.');
    return;
  }

  let backupSheet = ss.getSheetByName(S.PREFILL_BACKUP);
  if (!backupSheet) {
    backupSheet = ss.insertSheet(S.PREFILL_BACKUP);
  } else {
    backupSheet.clear();
  }

  backupSheet.getRange(1, 1, 1, 3).setValues([['Class Name', 'Academic Tier', 'Room Assigned']]);
  backupSheet.getRange(1, 5, 1, 5).setValues([['Teacher Name', 'Subject Specialization', 'Max Hours / Week', 'Days Unavailable', 'Total Hours Scheduled']]);
  backupSheet.getRange(1, 11, 1, 1).setValues([['Subject Name']]);
  backupSheet.getRange(1, 13, 1, 8).setValues([CONFIG.SCHEDULE_HEADERS]);

  if (classRows.length > 0) backupSheet.getRange(2, 1, classRows.length, 3).setValues(classRows);
  if (teacherRows.length > 0) backupSheet.getRange(2, 5, teacherRows.length, 5).setValues(teacherRows);
  if (subjectRows.length > 0) backupSheet.getRange(2, 11, subjectRows.length, 1).setValues(subjectRows);
  if (scheduleRows.length > 0) backupSheet.getRange(2, 13, scheduleRows.length, 8).setValues(scheduleRows);

  backupSheet.hideSheet();
  PropertiesService.getDocumentProperties().setProperty('PREFILL_SAVED_AT', new Date().toLocaleString());

  ui.alert(
    '✅ Pre-fill Import Data Saved Successfully!\n\n' +
    'The current Google Sheet data is now saved as a snapshot:\n' +
    '• Classes  : ' + classRows.length + '\n' +
    '• Teachers : ' + teacherRows.length + '\n' +
    '• Subjects : ' + subjectRows.length + '\n' +
    '• Schedule : ' + scheduleRows.length + ' entries\n\n' +
    '📥 Import Timetable will now ask which one to load: this snapshot or the built-in timetable.'
  );
}
