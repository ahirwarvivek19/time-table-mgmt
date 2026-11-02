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
  const ui = SpreadsheetApp.getUi();
  const backupSheet = ss.getSheetByName(CONFIG.SHEETS.PREFILL_BACKUP);
  const current = scheduleRowCount_();
  const warning = current
    ? '⚠️ This REPLACES everything on Teachers, Classes, Subjects and Master_Schedule (now ' + current +
      ' schedule rows), so every edit made in this sheet is lost. Columns you added yourself are kept; ' +
      'they stay next to the right rows only when you load a saved snapshot.\n\n'
    : '';
  const builtIn = 'the built-in timetable (' + RAW_CLASS_ROWS.length + ' classes, ' + RAW_SCHEDULE.length + ' periods)';

  if (!backupSheet || backupSheet.getLastRow() < 2) {
    // No snapshot: still ask, so one click never replaces hand-edited data
    const answer = ui.alert(
      'Replace the timetable?',
      warning + 'Import ' + builtIn + '?\n\nTo keep a copy of the current data first, choose NO and run 💾 Save Snapshot.',
      ui.ButtonSet.YES_NO
    );
    return answer === ui.Button.YES ? 'default' : null;
  }

  const savedAt = PropertiesService.getDocumentProperties().getProperty('PREFILL_SAVED_AT');
  const answer = ui.alert(
    'Which timetable should be imported?',
    warning +
    'A saved snapshot of this sheet exists' + (savedAt ? ' (saved ' + savedAt + ')' : '') + '.\n\n' +
    'YES  → import the saved snapshot\n' +
    'NO   → import ' + builtIn + '\n' +
    'CANCEL → do nothing',
    ui.ButtonSet.YES_NO_CANCEL
  );
  if (answer === ui.Button.YES) return 'backup';
  if (answer === ui.Button.NO) return 'default';
  return null;
}

/** Number of Master_Schedule rows that hold a day or a class (0 when the tab is missing). */
function scheduleRowCount_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.SCHEDULE);
  if (!sheet || sheet.getLastRow() < 2) return 0;
  const C = CONFIG.COL;
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, C.CLASS).getValues()
    .filter(r => String(r[C.DAY - 1]).trim() !== '' || String(r[C.CLASS - 1]).trim() !== '').length;
}

/** Number of rows with any content. */
function filledCount_(rows) {
  return rows.filter(r => r.some(v => String(v).trim() !== '')).length;
}

/** Rows up to the last non-blank one, so blank spacer rows keep their place. */
function trimTrailingBlankRows_(rows) {
  let n = rows.length;
  while (n > 0 && rows[n - 1].every(v => String(v).trim() === '')) n--;
  return rows.slice(0, n);
}

/** Reads the four data blocks from the snapshot sheet. */
function readPrefillBackup_() {
  const backupSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.PREFILL_BACKUP);
  const n = backupSheet.getLastRow() - 1;
  // Blocks have different lengths, so each is trimmed at its own end; blank rows inside a
  // block are kept so rows go back where they were (and stay next to the owner's own columns)
  return {
    classRows: trimTrailingBlankRows_(backupSheet.getRange(2, 1, n, 3).getValues()),
    teacherRows: trimTrailingBlankRows_(backupSheet.getRange(2, 5, n, 5).getValues()),
    subjectRows: trimTrailingBlankRows_(backupSheet.getRange(2, 11, n, 1).getValues()),
    scheduleRows: trimTrailingBlankRows_(backupSheet.getRange(2, 13, n, 8).getValues())
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

/**
 * Replaces the data rows (below the header) of the columns the timetable owns.
 * Columns to the right that the owner added (phone numbers, notes) are left alone.
 */
function replaceSheetRows_(sheetName, rows, width) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return 0;
  // Always clear, so an empty block in a snapshot does not leave old rows behind
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, width).clearContent();
  if (rows.length) sheet.getRange(2, 1, rows.length, width).setValues(rows);
  return filledCount_(rows);
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
  logChanges_('Import', '', '', [{ cls: '(all)', before: '(sheet before import)',
    after: (source === 'backup' ? 'saved snapshot' : 'built-in timetable') + ': ' + counts.schedule + ' schedule rows' }]);
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

  const read = (name, width) => {
    const sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() < 2) return [];
    // Blank rows inside the data are kept, so a restore puts every row back in its place
    return trimTrailingBlankRows_(sheet.getRange(2, 1, sheet.getLastRow() - 1, width).getValues());
  };
  const classRows = read(S.CLASSES, 3);
  const teacherRows = read(S.TEACHERS, 5);
  const subjectRows = read(S.SUBJECTS, 1);
  const scheduleRows = read(S.SCHEDULE, 8);

  if (classRows.length === 0 && teacherRows.length === 0 && scheduleRows.length === 0) {
    ui.alert('⚠️ No valid data found in the current sheets to save as pre-fill import data.');
    return;
  }
  // Count only the typed key columns: computed ones (Total Hours, Clash Status) still hold
  // values after the real data has been wiped
  const named = (rows, cols) => rows.filter(r => cols.some(c => String(r[c]).trim() !== '')).length;
  if (named(teacherRows, [0]) === 0 || named(scheduleRows, [CONFIG.COL.DAY - 1, CONFIG.COL.CLASS - 1]) === 0) {
    ui.alert('⚠️ Not saved: the Teachers or Master_Schedule tab is empty, and a snapshot of it would empty that tab when restored.');
    return;
  }

  let backupSheet = ss.getSheetByName(S.PREFILL_BACKUP);
  // There is only one snapshot: replacing it (e.g. right after a mistaken Import) loses the good copy
  if (backupSheet && backupSheet.getLastRow() >= 2) {
    const savedAt = PropertiesService.getDocumentProperties().getProperty('PREFILL_SAVED_AT');
    const answer = ui.alert('Replace the saved snapshot?',
      'A snapshot already exists' + (savedAt ? ' (saved ' + savedAt + ')' : '') + '. Saving now replaces it with the sheet as it is now.\n\nContinue?',
      ui.ButtonSet.YES_NO);
    if (answer !== ui.Button.YES) return;
  }
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
    '• Classes  : ' + filledCount_(classRows) + '\n' +
    '• Teachers : ' + filledCount_(teacherRows) + '\n' +
    '• Subjects : ' + filledCount_(subjectRows) + '\n' +
    '• Schedule : ' + filledCount_(scheduleRows) + ' entries\n\n' +
    '📥 Import Timetable will now ask which one to load: this snapshot or the built-in timetable.'
  );
}
