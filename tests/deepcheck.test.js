'use strict';
/**
 * Regression tests for the deep check of 2026-10-24 (live code 4bb8fea): one test per
 * critical / high finding. Each one failed on that commit.
 *   1  Import replaced every data tab without asking when no snapshot existed
 *   2  A paste into Master_Schedule (no e.oldValue) blanked the cell or dropped teachers
 *   3  Re-saving a combined lecture unchanged in a teacher view rewrote other sections' subjects
 *   4  Teacher_View moves to / from a misaligned row lost the teacher or saved a clash
 *   5  Class_View merged a class's duplicate rows and deleted one (see editing.test.js "Duplicate rows")
 *   6  Setup renamed a column the owner inserted, then Refresh / Check overwrote its values
 *   7  Restoring a snapshot cleared the owner's extra columns and removed blank spacer rows
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('path');
const { createEnvironment, loadProject, editEvent } = require('./harness/gas-mock');

const ROOT = path.resolve(__dirname, '..');
const plain = x => JSON.parse(JSON.stringify(x));
const STAFF = ['Asha', 'Bina', 'Chetan', 'Deepa'];

function small(schedule, teachers, classes) {
  const env = createEnvironment();
  const { api } = loadProject(ROOT, env);
  api.setupInitialSpreadsheet();
  const put = (name, rows, w) => rows.length && env.ss.getSheetByName(name).getRange(2, 1, rows.length, w).setValues(rows);
  put('Teachers', teachers.map(t => [t, '', 40, '', '']), 5);
  put('Classes', (classes || ['6th A', '6th B', '7th A']).map(c => [c, 'Middle', '']), 3);
  put('Master_Schedule', schedule, schedule.length ? schedule[0].length : 8);
  return { env, api, ss: env.ss };
}
function builtIn() {
  const env = createEnvironment();
  const { api } = loadProject(ROOT, env);
  api.setupInitialSpreadsheet();
  env.answers.push('YES');
  api.importExcelData();
  return { env, api, ss: env.ss };
}
const msValues = (ss, width) => {
  const sh = ss.getSheetByName('Master_Schedule');
  return plain(sh.getRange(2, 1, sh.getLastRow() - 1, width || 8).getValues());
};
const lastToast = env => env.ss.toasts[env.ss.toasts.length - 1] || {};
/** A Master_Schedule edit; `oldValue: undefined` is what Sheets sends for a paste. */
function scheduleEdit(env, api, row, col, value, opts) {
  const range = env.ss.getSheetByName('Master_Schedule').getRange(row, col);
  const old = range.getValue();
  range.setValue(value);
  const oldValue = opts && 'oldValue' in opts ? opts.oldValue : (old === '' ? undefined : old);
  api.onEdit(editEvent(env, range, { triggerUid: env.properties.EDIT_TRIGGER_UID, oldValue: oldValue }));
}

// ───────────────────────────────────────────── 1

test('1: Import asks before replacing data even with no snapshot; NO keeps every hand edit', () => {
  const { env, api, ss } = builtIn();
  const ms = ss.getSheetByName('Master_Schedule');
  ms.getRange(2, 5).setValue('HAND EDIT');
  const before = msValues(ss);
  const alerts = env.alerts.length;
  env.answers.push('NO');
  api.importExcelData();
  assert.deepEqual(msValues(ss), before, 'NO changes nothing');
  assert.match(env.alerts[alerts], /REPLACES everything[\s\S]*now 1344 schedule rows/);
  env.answers.push('YES');
  api.importExcelData();
  assert.notEqual(ms.getRange(2, 5).getValue(), 'HAND EDIT', 'YES replaces');
});

test('1: Setup / Repair on a sheet with data says the data was kept and not to import', () => {
  const { env, api } = builtIn();
  api.setupInitialSpreadsheet();
  const msg = env.alerts[env.alerts.length - 1];
  assert.match(msg, /Your timetable data was not changed/);
  assert.match(msg, /Do not run 📥 Import Timetable/);
  assert.doesNotMatch(msg, /Next step: run 📥 Import/);
  // A brand-new sheet still points at Import
  const fresh = createEnvironment();
  loadProject(ROOT, fresh).api.setupInitialSpreadsheet();
  assert.match(fresh.alerts[fresh.alerts.length - 1], /Next step: run 📥 Import Timetable/);
});

// ───────────────────────────────────────────── 2

test('2: a paste with no oldValue is checked but never blanks or rewrites the cell', () => {
  const { env, api, ss } = small([
    ['Monday', 1, '6th A', 'Middle', 'Maths', 'Asha', '', ''],
    ['Monday', 1, '7th A', 'Middle', 'Science', 'Deepa', '', ''],
  ], STAFF);
  const teacher = () => ss.getSheetByName('Master_Schedule').getRange(2, 6).getValue();
  // A clash pasted in: left as pasted (Sheets' own undo can restore it) and reported
  scheduleEdit(env, api, 2, 6, 'Deepa', { oldValue: undefined });
  assert.equal(teacher(), 'Deepa');
  assert.equal(lastToast(env).title, 'Check this cell');
  assert.match(lastToast(env).msg, /Deepa already teaches 7th A/);
  // An unknown name pasted in: same, never ''
  scheduleEdit(env, api, 2, 6, 'Ashaa', { oldValue: undefined });
  assert.equal(teacher(), 'Ashaa');
  assert.match(lastToast(env).msg, /unknown teacher "Ashaa"/);
  // Some runtimes pass {} instead of undefined
  scheduleEdit(env, api, 2, 6, 'Nobody', { oldValue: {} });
  assert.equal(teacher(), 'Nobody');
  // With the old value known, a refused edit still puts it back
  scheduleEdit(env, api, 2, 6, 'Chetan', { oldValue: 'Asha' });
  assert.equal(teacher(), 'Chetan', 'valid typed edit saved');
  scheduleEdit(env, api, 2, 6, 'Deepa', { oldValue: 'Chetan' });
  assert.equal(teacher(), 'Chetan', 'refused typed edit reverted to the reported old value');
});

test('2: the first dropdown pick into an empty split Teacher cell still fills the first group', () => {
  const { env, api, ss } = small([['Monday', 1, '6th A', 'Middle', 'Hindi / IT', '', '', '']], STAFF);
  const cell = () => ss.getSheetByName('Master_Schedule').getRange(2, 6).getValue();
  scheduleEdit(env, api, 2, 6, 'Asha', { oldValue: undefined });
  assert.equal(cell(), 'Asha / -');
  assert.match(lastToast(env).msg, /first empty group/);
  scheduleEdit(env, api, 2, 6, 'Bina');
  assert.equal(cell(), 'Asha / Bina');
});

// ───────────────────────────────────────────── 3

test('3: re-saving every teacher-view cell unchanged changes no Master_Schedule row (2026-27 data)', () => {
  const { api, ss } = builtIn();
  const before = JSON.stringify(msValues(ss, 7));
  let saved = 0;
  // Teacher_Day_View shows every teacher for a day; Teacher_View text is the same "Class\nSubject" form
  api.CONFIG.DAYS.forEach(day => {
    api.TeacherDayViewManager.renderTeacherDayView(day, 'All Classes');
    const sh = ss.getSheetByName('Teacher_Day_View');
    sh.getRange(6, 1, sh.getLastRow() - 5, api.viewWidth_()).getValues().forEach(row => row.slice(1).forEach((v, k) => {
      const period = api.periodOfViewColumn_(k + 2);
      if (!period || !v || v === 'FREE') return; // Prayer / Lunch columns are not periods
      if (!api.updateTeacherSlot_(row[0], day, period, v, [], [])) saved++;
    }));
  });
  assert.ok(saved > 1000, 'cells re-saved: ' + saved);
  assert.equal(JSON.stringify(msValues(ss, 7)), before);
});

test('3: a subject typed without its medium mark keeps the stored mark', () => {
  const { api } = small([], STAFF);
  const assign = (...a) => api.ScheduleParser.assignTeacher(...a);
  const res = assign('Asha / Bina', 'Hindi (E.M) / English (H.M)', 'Asha', 'Hindi', '');
  assert.deepEqual([res.subject, res.teacher], ['Hindi (E.M) / English (H.M)', 'Asha / Bina']);
  const res2 = assign('Asha', 'Mathematics', 'Asha', 'Mathematics (E.M)', '');
  assert.equal(res2.subject, 'Mathematics', 'no mark invented');
  const res3 = assign('Asha', 'Hindi (E.M)', 'Asha', 'Hindi (H.M)', '');
  assert.equal(res3.subject, 'Hindi (H.M)', 'a different medium is a real change');
});

// ───────────────────────────────────────────── 4

test('4: Teacher_View refuses a move to or from a row whose groups do not line up', () => {
  const rowsA = [
    ['Monday', 1, '6th A', 'Middle', 'Maths', 'Deepa', '', ''],
    ['Monday', 1, '7th A', 'Middle', 'Hindi / IT / Art', 'Asha / Bina', '', ''],
  ];
  let { env, api, ss } = small(rowsA, STAFF);
  let err = api.updateTeacherSlot_('Deepa', 'Monday', 1, '7th A - Art', [], []);
  assert.match(err, /7th A: its subjects and teacher groups don't line up/);
  assert.deepEqual(msValues(ss), plain(rowsA), 'Deepa keeps 6th A');

  const rowsB = [
    ['Monday', 1, '6th A', 'Middle', 'Hindi / IT / Art', 'Asha / Bina', '', ''],
    ['Monday', 1, '7th A', 'Middle', 'Maths', 'Deepa', '', ''],
  ];
  ({ env, api, ss } = small(rowsB, STAFF));
  err = api.updateTeacherSlot_('Asha', 'Monday', 1, '7th A - Maths', [], []);
  assert.match(err, /6th A: its subjects and teacher groups don't line up/);
  assert.deepEqual(msValues(ss), plain(rowsB), 'no clash saved');
  err = api.updateTeacherSlot_('Asha', 'Monday', 1, 'FREE', [], []);
  assert.match(err, /6th A: its subjects and teacher groups don't line up/, 'FREE reports it too');
});

// ───────────────────────────────────────────── 6

test('6: a column inserted before "Total Hours Scheduled" / "Clash Status" is never renamed or overwritten', () => {
  const { env, api, ss } = small([
    ['Monday', 1, '6th A', 'Middle', 'Maths', 'Asha', '', 'note 1', ''],
    ['Monday', 1, '6th B', 'Middle', 'Science', 'Asha', '', 'note 2', ''],
  ], STAFF);
  const teachers = ss.getSheetByName('Teachers');
  const ms = ss.getSheetByName('Master_Schedule');
  // The owner inserts "Phone" at Teachers E and "Notes" at Master_Schedule H
  teachers.getRange(1, 5, 1, 2).setValues([['Phone', 'Total Hours Scheduled']]);
  teachers.getRange(2, 5, 4, 1).setValues([['98765-1'], ['98765-2'], ['98765-3'], ['98765-4']]);
  ms.getRange(1, 8, 1, 2).setValues([['Notes', 'Clash Status']]);
  api.setupInitialSpreadsheet();
  api.refreshAllViews(true);
  env.alerts.length = 0;
  api.runValidation();

  assert.deepEqual(plain(teachers.getRange(1, 5, 5, 2).getValues()),
    [['Phone', 'Total Hours Scheduled'], ['98765-1', 1], ['98765-2', 0], ['98765-3', 0], ['98765-4', 0]]);
  assert.deepEqual(plain(ms.getRange(1, 8, 3, 1).getValues()), [['Notes'], ['note 1'], ['note 2']]);
  assert.match(env.alerts[0], /not written: Master_Schedule column H is not "Clash Status"/);
});

test('6: Setup still fills in a blank header and adds a missing one at the end', () => {
  const { api, ss } = small([], STAFF);
  const teachers = ss.getSheetByName('Teachers');
  teachers.getRange(1, 3).setValue('');                  // blank → refilled
  teachers.getRange(1, 5).setValue('My Notes');          // "Total Hours Scheduled" gone → appended
  api.setupInitialSpreadsheet();
  assert.deepEqual(plain(teachers.getRange(1, 1, 1, 6).getValues()[0]),
    ['Teacher Name', 'Subject Specialization', 'Max Hours / Week', 'Days Unavailable', 'My Notes', 'Total Hours Scheduled']);
});

// ───────────────────────────────────────────── 7

test('7: Save Snapshot → Import snapshot keeps extra columns and blank spacer rows in place', () => {
  const { env, api, ss } = builtIn();
  const ms = ss.getSheetByName('Master_Schedule');
  const classes = ss.getSheetByName('Classes');
  ms.getRange(1, 9).setValue('Notes');
  ms.getRange(10, 9).setValue('my note');
  classes.getRange(1, 4).setValue('Class Teacher');
  classes.getRange(2, 4).setValue('CT1');
  ms.getRange(20, 1, 1, 8).clearContent();               // a blank spacer row
  const keep = msValues(ss, 9);
  api.saveCurrentSheetAsPrefillData();
  ms.getRange(2, 5).setValue('EDITED AFTER SNAPSHOT');
  env.answers.push('YES');                              // YES = the saved snapshot
  api.importExcelData();
  const after = msValues(ss, 9);
  assert.deepEqual(after.map(r => r.slice(0, 7)), keep.map(r => r.slice(0, 7)), 'rows back in place, spacer included');
  assert.equal(ms.getRange(10, 9).getValue(), 'my note');
  assert.equal(ms.getRange(1, 9).getValue(), 'Notes');
  assert.equal(classes.getRange(2, 4).getValue(), 'CT1');
});
