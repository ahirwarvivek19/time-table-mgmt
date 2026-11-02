'use strict';
/**
 * Regression tests for the round-2 deep check of 2026-10-24 (real school scenarios, code
 * review, Cover Manager / web app / triggers). Each test failed before its fix.
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
  put('Teachers', teachers.map(t => Array.isArray(t) ? t : [t, '', 40, '', '']), 5);
  put('Classes', (classes || ['6th A', '6th B', '7th A']).map(c => [c, 'Middle', '']), 3);
  put('Master_Schedule', schedule, schedule.length ? schedule[0].length : 8);
  return { env, api, ss: env.ss };
}
const ms = ss => plain(ss.getSheetByName('Master_Schedule').getRange(2, 1, ss.getSheetByName('Master_Schedule').getLastRow() - 1, 8).getValues());
const lastToast = env => env.ss.toasts[env.ss.toasts.length - 1] || {};
const uid = env => ({ triggerUid: env.properties.EDIT_TRIGGER_UID });
function edit(env, api, sheetName, row, col, value, extra) {
  const range = env.ss.getSheetByName(sheetName).getRange(row, col);
  const old = range.getValue();
  range.setValue(value);
  api.onEdit(editEvent(env, range, Object.assign({ oldValue: old === '' ? undefined : old }, uid(env), extra || {})));
}
function classEdit(env, api, cls, row, col, value) {
  api.ClassViewManager.renderClassView(cls);
  edit(env, api, 'Class_View', row, col, value);
}
const logRows = ss => { const s = ss.getSheetByName('Change_Log'); return s ? plain(s.getRange(2, 1, Math.max(1, s.getLastRow() - 1), 8).getValues()) : []; };

// ───────────────────────────────────────────── editing

test('A teacher freed from a split period leaves their group with "-" (the subject is kept) and a notice', () => {
  const { env, api, ss } = small([['Monday', 1, '6th A', 'Middle', 'English / IT / Beauty', 'Asha / Bina / Chetan', '', '']], STAFF);
  assert.equal(api.updateTeacherSlot_('Chetan', 'Monday', 1, 'FREE', [], []), '');
  assert.deepEqual(ms(ss)[0].slice(4, 6), ['English / IT / Beauty', 'Asha / Bina / -']);
  // The replacement then takes the empty group instead of overwriting another subject
  const notices = [];
  assert.equal(api.updateTeacherSlot_('Deepa', 'Monday', 1, '6th A - Beauty', [], notices), '');
  assert.deepEqual(ms(ss)[0].slice(4, 6), ['English / IT / Beauty', 'Asha / Bina / Deepa']);
  // Through Teacher_Day_View the editor is told
  api.TeacherDayViewManager.renderTeacherDayView('Monday', 'All Classes');
  const tdv = ss.getSheetByName('Teacher_Day_View');
  const row = plain(tdv.getRange(6, 1, 4, 1).getValues()).findIndex(r => r[0] === 'Deepa') + 6;
  edit(env, api, 'Teacher_Day_View', row, 3, 'FREE');
  assert.match(lastToast(env).msg, /6th A Beauty group now has no teacher/);
});

test('Picking one subject on a split period is refused instead of merging the groups', () => {
  const { env, api, ss } = small([['Monday', 1, '6th A', 'Middle', 'Sanskrit / Beauty', 'Asha / Bina', '', '']], STAFF);
  classEdit(env, api, '6th A', 6, 3, 'Music');
  assert.deepEqual(ms(ss)[0].slice(4, 6), ['Sanskrit / Beauty', 'Asha / Bina']);
  assert.match(lastToast(env).msg, /split into 2 groups .* e\.g\. "Music \/ Beauty"/);
  // Master_Schedule dropdown pick: same rule
  edit(env, api, 'Master_Schedule', 2, 5, 'Music', { oldValue: 'Sanskrit / Beauty' });
  assert.equal(ms(ss)[0][4], 'Sanskrit / Beauty');
  // Typing every group works; a co-taught single subject can still change subject
  classEdit(env, api, '6th A', 6, 3, 'Sanskrit / Music');
  assert.equal(ms(ss)[0][4], 'Sanskrit / Music');
  const two = small([['Monday', 1, '6th A', 'Middle', 'Physics', 'Asha / Bina', '', '']], STAFF);
  classEdit(two.env, two.api, '6th A', 6, 3, 'Chemistry');
  assert.equal(ms(two.ss)[0][4], 'Chemistry');
});

test('Booking a teacher on their day off is saved with a warning', () => {
  const { env, api, ss } = small([['Monday', 1, '6th A', 'Middle', 'Maths', 'Asha', '', '']],
    [['Asha', '', 40, '', ''], ['Bina', '', 40, 'Mon-Wed', '']]);
  const notices = [];
  assert.equal(api.updateTeacherSlot_('Bina', 'Monday', 1, '6th A - Maths', [], notices), '');
  assert.ok(notices.some(n => /Bina is marked unavailable on Monday/.test(n)));
  classEdit(env, api, '6th A', 7, 3, 'Asha');
  classEdit(env, api, '6th A', 7, 3, 'Bina');
  assert.equal(ms(ss)[0][5], 'Bina');
  assert.match(lastToast(env).msg, /Bina is marked unavailable on Monday/);
});

test('Teacher_Day_View refuses a teacher name that is not on the Teachers tab', () => {
  const { api, ss } = small([['Monday', 1, '6th A', 'Middle', 'Maths', 'Asha', '', '']], STAFF);
  const err = api.updateTeacherSlot_('Binaa', 'Monday', 1, '6th A - Maths', [], []);
  assert.match(err, /"Binaa" is not on the Teachers tab/);
  assert.equal(ms(ss)[0][5], 'Asha');
});

test('An edit with B3 cleared tells the editor it was not saved', () => {
  const { env, api } = small([['Monday', 1, '6th A', 'Middle', 'Maths', 'Asha', '', '']], STAFF);
  api.ClassViewManager.renderClassView('6th A');
  env.ss.getSheetByName('Class_View').getRange('B3').setValue('');
  edit(env, api, 'Class_View', 7, 3, 'Bina');
  assert.match(lastToast(env).msg, /Pick a class in B3 first/);
  api.TeacherViewManager.renderTeacherView('Asha');
  env.ss.getSheetByName('Teacher_View').getRange('B3').setValue('');
  edit(env, api, 'Teacher_View', 6, 3, 'FREE');
  assert.match(lastToast(env).msg, /Pick a teacher in B3 first/);
});

// ───────────────────────────────────────────── Teachers / Classes tabs

test('Editing the Teachers tab refreshes DAY OFF, dropdowns and Clash Status; a rename says what still uses the old name', () => {
  const { env, api, ss } = small([['Monday', 1, '6th A', 'Middle', 'Maths', 'Asha', '', '']], STAFF);
  api.TeacherDayViewManager.renderTeacherDayView('Monday', 'All Classes');
  edit(env, api, 'Teachers', 3, 4, 'Monday'); // Bina off on Mondays
  const tdv = ss.getSheetByName('Teacher_Day_View');
  const binaRow = plain(tdv.getRange(6, 1, 4, 3).getValues()).find(r => r[0] === 'Bina');
  assert.equal(binaRow[1], 'PRAYER');
  assert.equal(binaRow[2], 'DAY OFF');
  // A new teacher shows up in the views and the dropdown straight away
  edit(env, api, 'Teachers', 6, 1, 'Esha');
  assert.ok(plain(tdv.getRange(6, 1, 6, 1).getValues()).some(r => r[0] === 'Esha'));
  // Renaming Asha: Master_Schedule is not rewritten, the editor is told
  edit(env, api, 'Teachers', 2, 1, 'Asha Rao');
  assert.equal(ms(ss)[0][5], 'Asha');
  assert.match(lastToast(env).msg, /1 period\(s\) in Master_Schedule still use the old name "Asha"/);
  assert.match(String(ss.getSheetByName('Master_Schedule').getRange(2, 8).getValue()), /Unregistered/i);
});

test('Days Unavailable is read in the ways people write it; unreadable text is reported', () => {
  const { api } = small([], STAFF);
  const d = t => plain(api.parseDaysList_(t));
  assert.deepEqual(d('Monday, Friday').days, ['Monday', 'Friday']);
  assert.deepEqual(d('Monday & Friday').days, ['Monday', 'Friday']);
  assert.deepEqual(d('monday and friday').days, ['Monday', 'Friday']);
  assert.deepEqual(d('Mondays').days, ['Monday']);
  assert.deepEqual(d('Mon-Wed').days, ['Monday', 'Tuesday', 'Wednesday']);
  assert.deepEqual(d('Thursday to Saturday').days, ['Thursday', 'Friday', 'Saturday']);
  assert.deepEqual(d('Monday Friday').days, ['Monday', 'Friday']);
  assert.deepEqual(d('sat ;').days, ['Saturday']);
  assert.deepEqual(d('').days, []);
  // Any unreadable part: use none of the value's days and report the whole value
  assert.deepEqual(d('Monday, Funday'), { days: [], unreadable: ['Monday, Funday'] });
  assert.deepEqual(d('Except Saturday').days, []);
  assert.deepEqual(d('Monday (morning)').days, []);
  assert.deepEqual(d('Thursday-Monday').days, ['Thursday', 'Friday', 'Saturday', 'Monday']);
  assert.deepEqual(d('N/A'), { days: [], unreadable: [] });
  assert.deepEqual(d('None'), { days: [], unreadable: [] });

  const t = small([['Monday', 1, '6th A', 'Middle', 'Maths', 'Asha', '', '']], [['Asha', '', 40, 'Mondays', ''], ['Bina', '', 40, 'every alternate day', '']]);
  t.api.runValidation();
  const msg = t.env.alerts[t.env.alerts.length - 1];
  assert.match(msg, /Unavailable day: 1/);
  assert.match(msg, /could not be read[\s\S]*Bina: "every alternate day"/);
});

// ───────────────────────────────────────────── Master_Schedule edits

test('Multi-cell Master_Schedule edits are written to Change_Log, one line per row', () => {
  const { env, api, ss } = small([
    ['Monday', 1, '6th A', 'Middle', 'Maths', 'Asha', '', ''],
    ['Monday', 1, '6th B', 'Middle', 'Science', 'Bina', '', ''],
  ], STAFF);
  const range = ss.getSheetByName('Master_Schedule').getRange(2, 6, 2, 1);
  range.setValues([['Chetan'], ['Deepa']]);
  api.onEdit(editEvent(env, range, uid(env)));
  const log = logRows(ss);
  assert.equal(log.length, 2);
  assert.deepEqual(log.map(r => [r[5], r[7]]), [['6th A', 'Teacher: Chetan'], ['6th B', 'Teacher: Deepa']]);
  assert.match(log[0][6], /previous values not reported/);
});

test('Notes in the owner columns of Master_Schedule skip the refresh', () => {
  const { env, api, ss } = small([['Monday', 1, '6th A', 'Middle', 'Maths', 'Asha', '', '']], STAFF);
  api.ClassViewManager.renderClassView('6th A');
  const cv = ss.getSheetByName('Class_View');
  ss.getSheetByName('Master_Schedule').getRange(2, 5).setValue('Art'); // changed behind the view's back
  edit(env, api, 'Master_Schedule', 2, 10, 'my note');
  assert.equal(cv.getRange(6, 3).getValue(), 'Maths', 'a note in column J does not re-render the views');
  edit(env, api, 'Master_Schedule', 2, 7, 'R12');
  assert.equal(cv.getRange(6, 3).getValue(), 'Art', 'a Room edit does');
});

test('When the installable trigger stops running, the editor is warned that edits are not saved', () => {
  const { env, api, ss } = small([['Monday', 1, '6th A', 'Middle', 'Maths', 'Asha', '', '']], STAFF);
  api.ClassViewManager.renderClassView('6th A');
  const range = ss.getSheetByName('Class_View').getRange(7, 3);
  range.setValue('Bina');
  // The simple trigger fires, the installable one never does (deleted / owner lost access)
  env.properties.LAST_SYNC_AT = String(Date.now() - 5 * 60000);
  api.onEdit(editEvent(env, range, { oldValue: 'Asha' }));
  assert.equal(lastToast(env).title, '⚠️ Sync not running');
  // When it does run, no warning
  const before = ss.toasts.length;
  api.onEdit(editEvent(env, range, Object.assign({ oldValue: 'Asha' }, uid(env))));
  api.onEdit(editEvent(env, range, { oldValue: 'Asha' }));
  assert.ok(!ss.toasts.slice(before).some(t => t.title === '⚠️ Sync not running'));
  // Edits on unrelated tabs never wait or warn
  const n = ss.toasts.length;
  env.properties.LAST_SYNC_AT = '0';
  const other = ss.getSheetByName('Cover_Manager').getRange(5, 1);
  other.setValue('x');
  api.onEdit(editEvent(env, other, {}));
  assert.equal(ss.toasts.length, n);
});

// ───────────────────────────────────────────── validation / names / grid

test('Check For Conflicts reads a Period typed as text "03" as period 3', () => {
  const { api, ss } = small([
    ['Monday', 3, '6th A', 'Middle', 'Maths', 'Asha', '', ''],
    ['Monday', '03', '6th B', 'Middle', 'Science', 'Asha', '', ''],
  ], STAFF);
  api.validateSchedule_();
  assert.match(String(ss.getSheetByName('Master_Schedule').getRange(3, 8).getValue()), /Teacher Clash/);
});

test('Names with a zero-width space or decomposed Hindi still match', () => {
  const { api } = small([], STAFF);
  const P = api.ScheduleParser;
  assert.ok(P.sameName('Bina\u200B', 'bina'));
  assert.ok(P.sameName('\u095B\u093E\u0928', '\u091C\u093C\u093E\u0928'));
});

test('Master Grid: a "6th A, 6th B" row shows under both classes; a duplicate day does not make an entry look daily', () => {
  const rows = ['Monday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']
    .map(d => [d, 1, '6th A', 'Middle', 'Maths', 'Asha', '', '']);
  rows.push(['Saturday', 2, '6th A, 6th B', 'Middle', 'Yoga', 'Bina', '', '']);
  const { api, ss } = small(rows, STAFF);
  api.generateMasterGrid(true);
  const g = plain(ss.getSheetByName('Master_Grid_View').getDataRange().getValues());
  const a = g.find(r => r[0] === '6th A'), b = g.find(r => r[0] === '6th B');
  assert.equal(a[1], 'PRAYER');
  assert.match(a[2], /Maths \(Asha\) \[/, 'not shown as every day');
  assert.match(a[3], /Yoga \(Bina\) \[Sat\]/);
  assert.match(b[3], /Yoga \(Bina\) \[Sat\]/);
  assert.ok(!g.some(r => r[0] === '6th A, 6th B'));
});

// ───────────────────────────────────────────── Cover Manager

test('Cover Manager: both absent teachers of a split period get their own cover', () => {
  const { api, ss } = small([['Monday', 2, '6th B', 'Middle', 'Sanskrit / Hindi', 'Asha / Bina', '', '']], STAFF);
  const date = '2026-10-26'; // a Monday
  api.apiAssignCover(date, 'Asha', 2, '6th B', 'Chetan');
  api.apiAssignCover(date, 'Bina', 2, '6th B', 'Deepa');
  assert.equal(ss.getSheetByName('Cover_Manager').getLastRow(), 3);
  assert.throws(() => api.apiAssignCover(date, 'Asha', 2, '6th B', 'Deepa'), /already covered/);
});

test('Cover Manager: one cover teacher can take every section of a combined lecture, and is offered first', () => {
  const { api } = small([
    ['Monday', 4, '6th A', 'Middle', 'Computer', 'Asha', '', ''],
    ['Monday', 4, '6th B', 'Middle', 'Computer', 'Asha', '', ''],
  ], STAFF);
  const date = '2026-10-26';
  api.apiAssignCover(date, 'Asha', 4, '6th A', 'Chetan');
  assert.equal(plain(api.apiGetAvailableTeachers('Monday', 4, date, 'Asha'))[0], 'Chetan');
  assert.ok(api.apiAssignCover(date, 'Asha', 4, '6th B', 'Chetan'));
  // ...but not a lesson of a different absent teacher
  assert.ok(!plain(api.apiGetAvailableTeachers('Monday', 4, date, 'Bina')).includes('Chetan'));
});

test('Cover Manager refuses a cover for a lesson the absent teacher does not teach', () => {
  const { api } = small([['Monday', 1, '6th A', 'Middle', 'Maths', 'Asha', '', '']], STAFF);
  assert.throws(() => api.apiAssignCover('2026-10-26', 'Asha', 1, '6th B', 'Chetan'), /Asha does not teach 6th B in period 1/);
});

// ───────────────────────────────────────────── snapshot / import / setup / web app

test('Save Snapshot asks before replacing the existing snapshot, and refuses an empty Teachers tab', () => {
  const { env, api, ss } = small([['Monday', 1, '6th A', 'Middle', 'Maths', 'Asha', '', '']], STAFF);
  api.saveCurrentSheetAsPrefillData();
  const backup = ss.getSheetByName('_Prefill_Data_Backup');
  ss.getSheetByName('Master_Schedule').getRange(2, 5).setValue('Art');
  env.answers.push('NO');
  api.saveCurrentSheetAsPrefillData();
  assert.equal(backup.getRange(2, 17).getValue(), 'Maths', 'NO keeps the old snapshot');
  env.answers.push('YES');
  api.saveCurrentSheetAsPrefillData();
  assert.equal(backup.getRange(2, 17).getValue(), 'Art');
  ss.getSheetByName('Teachers').getRange(2, 1, 4, 5).clearContent();
  api.saveCurrentSheetAsPrefillData();
  assert.match(env.alerts[env.alerts.length - 1], /Not saved: the Teachers or Master_Schedule tab is empty/);
});

test('Import writes one line to Change_Log', () => {
  const env = createEnvironment();
  const { api } = loadProject(ROOT, env);
  api.setupInitialSpreadsheet();
  env.answers.push('YES');
  api.importExcelData();
  const log = logRows(env.ss);
  assert.equal(log.length, 1);
  assert.match(log[0][7], /built-in timetable: 1344 schedule rows/);
});

test('Setup / Repair renames and hides the old Teacher_Availability_Grid tab, and leaves the owner\'s own tabs unstyled', () => {
  const { env, api, ss } = small([], STAFF);
  ss.insertSheet('Teacher_Availability_Grid').getRange(1, 1).setValue('my note');
  const own = ss.insertSheet('Exam Duty');
  own.getRange(1, 1, 2, 2).setValues([['Name', 'Room'], ['Asha', '12']]);
  own.setFrozenRows(3);
  api.setupInitialSpreadsheet();
  const old = ss.getSheetByName('Teacher_Availability_Grid (old)');
  assert.ok(old && old.isSheetHidden());
  assert.equal(old.getRange(1, 1).getValue(), 'my note');
  assert.match(env.alerts[env.alerts.length - 1], /renamed "Teacher_Availability_Grid \(old\)" and hidden/);
  assert.equal(own.getFrozenRows(), 3, 'frozen rows on the owner\'s tab kept');
});

test('Web app teacher page shows DAY OFF on a teacher\'s day off', () => {
  const { api } = small([['Tuesday', 1, '6th A', 'Middle', 'Maths', 'Bina', '', '']], [['Asha', '', 40, '', ''], ['Bina', '', 40, 'Monday', '']]);
  const t = plain(api.apiWebGetTimetable('teacher', 'Bina'));
  assert.equal(t.rows[0].cells[0], 'DAY OFF');
  assert.match(t.rows[1].cells[0], /6th A/);
  assert.equal(t.rows[2].cells[0], '');
});
