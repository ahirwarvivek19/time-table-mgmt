'use strict';
/**
 * Editing from the views: every gap found in the editing review has a test here that
 * reproduces it. Class_View cells: row 6 = Monday subject, row 7 = Monday teacher,
 * column 1 + period. Teacher_View cells: row 6 = Monday, column 1 + period.
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('path');
const { createEnvironment, loadProject, editEvent } = require('./harness/gas-mock');

const ROOT = path.resolve(__dirname, '..');
const plain = x => JSON.parse(JSON.stringify(x));

function small(schedule, teachers, classes) {
  const env = createEnvironment();
  const { api } = loadProject(ROOT, env);
  api.setupInitialSpreadsheet();
  const put = (name, rows, w) => rows.length && env.ss.getSheetByName(name).getRange(2, 1, rows.length, w).setValues(rows);
  put('Teachers', teachers.map(t => [t, '', 40, '', '']), 5);
  put('Classes', (classes || ['6th A', '6th B', '7th A']).map(c => [c, 'Middle', '']), 3);
  put('Master_Schedule', schedule, 8);
  return { env, api, ss: env.ss };
}
const rows = ss => { const sh = ss.getSheetByName('Master_Schedule'); return sh.getLastRow() < 2 ? [] : sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues(); };
const slot = (ss, cls, day, period) => rows(ss).find(r => r[2] === cls && r[0] === day && r[1] == period);
const lastToast = env => env.ss.toasts[env.ss.toasts.length - 1] || {};

/** Types into Class_View (one cell, or a block when `values` is 2-D). */
function classEdit(env, api, cls, row, col, values) {
  api.ClassViewManager.renderClassView(cls);
  const cv = env.ss.getSheetByName('Class_View');
  const block = Array.isArray(values) ? values : [[values]];
  const range = cv.getRange(row, col, block.length, block[0].length);
  range.setValues(block);
  api.onEdit(editEvent(env, range, { triggerUid: env.properties.EDIT_TRIGGER_UID }));
  return cv;
}
function teacherEdit(env, api, teacher, row, col, value) {
  api.TeacherViewManager.renderTeacherView(teacher);
  const range = env.ss.getSheetByName('Teacher_View').getRange(row, col);
  range.setValue(value);
  api.onEdit(editEvent(env, range, { triggerUid: env.properties.EDIT_TRIGGER_UID }));
}

const SPLIT = [['Monday', 1, '6th A', 'Middle', 'Hindi / IT', 'Asha / Bina + Chetan', '', '']];
const STAFF = ['Asha', 'Bina', 'Chetan', 'Deepa'];

// ───────────────────────────────────────────── Class_View: split periods

test('Class_View: one teacher picked on a split period whose groups are all taken is refused and the view is put back', () => {
  const { env, api, ss } = small(SPLIT, STAFF);
  const cv = classEdit(env, api, '6th A', 7, 2, 'Deepa');
  assert.deepEqual(slot(ss, '6th A', 'Monday', 1).slice(4, 6), ['Hindi / IT', 'Asha / Bina + Chetan'], 'nothing saved');
  assert.equal(lastToast(env).title, 'Not saved');
  assert.match(lastToast(env).msg, /all 2 groups already have a teacher/);
  assert.equal(cv.getRange(7, 2).getValue(), 'Asha / Bina + Chetan', 'typed text reverted');
});

test('Class_View: picking teachers one at a time from the dropdown fills the groups in subject order', () => {
  const { env, api, ss } = small([['Monday', 1, '6th A', 'Middle', 'Hindi / IT / Art', '', '', '']], STAFF);
  classEdit(env, api, '6th A', 7, 2, 'Asha');
  assert.equal(slot(ss, '6th A', 'Monday', 1)[5], 'Asha / - / -');
  classEdit(env, api, '6th A', 7, 2, 'Bina');
  classEdit(env, api, '6th A', 7, 2, 'Deepa');
  assert.equal(slot(ss, '6th A', 'Monday', 1)[5], 'Asha / Bina / Deepa');
  classEdit(env, api, '6th A', 7, 2, 'Bina'); // already there: no change, no error
  assert.equal(slot(ss, '6th A', 'Monday', 1)[5], 'Asha / Bina / Deepa');
  classEdit(env, api, '6th A', 7, 2, ''); // clearing the teacher cell starts again
  assert.equal(slot(ss, '6th A', 'Monday', 1)[5], '');
  // a single-subject period keeps the normal dropdown: the pick replaces the teacher
  classEdit(env, api, '6th B', 6, 2, [['Maths'], ['Asha']]);
  classEdit(env, api, '6th B', 7, 2, 'Chetan');
  assert.equal(slot(ss, '6th B', 'Monday', 1)[5], 'Chetan');
});

test('Teacher dropdowns list names A→Z', () => {
  const { api, ss } = small([], ['Zoya', 'asha', 'Bina']);
  api.ClassViewManager.renderClassView('6th A');
  assert.deepEqual(plain(ss.getSheetByName('Class_View').getRange(7, 2).getDataValidation().list), ['asha', 'Bina', 'Zoya']);
  api.TeacherViewManager.renderTeacherView('Zoya');
  assert.deepEqual(plain(ss.getSheetByName('Teacher_View').getRange('B3').getDataValidation().list), ['asha', 'Bina', 'Zoya']);
  assert.deepEqual(plain(api.apiGetTeachers()), ['asha', 'Bina', 'Zoya']);
});

test('Class_View: a third subject without a third teacher group is refused', () => {
  const { env, api, ss } = small(SPLIT, STAFF);
  classEdit(env, api, '6th A', 6, 2, 'Hindi / IT / Art');
  assert.equal(slot(ss, '6th A', 'Monday', 1)[4], 'Hindi / IT');
  assert.match(lastToast(env).msg, /3 groups .* but 2 teacher groups/);
});

test('Class_View: subject and teacher typed together in full form are saved aligned', () => {
  const { env, api, ss } = small(SPLIT, STAFF);
  classEdit(env, api, '6th A', 6, 2, [['Hindi / IT / Art'], ['Asha / Bina + Chetan / Deepa']]);
  const r = slot(ss, '6th A', 'Monday', 1);
  assert.deepEqual([r[4], r[5]], ['Hindi / IT / Art', 'Asha / Bina + Chetan / Deepa']);
  const lanes = plain(api.ScheduleParser.parseLanes({ Subject: r[4], Teacher: r[5] }).lanes).map(l => [l.subject, l.teachers]);
  assert.deepEqual(lanes, [['Hindi', ['Asha']], ['IT', ['Bina', 'Chetan']], ['Art', ['Deepa']]]);
  assert.ok(ss.getSheetByName('Subjects').getRange(1, 1, 50, 1).getValues().some(v => v[0] === 'Art'), 'new subject added to Subjects');
});

test('Class_View: "-" keeps an empty group in place', () => {
  const { env, api, ss } = small(SPLIT, STAFF);
  classEdit(env, api, '6th A', 7, 2, 'Asha / -');
  assert.equal(slot(ss, '6th A', 'Monday', 1)[5], 'Asha / -');
  assert.notEqual(lastToast(env).title, 'Not saved');
});

// ───────────────────────────────────────────── Class_View: names and clashes

test('Class_View: an unknown teacher name is refused', () => {
  const { env, api, ss } = small([['Monday', 1, '6th A', 'Middle', 'Maths', 'Asha', '', '']], STAFF);
  classEdit(env, api, '6th A', 7, 2, 'Ashaa');
  assert.equal(slot(ss, '6th A', 'Monday', 1)[5], 'Asha');
  assert.match(lastToast(env).msg, /unknown teacher "Ashaa"/);
});

test('Class_View: a teacher name typed into the subject row is refused', () => {
  const { env, api, ss } = small([['Monday', 1, '6th A', 'Middle', 'Maths', 'Asha', '', '']], STAFF);
  classEdit(env, api, '6th A', 6, 2, 'Bina');
  assert.equal(slot(ss, '6th A', 'Monday', 1)[4], 'Maths');
  assert.match(lastToast(env).msg, /"Bina" is a teacher name/);
});

test('Class_View: booking a busy teacher is refused; the same subject in the same grade is a combined lecture', () => {
  const { env, api, ss } = small([
    ['Monday', 1, '6th A', 'Middle', 'Maths', 'Asha', '', ''],
    ['Monday', 1, '6th B', 'Middle', 'Science', 'Bina', '', ''],
  ], STAFF);
  classEdit(env, api, '6th A', 7, 2, 'Bina');
  assert.equal(slot(ss, '6th A', 'Monday', 1)[5], 'Asha', 'clash refused');
  assert.match(lastToast(env).msg, /Bina already teaches 6th B \(Science\)/);
  // Same subject + same grade = combined lecture, allowed
  classEdit(env, api, '6th A', 6, 2, [['Science'], ['Bina']]);
  assert.deepEqual(slot(ss, '6th A', 'Monday', 1).slice(4, 6), ['Science', 'Bina']);
  const note = env.ss.getSheetByName('Class_View').getRange(6, 2).getNote();
  assert.match(note, /Taught together:\n• Science — Bina: with 6th B/, 'view redrawn with the new note');
});

test('Class_View: an edit next to an existing clash is not blocked by that clash', () => {
  const { env, api, ss } = small([
    ['Monday', 1, '6th A', 'Middle', 'Maths', 'Asha', '', ''],
    ['Monday', 1, '7th A', 'Middle', 'Maths', 'Asha', '', ''], // different grade = existing clash
  ], STAFF, ['6th A', '6th B', '7th A']);
  classEdit(env, api, '6th A', 6, 2, 'Mathematics');
  assert.equal(slot(ss, '6th A', 'Monday', 1)[4], 'Mathematics');
});

// ───────────────────────────────────────────── Class_View: clearing, colours, several classes

test('Class_View: clearing only the subject is refused; clearing both frees the period', () => {
  const { env, api, ss } = small([['Monday', 1, '6th A', 'Middle', 'Maths', 'Asha', '', '']], STAFF);
  classEdit(env, api, '6th A', 6, 2, '');
  assert.equal(slot(ss, '6th A', 'Monday', 1)[4], 'Maths');
  assert.match(lastToast(env).msg, /subject was cleared but the teacher was not/);
  classEdit(env, api, '6th A', 6, 2, [[''], ['']]);
  assert.equal(slot(ss, '6th A', 'Monday', 1), undefined, 'row removed');
});

test('Class_View redraws after an edit: a lecture that stops being combined loses its colour and note', () => {
  const { env, api } = small([
    ['Monday', 1, '6th A', 'Middle', 'Maths', 'Asha', '', ''],
    ['Monday', 1, '6th B', 'Middle', 'Maths', 'Asha', '', ''],
  ], STAFF);
  api.ClassViewManager.renderClassView('6th A');
  const cv = env.ss.getSheetByName('Class_View');
  assert.equal(cv.getRange(6, 2).getBackground(), '#F5EEF8');
  classEdit(env, api, '6th A', 6, 2, [['Art'], ['Deepa']]);
  assert.equal(cv.getRange(6, 2).getBackground(), '#EBF5FB');
  assert.equal(cv.getRange(6, 2).getNote(), '');
});

test('Class_View with two classes selected writes one row per class', () => {
  const { env, api, ss } = small([], STAFF);
  classEdit(env, api, '6th A, 6th B', 6, 3, [['Music'], ['Deepa']]);
  assert.deepEqual(slot(ss, '6th A', 'Monday', 2).slice(4, 6), ['Music', 'Deepa']);
  assert.deepEqual(slot(ss, '6th B', 'Monday', 2).slice(4, 6), ['Music', 'Deepa']);
  assert.equal(rows(ss).filter(r => /,/.test(r[2])).length, 0, 'no combined "6th A, 6th B" row');
});

// ───────────────────────────────────────────── Teacher_View

test('Teacher_View: typing a class with a new subject replaces the lesson instead of splitting the class', () => {
  const { env, api, ss } = small([['Monday', 1, '6th A', 'Middle', 'Maths', 'Asha', '', '']], STAFF);
  teacherEdit(env, api, 'Bina', 6, 2, '6th A - Science');
  assert.deepEqual(slot(ss, '6th A', 'Monday', 1).slice(4, 6), ['Science', 'Bina']);
  assert.equal(lastToast(env).title, 'Saved');
  assert.match(lastToast(env).msg, /6th A Maths — Asha replaced by Bina/);
});

test('Teacher_View: a bare class name takes over the lesson, and the old class is reported as uncovered', () => {
  const { env, api, ss } = small([
    ['Monday', 1, '6th A', 'Middle', 'Maths', 'Asha', '', ''],
    ['Monday', 1, '6th B', 'Middle', 'Science', 'Bina', '', ''],
  ], STAFF);
  teacherEdit(env, api, 'Asha', 6, 2, '6th B');
  assert.deepEqual(slot(ss, '6th B', 'Monday', 1).slice(4, 6), ['Science', 'Asha'], 'replaced, not co-teaching');
  assert.match(lastToast(env).msg, /6th A Maths now has no teacher/);
  assert.match(lastToast(env).msg, /Bina replaced by Asha/);
});

test('Teacher_View: a split subject typed here is refused (groups are set up in Class_View)', () => {
  const { env, api, ss } = small([['Monday', 1, '6th A', 'Middle', 'Maths', 'Asha', '', '']], STAFF);
  teacherEdit(env, api, 'Bina', 6, 2, '6th A - Maths / Science');
  assert.equal(slot(ss, '6th A', 'Monday', 1)[5], 'Asha');
  assert.match(lastToast(env).msg, /type one subject here/);
});

// ───────────────────────────────────────────── day names

test('Day names in any case show in every view, and an edit in Master_Schedule writes them back as "Monday"', () => {
  const { env, api, ss } = small([['monday', 1, '6th A', 'Middle', 'Maths', 'Asha', '', '']], STAFF);
  api.TeacherViewManager.renderTeacherView('Asha');
  assert.match(String(ss.getSheetByName('Teacher_View').getRange(6, 2).getValue()), /6th A/, 'Teacher_View');
  api.TeacherDayViewManager.renderTeacherDayView('Monday', 'All Classes');
  const tdv = ss.getSheetByName('Teacher_Day_View');
  assert.match(String(tdv.getRange(6, 2).getValue()), /6th A/, 'Teacher_Day_View');
  const web = plain(api.apiWebGetTimetable('teacher', 'Asha'));
  assert.match(web.rows[0].cells[0], /6th A · Maths/, 'web page');
  const ms = ss.getSheetByName('Master_Schedule');
  const range = ms.getRange(2, 1);
  range.setValue('mon');
  api.onEdit(editEvent(env, range, { triggerUid: env.properties.EDIT_TRIGGER_UID, oldValue: 'monday' }));
  assert.equal(ms.getRange(2, 1).getValue(), 'Monday');
});

test('Clash check treats "monday" and "Monday" as the same day', () => {
  const { api, ss } = small([
    ['monday', 1, '6th A', 'Middle', 'Maths', 'Asha', '', ''],
    ['Monday', 1, '6th B', 'Middle', 'Science', 'Asha', '', ''],
  ], STAFF);
  api.validateSchedule_();
  assert.match(String(slot(ss, '6th A', 'monday', 1)[7]), /Teacher Clash \(Asha\)/);
});

// ───────────────────────────────────────────── after every edit

test('After a view edit, Clash Status, teacher load and the change log are up to date', () => {
  const { env, api, ss } = small([
    ['Monday', 1, '6th A', 'Middle', 'Maths', 'Asha', '', ''],
    ['Monday', 1, '6th B', 'Middle', 'Science', 'Bina', '', ''],
    ['Monday', 2, '6th B', 'Middle', 'Science', 'Asha', '', ''],
  ], STAFF);
  api.validateSchedule_();
  teacherEdit(env, api, 'Asha', 6, 3, 'FREE'); // Monday P2
  const teachers = ss.getSheetByName('Teachers').getRange(2, 1, 4, 5).getValues();
  assert.equal(teachers.find(t => t[0] === 'Asha')[4], 1, 'Total Hours Scheduled recomputed');
  const log = ss.getSheetByName('Change_Log');
  assert.ok(log, 'Change_Log created');
  const last = log.getRange(log.getLastRow(), 1, 1, 8).getValues()[0];
  assert.deepEqual(last.slice(2), ['Teacher_View', 'Monday', 2, '6th B', 'Science — Asha', 'Science']);
});

test('Deleting a Master_Schedule row refreshes the views (change trigger)', () => {
  const { env, api, ss } = small([['Monday', 1, '6th A', 'Middle', 'Maths', 'Asha', '', '']], STAFF);
  api.ClassViewManager.renderClassView('6th A');
  const cv = ss.getSheetByName('Class_View');
  assert.equal(cv.getRange(6, 2).getValue(), 'Maths');
  ss.getSheetByName('Master_Schedule').deleteRow(2);
  api.onScheduleChange({ source: ss, changeType: 'OTHER', triggerUid: env.properties.CHANGE_TRIGGER_UID });
  assert.equal(cv.getRange(6, 2).getValue(), 'Maths', 'other change types are ignored');
  api.onScheduleChange({ source: ss, changeType: 'REMOVE_ROW', triggerUid: 'someone-else' });
  assert.equal(cv.getRange(6, 2).getValue(), 'Maths', 'unrecorded trigger ignored');
  api.onScheduleChange({ source: ss, changeType: 'REMOVE_ROW', triggerUid: env.properties.CHANGE_TRIGGER_UID });
  assert.equal(cv.getRange(6, 2).getValue(), '');
});

test('Refresh adds a warning-only protection to the generated tabs', () => {
  const { api, ss } = small([['Monday', 1, '6th A', 'Middle', 'Maths', 'Asha', '', '']], STAFF);
  api.refreshAllViews(true);
  const p = ss.getSheetByName('Master_Grid_View').getProtections();
  assert.equal(p.length, 1);
  assert.equal(p[0].warningOnly, true);
  api.refreshAllViews(true);
  assert.equal(ss.getSheetByName('Master_Grid_View').getProtections().length, 1, 'not added twice');
});

// ───────────────────────────────────────────── reported: 11th period 2

test('11th A–D period 2: the note says which group is shared with which sections', () => {
  const env = createEnvironment();
  const { api } = loadProject(ROOT, env);
  const slotRows = api.RAW_SCHEDULE.filter(r => r.Day === 'Monday' && r.Period === 2);
  const note = c => api.ScheduleParser.describeSlot(slotRows, c).note;
  // 11th C: IT and Beauty are shared with 11th B and 11th D; its English (E.M) group is its own
  assert.match(note('11th C'), /• IT — Divya Lodval: with 11th B, 11th D/);
  assert.doesNotMatch(note('11th C'), /English \(E\.M\) — Rajesh Kumar Bairagi: with/);
  // 11th A has no IT in period 2 in Timetable.xlsx (its IT is in period 3); its English is shared
  assert.doesNotMatch(note('11th A'), /IT —/);
  assert.match(note('11th A'), /• English — Anuradha Katare: with 11th B, 11th D/);
  // 11th B lists every section per group
  assert.match(note('11th B'), /• IT — Divya Lodval: with 11th C, 11th D/);
  assert.match(note('11th B'), /• English — Anuradha Katare: with 11th A, 11th D/);
});

// ───────────────────────────────────────────── Master_Schedule: same checks as Class_View

/** Types into one Master_Schedule cell the way Sheets does: the new value is in the cell, the old one in e.oldValue. */
function scheduleEdit(env, api, row, col, value) {
  const range = env.ss.getSheetByName('Master_Schedule').getRange(row, col);
  const old = range.getValue();
  range.setValue(value);
  api.onEdit(editEvent(env, range, { triggerUid: env.properties.EDIT_TRIGGER_UID, oldValue: old === '' ? undefined : old }));
}

test('Master_Schedule: picking teachers one at a time on a split row fills the groups in order', () => {
  const { env, api, ss } = small([['Monday', 1, '6th A', 'Middle', 'Hindi / IT', '', '', '']], STAFF);
  scheduleEdit(env, api, 2, 6, 'Asha');
  assert.equal(slot(ss, '6th A', 'Monday', 1)[5], 'Asha / -');
  scheduleEdit(env, api, 2, 6, 'Bina');
  assert.equal(slot(ss, '6th A', 'Monday', 1)[5], 'Asha / Bina');
  scheduleEdit(env, api, 2, 6, 'Deepa');
  assert.equal(slot(ss, '6th A', 'Monday', 1)[5], 'Asha / Bina', 'full: old value kept');
  assert.match(lastToast(env).msg, /all 2 groups already have a teacher/);
});

test('Master_Schedule: unknown names, new clashes and misaligned groups are refused; the old value stays', () => {
  const { env, api, ss } = small([
    ['Monday', 1, '6th A', 'Middle', 'Maths', 'Asha', '', ''],
    ['Monday', 1, '6th B', 'Middle', 'Science', 'Bina', '', ''],
    ['Monday', 2, '6th A', 'Middle', 'Hindi / IT', 'Asha / Bina', '', ''],
  ], STAFF);
  scheduleEdit(env, api, 2, 6, 'Ashaa');
  assert.equal(slot(ss, '6th A', 'Monday', 1)[5], 'Asha');
  assert.match(lastToast(env).msg, /unknown teacher "Ashaa"/);
  scheduleEdit(env, api, 2, 6, 'Bina');
  assert.equal(slot(ss, '6th A', 'Monday', 1)[5], 'Asha');
  assert.match(lastToast(env).msg, /Bina already teaches 6th B \(Science\)/);
  scheduleEdit(env, api, 4, 5, 'Hindi / IT / Art');
  assert.equal(slot(ss, '6th A', 'Monday', 2)[4], 'Hindi / IT');
  assert.match(lastToast(env).msg, /3 groups .* but 2 teacher groups/);
  // A valid edit goes through and is logged
  scheduleEdit(env, api, 2, 6, 'Chetan');
  assert.equal(slot(ss, '6th A', 'Monday', 1)[5], 'Chetan');
  const log = ss.getSheetByName('Change_Log');
  assert.deepEqual(log.getRange(log.getLastRow(), 3, 1, 6).getValues()[0], ['Master_Schedule', 'Monday', 1, '6th A', 'Maths — Asha', 'Maths — Chetan']);
});

test('Master_Schedule: clearing both cells keeps the row; a row still being filled in is not checked', () => {
  const { env, api, ss } = small([['Monday', 1, '6th A', 'Middle', 'Maths', 'Asha', '', '']], STAFF);
  scheduleEdit(env, api, 2, 6, '');
  scheduleEdit(env, api, 2, 5, '');
  assert.equal(rows(ss).length, 1, 'row kept');
  assert.deepEqual(rows(ss)[0].slice(4, 6), ['', '']);
  // New row typed left to right: no class yet, so the teacher is accepted as typed
  const ms = ss.getSheetByName('Master_Schedule');
  ms.getRange(3, 1, 1, 2).setValues([['Tuesday', 1]]);
  scheduleEdit(env, api, 3, 6, 'Anyone');
  assert.equal(ms.getRange(3, 6).getValue(), 'Anyone');
});

// ───────────────────────────────────────────── duplicate rows (two rows, one class, one period)

test('Duplicate rows: Class_View edits the merged period and saves it into one row; Master_Schedule edits only the row typed in', () => {
  const dupRows = () => [
    ['Monday', 1, '6th A', 'Middle', 'Maths', 'Asha', '', ''],
    ['Monday', 1, '6th A', 'Middle', 'Art', 'Deepa', '', ''],
  ];
  // Class_View: the view shows "Maths / Art"; editing the teacher row saves one row
  let { env, api, ss } = small(dupRows(), STAFF);
  api.ClassViewManager.renderClassView('6th A');
  assert.equal(ss.getSheetByName('Class_View').getRange(6, 2).getValue(), 'Maths / Art');
  classEdit(env, api, '6th A', 7, 2, 'Asha / Chetan');
  const left = rows(ss).filter(r => r[2] === '6th A');
  assert.equal(left.length, 1, 'duplicate merged away');
  assert.deepEqual(left[0].slice(4, 6), ['Maths / Art', 'Asha / Chetan']);
  // Master_Schedule: editing the second row leaves the first alone
  ({ env, api, ss } = small(dupRows(), STAFF));
  scheduleEdit(env, api, 3, 6, 'Chetan');
  assert.deepEqual(rows(ss).map(r => r[5]), ['Asha', 'Chetan']);
});
