'use strict';
/**
 * Regression tests for the round-3 deep check of 2026-10-24 (regression review of the
 * whole change + an office day that follows docs/SOP.md literally). Each test failed
 * before its fix.
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
const msSheet = ss => ss.getSheetByName('Master_Schedule');
const ms = (ss, w) => plain(msSheet(ss).getRange(2, 1, msSheet(ss).getLastRow() - 1, w || 8).getValues());
const lastToast = env => env.ss.toasts[env.ss.toasts.length - 1] || {};
const uid = env => ({ triggerUid: env.properties.EDIT_TRIGGER_UID });
function edit(env, api, sheetName, row, col, value, extra) {
  const range = env.ss.getSheetByName(sheetName).getRange(row, col);
  const old = range.getValue();
  range.setValue(value);
  api.onEdit(editEvent(env, range, Object.assign({ oldValue: old === '' ? undefined : old }, uid(env), extra || {})));
}
function classEdit(env, api, row, col, value) {
  api.ClassViewManager.renderClassView('6th A');
  edit(env, api, 'Class_View', row, col, value);
}
const logRows = ss => { const s = ss.getSheetByName('Change_Log'); return s ? plain(s.getRange(2, 1, Math.max(1, s.getLastRow() - 1), 8).getValues()) : []; };

test('Freeing a period in Class_View keeps its Master_Schedule row, room and the owner\'s own columns', () => {
  const { env, api, ss } = small([['Monday', 1, '6th A', 'Middle', 'Maths', 'Asha', 'Lab 2', '', '', 'my note']], STAFF);
  classEdit(env, api, 7, 3, '');   // teacher cell
  classEdit(env, api, 6, 3, '');   // subject cell
  const rows = ms(ss, 10);
  assert.equal(rows.length, 1, 'the row is not deleted');
  assert.deepEqual([rows[0][2], rows[0][4], rows[0][5], rows[0][6], rows[0][9]], ['6th A', '', '', 'Lab 2', 'my note']);
  // The free period shows as free and the next edit reuses the same row
  classEdit(env, api, 6, 3, 'Science');
  classEdit(env, api, 7, 3, 'Bina');
  assert.equal(ms(ss, 10).length, 1);
  assert.deepEqual(ms(ss, 10)[0].slice(4, 7), ['Science', 'Bina', 'Lab 2']);
});

test('Typing the split subjects first keeps the current teacher on the first group (SOP order works)', () => {
  const { env, api, ss } = small([['Monday', 1, '6th A', 'Middle', 'Maths', 'Asha', '', '']], STAFF);
  classEdit(env, api, 6, 3, 'Hindi / IT / Art');
  assert.deepEqual(ms(ss)[0].slice(4, 6), ['Hindi / IT / Art', 'Asha / - / -']);
  classEdit(env, api, 7, 3, 'Asha / Bina / Deepa');
  assert.deepEqual(ms(ss)[0].slice(4, 6), ['Hindi / IT / Art', 'Asha / Bina / Deepa']);
});

test('Pasting one subject over a split Master_Schedule row is flagged "Check this cell"', () => {
  const { env, api, ss } = small([['Monday', 1, '6th A', 'Middle', 'Sanskrit / Beauty', 'Asha / Bina', '', '']], STAFF);
  const range = msSheet(ss).getRange(2, 5);
  range.setValue('Music');
  api.onEdit(editEvent(env, range, uid(env))); // a paste: Sheets sends no oldValue
  assert.equal(ms(ss)[0][4], 'Music', 'a paste is never rewritten');
  assert.equal(lastToast(env).title, 'Check this cell');
  assert.match(lastToast(env).msg, /one subject with 2 teacher groups/);
});

test('A multi-cell Master_Schedule paste writes Change_Log once, with each row\'s own day and period', () => {
  const { env, api, ss } = small([
    ['Monday', 1, '6th A', 'Middle', 'Maths', 'Asha', '', ''],
    ['Tuesday', 3, '6th B', 'Middle', 'Science', 'Bina', '', ''],
  ], STAFF);
  const range = msSheet(ss).getRange(2, 6, 2, 1);
  range.setValues([['Chetan'], ['Deepa']]);
  let writes = 0;
  const log = ss.getSheetByName('Change_Log') || ss.insertSheet('Change_Log');
  const orig = log.getRange.bind(log);
  log.getRange = function () { const r = orig.apply(null, arguments); const sv = r.setValues.bind(r); r.setValues = v => { writes++; return sv(v); }; return r; };
  api.onEdit(editEvent(env, range, uid(env)));
  assert.equal(writes, 1);
  const all = plain(orig(1, 1, log.getLastRow(), 8).getValues());
  assert.deepEqual(all.slice(-2).map(r => [r[3], r[4], r[5]]), [['Monday', 1, '6th A'], ['Tuesday', 3, '6th B']]);
});

test('Change_Log text starting with "-" is stored as text, not read as a formula', () => {
  const { api } = small([], STAFF);
  assert.equal(api.asText_('- / IT — Asha'), "'- / IT — Asha");
  assert.equal(api.asText_('Maths'), 'Maths');
});

test('Cover Manager: a cover teacher is not booked into both classes of an absent teacher\'s clash', () => {
  const { api } = small([
    ['Monday', 2, '6th A', 'Middle', 'Science', 'Asha', '', ''],
    ['Monday', 2, '7th A', 'Middle', 'Science', 'Asha', '', ''],   // same subject, different grade: a clash
  ], STAFF.concat(['Esha']), ['6th A', '6th B', '7th A']);
  const date = '2026-10-26';
  api.apiAssignCover(date, 'Asha', 2, '6th A', 'Chetan');
  assert.ok(!plain(api.apiGetAvailableTeachers('Monday', 2, date, 'Asha')).includes('Chetan'));
  assert.throws(() => api.apiAssignCover(date, 'Asha', 2, '7th A', 'Chetan'), /already covering/);
  assert.ok(api.apiAssignCover(date, 'Asha', 2, '7th A', 'Deepa'));
});

test('Cover Manager: a section of an already-covered combined row cannot get a second cover', () => {
  const { api } = small([['Monday', 3, '6th A, 6th B', 'Middle', 'Art', 'Asha', '', '']], STAFF);
  const date = '2026-10-26';
  api.apiAssignCover(date, 'Asha', 3, '6th A, 6th B', 'Bina');
  assert.throws(() => api.apiAssignCover(date, 'Asha', 3, '6th A', 'Chetan'), /already covered/);
});

test('Registry owner columns and Class_View titles do not trigger a refresh or a false "Not saved"', () => {
  const { env, api, ss } = small([['Monday', 1, '6th A', 'Middle', 'Maths', 'Asha', '', '']], STAFF);
  api.ClassViewManager.renderClassView('6th A');
  const cv = ss.getSheetByName('Class_View');
  msSheet(ss).getRange(2, 5).setValue('Art'); // changed behind the view's back
  edit(env, api, 'Teachers', 2, 6, '98765 43210'); // phone number in an owner column
  assert.equal(cv.getRange(6, 3).getValue(), 'Maths', 'no refresh for an owner column');
  edit(env, api, 'Teachers', 2, 4, 'Saturday');
  assert.equal(cv.getRange(6, 3).getValue(), 'Art', 'Days Unavailable does refresh');

  cv.getRange('B3').setValue('');
  const before = ss.toasts.length;
  edit(env, api, 'Class_View', 1, 1, 'My title');
  assert.ok(!ss.toasts.slice(before).some(t => t.title === 'Not saved'));
});
