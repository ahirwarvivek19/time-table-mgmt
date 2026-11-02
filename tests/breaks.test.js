'use strict';
/**
 * Prayer and Lunch columns: every view shows Prayer before period 1 and Lunch after
 * period 4, as one grey cell down the grid. They are not periods: typing in them saves
 * nothing and the view is put back.
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('path');
const { createEnvironment, loadProject, editEvent } = require('./harness/gas-mock');

const ROOT = path.resolve(__dirname, '..');
const plain = x => JSON.parse(JSON.stringify(x));
const STAFF = ['Asha', 'Bina', 'Chetan', 'Deepa'];
const SCHEDULE = [
  ['Monday', 1, '6th A', 'Middle', 'Maths', 'Asha', '', ''],
  ['Monday', 4, '6th A', 'Middle', 'Science', 'Bina', '', ''],
  ['Monday', 5, '6th A', 'Middle', 'Hindi', 'Chetan', '', ''],
];

function small(schedule) {
  const env = createEnvironment();
  const { api } = loadProject(ROOT, env);
  api.setupInitialSpreadsheet();
  const put = (name, rows, w) => rows.length && env.ss.getSheetByName(name).getRange(2, 1, rows.length, w).setValues(rows);
  put('Teachers', STAFF.map(t => [t, '', 40, '', '']), 5);
  put('Classes', ['6th A', '6th B'].map(c => [c, 'Middle', '']), 3);
  put('Master_Schedule', schedule, 8);
  return { env, api, ss: env.ss };
}
const ms = ss => { const s = ss.getSheetByName('Master_Schedule'); return plain(s.getRange(2, 1, s.getLastRow() - 1, 8).getValues()); };
const lastToast = env => env.ss.toasts[env.ss.toasts.length - 1] || {};
function edit(env, sheetName, row, col, values) {
  const block = Array.isArray(values) ? values : [[values]];
  const range = env.ss.getSheetByName(sheetName).getRange(row, col, block.length, block[0].length);
  range.setValues(block);
  return range;
}
const fire = (env, api, range) => api.onEdit(editEvent(env, range, { triggerUid: env.properties.EDIT_TRIGGER_UID }));

test('layout: Prayer, periods 1–4, Lunch, periods 5–8 (columns B–K); the fallback matches PrefillData', () => {
  const { api } = small(SCHEDULE);
  const slots = plain(api.viewSlots_());
  assert.deepEqual(slots.map(s => s.period || s.name), ['Prayer', 1, 2, 3, 4, 'Lunch', 5, 6, 7, 8]);
  assert.equal(api.viewWidth_(), 11);
  assert.deepEqual([1, 4, 5, 8].map(p => api.viewColumnOfPeriod_(p)), [3, 6, 8, 11]);
  assert.deepEqual([1, 2, 3, 7, 8, 12].map(c => api.periodOfViewColumn_(c)), [null, null, 1, null, 5, null]);
  assert.deepEqual(plain(api.viewBreaks_()), plain(api.CONFIG.BREAKS), 'CONFIG.BREAKS is in step with RAW_BREAK_ROWS');
});

test('every view shows the Prayer and Lunch columns as one grey cell with no dropdown or note', () => {
  const { api, ss } = small(SCHEDULE);
  const views = [
    ['Class_View', () => api.ClassViewManager.renderClassView('6th A'), 5, 12],
    ['Teacher_View', () => api.TeacherViewManager.renderTeacherView('Asha'), 5, 6],
    ['Teacher_Day_View', () => api.TeacherDayViewManager.renderTeacherDayView('Monday', 'All Classes'), 5, STAFF.length],
    ['Master_Grid_View', () => api.generateMasterGrid(true), 1, 1],
  ];
  views.forEach(([name, render, headerRow, bodyRows]) => {
    render();
    const sh = ss.getSheetByName(name);
    assert.equal(sh.getRange(headerRow, 2).getValue(), 'Prayer\n9:10–10:00', name);
    assert.equal(sh.getRange(headerRow, 7).getValue(), 'Lunch\n12:50–1:20', name);
    assert.equal(sh.getRange(headerRow, 3).getValue(), 'Period 1\n10:00–10:50', name);
    assert.equal(sh.getRange(headerRow, 8).getValue(), 'Period 5\n1:20–2:00', name);
    assert.equal(sh.getRange(headerRow + 1, 2).getValue(), 'PRAYER', name);
    assert.equal(sh.getRange(headerRow + 1, 7).getValue(), 'LUNCH', name);
    [2, 7].forEach(col => {
      assert.equal(sh.getRange(headerRow + 1, col).getBackground(), '#EAECEE', name + ' grey');
      assert.equal(sh.getRange(headerRow + 1, col).getDataValidation(), null, name + ' no dropdown');
      assert.equal(sh.getRange(headerRow + 1, col).getNote(), '', name + ' no note');
      if (bodyRows > 1) {
        assert.ok(sh._merges.some(m => m[0] === headerRow + 1 && m[1] === col && m[2] === bodyRows && m[3] === 1), name + ' merged down the grid');
      }
    });
  });
  // The periods next to a break are still the real lessons
  api.ClassViewManager.renderClassView('6th A');
  const cv = ss.getSheetByName('Class_View');
  assert.equal(cv.getRange(6, 3).getValue(), 'Maths');
  assert.equal(cv.getRange(6, 6).getValue(), 'Science');
  assert.equal(cv.getRange(6, 8).getValue(), 'Hindi');
  assert.ok(cv.getRange(7, 3).getDataValidation(), 'period cells keep their teacher dropdown');
});

test('typing in a Prayer / Lunch cell saves nothing, puts the view back and says why', () => {
  [
    ['Class_View', api => api.ClassViewManager.renderClassView('6th A'), 6],
    ['Teacher_View', api => api.TeacherViewManager.renderTeacherView('Asha'), 6],
    ['Teacher_Day_View', api => api.TeacherDayViewManager.renderTeacherDayView('Monday', 'All Classes'), 6],
  ].forEach(([name, render, row]) => {
    [2, 7].forEach(col => {
      const { env, api, ss } = small(SCHEDULE);
      render(api);
      fire(env, api, edit(env, name, row, col, '6th A - Maths'));
      assert.deepEqual(ms(ss), plain(SCHEDULE), name + ' col ' + col + ': Master_Schedule unchanged');
      assert.equal(ss.getSheetByName(name).getRange(row, col).getValue(), col === 2 ? 'PRAYER' : 'LUNCH', name + ' put back');
      assert.match(lastToast(env).msg, /Prayer and Lunch columns are not part of the timetable/, name);
    });
  });
});

test('a paste across period 4, Lunch and period 5 saves both periods and skips Lunch', () => {
  const { env, api, ss } = small(SCHEDULE);
  api.ClassViewManager.renderClassView('6th A');
  fire(env, api, edit(env, 'Class_View', 6, 6, [['Art', 'xx', 'Music'], ['Deepa', 'yy', 'Asha']]));
  const rows = ms(ss);
  assert.deepEqual(rows.find(r => r[1] === 4).slice(4, 6), ['Art', 'Deepa']);
  assert.deepEqual(rows.find(r => r[1] === 5).slice(4, 6), ['Music', 'Asha']);
  assert.equal(rows.length, 3, 'no row for the Lunch column');
  assert.equal(ss.getSheetByName('Class_View').getRange(6, 7).getValue(), 'LUNCH');
  // A Teacher_View paste across Prayer and period 1
  api.TeacherViewManager.renderTeacherView('Bina');
  fire(env, api, edit(env, 'Teacher_View', 7, 2, [['zz', '6th B - Science']]));
  assert.ok(ms(ss).some(r => r[0] === 'Tuesday' && r[1] === 1 && r[2] === '6th B' && r[5] === 'Bina'));
});

test('a Master Grid drawn before the break columns is redrawn in full by the next refresh', () => {
  const { api, ss } = small(SCHEDULE);
  const grid = ss.insertSheet('Master_Grid_View');
  grid.getRange(1, 1, 2, 9).setValues([['Class', 'Period 1', '', '', '', '', '', '', 'Period 8'], ['6th A', 'old', '', '', '', '', '', '', '']]);
  api.refreshMasterGridData_();
  assert.equal(grid.getRange(1, 2).getValue(), 'Prayer\n9:10–10:00');
  assert.equal(grid.getLastColumn(), 11);
  assert.match(String(grid.getRange(2, 3).getValue()), /Maths \(Asha\)/);
});

test('teacher web page and local preview list the columns in the same order', () => {
  const { api } = small(SCHEDULE);
  const slots = plain(api.viewSlots_()).map(s => s.period ? 'P' + s.period : s.name);
  const opts = plain(api.apiWebGetOptions());
  assert.deepEqual(opts.slots.map(s => s.period ? 'P' + s.period : s.name), slots);
  assert.equal(opts.slots[0].time, '9:10–10:00');
  assert.equal(opts.periods.length, 8, 'periods list unchanged for older pages');
});
