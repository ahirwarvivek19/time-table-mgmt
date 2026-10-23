'use strict';
/**
 * Upgrading a sheet that is already in use: the data a school typed must survive.
 *
 * tests/fixtures/legacy-sheet.json is a sheet built by the code that was on GitHub before
 * this rewrite (25 classes incl. "11th C (H)", 1,200 periods), plus hand edits: a renamed
 * subject, legacy co-teachers, a lowercase day, a typed room, a combined "11th A, 11th B"
 * row, a split row with an empty group, a legacy one-teacher/two-subject row, an added
 * teacher with unavailable days, a class room and a Cover_Manager entry.
 *
 * The upgrade path is: clasp push → reload (onOpen) → Setup / Repair → Refresh All Views
 * → Check For Conflicts → use the views. No import. Nothing the school typed may change;
 * only the two computed columns (Total Hours Scheduled, Clash Status) are rewritten.
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('path');
const { createEnvironment, loadProject, editEvent } = require('./harness/gas-mock');

const ROOT = path.resolve(__dirname, '..');
const LEGACY = require('./fixtures/legacy-sheet.json');
const DATA_TABS = ['Teachers', 'Subjects', 'Classes', 'Master_Schedule', 'Cover_Manager'];
// Columns the script computes (1-based); everything else is user data
const COMPUTED = { Teachers: [5], Master_Schedule: [8] };

function legacySheet() {
  const env = createEnvironment();
  const { api } = loadProject(ROOT, env);
  Object.keys(LEGACY).forEach(name => {
    const sh = env.ss.insertSheet(name);
    const v = LEGACY[name];
    sh.getRange(1, 1, v.length, v[0].length).setValues(v);
  });
  return { env, api, ss: env.ss };
}
const read = (ss, name) => {
  const sh = ss.getSheetByName(name);
  const want = LEGACY[name];
  return sh.getRange(1, 1, Math.max(sh.getLastRow(), want.length), want[0].length).getValues();
};
const userData = (name, grid) => grid.map(r => r.map((v, c) => ((COMPUTED[name] || []).indexOf(c + 1) !== -1 ? '<computed>' : v)));

function assertDataUnchanged(ss, label) {
  DATA_TABS.forEach(name => {
    assert.deepEqual(userData(name, read(ss, name)), userData(name, LEGACY[name]), label + ': ' + name + ' changed');
  });
}

test('Upgrade without importing: Setup / Repair, Refresh, Check For Conflicts and every view leave the data as typed', () => {
  const { env, api, ss } = legacySheet();
  api.onOpen();
  assertDataUnchanged(ss, 'onOpen');
  api.setupInitialSpreadsheet();
  assertDataUnchanged(ss, 'Setup / Repair');
  api.refreshAllViews();
  assertDataUnchanged(ss, 'Refresh All Views');
  api.runValidation();
  assertDataUnchanged(ss, 'Check For Conflicts');
  api.getClassNames_().forEach(c => api.ClassViewManager.renderClassView(c));
  api.getTeacherNames_().forEach(t => api.TeacherViewManager.renderTeacherView(t));
  api.TeacherDayViewManager.renderTeacherDayView('Friday', 'All Classes');
  api.renderTeacherAvailabilityGrid_('Monday');
  api.generateMasterGrid(true);
  api.apiGetTeachers();
  assertDataUnchanged(ss, 'opening views');
  assert.equal(env.alerts.filter(a => /Import Complete/.test(a)).length, 0, 'nothing was imported');
  assert.equal(ss.getSheetByName('_Prefill_Data_Backup'), null, 'no snapshot written');
});

test('Upgrade: hand-edited and legacy rows still show correctly in the views', () => {
  const { api, ss } = legacySheet();
  api.setupInitialSpreadsheet();
  // legacy: one teacher for two subjects
  api.TeacherViewManager.renderTeacherView('Indira Mathur');
  assert.match(String(ss.getSheetByName('Teacher_View').getRange(10, 9).getValue()), /6th B/);
  // a hand-added second row for a period that already had one: both show (merged),
  // keeping subject ↔ teacher order, and Check For Conflicts points at the duplicate
  api.ClassViewManager.renderClassView('6th A');
  const cv = ss.getSheetByName('Class_View');
  assert.equal(cv.getRange(14, 9).getValue(), 'Sanskrit / Hindi / IT / Art');
  assert.match(String(cv.getRange(15, 9).getValue()), / \/ Ekta Gour \/ - \/ Manjula Marmat$/);
  api.validateSchedule_();
  const dup = read(ss, 'Master_Schedule').filter(r => r[2] === '6th A' && r[0] === 'Friday' && r[1] === 8);
  assert.equal(dup.length, 2);
  dup.forEach(r => assert.match(String(r[7]), /Duplicate Period/));
  // combined row shows in both classes (next to what 11th B already had); lowercase day still shows on Monday
  api.ClassViewManager.renderClassView('11th B');
  assert.match(String(cv.getRange(16, 9).getValue()), /(^| \/ )Yoga$/);
  api.ClassViewManager.renderClassView('11th A');
  assert.match(String(cv.getRange(16, 9).getValue()), /(^| \/ )Yoga$/);
  const lower = LEGACY.Master_Schedule[3];
  api.ClassViewManager.renderClassView(lower[2]);
  assert.equal(cv.getRange(6, 1 + Number(lower[1])).getValue(), lower[4], 'the "monday" row shows on Monday');
  // the added teacher is in the dropdowns
  assert.ok(api.getTeacherNamesSorted_().indexOf('New Teacher Added') !== -1);
});

test('Upgrade: one edit in a view changes only that period — every other row stays exactly as typed', () => {
  const { env, api, ss } = legacySheet();
  api.setupInitialSpreadsheet();
  api.ClassViewManager.renderClassView('6th A');
  const range = ss.getSheetByName('Class_View').getRange(13, 9); // Thursday P8 teacher row
  const before = read(ss, 'Master_Schedule');
  range.setValue(range.getValue() || 'Ruchika Vyas');
  api.onEdit(editEvent(env, range, { triggerUid: env.properties.EDIT_TRIGGER_UID }));
  const after = read(ss, 'Master_Schedule');
  const changed = after.filter((r, i) => before[i] && r.slice(0, 7).join('|') !== before[i].slice(0, 7).join('|'));
  assert.ok(changed.length <= 1, 'at most the edited row changed, got ' + changed.length);
  if (changed.length) assert.equal(changed[0][2], '6th A');
  // the lowercase day elsewhere was not rewritten by an unrelated edit
  assert.equal(after[3][0], 'monday');
});
