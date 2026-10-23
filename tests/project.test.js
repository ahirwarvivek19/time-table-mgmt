'use strict';
/**
 * End-to-end tests: the real .gs files run against the in-memory Sheets mock.
 * Run with `npm test`.
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { createEnvironment, loadProject, editEvent } = require('./harness/gas-mock');

const ROOT = path.resolve(__dirname, '..');
// Objects created inside the VM context have different prototypes; compare plain copies.
const plain = x => JSON.parse(JSON.stringify(x));

/** Fresh spreadsheet with the built-in timetable imported. */
function imported() {
  const env = createEnvironment();
  const { api } = loadProject(ROOT, env);
  api.setupInitialSpreadsheet();
  api.importExcelData();
  return { env, api, ss: env.ss };
}

/** Fresh spreadsheet with tiny hand-made data. */
function small(schedule, teachers, classes) {
  const env = createEnvironment();
  const { api } = loadProject(ROOT, env);
  api.setupInitialSpreadsheet();
  const put = (name, rows, w) => rows.length && env.ss.getSheetByName(name).getRange(2, 1, rows.length, w).setValues(rows);
  put('Teachers', teachers.map(t => (Array.isArray(t) ? t : [t, '', 40, '', ''])), 5);
  put('Classes', (classes || ['6th A', '6th B', '7th A', '7th B']).map(c => [c, 'Middle', '']), 3);
  put('Master_Schedule', schedule, 8);
  return { env, api, ss: env.ss };
}

const scheduleRows = ss => {
  const sh = ss.getSheetByName('Master_Schedule');
  return sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues();
};
const slot = (ss, cls, day, period) => scheduleRows(ss).find(r => r[2] === cls && r[0] === day && r[1] == period);

// ───────────────────────────────────────────── import

test('Import Timetable loads the 2026-27 timetable and renders every view', () => {
  const { ss, env } = imported();
  assert.equal(ss.getSheetByName('Master_Schedule').getLastRow() - 1, 1344);
  assert.equal(ss.getSheetByName('Classes').getLastRow() - 1, 28);
  assert.equal(ss.getSheetByName('Teachers').getLastRow() - 1, 78);
  // views were refreshed by the menu import (the duplicate importExcelData skipped this)
  assert.equal(ss.getSheetByName('Class_View').getRange('B3').getValue(), '6th A');
  assert.equal(ss.getSheetByName('Teacher_View').getRange('B3').getValue(), 'Shiv Singh Ahirwar');
  assert.ok(ss.getSheetByName('Master_Schedule').getRange(2, 5).getDataValidation(), 'subject dropdown applied');
  assert.ok(ss.getSheetByName('Master_Grid_View').getLastRow() > 1, 'master grid generated');
  assert.equal(ss.getSheetByName('Teacher_Day_View').getRange('B3').getValue(), 'Monday');
  assert.equal(ss.getSheetByName('Teacher_Day_View').getLastRow(), 5 + 78, 'one row per teacher');
  assert.equal(env.alerts.filter(a => /Import Complete/.test(a)).length, 1, 'exactly one summary alert');
});

test('2026-27 import: new bell times in the headers; 11th C medium groups show as a split class', () => {
  const { api, ss } = imported();
  api.ClassViewManager.renderClassView('11th C');
  const cv = ss.getSheetByName('Class_View');
  assert.equal(cv.getRange(5, 2).getValue(), 'Period 1\n9:30–10:20');
  assert.equal(cv.getRange(5, 6).getValue(), 'Period 5\n12:50–1:30');
  assert.equal(cv.getRange(5, 9).getValue(), 'Period 8\n2:50–3:30');
  assert.equal(cv.getRange(6, 2).getValue(), 'Hindi (E.M) / English (H.M)');
  assert.equal(cv.getRange(7, 2).getValue(), 'Ekta Gour / Rajesh Kumar Bairagi');
  const note = cv.getRange(6, 2).getNote();
  assert.match(note, /Split class — 2 groups/);
  assert.match(note, /Taught together:\n• Hindi \(E\.M\) — Ekta Gour: with 11th B, 11th D/);
  assert.doesNotMatch(note, /Clash/, 'a medium mark on one section only is the same lecture');
  // the fallback bell table in Config.gs matches the generated one
  api.CONFIG.PERIODS.forEach(p => assert.equal(api.periodTime_(p), api.CONFIG.PERIOD_TIMES[p], 'period ' + p));
});

test('Teachers sheet gets Total Hours Scheduled (combined lecture counted once)', () => {
  const { ss } = imported();
  const rows = ss.getSheetByName('Teachers').getRange(2, 1, 78, 5).getValues();
  const krishna = rows.find(r => r[0] === 'Krishna Patale');
  assert.ok(krishna[4] > 0 && krishna[4] <= 48, 'between 1 and 48 periods: ' + krishna[4]);
});

test('importing asks which data to use when a saved snapshot exists', () => {
  const { api, env, ss } = imported();
  api.saveCurrentSheetAsPrefillData();
  ss.getSheetByName('Master_Schedule').getRange(2, 5).setValue('EDITED');
  env.answers.push('CANCEL');
  api.importExcelData();
  assert.equal(ss.getSheetByName('Master_Schedule').getRange(2, 5).getValue(), 'EDITED', 'cancel keeps the sheet');
  env.answers.push('NO');
  api.importExcelData();
  assert.notEqual(ss.getSheetByName('Master_Schedule').getRange(2, 5).getValue(), 'EDITED', 'NO loads built-in data');
});

// ───────────────────────────────────────────── validation

test('validation: combined lecture is not a clash, real double booking flags BOTH rows', () => {
  const { api, ss } = small([
    ['Monday', 1, '6th A', 'Middle', 'Hindi / IT', 'A Teacher / B Teacher', '', ''],
    ['Monday', 1, '6th B', 'Middle', 'Hindi / Beauty', 'A Teacher / C Teacher', '', ''], // A teaches Hindi in both → combined
    ['Monday', 2, '6th A', 'Middle', 'Maths', 'D Teacher', '', ''],
    ['Monday', 2, '7th A', 'Middle', 'Science', 'D Teacher', '', ''],                  // D double-booked
  ], ['A Teacher', 'B Teacher', 'C Teacher', 'D Teacher']);
  const res = api.validateSchedule_();
  const status = scheduleRows(ss).map(r => r[7]);
  assert.equal(status[0], '');
  assert.equal(status[1], '');
  assert.match(status[2], /Teacher Clash \(D Teacher\)/);
  assert.match(status[3], /Teacher Clash \(D Teacher\)/);
  assert.equal(res.counts['Teacher clash'], 1);
});

test('validation: "N/A" rooms are not a room clash, a real shared room is', () => {
  const { api, ss } = small([
    ['Monday', 1, '6th A', 'Middle', 'Hindi', 'A Teacher', 'N/A', ''],
    ['Monday', 1, '6th B', 'Middle', 'Maths', 'B Teacher', 'N/A', ''],
    ['Monday', 2, '6th A', 'Middle', 'Hindi', 'A Teacher', '101', ''],
    ['Monday', 2, '6th B', 'Middle', 'Maths', 'B Teacher', '101', ''],
  ], ['A Teacher', 'B Teacher']);
  api.validateSchedule_();
  const status = scheduleRows(ss).map(r => r[7]);
  assert.deepEqual(status.slice(0, 2), ['', '']);
  assert.match(status[2], /Room Clash/);
  assert.match(status[3], /Room Clash/);
});

test('validation: unavailable day is case-insensitive; unregistered teacher and unknown class are reported', () => {
  const { api, ss } = small([
    ['Monday', 1, '6th A', 'Middle', 'Hindi', 'A Teacher', '', ''],
    ['Monday', 2, '6th A', 'Middle', 'Hindi', 'Ghost Teacher', '', ''],
    ['Monday', 3, '9th Z', 'Middle', 'Hindi', 'A Teacher', '', ''],
  ], [['A Teacher', '', 40, 'monday, Friday', '']]);
  api.validateSchedule_();
  const status = scheduleRows(ss).map(r => r[7]);
  assert.match(status[0], /Unavailable Day/);
  assert.match(status[1], /Unregistered Teacher \(Ghost Teacher\)/);
  assert.match(status[2], /Unknown Class \(9th Z\)/);
});

// ───────────────────────────────────────────── teacher view write-back

function teacherViewEdit(env, api, teacher, row, col, value) {
  api.TeacherViewManager.renderTeacherView(teacher);
  const sheet = env.ss.getSheetByName('Teacher_View');
  const range = sheet.getRange(row, col);
  range.setValue(value);
  api.onEdit(editEvent(env, range, { triggerUid: env.properties.EDIT_TRIGGER_UID }));
}

test('Teacher_View: moving a period removes the teacher from the old class', () => {
  const { env, api, ss } = small([
    ['Monday', 1, '6th A', 'Middle', 'Maths', 'A Teacher', '', ''],
    ['Monday', 1, '7th B', 'Middle', 'Science', 'B Teacher', '', ''],
  ], ['A Teacher', 'B Teacher']);
  teacherViewEdit(env, api, 'A Teacher', 6, 2, '7th B - Maths'); // Monday P1
  assert.equal(slot(ss, '6th A', 'Monday', 1)[5], '', 'no longer in 6th A');
  // A takes the 7th B period: B Teacher is replaced (and told), not silently paired up
  assert.equal(slot(ss, '7th B', 'Monday', 1)[5], 'A Teacher');
  assert.equal(slot(ss, '7th B', 'Monday', 1)[4], 'Maths');
  assert.ok(env.ss.toasts.some(t => /B Teacher replaced by A Teacher/.test(t.msg) && /6th A Maths now has no teacher/.test(t.msg)));
});

test('Teacher_View: changing one teacher\'s subject keeps the other optional subjects', () => {
  const { env, api, ss } = small([
    ['Monday', 1, '6th A', 'Middle', 'English / Sanskrit / IT', 'A Teacher / B Teacher / C Teacher', '', ''],
  ], ['A Teacher', 'B Teacher', 'C Teacher']);
  teacherViewEdit(env, api, 'C Teacher', 6, 2, '6th A - Beauty');
  assert.equal(slot(ss, '6th A', 'Monday', 1)[4], 'English / Sanskrit / Beauty');
});

test('Teacher_View: FREE removes the teacher and their paired subject', () => {
  const { env, api, ss } = small([
    ['Monday', 1, '6th A', 'Middle', 'Hindi / IT', 'A Teacher / B Teacher', '', ''],
  ], ['A Teacher', 'B Teacher']);
  teacherViewEdit(env, api, 'B Teacher', 6, 2, 'FREE');
  const r = slot(ss, '6th A', 'Monday', 1);
  assert.equal(r[5], 'A Teacher');
  assert.equal(r[4], 'Hindi');
});

test('Teacher_View: an unknown class is rejected instead of creating a phantom row', () => {
  const { env, api, ss } = small([['Monday', 1, '6th A', 'Middle', 'Maths', 'A Teacher', '', '']], ['A Teacher']);
  teacherViewEdit(env, api, 'A Teacher', 6, 2, '6A - Maths');
  assert.equal(scheduleRows(ss).length, 1);
  assert.equal(slot(ss, '6th A', 'Monday', 1)[5], 'A Teacher');
  assert.ok(env.ss.toasts.some(t => /unknown class/.test(t.msg)));
});

test('Teacher_Day_View uses the same write-back rules', () => {
  const { env, api, ss } = small([
    ['Tuesday', 3, '6th A', 'Middle', 'Maths', 'A Teacher', '', ''],
  ], ['A Teacher', 'B Teacher']);
  api.TeacherDayViewManager.renderTeacherDayView('Tuesday', 'All Classes');
  const sheet = ss.getSheetByName('Teacher_Day_View');
  const rowB = 6 + 1; // second teacher row
  assert.equal(sheet.getRange(rowB, 1).getValue(), 'B Teacher');
  const range = sheet.getRange(rowB, 4); // Period 3
  range.setValue('6th A - Maths');
  api.onEdit(editEvent(env, range, { triggerUid: env.properties.EDIT_TRIGGER_UID }));
  assert.equal(slot(ss, '6th A', 'Tuesday', 3)[5], 'B Teacher', 'B takes the period (same replace rule as Teacher_View)');
});

// ───────────────────────────────────────────── class view & multi-cell edits

test('Class_View: pasting a block syncs every cell, not just the first', () => {
  const { env, api, ss } = small([
    ['Monday', 1, '6th A', 'Middle', 'Maths', 'A Teacher', '', ''],
    ['Monday', 2, '6th A', 'Middle', 'Hindi', 'B Teacher', '', ''],
  ], ['A Teacher', 'B Teacher']);
  api.ClassViewManager.renderClassView('6th A');
  const cv = ss.getSheetByName('Class_View');
  const range = cv.getRange(6, 2, 1, 3); // Monday subject row, P1–P3
  range.setValues([['Science', 'English', 'Art']]);
  api.onEdit(editEvent(env, range, { triggerUid: env.properties.EDIT_TRIGGER_UID }));
  assert.equal(slot(ss, '6th A', 'Monday', 1)[4], 'Science');
  assert.equal(slot(ss, '6th A', 'Monday', 2)[4], 'English');
  assert.equal(slot(ss, '6th A', 'Monday', 3)[4], 'Art', 'missing slot row created');
});

test('Class_View: teacher typed into an empty slot is saved', () => {
  const { env, api, ss } = small([], ['A Teacher']);
  api.ClassViewManager.renderClassView('6th A');
  const range = ss.getSheetByName('Class_View').getRange(9, 3); // Tuesday teacher row, P2
  range.setValue('A Teacher');
  api.onEdit(editEvent(env, range, { triggerUid: env.properties.EDIT_TRIGGER_UID }));
  assert.equal(slot(ss, '6th A', 'Tuesday', 2)[5], 'A Teacher');
});

test('Class_View lists classes in school order, not alphabetically', () => {
  const { ss } = imported();
  const list = ss.getSheetByName('Class_View').getRange('B3').getDataValidation().list;
  assert.equal(list[0], '6th A');
  assert.equal(list[list.length - 1], '12th D');
});

// ───────────────────────────────────────────── triggers

test('simple onEdit steps aside once the installable trigger exists (no double processing)', () => {
  const { env, api, ss } = small([['Monday', 1, '6th A', 'Middle', 'Maths', 'A Teacher', '', '']], ['A Teacher']);
  assert.ok(env.properties.EDIT_TRIGGER_UID, 'setup installed and recorded the trigger');
  assert.equal(env.triggers.filter(t => t.getHandlerFunction() === 'onEdit').length, 1);
  api.installEditTrigger_();
  assert.equal(env.triggers.filter(t => t.getHandlerFunction() === 'onEdit').length, 1, 're-install does not duplicate');
  assert.equal(env.triggers.filter(t => t.getHandlerFunction() === 'onScheduleChange').length, 1, 'one change trigger');

  api.ClassViewManager.renderClassView('6th A');
  const range = ss.getSheetByName('Class_View').getRange(6, 2);
  range.setValue('Science');
  api.onEdit(editEvent(env, range)); // simple trigger call: no triggerUid
  assert.equal(slot(ss, '6th A', 'Monday', 1)[4], 'Maths', 'simple trigger ignored');
  api.onEdit(editEvent(env, range, { triggerUid: env.properties.EDIT_TRIGGER_UID }));
  assert.equal(slot(ss, '6th A', 'Monday', 1)[4], 'Science', 'installable trigger processed it');
});

test('Master_Schedule edit refreshes a Teacher_View showing a multi-teacher cell', () => {
  const { env, api, ss } = small([['Monday', 1, '6th A', 'Middle', 'Hindi', 'A Teacher', '', '']], ['A Teacher', 'B Teacher']);
  api.TeacherViewManager.renderTeacherView('B Teacher');
  const ms = ss.getSheetByName('Master_Schedule');
  const range = ms.getRange(2, 6);
  range.setValue('A Teacher / B Teacher');
  api.onEdit(editEvent(env, range, { triggerUid: env.properties.EDIT_TRIGGER_UID }));
  assert.match(String(ss.getSheetByName('Teacher_View').getRange(6, 2).getValue()), /6th A/);
});

// ───────────────────────────────────────────── other views

test('Master grid shows the day list when an entry runs on 5 of 6 days', () => {
  const days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'];
  const rows = days.map(d => [d, 1, '6th A', 'Middle', 'Maths', 'A Teacher', '', '']).concat([['Saturday', 1, '6th A', 'Middle', 'Bal Sabha', 'A Teacher', '', '']]);
  const { api, ss } = small(rows, ['A Teacher']);
  api.generateMasterGrid(true);
  const cell = ss.getSheetByName('Master_Grid_View').getRange(2, 2).getValue();
  assert.match(cell, /Maths \(A Teacher\) \[Mon,Tue,Wed,Thu,Fri\]/);
  assert.match(cell, /Bal Sabha \(A Teacher\) \[Sat\]/);
});

test('Teacher availability accepts "monday" and marks unavailable teachers', () => {
  const { env, api, ss } = small([['Monday', 1, '6th A', 'Middle', 'Maths', 'A Teacher', '', '']],
    [['A Teacher', '', 40, '', ''], ['B Teacher', '', 40, 'Monday', '']]);
  env.prompts.push('monday');
  api.generateTeacherAvailabilityGrid();
  const grid = ss.getSheetByName('Teacher_Availability_Grid').getRange(1, 1, 3, 3).getValues();
  assert.equal(grid[0][0], 'Teacher (Monday)');
  assert.equal(grid[1][1], 'BUSY (6th A)');
  assert.equal(grid[1][2], 'FREE');
  assert.equal(grid[2][1], 'UNAVAILABLE');
  env.prompts.push('Funday');
  api.generateTeacherAvailabilityGrid();
  assert.match(env.alerts[env.alerts.length - 1], /Please enter a school day/);
});

// ───────────────────────────────────────────── cover manager

test('Cover manager: available teachers exclude busy, unavailable, absent and already-covering teachers', () => {
  const { api } = small([
    ['Monday', 1, '6th A', 'Middle', 'Maths', 'Absent T', '', ''],
    ['Monday', 1, '6th B', 'Middle', 'Hindi', 'Busy T', '', ''],
    ['Monday', 1, '7th A', 'Middle', 'English', 'Other Absent', '', ''],
  ], [['Absent T', '', 40, '', ''], ['Busy T', '', 40, '', ''], ['Off T', '', 40, 'Monday', ''], ['Free T', '', 40, '', ''], ['Free U', '', 40, '', ''], ['Other Absent', '', 40, '', '']]);
  const date = '2026-10-12'; // a Monday
  assert.equal(api.apiGetDayForDate(date), 'Monday');
  assert.deepEqual(plain(api.apiGetAvailableTeachers('Monday', 1, date, 'Absent T')), ['Free T', 'Free U']);
  api.apiAssignCover(date, 'Other Absent', 1, '7th A', 'Free T');
  assert.deepEqual(plain(api.apiGetAvailableTeachers('Monday', 1, date, 'Absent T')), ['Free U']);
  assert.throws(() => api.apiAssignCover(date, 'Other Absent', 1, '7th A', 'Free U'), /already covered/);
  assert.throws(() => api.apiAssignCover(date, 'Absent T', 1, '6th A', 'Free T'), /already covering/);
  assert.deepEqual(plain(api.apiGetTeacherSchedule('Absent T', 'Monday')), [{ Period: 1, Class: '6th A', Subject: 'Maths' }]);
});

test('HTML pages: scripts parse and names are only ever inserted as text', () => {
  ['CoverManagerUI.html', 'TimetableWeb.html'].forEach(file => {
    const html = fs.readFileSync(path.join(ROOT, file), 'utf8');
    const script = html.match(/<script>([\s\S]*?)<\/script>/)[1];
    assert.doesNotThrow(() => new vm.Script(script), file + ' script parses');
    // innerHTML is only used to clear containers; no other HTML-parsing sinks
    (script.match(/innerHTML\s*[+]?=\s*[^;]+/g) || []).forEach(m => assert.match(m, /innerHTML\s*=\s*''$/, file + ': ' + m));
    assert.ok(!/insertAdjacentHTML|outerHTML|document\.write/.test(script), file + ': no other HTML sinks');
  });
});

// ───────────────────────────────────────────── parser helpers

test('ScheduleParser.assignTeacher / removeTeacher keep teacher↔subject pairs aligned', () => {
  const { api } = small([], []);
  const P0 = api.ScheduleParser;
  const ts = o => { const r = { teacher: o.teacher, subject: o.subject }; if ('removed' in o) r.removed = o.removed; return r; };
  const P = { assignTeacher: (...a) => ts(P0.assignTeacher(...a)), removeTeacher: (...a) => ts(P0.removeTeacher(...a)), rowIncludesTeacher: (...a) => P0.rowIncludesTeacher(...a) };
  assert.deepEqual(plain(P.assignTeacher('', 'Maths', 'A', '')), { teacher: 'A', subject: 'Maths' });
  assert.deepEqual(plain(P.assignTeacher('A', 'Hindi', 'B', 'Hindi')), { teacher: 'B', subject: 'Hindi' });
  assert.deepEqual(plain(P.assignTeacher('A', 'Hindi', 'B', 'IT')), { teacher: 'B', subject: 'IT' });
  assert.deepEqual(plain(P.assignTeacher('A', 'Hindi', 'B', '')), { teacher: 'B', subject: 'Hindi' });
  assert.deepEqual(plain(P.assignTeacher('A / B', 'Hindi / IT', 'b', 'Beauty')), { teacher: 'A / B', subject: 'Hindi / Beauty' });
  assert.deepEqual(plain(P.removeTeacher('A / B', 'Hindi / IT', 'A')), { teacher: 'B', subject: 'IT', removed: true });
  assert.deepEqual(plain(P.removeTeacher('A / B', 'Science', 'B')), { teacher: 'A', subject: 'Science', removed: true });
  assert.equal(P.removeTeacher('A', 'X', 'Z').removed, false);
  assert.ok(P.rowIncludesTeacher({ Teacher: 'Neha  Jhala', Subject: 'Maths' }, 'neha jhala'));
});

test('converter output reads back identically in the sheet (same cell grammar)', () => {
  const { buildAll } = require('../tools/convert/build');
  const { api, ss } = imported();
  const built = buildAll().schedule;
  const key = r => `${r[2]}|${r[0]}|${r[1]}`;
  const sheet = new Map(scheduleRows(ss).map(r => [key(r), r]));
  built.forEach(b => {
    const r = sheet.get(`${b.Class}|${b.Day}|${b.Period}`);
    const a1 = plain(api.ScheduleParser.parseLanes({ Subject: r[4], Teacher: r[5], Room: r[6] }));
    const a2 = plain(api.ScheduleParser.parseLanes(b));
    assert.deepEqual(a1, a2, key(r));
  });
});

test('read-only web page API returns class and teacher grids and rejects unknown names', () => {
  const { api } = imported();
  const opts = plain(api.apiWebGetOptions());
  assert.equal(opts.classes[0], '6th A');
  const cls = plain(api.apiWebGetTimetable('class', '9th A'));
  assert.equal(cls.rows.length, 6);
  assert.equal(cls.rows[0].cells[0], 'Sanskrit · Neha Jamidar');
  const t = plain(api.apiWebGetTimetable('teacher', 'neha jamidar'));
  assert.equal(t.title, 'Neha Jamidar');
  assert.match(t.rows[0].cells[0], /9th A · Sanskrit/);
  assert.throws(() => api.apiWebGetTimetable('class', '<script>'), /Unknown class/);
  const html = fs.readFileSync(path.join(ROOT, 'TimetableWeb.html'), 'utf8');
  assert.doesNotThrow(() => new vm.Script(html.match(/<script>([\s\S]*?)<\/script>/)[1]));
});

test('only the recorded installable trigger processes edits (duplicates and legacy triggers are ignored)', () => {
  const { env, api, ss } = small([['Monday', 1, '6th A', 'Middle', 'Maths', 'A Teacher', '', '']], ['A Teacher']);
  api.ClassViewManager.renderClassView('6th A');
  const range = ss.getSheetByName('Class_View').getRange(6, 2);
  range.setValue('Science');
  api.onEdit(editEvent(env, range, { triggerUid: 'someone-elses-trigger' }));
  assert.equal(slot(ss, '6th A', 'Monday', 1)[4], 'Maths', 'unrecorded trigger ignored');
  // Before any install is recorded, a stray installable trigger must not double-process either
  delete env.properties.EDIT_TRIGGER_UID;
  api.onEdit(editEvent(env, range, { triggerUid: 'legacy' }));
  assert.equal(slot(ss, '6th A', 'Monday', 1)[4], 'Maths');
  api.onEdit(editEvent(env, range)); // simple trigger handles it when nothing is installed
  assert.equal(slot(ss, '6th A', 'Monday', 1)[4], 'Science');
});

test('Teacher_View: clash cells and partial edits of a combined row are rejected; lowercase days still match', () => {
  const { env, api, ss } = small([
    ['monday ', 1, '6th A, 6th B', 'Middle', 'Maths', 'A Teacher', '', ''],
  ], ['A Teacher']);
  teacherViewEdit(env, api, 'A Teacher', 6, 2, '6th A - Maths');
  assert.equal(slot(ss, '6th A, 6th B', 'monday ', 1)[5], 'A Teacher', 'combined row untouched');
  assert.ok(env.ss.toasts.some(t => /combined row/.test(t.msg)));
  teacherViewEdit(env, api, 'A Teacher', 6, 2, '6th A\nMaths\n---\n7th A\nScience');
  assert.ok(env.ss.toasts.some(t => /has a clash/.test(t.msg)));
  teacherViewEdit(env, api, 'A Teacher', 6, 2, 'FREE');
  assert.equal(slot(ss, '6th A, 6th B', 'monday ', 1)[5], '', '"monday " row matched by FREE');
});

test('assignTeacher leaves a misaligned row (3 teachers / 2 subjects) unchanged', () => {
  const { api } = small([], []);
  const r = api.ScheduleParser.assignTeacher('A / B / C', 'Hindi / IT', 'B', 'Beauty');
  assert.deepEqual({ teacher: r.teacher, subject: r.subject }, { teacher: 'A / B / C', subject: 'Hindi / IT' });
});

test('Cover manager rejects non-school dates, busy cover teachers and unknown names; reads Date cells', () => {
  const { api, ss } = small([
    ['Monday', 1, '6th A', 'Middle', 'Maths', 'Absent T', '', ''],
    ['Monday', 1, '6th B', 'Middle', 'Hindi', 'Busy T', '', ''],
  ], ['Absent T', 'Busy T', 'Free T']);
  assert.throws(() => api.apiAssignCover('2026-10-11', 'Absent T', 1, '6th A', 'Free T'), /not a school day/); // Sunday
  assert.throws(() => api.apiAssignCover('2026-02-31', 'Absent T', 1, '6th A', 'Free T'), /not a school day/);
  assert.throws(() => api.apiAssignCover('2026-10-12', 'Absent T', 1, '6th A', 'Busy T'), /not free/);
  assert.throws(() => api.apiAssignCover('2026-10-12', 'Nobody', 1, '6th A', 'Free T'), /Unknown absent/);
  assert.throws(() => api.apiAssignCover('2026-10-12', 'Absent T', 1, '=HYPERLINK("x")', 'Free T'), /Unknown class/);
  // Sheets stores typed dates as Date objects
  ss.getSheetByName('Cover_Manager').appendRow([new Date(2026, 9, 12), 'Absent T', 1, '6th A', '', 'Free T']);
  assert.deepEqual(plain(api.apiGetCoversForDate('2026-10-12')), [{ period: 1, classToCover: '6th A', absentTeacher: 'Absent T', coverTeacher: 'Free T' }]);
  assert.equal(api.dayNameForDate_('2026-02-31'), '');
});

// ───────────────────────────────────────────── menu and styling

test('Menu is trimmed to the everyday actions, and every item points at a real function', () => {
  const env = createEnvironment();
  const { api } = loadProject(ROOT, env);
  api.onOpen();
  assert.equal(env.menus.length, 1, 'one menu, no submenus');
  const items = env.menus[0].items.filter(i => i !== '---');
  assert.deepEqual(items.map(i => i[1]), [
    'refreshAllViews', 'runValidation', 'generateTeacherAvailabilityGrid', 'openCoverManagerUI',
    'importExcelData', 'saveCurrentSheetAsPrefillData', 'setupInitialSpreadsheet'
  ]);
  items.forEach(([label, fn]) => assert.equal(typeof api[fn], 'function', label + ' → ' + fn));
});

test('Data tabs are styled down to the last sheet row, and refresh keeps the clash highlights', () => {
  const { api, ss } = imported();
  const res = api.validateSchedule_();
  assert.ok(res.rowsFlagged > 0, 'the built-in data has clashes to highlight');
  api.refreshAllViews(true);

  const ms = ss.getSheetByName('Master_Schedule');
  const status = ms.getRange(2, 8, ms.getLastRow() - 1, 1).getValues();
  let red = 0;
  status.forEach(([s], i) => {
    const bg = ms.getRange(i + 2, 8).getBackground();
    if (String(s).trim()) { red++; assert.equal(bg, '#FDEDEC', 'flagged row ' + (i + 2)); }
    else assert.equal(bg, i % 2 === 0 ? '#F8F9FA' : '#FFFFFF', 'clean row ' + (i + 2) + ' keeps its zebra colour');
  });
  assert.equal(red, res.rowsFlagged);

  // Rows below the data (where view edits and imports append) are already styled
  const below = ms.getLastRow() + 3;
  assert.equal(ms._bg.get(below + ',1'), below % 2 === 0 ? '#F8F9FA' : '#FFFFFF');
  const teachers = ss.getSheetByName('Teachers');
  assert.ok(teachers._bg.has((teachers.getLastRow() + 10) + ',5'), 'Teachers body styled past the last name');
});

test('Refresh rebuilds the Master Grid, so its rows match the classes after one is removed', () => {
  const { api, ss } = imported();
  const grid = ss.getSheetByName('Master_Grid_View');
  const before = grid.getLastRow();
  const classes = ss.getSheetByName('Classes');
  const last = classes.getLastRow();
  const removed = classes.getRange(last, 1).getValue();
  classes.getRange(last, 1, 1, 3).clearContent();
  const ms = ss.getSheetByName('Master_Schedule');
  const rows = ms.getRange(2, 1, ms.getLastRow() - 1, 8).getValues().filter(r => r[2] !== removed);
  ms.getRange(2, 1, ms.getLastRow() - 1, 8).clearContent();
  ms.getRange(2, 1, rows.length, 8).setValues(rows);
  api.refreshAllViews(true);
  assert.ok(grid.getLastRow() < before, 'grid shrank');
  const text = grid.getRange(1, 1, grid.getLastRow(), grid.getLastColumn()).getValues().flat().join('|');
  assert.ok(!text.includes(removed), removed + ' is gone from the grid');
});

test('Teacher Availability accepts schedule days written in any case', () => {
  const { api, ss, env } = small([['monday', 1, '6th A', 'Middle', 'Maths', 'T One', '', '']], ['T One', 'T Two']);
  env.prompts.push('Monday');
  api.generateTeacherAvailabilityGrid();
  const g = ss.getSheetByName('Teacher_Availability_Grid');
  const row = g.getRange(1, 1, g.getLastRow(), 2).getValues().find(r => r[0] === 'T One');
  assert.match(String(row[1]), /^BUSY/);
});
