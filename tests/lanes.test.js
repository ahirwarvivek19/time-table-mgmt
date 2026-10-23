'use strict';
/**
 * Split periods (one class divided into groups) and combined lectures (one
 * teacher, several sections) — parser, validation and view behaviour.
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
  put('Classes', (classes || ['9th C', '9th D', '11th D', '11th C (E)', '12th D']).map(c => [c, 'Secondary', '']), 3);
  put('Master_Schedule', schedule, 8);
  return { env, api, ss: env.ss };
}
const rows = ss => { const sh = ss.getSheetByName('Master_Schedule'); return sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues(); };
const P = () => small([], []).api.ScheduleParser;

// ───────────────────────────────────────────── cell grammar

test('parseLanes: split period with a co-taught group and per-group rooms', () => {
  const parsed = plain(P().parseLanes({ Subject: 'Hindi / Beauty / IT', Teacher: 'Ekta Gour / Krishna Patale / Divya Lodval + Reena Shivhare', Room: '12 / 14 / Lab' }));
  assert.equal(parsed.mismatch, false);
  assert.deepEqual(parsed.lanes, [
    { subject: 'Hindi', teachers: ['Ekta Gour'], room: '12' },
    { subject: 'Beauty', teachers: ['Krishna Patale'], room: '14' },
    { subject: 'IT', teachers: ['Divya Lodval', 'Reena Shivhare'], room: 'Lab' },
  ]);
});

test('parseLanes: legacy cells keep their meaning; "-" holds an empty group; N/A is no room', () => {
  const p = P();
  assert.deepEqual(plain(p.parseLanes({ Subject: 'Science', Teacher: 'A / B' }).lanes), [{ subject: 'Science', teachers: ['A', 'B'], room: '' }], 'one subject + 2 teachers = co-teaching');
  assert.deepEqual(plain(p.parseLanes({ Subject: 'Library / Computer', Teacher: '- / Reena Shivhare' }).lanes).map(l => l.teachers), [[], ['Reena Shivhare']]);
  assert.deepEqual(plain(p.parseLanes({ Subject: 'Hindi / IT', Teacher: 'A / B', Room: 'N/A' }).lanes).map(l => l.room), ['', '']);
  assert.deepEqual(plain(p.parseLanes({ Subject: 'Hindi / IT', Teacher: 'A / B', Room: '207' }).lanes).map(l => l.room), ['207', '207'], 'one room = whole slot');
  assert.equal(p.parseLanes({ Subject: 'Hindi / IT / Beauty', Teacher: 'A / B' }).mismatch, true);
  assert.equal(p.parseLanes({ Subject: 'Hindi / IT', Teacher: 'A', Room: '1 / 2 / 3' }).roomMismatch, true);
});

test('formatLanes is the inverse of parseLanes', () => {
  const p = P();
  [
    { Subject: 'Hindi / Beauty / IT', Teacher: 'Ekta Gour / Krishna Patale / Divya Lodval + Reena Shivhare', Room: '' },
    { Subject: 'Library / Computer', Teacher: '- / Reena Shivhare', Room: '' },
    { Subject: 'Science', Teacher: 'A / B', Room: '207' },
    { Subject: 'Accountancy / Mathematics', Teacher: 'Sayyed Wahid Hussain / Rakesh Kumar Jaiswal', Room: '301 / 206' },
  ].forEach(row => {
    const f = p.formatLanes(p.parseLanes(row).lanes);
    assert.deepEqual(plain({ Subject: f.subject, Teacher: f.teacher, Room: f.room }), row);
  });
});

test('assignTeacher / removeTeacher on split periods: groups, co-teachers and rooms stay aligned', () => {
  const p = P();
  const pick = o => ({ teacher: o.teacher, subject: o.subject, room: o.room });
  // a new teacher taking the IT group replaces its teacher (no silent co-teacher)
  const took = p.assignTeacher('A / B', 'Hindi / IT', 'C', 'IT', '1 / 2');
  assert.deepEqual(plain(pick(took)), { teacher: 'A / C', subject: 'Hindi / IT', room: '1 / 2' });
  assert.deepEqual(plain(took.replaced), [{ teacher: 'B', subject: 'IT' }]);
  // removing a co-teacher keeps the group
  assert.deepEqual(plain(pick(p.removeTeacher('A / B + C', 'Hindi / IT', 'C', '1 / 2'))), { teacher: 'A / B', subject: 'Hindi / IT', room: '1 / 2' });
  // removing the only teacher of a group removes the group and its room
  assert.deepEqual(plain(pick(p.removeTeacher('A / B', 'Hindi / IT', 'A', '1 / 2'))), { teacher: 'B', subject: 'IT', room: '2' });
  // a subject the split period does not have is refused (groups are set up in Class_View)
  assert.match(p.assignTeacher('A / B', 'Hindi / IT', 'C', 'Beauty', '').error, /type one of those subjects/);
  // no subject on a split period cannot be placed
  assert.match(p.assignTeacher('A / B', 'Hindi / IT', 'C', '', '').error, /split into 2 groups/);
  // a co-teacher moving to the other group takes it over
  assert.deepEqual(plain(pick(p.assignTeacher('A / B + C', 'Hindi / IT', 'C', 'Hindi', ''))), { teacher: 'C / B', subject: 'Hindi / IT', room: '' });
  // nobody assigned yet on a split period: the teacher takes the group they name, not the whole slot
  assert.deepEqual(plain(pick(p.assignTeacher('', 'Business Studies / Accountancy', 'C', 'Accountancy', ''))), { teacher: '- / C', subject: 'Business Studies / Accountancy', room: '' });
});

test('classifyTeacherSlot: same grade combines, different grade or subject clashes, activities may span grades', () => {
  const p = P();
  assert.deepEqual(plain(p.classifyTeacherSlot([{ cls: '9th C', subject: 'Social Science (H)' }, { cls: '9th D', subject: 'social science (h)' }])), { clash: false, combined: true, reason: '' });
  assert.equal(p.classifyTeacherSlot([{ cls: '11th D', subject: 'Mathematics' }, { cls: '12th D', subject: 'Mathematics' }]).reason, 'same subject in different grades');
  assert.equal(p.classifyTeacherSlot([{ cls: '9th C', subject: 'Hindi' }, { cls: '9th D', subject: 'English' }]).reason, 'different subjects');
  assert.equal(p.classifyTeacherSlot([{ cls: '11th B', subject: 'Sports & Yoga (Fit India)' }, { cls: '12th A', subject: 'Sports & Yoga (Fit India)' }]).clash, false);
  assert.equal(p.classifyTeacherSlot([{ cls: '11th C (H), 11th C (E)', subject: 'Accountancy' }]).combined, true, 'one row shared by two sections');
});

test('classifyTeacherSlot: medium marks — one marked section joins an unmarked one; H.M + E.M at once clashes', () => {
  const p = P();
  assert.deepEqual(plain(p.splitMedium('Physics (H.M)')), { base: 'Physics', medium: 'H' });
  assert.deepEqual(plain(p.splitMedium('Economics (E)')), { base: 'Economics', medium: 'E' }, 'older mark');
  assert.deepEqual(plain(p.splitMedium('Chemistry (Hindi)')), { base: 'Chemistry', medium: 'H' }, 'older mark');
  assert.deepEqual(plain(p.splitMedium('Life Skills Umang (Aadha Phool)')), { base: 'Life Skills Umang (Aadha Phool)', medium: '' });
  assert.deepEqual(plain(p.classifyTeacherSlot([{ cls: '11th B', subject: 'Hindi' }, { cls: '11th C', subject: 'Hindi (E.M)' }, { cls: '11th D', subject: 'Hindi (E.M)' }])), { clash: false, combined: true, reason: '' });
  assert.equal(p.classifyTeacherSlot([{ cls: '12th B', subject: 'Physics (H.M)' }, { cls: '12th D', subject: 'Physics (E.M)' }]).reason, 'Hindi- and English-medium groups at once');
  assert.equal(p.classifyTeacherSlot([{ cls: '11th B', subject: 'Physics (H.M)' }, { cls: '12th B', subject: 'Physics' }]).reason, 'same subject in different grades');
});

// ───────────────────────────────────────────── validation

test('validation: split + combined lecture is not flagged; cross-grade and mismatched groups are', () => {
  const T = ['Rubal Verma', 'Indira Mathur', 'Rakesh Kumar Jaiswal', 'Sayyed Wahid Hussain', 'X'];
  const { api, ss } = small([
    // 9th C + 9th D: one Social Science lecture split by medium (H with Rubal, E with Indira)
    ['Monday', 8, '9th C', 'Secondary', 'Social Science (H) / Social Science (E)', 'Rubal Verma / Indira Mathur', '', ''],
    ['Monday', 8, '9th D', 'Secondary', 'Social Science (H) / Social Science (E)', 'Rubal Verma / Indira Mathur', '', ''],
    // 11th C (E) maths group joins 11th D
    ['Monday', 6, '11th D', 'Secondary', 'Mathematics', 'Rakesh Kumar Jaiswal', '', ''],
    ['Monday', 6, '11th C (E)', 'Secondary', 'Economics (E) / Mathematics', 'Indira Mathur / Rakesh Kumar Jaiswal', '', ''],
    // …but the same teacher in 12th D at the same time is a clash
    ['Tuesday', 6, '11th D', 'Secondary', 'Mathematics', 'Rakesh Kumar Jaiswal', '', ''],
    ['Tuesday', 6, '12th D', 'Secondary', 'Mathematics', 'Rakesh Kumar Jaiswal', '', ''],
    // groups that don't line up
    ['Wednesday', 1, '9th C', 'Secondary', 'Hindi / IT / Beauty', 'X / Indira Mathur', '', ''],
    // two groups of different classes in one explicit room
    ['Thursday', 1, '11th C (E)', 'Secondary', 'Accountancy / Mathematics', 'Sayyed Wahid Hussain / Rakesh Kumar Jaiswal', '301 / 206', ''],
    ['Thursday', 1, '12th D', 'Secondary', 'Physics', 'X', '206', ''],
  ], T);
  const res = api.validateSchedule_();
  const st = rows(ss).map(r => r[7]);
  assert.deepEqual(st.slice(0, 4), ['', '', '', '']);
  assert.match(st[4], /Teacher Clash \(Rakesh Kumar Jaiswal: same subject in different grades\)/);
  assert.match(st[5], /Teacher Clash/);
  assert.match(st[6], /Lane Mismatch/);
  assert.match(st[7], /Room Clash/);
  assert.match(st[8], /Room Clash/);
  assert.ok(res.combinedLectures >= 3, 'combined lectures counted');
  assert.ok(res.splitSlots >= 4, 'split periods counted');
  assert.match(api.formatValidationSummary_(res), /combined lecture/);
});

test('validation: one room for the whole split period is not a room clash with itself', () => {
  const { api, ss } = small([
    ['Monday', 1, '9th C', 'Secondary', 'Hindi / IT', 'Rubal Verma / Indira Mathur', '207', ''],
  ], ['Rubal Verma', 'Indira Mathur']);
  api.validateSchedule_();
  assert.equal(rows(ss)[0][7], '');
});

// ───────────────────────────────────────────── views

function teacherViewEdit(env, api, teacher, row, col, value) {
  api.TeacherViewManager.renderTeacherView(teacher);
  const sheet = env.ss.getSheetByName('Teacher_View');
  const range = sheet.getRange(row, col);
  range.setValue(value);
  api.onEdit(editEvent(env, range, { triggerUid: env.properties.EDIT_TRIGGER_UID }));
}

test('Teacher_View shows a combined lecture once with all sections, and a cross-grade one as a clash', () => {
  const { env, api } = small([
    ['Monday', 8, '9th C', 'Secondary', 'Social Science (H) / Social Science (E)', 'Rubal Verma / Indira Mathur', '', ''],
    ['Monday', 8, '9th D', 'Secondary', 'Social Science (H) / Social Science (E)', 'Rubal Verma / Indira Mathur', '', ''],
    ['Tuesday', 6, '11th D', 'Secondary', 'Mathematics', 'Rakesh Kumar Jaiswal', '', ''],
    ['Tuesday', 6, '12th D', 'Secondary', 'Mathematics', 'Rakesh Kumar Jaiswal', '', ''],
  ], ['Rubal Verma', 'Indira Mathur', 'Rakesh Kumar Jaiswal']);
  api.TeacherViewManager.renderTeacherView('Indira Mathur');
  const tv = env.ss.getSheetByName('Teacher_View');
  assert.equal(tv.getRange(6, 9).getValue(), '9th C, 9th D\nSocial Science (E)');
  api.TeacherViewManager.renderTeacherView('Rakesh Kumar Jaiswal');
  assert.match(String(tv.getRange(7, 7).getValue()), /11th D\nMathematics\n---\n12th D\nMathematics/);
});

test('Teacher_View edit on a split period: a teacher takes the group they name; unknown subject or bare class is refused', () => {
  const { env, api, ss } = small([
    ['Monday', 1, '9th C', 'Secondary', 'Hindi / IT', 'Rubal Verma / Indira Mathur', '', ''],
    ['Monday', 1, '9th D', 'Secondary', 'English', 'Y', '', ''],
  ], ['Rubal Verma', 'Indira Mathur', 'X', 'Y']);
  // A subject the period does not have is refused, not added as a new group
  teacherViewEdit(env, api, 'X', 6, 2, '9th C - Beauty');
  let r = rows(ss)[0];
  assert.deepEqual([r[4], r[5]], ['Hindi / IT', 'Rubal Verma / Indira Mathur']);
  assert.ok(env.ss.toasts.some(t => /type one of those subjects/.test(t.msg)));
  // Naming an existing group replaces its teacher, and says so
  teacherViewEdit(env, api, 'X', 6, 2, '9th C - IT');
  r = rows(ss)[0];
  assert.deepEqual([r[4], r[5]], ['Hindi / IT', 'Rubal Verma / X']);
  assert.ok(env.ss.toasts.some(t => t.title === 'Saved' && /Indira Mathur replaced by X/.test(t.msg)));
  // Y has 9th D; typing just "9th C" cannot pick a group, so nothing changes anywhere
  teacherViewEdit(env, api, 'Y', 6, 2, '9th C');
  r = rows(ss);
  assert.equal(r[1][5], 'Y', 'Y was not removed from 9th D');
  assert.equal(r[0][5], 'Rubal Verma / X');
  assert.ok(env.ss.toasts.some(t => /split into 2 groups/.test(t.msg)));
});

// ───────────────────────────────────────────── Class_View markers

test('describeSlot: split groups, combined sections and clashes for one class cell', () => {
  const p = P();
  const slot = [
    { Class: '11th A', Subject: 'Hindi / IT', Teacher: 'Ekta Gour / Divya Lodval + Reena Shivhare', Room: '12 / Lab' },
    { Class: '11th B', Subject: 'Hindi / IT', Teacher: 'Ekta Gour / Divya Lodval + Reena Shivhare', Room: '12 / Lab' },
    { Class: '10th B', Subject: 'Hindi', Teacher: 'Ekta Gour', Room: '' },
  ];
  const d = plain(p.describeSlot(slot, '11th A'));
  assert.equal(d.split, true);
  assert.deepEqual(d.combinedWith, ['11th B']);
  assert.equal(d.clashes.length, 1);
  assert.match(d.clashes[0], /^Ekta Gour also has 11th B \(Hindi\), 10th B \(Hindi\)$/);
  assert.match(d.note, /• IT — Divya Lodval \+ Reena Shivhare · Room Lab/);
  // viewing two classes at once: the shared lecture is shown once, not duplicated
  const both = plain(p.describeSlot(slot.slice(0, 2), '11th A, 11th B'));
  assert.equal(both.lanes.length, 2);
  assert.deepEqual(both.combinedWith, []);
  // a plain period with nobody else → no note
  assert.equal(plain(p.describeSlot([{ Class: '9th A', Subject: 'Sanskrit', Teacher: 'Neha Jamidar' }], '9th A')).note, '');
});

test('Class_View colours split / combined / clash cells and explains them in a note', () => {
  const { env, api } = small([
    ['Monday', 1, '9th C', 'Secondary', 'Social Science', 'Rubal Verma', '', ''],
    ['Monday', 1, '9th D', 'Secondary', 'Social Science', 'Rubal Verma', '', ''],
    ['Monday', 2, '9th C', 'Secondary', 'Hindi / IT', 'Indira Mathur / X', '', ''],
    ['Monday', 3, '9th C', 'Secondary', 'English', 'Y', '', ''],
    ['Monday', 3, '9th D', 'Secondary', 'Mathematics', 'Y', '', ''],
    ['Monday', 4, '9th C', 'Secondary', 'Sanskrit', 'X', '', ''],
  ], ['Rubal Verma', 'Indira Mathur', 'X', 'Y']);
  api.ClassViewManager.renderClassView('9th C');
  const cv = env.ss.getSheetByName('Class_View');
  const cell = p => cv.getRange(6, 1 + p); // Monday subject row
  assert.equal(cell(1).getBackground(), '#F5EEF8');
  assert.equal(cell(1).getNote(), 'Taught together:\n• Social Science — Rubal Verma: with 9th D');
  assert.equal(cell(2).getBackground(), '#FEF5E7');
  assert.equal(cell(2).getValue(), 'Hindi / IT');
  assert.equal(cv.getRange(7, 3).getValue(), 'Indira Mathur / X');
  assert.match(cell(2).getNote(), /Split class — 2 groups:\n• Hindi — Indira Mathur\n• IT — X/);
  assert.equal(cell(3).getBackground(), '#FADBD8');
  assert.match(cell(3).getNote(), /⚠ Clash: Y also has 9th D \(Mathematics\)/);
  assert.equal(cell(4).getBackground(), '#EBF5FB');
  assert.equal(cell(4).getNote(), '');
});
