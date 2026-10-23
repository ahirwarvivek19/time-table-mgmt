'use strict';
/**
 * Randomised editing on the real 2026-27 timetable. A fixed seed makes every run the same.
 * After every edit (valid or not) these must hold:
 *   - nothing throws ("Timetable sync failed" never appears)
 *   - the number of teacher clashes never goes up (edits that would add one are refused;
 *     the clashes already in Timetable.xlsx may only go down)
 *   - no row ends up with subjects and teacher groups that don't line up
 *   - every teacher named in Master_Schedule is on the Teachers tab
 *   - Change_Log only grows
 * FUZZ_OPS=500 npm test runs a longer pass.
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('path');
const { createEnvironment, loadProject, editEvent } = require('./harness/gas-mock');

const ROOT = path.resolve(__dirname, '..');
const OPS = Number(process.env.FUZZ_OPS || 60);

function rng(seed) { // mulberry32
  return () => {
    seed |= 0; seed = (seed + 0x6D2B79F5) | 0;
    let t = Math.imul(seed ^ (seed >>> 15), 1 | seed);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

function builtIn() {
  const env = createEnvironment();
  const { api } = loadProject(ROOT, env);
  api.setupInitialSpreadsheet();
  env.answers.push('NO');
  api.importExcelData();
  return { env, api };
}

/** The sheet as the previous code left it, with hand edits (see upgrade.test.js) — opened by the new code, no import. */
function legacy() {
  const LEGACY = require('./fixtures/legacy-sheet.json');
  const env = createEnvironment();
  const { api } = loadProject(ROOT, env);
  Object.keys(LEGACY).forEach(name => {
    const v = LEGACY[name];
    env.ss.insertSheet(name).getRange(1, 1, v.length, v[0].length).setValues(v);
  });
  api.setupInitialSpreadsheet();
  return { env, api };
}

[['2026-27 built-in timetable', builtIn, OPS], ['an existing hand-edited sheet (no import)', legacy, Math.max(30, Math.round(OPS / 2))]].forEach(([label, load, count]) =>
test('Random edits through every view keep ' + label + ' consistent (' + count + ' edits)', () => {
  const { env, api } = load();
  const OPS_N = count;
  const ss = env.ss;
  const rand = rng(Number(process.env.FUZZ_SEED || 20261023));
  const pick = list => list[Math.floor(rand() * list.length)];
  const teachers = api.getTeacherNames_();
  const classes = api.getClassNames_();
  const days = api.CONFIG.DAYS;
  const ms = () => ss.getSheetByName('Master_Schedule');
  const msRows = () => ms().getRange(2, 1, ms().getLastRow() - 1, 8).getValues();
  const uid = env.properties.EDIT_TRIGGER_UID;

  const clashCount = () => (api.validateSchedule_().counts['Teacher clash'] || 0);
  const mismatches = () => msRows().filter(r => api.ScheduleParser.parseLanes({ Subject: r[4], Teacher: r[5], Room: r[6] }).mismatch).length;
  const unknownTeachers = () => msRows().filter(r => api.ScheduleParser.parseRowAssignments({ Subject: r[4], Teacher: r[5] })
    .some(a => a.teacher && !teachers.some(t => api.ScheduleParser.sameName(t, a.teacher)))).length;
  const logRows = () => { const s = ss.getSheetByName('Change_Log'); return s ? s.getLastRow() : 0; };

  let clashes = clashCount();
  const mismatchStart = mismatches();
  const unknownStart = unknownTeachers();
  let logs = logRows();
  const stats = { accepted: 0, refused: 0 };

  const ops = [
    // Class_View: pick one teacher on a random period
    () => {
      const cls = pick(classes), d = Math.floor(rand() * 6), p = 1 + Math.floor(rand() * 8);
      api.ClassViewManager.renderClassView(cls);
      const r = ss.getSheetByName('Class_View').getRange(7 + d * 2, 1 + p);
      r.setValue(pick(teachers));
      api.onEdit(editEvent(env, r, { triggerUid: uid }));
    },
    // Class_View: subject + teacher typed together (sometimes split, sometimes misaligned)
    () => {
      const cls = pick(classes), d = Math.floor(rand() * 6), p = 1 + Math.floor(rand() * 8);
      api.ClassViewManager.renderClassView(cls);
      const n = 1 + Math.floor(rand() * 3), m = rand() < 0.2 ? n + 1 : n;
      const subj = Array.from({ length: n }, () => pick(['Hindi', 'IT', 'Art', 'Maths', 'Science'])).join(' / ');
      const tch = Array.from({ length: m }, () => pick(teachers)).join(' / ');
      const r = ss.getSheetByName('Class_View').getRange(6 + d * 2, 1 + p, 2, 1);
      r.setValues([[subj], [tch]]);
      api.onEdit(editEvent(env, r, { triggerUid: uid }));
    },
    // Teacher_View: take a period, or free it
    () => {
      const t = pick(teachers), d = Math.floor(rand() * 6), p = 1 + Math.floor(rand() * 8);
      api.TeacherViewManager.renderTeacherView(t);
      const r = ss.getSheetByName('Teacher_View').getRange(6 + d, 1 + p);
      r.setValue(rand() < 0.3 ? 'FREE' : pick(classes) + (rand() < 0.5 ? ' - ' + pick(['Maths', 'Hindi', 'IT']) : ''));
      api.onEdit(editEvent(env, r, { triggerUid: uid }));
    },
    // Master_Schedule: pick a teacher in a random row
    () => {
      const row = 2 + Math.floor(rand() * (ms().getLastRow() - 1));
      const r = ms().getRange(row, 6);
      const old = r.getValue();
      r.setValue(pick(teachers));
      api.onEdit(editEvent(env, r, { triggerUid: uid, oldValue: old === '' ? undefined : old }));
    },
    // Garbage: unknown names, a teacher in the subject row, clearing only the subject
    () => {
      const cls = pick(classes), d = Math.floor(rand() * 6), p = 1 + Math.floor(rand() * 8);
      api.ClassViewManager.renderClassView(cls);
      const cv = ss.getSheetByName('Class_View');
      const kind = Math.floor(rand() * 3);
      const r = kind === 1 ? cv.getRange(7 + d * 2, 1 + p) : cv.getRange(6 + d * 2, 1 + p);
      r.setValue(kind === 0 ? pick(teachers) : kind === 1 ? 'Nobody ' + Math.floor(rand() * 99) : '');
      api.onEdit(editEvent(env, r, { triggerUid: uid }));
    },
  ];

  for (let k = 0; k < OPS_N; k++) {
    const toastsBefore = ss.toasts.length;
    pick(ops)();
    const fresh = ss.toasts.slice(toastsBefore);
    assert.ok(!fresh.some(t => /sync failed/.test(t.title)), 'edit ' + k + ' threw: ' + JSON.stringify(fresh));
    if (fresh.some(t => t.title === 'Not saved')) stats.refused++; else stats.accepted++;
    const now = clashCount();
    assert.ok(now <= clashes, 'edit ' + k + ' added a teacher clash (' + clashes + ' → ' + now + ')');
    clashes = now;
    assert.ok(mismatches() <= mismatchStart, 'edit ' + k + ' left subjects and teacher groups misaligned');
    assert.ok(unknownTeachers() <= unknownStart, 'edit ' + k + ' saved an unknown teacher name');
    const l = logRows();
    assert.ok(l >= logs, 'Change_Log shrank');
    logs = l;
  }
  // every view still renders
  classes.forEach(c => api.ClassViewManager.renderClassView(c));
  teachers.slice(0, 20).forEach(t => api.TeacherViewManager.renderTeacherView(t));
  days.forEach(d => api.TeacherDayViewManager.renderTeacherDayView(d, 'All Classes'));
  api.refreshAllViews(true);
  assert.ok(stats.accepted > 0 && stats.refused > 0, 'both paths exercised: ' + JSON.stringify(stats));
}));
