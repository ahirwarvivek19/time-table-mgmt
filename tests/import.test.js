'use strict';
/**
 * Tests for the reusable import utility: the flat format, source lists, class
 * replacement, and split / combined lectures through the flat path.
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const os = require('os');
const path = require('path');
const XLSX = require('xlsx');
const { execFileSync, spawnSync } = require('child_process');
const { buildAll, parseArgs } = require('../tools/convert/build');
const { parseFlat, toFlatCsv, dayOf, resolveTeachers, teacherIndex } = require('../tools/convert/flat');
const { analyze, canonicalize } = require('../tools/convert/lib/analyze');

const ROOT = path.resolve(__dirname, '..');
const BUILD = path.join(ROOT, 'tools/convert/build.js');
const tmp = () => fs.mkdtempSync(path.join(os.tmpdir(), 'tt-import-'));
const slot = r => `${r.Class}|${r.Day}|${r.Period}`;
const writeCsv = (dir, name, lines) => { const f = path.join(dir, name); fs.writeFileSync(f, lines.join('\n') + '\n'); return f; };
const sourcesFile = (dir, list) => { const f = path.join(dir, 'sources.json'); fs.writeFileSync(f, JSON.stringify({ label: 'test', sources: list })); return f; };

let base;
const baseBuild = () => (base = base || buildAll());

test('flat CSV round trip: export the 2026-27 build, re-import it alone, get the same timetable', () => {
  const dir = tmp();
  const built = baseBuild();
  const csv = path.join(dir, 'all.csv');
  fs.writeFileSync(csv, toFlatCsv(built.schedule));
  const again = buildAll({ sources: sourcesFile(dir, [{ format: 'flat', file: 'all.csv' }]) });
  assert.deepEqual(again.errors, []);
  const pick = r => [slot(r), r.Subject, r.Teacher, r.Room].join('|');
  assert.deepEqual(again.schedule.map(pick).sort(), built.schedule.map(pick).sort());
  assert.equal(again.report.clashes.length, built.report.clashes.length);
  assert.equal(again.report.combined.length, built.report.combined.length);
  assert.equal(again.report.splits.length, built.report.splits.length);
});

test('--add flat file replaces only the classes it covers, and says so', () => {
  const dir = tmp();
  const f = writeCsv(dir, '9a.csv', ['Class,Day,Period,Subject,Teacher,Room', '9th A,Mon,1,Mathematics,Deepti Tiwari,207']);
  const built = buildAll({ add: ['flat:' + f] });
  assert.deepEqual(built.errors, []);
  const nineA = built.schedule.filter(r => r.Class === '9th A');
  assert.equal(nineA.length, 48, 'missing slots are filled as free periods');
  assert.equal(nineA.filter(r => r.Subject).length, 1);
  assert.equal(nineA.find(r => r.Day === 'Monday' && r.Period === 1).Teacher, 'Deepti Tiwari');
  assert.ok(built.replaced.some(x => /9th A from Class-wise sheet \(Timetable\.xlsx, sheet 9-12\) replaced by a later source/.test(x)));
  // every other class is untouched
  const other = r => r.Class !== '9th A';
  const pick = r => [slot(r), r.Subject, r.Teacher].join('|');
  assert.deepEqual(built.schedule.filter(other).map(pick).sort(), baseBuild().schedule.filter(other).map(pick).sort());
});

test('flat reader: day forms, teacher short names, split + co-teacher syntax, rooms per group', () => {
  assert.equal(dayOf('mon'), 'Monday');
  assert.equal(dayOf('THU.'), 'Thursday');
  assert.equal(dayOf('शनिवार'), 'Saturday');
  assert.equal(dayOf('Funday'), '');
  const idx = teacherIndex();
  assert.deepEqual(resolveTeachers('ekta gour / Divya Lodval+reena shivhare / -', idx), { text: 'Ekta Gour / Divya Lodval + Reena Shivhare / -', unknown: [] });
  assert.deepEqual(resolveTeachers('Ruchika V.', idx), { text: 'Ruchika Vyas', unknown: [] });

  const dir = tmp();
  const f = writeCsv(dir, 'x.csv', [
    'room,TEACHER,Subject,period,Day,class', // any column order, any case
    '12 / Lab / 14,Ekta Gour / Divya Lodval + Reena Shivhare / Krishna Patale,Hindi / IT / Beauty,1,Monday,11th A',
  ]);
  const { rows, issues } = parseFlat(f);
  assert.deepEqual(issues, []);
  const r = rows.find(x => x.Class === '11th A' && x.Day === 'Monday' && x.Period === 1);
  assert.equal(r.Subject, 'Hindi / IT / Beauty');
  assert.equal(r.Teacher, 'Ekta Gour / Divya Lodval + Reena Shivhare / Krishna Patale');
  assert.equal(r.Room, '12 / Lab / 14');
  assert.equal(rows.length, 48);
});

test('flat reader reports every bad line instead of guessing', () => {
  const dir = tmp();
  const f = writeCsv(dir, 'bad.csv', [
    'Class,Day,Period,Subject,Teacher',
    '13th Z,Mon,1,Maths,Deepti Tiwari',
    '9th A,Someday,1,Maths,Deepti Tiwari',
    '9th A,Mon,9,Maths,Deepti Tiwari',
    '9th A,Mon,1,Maths,Mr Nobody',
    '9th A,Mon,1,Maths,Deepti Tiwari',
    '9th A,Mon,2,,Deepti Tiwari',
  ]);
  const { issues } = parseFlat(f);
  assert.match(issues.join('\n'), /line 2: unknown class "13th Z"/);
  assert.match(issues.join('\n'), /line 3: unknown day "Someday"/);
  assert.match(issues.join('\n'), /line 4: period "9" must be 1–8/);
  assert.match(issues.join('\n'), /line 5: unknown teacher "Mr Nobody"/);
  assert.match(issues.join('\n'), /line 6: 9th A Monday P1 already given on line 5/);
  assert.match(issues.join('\n'), /line 7: teacher "Deepti Tiwari" given without a subject/);
  assert.throws(() => parseFlat(writeCsv(dir, 'nohdr.csv', ['Class,Day,Period', '9th A,Mon,1'])), /missing column\(s\) subject, teacher/);
});

test('flat import: combined lecture is not a clash; same teacher with another subject is; split groups line up', () => {
  const dir = tmp();
  const f = writeCsv(dir, 'mix.csv', [
    'Class,Day,Period,Subject,Teacher',
    '9th C,Mon,1,Social Science,Rubal Verma',      // combined: 9th C + 9th D together
    '9th D,Mon,1,Social Science,Rubal Verma',
    '9th E,Tue,1,Social Science,Rubal Verma',      // clash: same teacher, other subject, other class
    '10th A,Tue,1,Geography,Rubal Verma',
    '11th A,Wed,1,Hindi / IT,Ekta Gour / Divya Lodval', // split period
    '11th B,Wed,1,Hindi / IT,Ekta Gour / Divya Lodval', // ...combined with 11th B (same grade)
    '10th B,Wed,1,Hindi,Ekta Gour',                     // Ekta Gour in grade 10 at the same time → clash
  ]);
  const parsed = parseFlat(path.join(dir, 'mix.csv'));
  assert.deepEqual(parsed.issues, []);
  const canon = canonicalize(parsed.rows);
  assert.deepEqual(canon.errors, []);
  const { clashes, combined, splits } = analyze(canon.rows);
  assert.ok(combined.some(c => c.teacher === 'Rubal Verma' && c.day === 'Monday' && c.classes.join() === '9th C,9th D'));
  assert.ok(combined.some(c => c.teacher === 'Divya Lodval' && c.day === 'Wednesday' && c.classes.join() === '11th A,11th B'));
  assert.ok(clashes.some(c => c.teacher === 'Rubal Verma' && c.day === 'Tuesday'));
  assert.ok(clashes.some(c => c.teacher === 'Ekta Gour' && c.day === 'Wednesday'));
  assert.ok(!clashes.some(c => c.teacher === 'Divya Lodval'));
  assert.equal(splits.filter(s => s.day === 'Wednesday').length, 2);
});

test('flat import: subjects and teacher groups that do not line up fail the build', () => {
  const dir = tmp();
  writeCsv(dir, 'mis.csv', ['Class,Day,Period,Subject,Teacher', '11th A,Mon,1,Hindi / IT / Beauty,Ekta Gour / Divya Lodval']);
  const built = buildAll({ add: ['flat:' + path.join(dir, 'mis.csv')] });
  assert.match(built.errors.join('\n'), /3 subjects but 2 teacher groups/);
});

test('build CLI: unknown options and missing files fail loudly; --check writes nothing', () => {
  assert.throws(() => parseArgs(['--chek']), /Unknown option --chek/);
  assert.throws(() => parseArgs(['--add']), /--add needs a value/);
  const r = spawnSync(process.execPath, [BUILD, '--check', '--add', 'flat:/no/such/file.csv'], { encoding: 'utf8' });
  assert.equal(r.status, 1);
  assert.match(r.stderr, /Source file not found/);
  const before = fs.statSync(path.join(ROOT, 'PrefillData.gs')).mtimeMs;
  execFileSync(process.execPath, [BUILD, '--check'], { encoding: 'utf8' });
  assert.equal(fs.statSync(path.join(ROOT, 'PrefillData.gs')).mtimeMs, before);
});

test('PrefillData.gs is up to date with the sources (re-run npm run convert after changing the converter)', () => {
  const built = baseBuild();
  const gs = fs.readFileSync(path.join(ROOT, 'PrefillData.gs'), 'utf8');
  const body = gs.slice(gs.indexOf('const RAW_SCHEDULE'));
  const rows = body.split('\n').filter(l => l.startsWith('  {')).map(l => JSON.parse(l.trim().replace(/,$/, '')));
  const pick = r => [slot(r), r.Subject, r.Teacher, r.Room].join('|');
  assert.deepEqual(rows.map(pick).sort(), built.schedule.map(pick).sort());
});

test('templates/timetable-template.csv is a valid flat file that shows split, co-taught and combined periods', () => {
  const { rows, issues } = parseFlat(path.join(ROOT, 'templates/timetable-template.csv'));
  assert.deepEqual(issues, []);
  const canon = canonicalize(rows);
  assert.deepEqual(canon.errors, []);
  const rep = analyze(canon.rows);
  assert.ok(rep.splits.some(s => s.cls === '11th A' && s.lanes.length === 3));
  assert.ok(rep.combined.some(c => c.teacher === 'Rubal Verma' && c.classes.join() === '9th C,9th D'));
  assert.deepEqual(rep.clashes, []);
});

test('documented commands run: --add flat Excel with #sheet, and convert:export-flat output re-imports', () => {
  const dir = tmp();
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['note: header may be below a title'], [], ['Class', 'Day', 'Period', 'Subject', 'Teacher'], ['9th B', 'Tue', 3, 'Mathematics', 'ritcha choudhary']]), 'Sheet1');
  const f = path.join(dir, 'n.xlsx');
  XLSX.writeFile(wb, f);
  const out = execFileSync(process.execPath, [BUILD, '--check', '--add', `flat:${f}#Sheet1`], { encoding: 'utf8' });
  assert.match(out, /9th B from Class-wise sheet \(Timetable\.xlsx, sheet 9-12\) replaced by a later source/);
  const csv = path.join(dir, 'flat.csv');
  execFileSync(process.execPath, [BUILD, '--check', '--export-flat', csv], { encoding: 'utf8' });
  assert.equal(fs.readFileSync(csv, 'utf8').split('\n').filter(Boolean).length, 1 + 28 * 48);
});
