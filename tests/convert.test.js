'use strict';
/**
 * Tests for the Excel readers: the cell grammar, the class-wise sheets (6-8, 9-12)
 * and the teacher-wise sheet that fills in and cross-checks teachers.
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const os = require('os');
const path = require('path');
const XLSX = require('xlsx');
const { execFileSync } = require('child_process');
const { parseCell } = require('../tools/convert/lib/cell-grammar');
const { parseClassBlocks, readTimes } = require('../tools/convert/class-blocks');
const { parseTeacherwise, parseClaimCell, fillFromClaims, compare } = require('../tools/convert/teacherwise');
const { buildAll } = require('../tools/convert/build');
const { TEACHERS, CLASSES, PERIODS } = require('../tools/convert/registry');

const ROOT = path.resolve(__dirname, '..');
const BUILD = path.join(ROOT, 'tools/convert/build.js');
const EXCEL = path.join(ROOT, 'Timetable.xlsx');
const read = raw => parseCell(raw).lanes.map(l => [l.subject, l.teachers.join(' + ')]);
const tmp = () => fs.mkdtempSync(path.join(os.tmpdir(), 'tt-convert-'));
const slot = (rows, c, d, p) => rows.find(r => r.Class === c && r.Day === d && r.Period === p);

let base;
const baseBuild = () => (base = base || buildAll());

/** A copy of Timetable.xlsx with some cells changed, plus a sources.json that reads it. */
function editedCopy(edits) {
  const dir = tmp();
  const wb = XLSX.readFile(EXCEL);
  Object.entries(edits).forEach(([ref, v]) => { const [sheet, a] = ref.split('!'); wb.Sheets[sheet][a] = { t: 's', v }; });
  XLSX.writeFile(wb, path.join(dir, 'Timetable.xlsx'));
  const cfg = JSON.parse(fs.readFileSync(path.join(ROOT, 'tools/convert/sources.json'), 'utf8'));
  fs.writeFileSync(path.join(dir, 'sources.json'), JSON.stringify(cfg));
  return path.join(dir, 'sources.json');
}

test('converter --check passes: 28 classes × 48 periods, every teacher registered', () => {
  const out = execFileSync(process.execPath, [BUILD, '--check'], { encoding: 'utf8' });
  assert.match(out, /✓ 28 classes/);
  assert.match(out, /1344 schedule rows/);
  assert.match(out, /0 blank slots/);
});

test('cell grammar reads every cell shape in both class sheets', () => {
  // classes 6–8
  assert.deepEqual(read('ENG \nRuchika V.'), [['English', 'Ruchika Vyas']]);
  assert.deepEqual(read('\nRuchika V./UMANG ADDHA PHOOL '), [['Life Skills Umang (Aadha Phool)', 'Ruchika Vyas']], 'teacher written first');
  assert.deepEqual(read('UMANG-ADDHA PHOOL\nLalita / Farhana'), [['Life Skills Umang (Aadha Phool)', 'Lalita Tiwari + Farhana Khan']]);
  assert.deepEqual(read('Sanskrit Lalita Tiwari/ M. Sharma '), [['Sanskrit', 'Lalita Tiwari + Mahesh Chandra Sharma']], 'one subject, two teachers = co-teaching');
  assert.deepEqual(read('Library/ COMPUTER'), [['Library', ''], ['Computer', '']], 'split period, no teacher written');
  assert.deepEqual(read('Library/ '), [['Library', '']]);
  assert.deepEqual(read('DAKSHATA SO.SCIENCE '), [['Proficiency Social Science', '']]);
  assert.deepEqual(read('DAKASHTA ENGLISH'), [['Proficiency English', '']], 'misspelling');
  assert.deepEqual(read('Maths seema parmar.'), [['Mathematics', 'Seema Parmar']]);
  assert.deepEqual(read('DANCE RUPALI'), [['Dance', 'Rupali Chuliwal']]);
  // classes 9–12
  assert.deepEqual(read('SANSKRIT  NEHA JAMIDAR'), [['Sanskrit', 'Neha Jamidar']]);
  assert.deepEqual(read('SANSKRIT NEHA.J          IT-DIVYA BEAUTY-KRISHNA'), [['Sanskrit', 'Neha Jamidar'], ['IT', 'Divya Lodval'], ['Beauty', 'Krishna Patale']], 'groups without "/"');
  assert.deepEqual(read('IT- DIVYA/ POLY.SCI/ BEAUTY KRISHNA'), [['IT', 'Divya Lodval'], ['Political Science', ''], ['Beauty', 'Krishna Patale']]);
  assert.deepEqual(read('SO.SCI RUBAL VERMA(H.M)/ I.MATHUR (E.M)'), [['Social Science (H.M)', 'Rubal Verma'], ['Social Science (E.M)', 'Indira Mathur']], 'one subject, a group per medium');
  assert.deepEqual(read('SO.SCI ROSHNI SUMAN (H.M)  INDRA.M (E.M)'), [['Social Science (H.M)', 'Roshini Suman'], ['Social Science (E.M)', 'Indira Mathur']]);
  assert.deepEqual(read('PHYSICS GK.DIXIT (E.M) / CHEMISTRY D.BHATI (H.M)'), [['Physics (E.M)', 'Gautam Kumar Dixit'], ['Chemistry (H.M)', 'Dinesh Bhati']]);
  assert.deepEqual(read('ACCOUNTS (H.M)     ECO I.MATHUR (E.M)'), [['Accountancy (H.M)', ''], ['Economics (E.M)', 'Indira Mathur']]);
  assert.deepEqual(read('B.STUD(E.M)/ECO AK.PANDIT  (H.M)'), [['Business Studies (E.M)', ''], ['Economics (H.M)', 'Arun Kumar Pandit']]);
  assert.deepEqual(read('CCLE INDRA.M/ KRISHNA.P'), [['CCLE', 'Indira Mathur + Krishna Patale']]);
  assert.deepEqual(read('MATHS /CCLE NEHA JHALA'), [['Mathematics & CCLE', 'Neha Jhala']], 'class teacher takes both');
  assert.deepEqual(read('PLOY.SCI (RO.S)/ HINDI  (ARCHANA.P) '), [['Political Science', 'Roshini Suman'], ['Hindi', 'Archana Pandey']], 'teacher in brackets');
  assert.deepEqual(read('GEOGRAPHY (PRACTICAL)'), [['Geography Practical', '']]);
  assert.deepEqual(read('UHV A.KARTARE'), [['UHV', 'Anuradha Katare']]);
  assert.deepEqual(read('ACCOUNTS  WAHID SIR'), [['Accountancy', 'Sayyed Wahid Hussain']]);
  assert.deepEqual(read(''), []);
  assert.match(parseCell('MATHS MR NOBODY').problems.join(), /unknown word "MR"/, 'unknown words are reported, not guessed');
});

test('class sheets: every cell reads cleanly, 6 days per class, bell times and lunch come from the header', () => {
  for (const [sheet, n] of [['6-8', 12], ['9-12', 16]]) {
    const r = parseClassBlocks(EXCEL, sheet);
    assert.deepEqual(r.issues, [], sheet);
    assert.equal(new Set(r.rows.map(x => x.Class)).size, n);
    assert.equal(r.rows.length, n * 48);
    assert.deepEqual(r.bell, PERIODS);
    assert.equal(r.lunch, '12:20–12:50');
    const registered = new Set(TEACHERS.map(t => t[0]));
    r.rows.forEach(row => row.Teacher.split(/\s*[/+]\s*/).filter(t => t && t !== '-').forEach(t => assert.ok(registered.has(t), `${row.Source}: ${t}`)));
  }
  assert.deepEqual(readTimes('12:50- 01:30'), ['12:50', '13:30']);
  assert.deepEqual(readTimes('02:50-03:30'), ['14:50', '15:30']);
  assert.deepEqual(readTimes('09:30-10:20'), ['09:30', '10:20']);
  assert.ok(CLASSES.some(c => c[0] === '11th C') && !CLASSES.some(c => c[0] === '6th E'));
});

test('teacher-wise cells: unnamed days mean the other days; days may follow the class; hints and medium', () => {
  const days = text => {
    const by = {};
    parseClaimCell(text).claims.forEach(c => { const k = c.cls + (c.subject ? ` [${c.subject}]` : '') + (c.medium ? ` (${c.medium})` : ''); (by[k] = by[k] || []).push(c.day.slice(0, 3)); });
    return by;
  };
  assert.deepEqual(days('11TH A SAT 9TH A'), { '11th A': ['Mon', 'Tue', 'Wed', 'Thu', 'Fri'], '9th A': ['Sat'] });
  assert.deepEqual(days('12TH C (ACC) MON TO THU'), { '12th C [Accountancy]': ['Mon', 'Tue', 'Wed', 'Thu'] });
  assert.deepEqual(days('MON-TUE12TH D (HM) WED-THU12TH B(EM)'), { '12th D (H)': ['Mon', 'Tue'], '12th B (E)': ['Wed', 'Thu'] });
  assert.deepEqual(days('WED-THU-10TH B+C'), { '10th B': ['Wed', 'Thu'], '10th C': ['Wed', 'Thu'] });
  assert.deepEqual(days('SAT-9THE(UHV)'), { '9th E [UHV]': ['Sat'] });
  assert.deepEqual(days('11TH C B.STU(E.M)'), { '11th C [Business Studies] (E)': ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'] });
  assert.match(parseClaimCell('LIBRARY').problems.join(), /without a class/);
});

test('teacher-wise fill: one fitting teacher fills the group; wrong subject, wrong medium, busy or tied candidates do not', () => {
  const row = (Class, Subject, Teacher, Day = 'Monday', Period = 3) => ({ Day, Period, Class, Tier: 'Middle', Subject, Teacher, Room: '' });
  const rows = [
    row('6th A', 'Proficiency English', ''),
    row('6th A', 'English', 'Ruchika Vyas', 'Tuesday', 1),          // Ruchika teaches 6th A English → strong fit
    row('11th C', 'Accountancy (H.M) / Economics (E.M)', '- / Indira Mathur'),
    row('12th C', 'Accountancy', 'Sayyed Wahid Hussain', 'Monday', 4),
    row('11th C', 'Business Studies (H.M)', '', 'Monday', 4),      // Wahid busy in 12th C at Mon P4
    row('7th B', 'Library', ''),
  ];
  const claim = (teacher, cls, extra = {}) => ({ teacher, day: 'Monday', period: extra.period || 3, cls, subject: '', medium: '', cell: 'T!X1', text: '…', ...extra });
  const tw = { teachers: [], claims: [
    claim('Ruchika Vyas', '6th A'), claim('Manisha Lovanshi', '6th A'),           // two claims: subject affinity picks Ruchika
    claim('Sayyed Wahid Hussain', '11th C', { subject: 'Accountancy' }),
    claim('Indira Mathur', '11th C', { subject: 'Economics', medium: 'E' }),
    claim('Sayyed Wahid Hussain', '11th C', { period: 4, subject: 'Business Studies', medium: 'H' }),
    claim('Somesh Vyas', '7th B'), claim('Roshi Jahan', '7th B'),                 // nobody teaches Library → no fill
  ] };
  const { rows: out, filled, unfilled } = fillFromClaims(rows, tw);
  assert.equal(slot(out, '6th A', 'Monday', 3).Teacher, 'Ruchika Vyas');
  assert.equal(slot(out, '11th C', 'Monday', 3).Teacher, 'Sayyed Wahid Hussain / Indira Mathur');
  assert.equal(slot(out, '11th C', 'Monday', 4).Teacher, '', 'busy teacher is not used');
  assert.match(unfilled.find(u => u.cls === '11th C').candidates.join(), /class-wise sheet has them in 12th C Accountancy/);
  assert.equal(slot(out, '7th B', 'Monday', 3).Teacher, '');
  assert.equal(filled.length, 2);
  const hindi = fillFromClaims([row('9th A', 'Hindi (H.M)', '')], { teachers: [], claims: [claim('Ekta Gour', '9th A', { medium: 'E' })] });
  assert.equal(hindi.rows[0].Teacher, '', 'a claim for the other medium is not used');
});

test('2026-27 build: teachers filled from the teacher-wise sheet, defaults and the two-sheet comparison', () => {
  const b = baseBuild();
  const at = (c, d, p) => slot(b.schedule, c, d, p);
  assert.equal(at('6th A', 'Monday', 3).Teacher, 'Ruchika Vyas', 'Proficiency English ← Teacherwise!L27');
  assert.equal(at('11th A', 'Monday', 1).Teacher, 'Rubal Verma', 'Geography ← Teacherwise!J21');
  assert.equal(at('12th C', 'Monday', 3).Teacher, 'Sayyed Wahid Hussain / Indira Mathur / Rakesh Kumar Jaiswal');
  assert.equal(at('6th A', 'Saturday', 4).Teacher, '- / Reena Shivhare', 'unnamed Computer → the only computer teacher');
  assert.ok(b.missing.some(m => m.cls === '11th C' && m.period === 6 && /class-wise sheet has them in 12th C/.test(m.candidates.join())));
  // the two sheets disagree about 8th C: class sheet P2 Science / P5 Sanskrit, teacher sheet the other way round
  const tw = b.diff.teacherOnly.filter(x => x.cls === '8th C' && x.day === 'Monday');
  assert.ok(tw.some(x => x.teacher === 'Jitendra Dhawan' && x.period === 5));
  assert.ok(tw.some(x => x.teacher === 'Sonali Sharma' && x.period === 2));
  assert.ok(b.diff.classOnly.some(x => x.teacher === 'Jitendra Dhawan' && x.cls === '8th C' && x.period === 2));
  assert.ok(b.report.clashes.some(c => c.teacher === 'Jitendra Dhawan' && c.period === 2), 'the swap double-books Jitendra');
  // medium groups of one section join another section's lecture: no clash
  assert.ok(!b.report.clashes.some(c => c.teacher === 'Ekta Gour' && c.period === 1));
  assert.ok(b.report.combined.some(c => c.teacher === 'Ekta Gour' && c.period === 1 && c.classes.join() === '11th B,11th C,11th D'));
  assert.ok(b.tw.notes.some(n => /Dinesh Kumar Pandya is listed 2 times/.test(n)));
  assert.equal(b.tw.classTeachers.get('9th A')[0], 'Neha Jamidar');
});

test('a changed bell time, an unknown class or an unknown word stops the build and says where', () => {
  const bell = buildAll({ sources: editedCopy({ '9-12!D4': '09:40-10:20' }) });
  assert.match(bell.errors.join('\n'), /sheet 9-12\): bell times P1 09:40–10:20.*differ from PERIODS in registry\.js/);
  const cls = buildAll({ sources: editedCopy({ '6-8!B5': 'कक्षा 6 F' }) });
  assert.match(cls.errors.join('\n'), /B5: class "कक्षा 6 F" → "6th F" is not in CLASSES/);
  assert.match(cls.errors.join('\n'), /6th A: no source gives a timetable/);
  const word = buildAll({ sources: editedCopy({ '9-12!D5': 'SANSKRIT NEHA JAMIDAR XYZ' }) });
  assert.match(word.errors.join('\n'), /9th A Monday P1 \(9-12!D5\): unknown word "XYZ"/);
});

test('teacher-wise sheet is optional: a source list without it still builds', () => {
  const dir = tmp();
  fs.writeFileSync(path.join(dir, 's.json'), JSON.stringify({ label: 't', sources: [{ format: 'class-blocks', file: EXCEL, sheet: '6-8' }, { format: 'class-blocks', file: EXCEL, sheet: '9-12' }] }));
  const b = buildAll({ sources: path.join(dir, 's.json') });
  assert.deepEqual(b.errors, []);
  assert.equal(b.tw, null);
  assert.equal(slot(b.schedule, '6th A', 'Monday', 3).Teacher, '', 'nothing is filled without it');
  assert.equal(slot(b.schedule, '6th A', 'Saturday', 4).Teacher, '- / Reena Shivhare');
});
