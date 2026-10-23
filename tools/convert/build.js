#!/usr/bin/env node
/**
 * build.js — converts the Excel timetables listed in sources.json into Apps Script pre-fill data.
 *
 *   npm run convert                         → writes PrefillData.gs + docs/TIMETABLE_REVIEW_<label>.md
 *   npm run convert:check                   → only validates, writes nothing (exit 1 on errors)
 *   node tools/convert/build.js --sources my-sources.json      → use another source list
 *   node tools/convert/build.js --add flat:path/to/file.csv    → add a source (replaces the classes it covers)
 *   node tools/convert/build.js --export-flat out.csv          → also write the result as a flat CSV
 *
 * See docs/IMPORTING.md for the formats, the commands and how to handle a new Excel file.
 */
'use strict';

const fs = require('fs');
const path = require('path');
const { DAYS, TEACHERS, CLASSES, PERIODS } = require('./registry');
const { parseClassBlocks } = require('./class-blocks');
const { parseFlat, toFlatCsv } = require('./flat');
const { parseTeacherwise, fillFromClaims, compare, baseOf } = require('./teacherwise');
const { DEFAULT_TEACHER } = require('./lib/vocabulary');
const { canonicalize, analyze, collapseDays, dayList } = require('./lib/analyze');
const { loadGrammar, ROOT } = require('./lib/gas');

const OUT_GS = path.join(ROOT, 'PrefillData.gs');
const DEFAULT_SOURCES = path.join(__dirname, 'sources.json');
const FORMAT_NAMES = { 'class-blocks': 'Class-wise sheet', flat: 'Flat table' };
const P = () => loadGrammar().ScheduleParser;

/** Every teacher in a row, across all groups. */
const teachersOf = row => P().parseRowAssignments(row).map(a => a.teacher).filter(Boolean);
const classOrder = new Map(CLASSES.map((c, i) => [c[0], i]));
const bySlot = (a, b) => (classOrder.get(a.cls) - classOrder.get(b.cls)) || DAYS.indexOf(a.day) - DAYS.indexOf(b.day) || a.period - b.period;

/** Reads one source and returns a uniform result. */
function readSource(src, opts) {
  const file = path.resolve(opts.baseDir, src.file);
  if (!fs.existsSync(file)) throw new Error(`Source file not found: ${file}`);
  const empty = { notes: [], ignored: [], blanks: [], issues: [], bell: [], cells: new Map() };
  if (src.format === 'class-blocks') return { ...empty, ...parseClassBlocks(file, src.sheet) };
  if (src.format === 'flat') return { ...empty, ...parseFlat(file, src.sheet) };
  throw new Error(`Unknown source format "${src.format}" — use class-blocks or flat`);
}

/** Loads the source list (sources.json, or --sources) plus any --add format:file entries. */
function loadSources(opts) {
  const cfgFile = path.resolve(opts.sources || DEFAULT_SOURCES);
  const cfg = JSON.parse(fs.readFileSync(cfgFile, 'utf8'));
  const list = [...cfg.sources];
  (opts.add || []).forEach(spec => {
    const m = spec.match(/^([a-z0-9-]+):(.+?)(?:#(.+))?$/);
    if (!m) throw new Error(`--add expects format:file[#sheet], got "${spec}"`);
    list.push({ format: m[1], file: path.resolve(m[2]), sheet: m[3] });
  });
  // Files in the default sources.json are relative to the repo root; elsewhere to the JSON file.
  const baseDir = cfgFile === DEFAULT_SOURCES ? ROOT : path.dirname(cfgFile);
  const tw = cfg.teacherwise ? { ...cfg.teacherwise, file: path.resolve(baseDir, cfg.teacherwise.file) } : null;
  return { label: cfg.label || 'timetable', sources: list, baseDir, teacherwise: tw };
}

/** Gives teacher-less groups of a subject in DEFAULT_TEACHER their only possible teacher. */
function applyDefaults(rows) {
  const applied = [];
  const out = rows.map(r => {
    const parsed = P().parseLanes(r);
    let changed = false;
    parsed.lanes.forEach(l => {
      const t = DEFAULT_TEACHER[l.subject];
      if (t && !l.teachers.length) { l.teachers = [t]; changed = true; applied.push({ cls: r.Class, day: r.Day, period: r.Period, subject: l.subject, teacher: t, source: r.Source || '' }); }
    });
    if (!changed) return r;
    const f = P().formatLanes(parsed.lanes);
    return { ...r, Subject: f.subject, Teacher: f.teacher };
  });
  return { rows: out, applied };
}

/** Groups that still have no teacher, with the teacher of that subject for the class (or school) as a suggestion. */
function teacherless(rows, unfilled) {
  const teaches = new Map(); // class|base → Set(teacher)
  const anywhere = new Map(); // base → Set(teacher)
  rows.forEach(r => P().parseLanes(r).lanes.forEach(l => l.teachers.forEach(t => {
    const b = baseOf(l.subject), k = `${r.Class}|${b}`;
    (teaches.get(k) || teaches.set(k, new Set()).get(k)).add(t);
    (anywhere.get(b) || anywhere.set(b, new Set()).get(b)).add(t);
  })));
  const cand = new Map(unfilled.map(u => [`${u.cls}|${u.day}|${u.period}|${u.subject}`, u.candidates]));
  const out = [];
  rows.forEach(r => P().parseLanes(r).lanes.forEach(l => {
    if (!String(l.subject || '').trim() || l.teachers.length) return;
    const b = baseOf(l.subject);
    const own = [...(teaches.get(`${r.Class}|${b}`) || [])];
    const other = [...(anywhere.get(b) || [])];
    out.push({
      cls: r.Class, day: r.Day, period: r.Period, subject: l.subject, source: r.Source || '',
      suggestion: own.length ? own.join(', ') : other.length ? `${other.join(', ')} (in other classes)` : '',
      candidates: cand.get(`${r.Class}|${r.Day}|${r.Period}|${l.subject}`) || [],
    });
  }));
  return out.sort(bySlot);
}

/**
 * Builds and validates the full timetable.
 * A class given by a later source replaces that class from earlier sources.
 * @param {{sources?: string, add?: string[]}} opts
 */
function buildAll(opts = {}) {
  const errors = [], warnings = [];
  const { label, sources, baseDir, teacherwise } = loadSources(opts);
  const parts = sources.map(src => ({ src, ...readSource(src, { ...opts, baseDir }) }));
  const owner = new Map(); // class → index of the last source that supplies it
  parts.forEach((p, i) => new Set(p.rows.map(r => r.Class)).forEach(c => owner.set(c, i)));
  const replaced = [];
  const classOfBlank = b => b.replace(/ (Monday|Tuesday|Wednesday|Thursday|Friday|Saturday) P\d$/, '');
  const fmtBell = b => b.map(([n, s, e]) => `P${n} ${s}–${e}`).join(', ');
  parts.forEach((p, i) => {
    p.name = `${FORMAT_NAMES[p.src.format] || p.src.format} (${path.basename(p.src.file)}${p.src.sheet ? ', sheet ' + p.src.sheet : ''})`;
    p.issues.forEach(x => errors.push(`${p.name}: ${x}`));
    if (p.bell.length && JSON.stringify(p.bell) !== JSON.stringify(PERIODS)) {
      errors.push(`${p.name}: bell times ${fmtBell(p.bell)} differ from PERIODS in registry.js (${fmtBell(PERIODS)}) — update PERIODS`);
    }
    const dropped = [...new Set(p.rows.filter(r => owner.get(r.Class) !== i).map(r => r.Class))];
    if (dropped.length) replaced.push(`${dropped.join(', ')} from ${p.name} replaced by a later source`);
    p.rows = p.rows.filter(r => owner.get(r.Class) === i);
    p.blanks = p.blanks.filter(b => owner.get(classOfBlank(b)) === i);
  });

  const canon = canonicalize(parts.flatMap(p => p.rows));
  canon.errors.forEach(e => errors.push(e));
  let schedule = canon.rows;

  // ── Teacher-wise sheet: fill missing teachers, then compare the two sheets ──
  let tw = null, fills = { filled: [], unfilled: [] }, diff = { classOnly: [], teacherOnly: [] };
  if (teacherwise) {
    if (!fs.existsSync(teacherwise.file)) throw new Error(`Teacher-wise file not found: ${teacherwise.file}`);
    tw = parseTeacherwise(teacherwise.file, teacherwise.sheet);
    tw.name = `Teacher-wise sheet (${path.basename(teacherwise.file)}, sheet ${teacherwise.sheet})`;
    tw.issues.forEach(x => warnings.push(x));
    fills = fillFromClaims(schedule, tw);
    schedule = fills.rows;
    diff = compare(schedule, tw);
  }
  const defaults = applyDefaults(schedule);
  schedule = defaults.rows;
  const missing = teacherless(schedule, fills.unfilled);

  // ── Referential checks ────────────────────────────────────────────────────
  const teacherSet = new Set(TEACHERS.map(t => t[0]));
  const classSet = new Set(CLASSES.map(c => c[0]));
  if (teacherSet.size !== TEACHERS.length) errors.push('Duplicate teacher names in registry');
  schedule.forEach(r => {
    if (!classSet.has(r.Class)) errors.push(`Unknown class ${r.Class}`);
    teachersOf(r).forEach(t => { if (!teacherSet.has(t)) errors.push(`Unregistered teacher "${t}" in ${r.Class} ${r.Day} P${r.Period}`); });
  });
  CLASSES.forEach(([c]) => {
    const n = schedule.filter(r => r.Class === c).length;
    if (n === 0) errors.push(`${c}: no source gives a timetable for this class — add a source for it, or remove it from CLASSES in registry.js`);
    else if (n !== DAYS.length * PERIODS.length) errors.push(`${c}: expected ${DAYS.length * PERIODS.length} rows (6 days × 8 periods), got ${n}`);
  });
  const keys = new Set();
  schedule.forEach(r => { const k = `${r.Class}|${r.Day}|${r.Period}`; if (keys.has(k)) errors.push('Duplicate slot ' + k); keys.add(k); });

  const report = analyze(schedule);
  return { schedule, errors: [...new Set(errors)], warnings, parts, report, label, replaced, tw, fills, diff, defaults: defaults.applied, missing };
}

/** One line per teacher / period / set of classes where the two sheets disagree, days merged. */
function disagreements(diff) {
  const slots = new Map(); // teacher|day|period → {cls: Set, tw: Set, cells: Set}
  const at = (t, d, p) => { const k = `${t}|${d}|${p}`; if (!slots.has(k)) slots.set(k, { teacher: t, day: d, period: p, cls: new Set(), tw: new Set(), cells: new Set(), subjects: new Set() }); return slots.get(k); };
  diff.classOnly.forEach(x => { const s = at(x.teacher, x.day, x.period); s.cls.add(x.cls); s.subjects.add(x.subject); });
  diff.teacherOnly.forEach(x => { const s = at(x.teacher, x.day, x.period); s.tw.add(x.cls); s.cells.add(x.cell.replace(/^.*!/, '')); });
  const list = [...slots.values()].map(s => ({ ...s, a: [...s.cls].join(' + ') || '—', b: [...s.tw].join(' + ') || '—', subj: [...s.subjects].join(', '), cellText: [...s.cells].join(', ') }));
  return collapseDays(list, s => `${s.teacher}|${s.period}|${s.a}|${s.b}|${s.subj}`)
    .sort((x, y) => x.teacher.localeCompare(y.teacher) || x.period - y.period || DAYS.indexOf(x.days[0]) - DAYS.indexOf(y.days[0]));
}

function writeOutputs(built) {
  const { schedule, parts, report, label, replaced, tw, fills, diff, defaults, missing, warnings } = built;
  const OUT_MD = path.join(ROOT, 'docs', `TIMETABLE_REVIEW_${label}.md`);
  const subjects = [...new Set(schedule.map(r => r.Subject).filter(Boolean))].sort((a, b) => a.localeCompare(b));
  const blanks = parts.flatMap(p => p.blanks);
  const ignored = parts.flatMap(p => p.ignored);
  const sourceLines = parts.map(p => `${p.name} → ${[...new Set(p.rows.map(r => r.Class))].length} classes`);
  const used = new Set(schedule.flatMap(teachersOf));
  const bell = PERIODS;

  // ── PrefillData.gs ────────────────────────────────────────────────────────
  const j = v => JSON.stringify(v);
  const gs = [
    '/**',
    ' * PrefillData.gs — GENERATED by tools/convert/build.js. Do not edit by hand.',
    ...parts.map((p, i) => (i ? ' *          ' : ' * Sources: ') + p.name),
    ...(tw ? [' *          ' + tw.name + ' — fills teachers the class sheets leave out'] : []),
    ` * Classes: ${CLASSES.length} | Teachers: ${TEACHERS.length} | Subjects: ${subjects.length} | Schedule rows: ${schedule.length}`,
    ` * Review the conversion in docs/${path.basename(OUT_MD)} before importing.`,
    ' */',
    '',
    '// [Class Name, Academic Tier, Room Assigned]',
    'const RAW_CLASS_ROWS = [\n' + CLASSES.map(c => '  ' + j(c)).join(',\n') + '\n];',
    '',
    '// [Teacher Name, Subject Specialization, Max Hours / Week, Days Unavailable, Total Hours Scheduled]',
    'const RAW_TEACHER_ROWS = [\n' + TEACHERS.map(t => '  ' + j([...t, ''])).join(',\n') + '\n];',
    '',
    '// [Subject Name] — every distinct Subject cell value used in RAW_SCHEDULE',
    'const RAW_SUBJECT_ROWS = [\n' + subjects.map(s => '  ' + j([s])).join(',\n') + '\n];',
    '',
    '// [Period, Start, End] — bell schedule (24-hour) shown in the view headers',
    'const RAW_PERIOD_ROWS = [\n' + bell.map(p => '  ' + j(p)).join(',\n') + '\n];',
    '',
    '// One row per class / day / period. Split periods: "Hindi / IT" ↔ "Teacher A / Teacher B"',
    '// (groups line up; "+" joins co-teachers of one group; "(H.M)" / "(E.M)" = medium). See spreadsheet_schema.md.',
    'const RAW_SCHEDULE = [\n' + schedule.map(r => '  ' + j({ Day: r.Day, Period: r.Period, Class: r.Class, Tier: r.Tier, Subject: r.Subject, Teacher: r.Teacher, Room: r.Room })).join(',\n') + '\n];',
    '',
  ].join('\n');
  fs.writeFileSync(OUT_GS, gs);

  // ── Review report ─────────────────────────────────────────────────────────
  const md = [];
  const slotText = x => `${x.cls} | ${x.day.slice(0, 3)} | P${x.period}`;
  const combined = collapseDays(report.combined, e => `${e.teacher}|${e.period}|${e.subject}|${e.classes.join('+')}`)
    .sort((a, b) => a.period - b.period || a.classes[0].localeCompare(b.classes[0]));
  const splitRows = collapseDays(report.splits.map(s => ({ ...s, text: s.lanes.map(l => `${l.subject} — ${l.teachers.join(' + ') || '(no teacher)'}`).join('; ') })),
    s => `${s.cls}|${s.period}|${s.text}`);
  const dis = tw ? disagreements(diff) : [];
  md.push(`# Timetable ${label} — conversion review`, '');
  md.push('Generated by `tools/convert/build.js` (`npm run convert`). Check this against the Excel before running **📥 Import Timetable** in the sheet.', '');
  md.push('## Summary', '');
  md.push('| Item | Value |', '|---|---|');
  sourceLines.forEach((l, i) => md.push(`| Source ${i + 1} | ${l} |`));
  if (tw) md.push(`| Cross-check | ${tw.name}: ${tw.teachers.length} teachers, ${tw.claims.length} teacher / day / period entries |`);
  md.push(`| Classes | ${CLASSES.length} (${CLASSES.filter(c => c[1] === 'Middle').length} middle, ${CLASSES.filter(c => c[1] !== 'Middle').length} secondary / higher secondary) |`);
  const lunch = [...new Set(parts.map(p => p.lunch).filter(Boolean))];
  md.push(`| Bell schedule | ${bell.map(([n, s, e]) => `P${n} ${s}–${e}`).join(' · ')}${lunch.length ? ` · lunch ${lunch.join(', ')}` : ''} |`);
  md.push(`| Teachers | ${TEACHERS.length} in registry, ${used.size} with periods |`);
  md.push(`| Schedule rows | ${schedule.length} |`);
  md.push(`| Split periods (class divided into groups) | ${report.splits.length} |`);
  md.push(`| Combined lectures (one teacher, several sections) | ${report.combined.length} |`);
  md.push(`| Teachers filled in from the teacher-wise sheet | ${fills.filled.length} groups |`);
  md.push(`| Groups still without a teacher | ${missing.length} |`);
  md.push(`| Blank slots | ${blanks.length} |`);
  md.push(`| Teacher clashes | ${report.clashes.length} (teacher / day / period) |`);
  md.push(`| Room clashes | ${report.roomClashes.length} |`);
  if (tw) md.push(`| Class-wise vs teacher-wise disagreements | ${dis.length} (teacher / period lines) |`);
  md.push('');

  md.push('## How split and combined lectures are stored', '');
  md.push('- **Split period** — one class, several groups at once (e.g. Sanskrit / IT / Beauty electives, or Hindi-medium and English-medium groups). One Master_Schedule row; the Subject, Teacher and Room cells list the groups in the same order, separated by ` / `. Co-teachers of one group are joined with ` + `. A group with no teacher is `-`.');
  md.push('- **Medium** — `(H.M)` / `(E.M)` after a subject marks the Hindi- or English-medium group, as in the Excel (e.g. `Physics (H.M) / Chemistry (E.M)`). 11th C and 12th C are one class each; their medium groups are split periods.');
  md.push('- **Combined lecture** — one teacher teaches the same subject to several sections at once (e.g. 12th B + 12th D Physics). Each section keeps its own row with the same subject; the clash check treats same teacher + same subject + same period + same grade as one lecture. A medium mark on only one of the sections (`Hindi` in 11th B, `Hindi (E.M)` in 11th C) is still the same lecture; `(H.M)` and `(E.M)` groups at once are a clash.');
  md.push('- Same teacher + same period + **different** subjects is a real clash.', '');

  md.push('## How the cells were read', '');
  md.push('- Every subject word starts a group; a teacher belongs to the subject before it: `SANSKRIT NEHA.J IT-DIVYA BEAUTY-KRISHNA` = Sanskrit (Neha Jamidar) / IT (Divya Lodval) / Beauty (Krishna Patale).');
  md.push('- `SO.SCI RUBAL VERMA(H.M)/ I.MATHUR (E.M)` = Social Science (H.M) with Rubal Verma / Social Science (E.M) with Indira Mathur.');
  md.push('- `CCLE INDRA.M/ KRISHNA.P` and `SST Lalita / Farhana` = one group taught by both teachers.');
  md.push('- `MATHS /CCLE NEHA JHALA` = one group "Mathematics & CCLE" taken by the class teacher.');
  md.push('- `DAKSHATA X` = "Proficiency X". `COMPUTER` with no name = Reena Shivhare (the only computer teacher) when the teacher-wise sheet does not say otherwise.');
  md.push('- `A.KARTARE` is read as Anuradha Katare and `DANCE RUPALI` as Rupali Chuliwal (the dance teacher).');
  md.push('- Every distinct cell and how it was read is listed at the end of this file.', '');

  md.push('## Needs confirmation', '');
  const notes = [...new Set([...parts.flatMap(p => p.notes).filter(n => !/^Lunch /.test(n)), ...(tw ? tw.notes : [])])];
  notes.forEach(n => md.push('- ' + n));
  replaced.forEach(n => md.push('- ' + n));
  warnings.forEach(n => md.push('- ' + n));
  if (defaults.length) md.push(`- ${defaults.length} Computer period(s) name no teacher and the teacher-wise sheet does not list them; they are given to **Reena Shivhare**: ${collapseDays(defaults, d => `${d.cls}|${d.period}`).map(d => `${d.cls} ${d.dayText} P${d.period}`).join(', ')}.`);
  md.push('');

  if (fills.filled.length) {
    md.push('## Teachers filled in from the teacher-wise sheet', '');
    md.push('The class-wise cell names no teacher for these groups; the teacher-wise sheet puts exactly one fitting teacher in that class at that time.', '');
    md.push('| Class | Period | Days | Subject | Teacher | Teacher-wise cell |', '|---|---|---|---|---|---|');
    collapseDays(fills.filled.map(f => ({ ...f, cell: f.evidence.replace(/^.*?!/, '') })), f => `${f.cls}|${f.period}|${f.subject}|${f.teacher}|${f.cell}`)
      .sort((a, b) => classOrder.get(a.cls) - classOrder.get(b.cls) || a.period - b.period || DAYS.indexOf(a.days[0]) - DAYS.indexOf(b.days[0]))
      .forEach(f => md.push(`| ${f.cls} | ${f.period} | ${f.dayText} | ${f.subject} | ${f.teacher} | ${f.cell} |`));
    md.push('');
  }

  md.push('## Groups still without a teacher', '');
  if (missing.length) {
    md.push('Neither sheet names a teacher. "Teaches this subject to the class" is a suggestion only — nothing was filled from it.', '');
    md.push('| Class | Days | Period | Subject | Teaches this subject to the class | Teacher-wise sheet has here |', '|---|---|---|---|---|---|');
    collapseDays(missing, m => `${m.cls}|${m.period}|${m.subject}|${m.suggestion}|${m.candidates.join(';')}`)
      .sort((a, b) => classOrder.get(a.cls) - classOrder.get(b.cls) || a.period - b.period)
      .forEach(m => md.push(`| ${m.cls} | ${m.dayText} | ${m.period} | ${m.subject} | ${m.suggestion || '—'} | ${m.candidates.join(', ') || '—'} |`));
  } else md.push('None.');
  md.push('');

  if (tw) {
    md.push('## Class-wise sheet vs teacher-wise sheet', '');
    md.push('Where the two sheets put a teacher in different classes. The class-wise sheets are what gets imported; use this list to find mistakes in either sheet.', '');
    if (dis.length) {
      md.push('| Teacher | Period | Days | Class-wise sheet | Teacher-wise sheet | Teacher-wise cell |', '|---|---|---|---|---|---|');
      dis.forEach(d => md.push(`| ${d.teacher} | ${d.period} | ${d.dayText} | ${d.a}${d.subj ? ' (' + d.subj + ')' : ''} | ${d.b} | ${d.cellText || '—'} |`));
    } else md.push('The two sheets agree.');
    md.push('');
    if (tw.classTeachers.size) {
      md.push('## Class teachers (teacher-wise sheet)', '');
      md.push('| Class | Class teacher |', '|---|---|');
      CLASSES.forEach(([c]) => md.push(`| ${c} | ${(tw.classTeachers.get(c) || ['—']).join(', ')} |`));
      md.push('');
    }
  }

  md.push('## Teacher load (periods per week)', '');
  md.push('A combined lecture counts once. 48 = every period of the week. **Check For Conflicts** reports teachers above `Max Hours / Week` (40 in the registry).', '');
  const load = new Map();
  schedule.forEach(r => teachersOf(r).forEach(t => (load.get(t) || load.set(t, new Set()).get(t)).add(`${r.Day}|${r.Period}`)));
  const loads = [...load.entries()].map(([t, set]) => [t, set.size]).sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]));
  const loadCell = x => (x ? `${x[0]} | ${x[1] > 40 ? '**' + x[1] + '**' : x[1]}` : ' | ');
  md.push('| Teacher | Periods | Teacher | Periods | Teacher | Periods |', '|---|---|---|---|---|---|');
  for (let i = 0; i < loads.length; i += 3) md.push(`| ${[loads[i], loads[i + 1], loads[i + 2]].map(loadCell).join(' | ')} |`);
  md.push('');

  md.push('## Teacher clashes', '');
  md.push('Same teacher, same day and period, with different subjects, different grades or both media at once. These come from the Excel; fix them there (or in Master_Schedule — **Check For Conflicts** marks them).', '');
  if (report.clashes.length) {
    md.push('| Teacher | Period | Days | Classes | Why | Excel cells |', '|---|---|---|---|---|---|');
    collapseDays(report.clashes, c => `${c.teacher}|${c.period}|${c.detail}|${c.reason}`).forEach(c => {
      const first = report.clashes.find(x => x.teacher === c.teacher && x.period === c.period && x.detail === c.detail && x.day === c.days[0]);
      const more = c.days.length > 1 ? ` (${c.days[0].slice(0, 3)}; the other days are in the rows below)` : '';
      md.push(`| ${c.teacher} | ${c.period} | ${c.dayText} | ${c.detail} | ${c.reason} | ${first.cells}${more} |`);
    });
  } else md.push('None.');
  md.push('');
  if (report.roomClashes.length) {
    md.push('## Room clashes', '', '| Room | Day | Period | Classes |', '|---|---|---|---|');
    report.roomClashes.forEach(c => md.push(`| ${c.room} | ${c.day} | ${c.period} | ${c.detail} |`));
    md.push('');
  }

  md.push('## Split periods', '');
  md.push('| Class | Period | Days | Groups (subject — teacher) |', '|---|---|---|---|');
  splitRows.sort((a, b) => classOrder.get(a.cls) - classOrder.get(b.cls) || a.period - b.period)
    .forEach(s => md.push(`| ${s.cls} | ${s.period} | ${s.dayText} | ${s.text} |`));
  md.push('');

  md.push('## Combined lectures (not clashes)', '');
  md.push('| Period | Days | Teacher | Subject | Sections together |', '|---|---|---|---|---|');
  combined.forEach(c => md.push(`| ${c.period} | ${c.dayText} | ${c.teacher} | ${c.subject} | ${c.classes.join(' + ')} |`));
  md.push('');

  md.push('## Blank slots (free periods)', '');
  if (blanks.length) {
    const blankGroups = {};
    blanks.forEach(b => { const m = b.match(/^(.*?) (Monday|Tuesday|Wednesday|Thursday|Friday|Saturday) (P\d)$/); (blankGroups[m[1]] = blankGroups[m[1]] || []).push(`${m[2].slice(0, 3)} ${m[3]}`); });
    md.push('| Class | Slots |', '|---|---|');
    Object.entries(blankGroups).forEach(([c, list]) => md.push(`| ${c} | ${list.join(', ')} |`));
  } else md.push('None — every period of every class has an entry.');
  md.push('');

  const cellRows = parts.flatMap(p => [...p.cells.values()].map(c => ({ ...c, sheet: p.src.sheet || path.basename(p.src.file) })));
  if (cellRows.length) {
    md.push('## Every distinct cell and how it was read', '');
    md.push('<details><summary>' + cellRows.length + ' distinct cells</summary>', '');
    md.push('| Sheet | Cell text | Times | First at | Subject | Teacher |', '|---|---|---|---|---|---|');
    cellRows.forEach(c => md.push(`| ${c.sheet} | ${c.text.replace(/\|/g, '\\|')} | ${c.count} | ${c.at.join(', ')} | ${c.subject} | ${c.teacher || '—'} |`));
    md.push('', '</details>', '');
  }

  md.push('## Ignored cells (outside the printed tables)', '');
  md.push('<details><summary>' + ignored.length + ' cells</summary>', '');
  ignored.forEach(i => md.push('- ' + i.replace(/\n/g, ' ')));
  md.push('', '</details>', '');

  md.push('## Class timetables', '');
  CLASSES.forEach(([c]) => {
    md.push(`### ${c}`, '');
    md.push('| Day | ' + bell.map(([n, s]) => `P${n} ${s}`).join(' | ') + ' |', '|---|---|---|---|---|---|---|---|---|');
    DAYS.forEach(d => {
      const cells = [1, 2, 3, 4, 5, 6, 7, 8].map(p => {
        const r = schedule.find(x => x.Class === c && x.Day === d && x.Period === p);
        return r && r.Subject ? `${r.Subject}${r.Teacher ? '<br>_' + r.Teacher + '_' : ''}` : '—';
      });
      md.push(`| ${d.slice(0, 3)} | ${cells.join(' | ')} |`);
    });
    md.push('');
  });
  fs.mkdirSync(path.dirname(OUT_MD), { recursive: true });
  fs.writeFileSync(OUT_MD, md.join('\n'));
  console.log(`  wrote ${path.relative(ROOT, OUT_GS)} and ${path.relative(ROOT, OUT_MD)}`);
}

/** Parses the command line; unknown flags are an error so typos do not pass silently. */
function parseArgs(argv) {
  const opts = { add: [] };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    const val = () => { if (i + 1 >= argv.length) throw new Error(`${a} needs a value`); return argv[++i]; };
    if (a === '--check') opts.check = true;
    else if (a === '--sources') opts.sources = val();
    else if (a === '--add') opts.add.push(val());
    else if (a === '--export-flat') opts.exportFlat = val();
    else if (a === '--help' || a === '-h') opts.help = true;
    else throw new Error(`Unknown option ${a} (see --help)`);
  }
  return opts;
}

function main() {
  let opts;
  try { opts = parseArgs(process.argv.slice(2)); } catch (e) { console.error('✗ ' + e.message); process.exit(2); }
  if (opts.help) { console.log(fs.readFileSync(__filename, 'utf8').split('*/')[0].replace(/^#!.*\n\/\*\*\n?/, '').replace(/^ \* ?/gm, '')); return; }
  let built;
  try { built = buildAll(opts); } catch (e) { console.error('✗ ' + e.message); process.exit(1); }
  if (built.errors.length) {
    console.error('✗ Conversion errors:\n  ' + built.errors.join('\n  '));
    process.exit(1);
  }
  const { schedule, parts, report, fills, missing, diff, tw } = built;
  const subjects = new Set(schedule.map(r => r.Subject).filter(Boolean));
  const used = new Set(schedule.flatMap(teachersOf));
  const blanks = parts.reduce((n, p) => n + p.blanks.length, 0);
  console.log(`✓ ${CLASSES.length} classes, ${TEACHERS.length} teachers (${used.size} with periods), ${subjects.size} subjects, ${schedule.length} schedule rows`);
  console.log(`  ${report.splits.length} split periods, ${report.combined.length} combined lectures, ${blanks} blank slots, ${report.clashes.length} teacher clashes, ${report.roomClashes.length} room clashes`);
  if (tw) console.log(`  teacher-wise sheet: ${fills.filled.length} teachers filled in, ${missing.length} groups still without a teacher, ${disagreements(diff).length} disagreement lines`);
  built.warnings.forEach(w => console.log('  warning: ' + w));
  built.replaced.forEach(r => console.log('  note: ' + r));
  if (opts.exportFlat) { fs.writeFileSync(opts.exportFlat, toFlatCsv(schedule)); console.log(`  wrote flat CSV ${opts.exportFlat}`); }
  if (opts.check) return;
  writeOutputs(built);
}

module.exports = { buildAll, parseArgs, disagreements, dayList };
if (require.main === module) main();
