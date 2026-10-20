/**
 * flat.js — reads a timetable written as one row per class / day / period.
 *
 * This is the format to use for any timetable whose layout the other readers do
 * not know: a CSV or an Excel sheet with a header row and these columns
 * (header names are case-insensitive, column order does not matter):
 *
 *   Class   | Day     | Period | Subject              | Teacher                         | Room
 *   9th A   | Monday  | 1      | Sanskrit             | Neha Jamidar                    |
 *   11th A  | Mon     | 1      | Hindi / IT / Beauty  | Ekta Gour / Divya Lodval / Krishna Patale | 12 / Lab / 14
 *
 *  - Day: English name or 3-letter short form, or the Hindi name (सोमवार …).
 *  - Period: 1–8.
 *  - Split period: subjects, teachers and rooms are listed in the same order,
 *    separated by "/". Co-teachers of one group are joined with "+". A group with
 *    no teacher is "-".
 *  - Combined lecture: give every section its own row with the same subject and teacher.
 *  - Teacher names must match the registry (registry.js); case, extra spaces and the
 *    short forms in TEACHER_ALIASES (lib/vocabulary.js) are accepted and written back in full.
 *  - A blank Subject is a free period. Missing slots are filled as free periods
 *    and reported.
 */
'use strict';

const fs = require('fs');
const XLSX = require('xlsx');
const { DAYS, TEACHERS, CLASSES } = require('./registry');
const { TEACHER_ALIASES } = require('./lib/vocabulary');

const REQUIRED = ['class', 'day', 'period', 'subject', 'teacher'];
const HINDI_DAYS = { 'सोमवार': 'Monday', 'मंगलवार': 'Tuesday', 'बुधवार': 'Wednesday', 'गुरुवार': 'Thursday', 'गुरूवार': 'Thursday', 'शुक्रवार': 'Friday', 'शनिवार': 'Saturday' };
const clean = s => String(s == null ? '' : s).replace(/\s+/g, ' ').trim();
const key = s => clean(s).toLowerCase().replace(/\.$/, '');

/** "mon", "Monday", "MON." → "Monday"; '' when unknown. */
function dayOf(raw) {
  const v = clean(raw);
  if (HINDI_DAYS[v]) return HINDI_DAYS[v];
  const k = v.toLowerCase().replace(/\.$/, '');
  return DAYS.find(d => d.toLowerCase() === k || (k.length >= 3 && d.toLowerCase().startsWith(k))) || '';
}

/** Lookup from any accepted spelling to the registry name. */
function teacherIndex() {
  const idx = new Map();
  TEACHERS.forEach(([n]) => idx.set(key(n), n));
  Object.entries(TEACHER_ALIASES).forEach(([a, n]) => idx.set(key(a), n));
  return idx;
}

/**
 * Rewrites every name in a Teacher cell to its registry spelling, keeping the
 * "/" (group) and "+" (co-teacher) structure. Unknown names are returned in `unknown`.
 */
function resolveTeachers(cell, idx) {
  const unknown = [];
  const text = clean(cell);
  if (!text) return { text: '', unknown };
  const out = text.split(/(\s*[\/,|]\s*|\s*\+\s*|\s+&\s+)/).map((part, i) => {
    if (i % 2 === 1) return /\+|&/.test(part) ? ' + ' : ' / ';
    const t = clean(part);
    if (!t || ['-', '—', '–', '?'].includes(t)) return t;
    const hit = idx.get(key(t));
    if (!hit) unknown.push(t);
    return hit || t;
  }).join('');
  return { text: out, unknown };
}

/**
 * @param {string} file .csv / .xlsx / .xls
 * @param {string=} sheetName first sheet when omitted
 * @returns {{rows: Object[], issues: string[], blanks: string[], ignored: string[], classes: string[]}}
 */
function parseFlat(file, sheetName) {
  // CSV must be read as UTF-8 text: XLSX.readFile treats it as Latin-1 and garbles
  // Hindi names and dashes. A leading BOM (Excel's "CSV UTF-8") is dropped.
  const wb = /\.(csv|tsv|txt)$/i.test(file)
    ? XLSX.read(fs.readFileSync(file, 'utf8').replace(/^\uFEFF/, ''), { type: 'string', raw: true })
    : XLSX.readFile(file, { raw: false });
  const sheet = sheetName || wb.SheetNames[0];
  const ws = wb.Sheets[sheet];
  if (!ws) throw new Error(`Sheet "${sheet}" not found in ${file} (sheets: ${wb.SheetNames.join(', ')})`);
  const grid = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false });

  const rows = [], issues = [], blanks = [], ignored = [];
  const headerRow = grid.findIndex(r => r.some(c => key(c) === 'class') && r.some(c => key(c) === 'day'));
  if (headerRow < 0) throw new Error(`${file}: no header row with "Class" and "Day" columns`);
  const header = grid[headerRow].map(key);
  const col = name => header.indexOf(name);
  const missing = REQUIRED.filter(h => col(h) < 0);
  if (missing.length) throw new Error(`${file}: missing column(s) ${missing.join(', ')}`);

  const classTier = new Map(CLASSES.map(c => [key(c[0]), c]));
  const idx = teacherIndex();
  const seen = new Map();

  for (let i = headerRow + 1; i < grid.length; i++) {
    const r = grid[i];
    const line = i + 1;
    if (r.every(c => !clean(c))) continue;
    const get = name => (col(name) >= 0 ? clean(r[col(name)]) : '');
    const cls = classTier.get(key(get('class')));
    const day = dayOf(get('day'));
    const period = Number(get('period'));
    if (!cls) { issues.push(`line ${line}: unknown class "${get('class')}" — add it to CLASSES in registry.js`); continue; }
    if (!day) { issues.push(`line ${line}: unknown day "${get('day')}"`); continue; }
    if (!Number.isInteger(period) || period < 1 || period > 8) { issues.push(`line ${line}: period "${get('period')}" must be 1–8`); continue; }
    const slot = `${cls[0]}|${day}|${period}`;
    if (seen.has(slot)) { issues.push(`line ${line}: ${cls[0]} ${day} P${period} already given on line ${seen.get(slot)}`); continue; }
    seen.set(slot, line);
    const teacher = resolveTeachers(get('teacher'), idx);
    teacher.unknown.forEach(t => issues.push(`line ${line}: unknown teacher "${t}" — add to TEACHERS (or TEACHER_ALIASES) in the converter`));
    const subject = get('subject');
    if (!subject && teacher.text && !/^[-—–?\s/]*$/.test(teacher.text)) issues.push(`line ${line}: teacher "${teacher.text}" given without a subject`);
    rows.push({ Day: day, Period: period, Class: cls[0], Tier: cls[1], Subject: subject, Teacher: subject ? teacher.text : '', Room: get('room') || cls[2], Source: `${sheet}!${line}` });
  }

  // Fill holes so each class keeps 48 slots; report them as blanks.
  const classes = [...new Set(rows.map(r => r.Class))];
  classes.forEach(c => {
    const cls = classTier.get(key(c));
    DAYS.forEach(d => [1, 2, 3, 4, 5, 6, 7, 8].forEach(p => {
      if (!seen.has(`${c}|${d}|${p}`)) rows.push({ Day: d, Period: p, Class: c, Tier: cls[1], Subject: '', Teacher: '', Room: cls[2], Source: `${sheet} (missing)` });
    }));
  });
  rows.forEach(r => { if (!r.Subject) blanks.push(`${r.Class} ${r.Day} P${r.Period}`); });
  return { rows, issues, blanks, ignored, classes };
}

/** Writes canonical rows as a flat CSV (same columns parseFlat reads). */
function toFlatCsv(rows) {
  const q = v => { const s = String(v == null ? '' : v); return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; };
  const order = new Map(CLASSES.map((c, i) => [c[0], i]));
  const sorted = [...rows].sort((a, b) => order.get(a.Class) - order.get(b.Class) || DAYS.indexOf(a.Day) - DAYS.indexOf(b.Day) || a.Period - b.Period);
  return ['Class,Day,Period,Subject,Teacher,Room', ...sorted.map(r => [r.Class, r.Day, r.Period, r.Subject, r.Teacher, r.Room].map(q).join(','))].join('\n') + '\n';
}

module.exports = { parseFlat, toFlatCsv, dayOf, resolveTeachers, teacherIndex };
