/**
 * class-blocks.js — reads the school's class-wise timetable sheets ("6-8" and
 * "9-12" in Timetable.xlsx; the older "CLASS WISE new" sheet has the same layout).
 *
 * Layout of one class block:
 *
 *   row  | B          | C        | D  E  F  G      | H (lunch)  | I  J  K  L
 *   head |            |          | 1  2  3  4      |            | 5  6  7  8      ← period numbers
 *   head |            |          | 09:30-10:20 …   | 12:20-12:50| 12:50- 01:30 …  ← bell times
 *   Mon  | कक्षा 9 A  | सोमवार   | one cell per period          ← "SUBJECT TEACHER", see cell-grammar.js
 *   …    |            | मंगलवार … शनिवार
 *
 * The period columns and bell times are read from the header rows, not assumed.
 * Anything outside the printed table (other columns, stray text) is listed as ignored.
 */
'use strict';

const XLSX = require('xlsx');
const { DAYS, CLASSES } = require('./registry');
const { parseCell } = require('./lib/cell-grammar');
const { loadGrammar } = require('./lib/gas');

const HINDI_DAYS = { 'सोमवार': 'Monday', 'मंगलवार': 'Tuesday', 'बुधवार': 'Wednesday', 'गुरूवार': 'Thursday', 'गुरुवार': 'Thursday', 'शुक्रवार': 'Friday', 'शनिवार': 'Saturday' };
const clean = s => String(s == null ? '' : s).replace(/\s+/g, ' ').trim();
const CLASS_LABEL = /^कक्षा\s*(\d{1,2})\s*([A-Z])$/;
const TITLE = /सांदिपनी|समय सारणी|timetable|^स\.?\s*क्र\.?$|^कक्षा$|^वार$/i;

/** "12:50- 01:30" → ['12:50', '13:30'] (afternoon hours written 01–06 are read as 13–18). */
function readTimes(text) {
  const m = clean(text).match(/^(\d{1,2})[:.](\d{2})\s*[-–]\s*(\d{1,2})[:.](\d{2})$/);
  if (!m) return null;
  const hhmm = (h, mm, after) => {
    let hour = Number(h);
    if (hour < 7 || (after && hour < after)) hour += 12;
    return `${String(hour).padStart(2, '0')}:${mm}`;
  };
  const start = hhmm(m[1], m[2]);
  return [start, hhmm(m[3], m[4], Number(start.slice(0, 2)))];
}

/**
 * @param {string} file
 * @param {string} sheetName
 * @returns {{rows: Object[], issues: string[], blanks: string[], ignored: string[], notes: string[],
 *            bell: Array<[number, string, string]>, lunch: string, cells: Map<string, Object>}}
 */
function parseClassBlocks(file, sheetName) {
  const wb = XLSX.readFile(file);
  const ws = wb.Sheets[sheetName];
  if (!ws) throw new Error(`Sheet "${sheetName}" not found in ${file} (sheets: ${wb.SheetNames.join(', ')})`);
  const grid = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', blankrows: true, raw: false });
  const P = loadGrammar().ScheduleParser;
  const registry = new Map(CLASSES.map(c => [c[0], c]));
  const addr = (r, c) => XLSX.utils.encode_cell({ r, c });

  const rows = [], issues = [], blanks = [], ignored = [], notes = [];
  const cells = new Map(); // distinct cell text → how it was read
  const seenClasses = new Set();
  let periodCols = null, bell = null, lunch = '', lunchCol = -1;
  let current = null; // {name, tier, room, dayIdx}

  grid.forEach((r, i) => {
    // Header row with the period numbers 1..8 → period columns; the next row has the bell times.
    const nums = r.map(v => clean(v)).map((v, c) => (/^[1-8]$/.test(v) ? [Number(v), c] : null)).filter(Boolean);
    if (nums.length === 8 && nums.every((n, k) => n[0] === k + 1)) {
      const cols = nums.map(n => n[1]);
      if (periodCols && cols.join() !== periodCols.join()) issues.push(`row ${i + 1}: period columns differ from the first header`);
      periodCols = cols;
      lunchCol = cols[3] + 1 < cols[4] ? cols[3] + 1 : -1;
      const times = grid[i + 1] ? cols.map(c => readTimes(grid[i + 1][c])) : [];
      if (times.length === 8 && times.every(Boolean)) {
        const found = times.map((t, k) => [k + 1, t[0], t[1]]);
        if (bell && JSON.stringify(bell) !== JSON.stringify(found)) issues.push(`row ${i + 2}: bell times differ from the first header`);
        bell = bell || found;
        if (lunchCol >= 0) { const lt = readTimes(grid[i + 1][lunchCol]); if (lt) lunch = lt.join('–'); }
      }
      current = null;
      return;
    }

    const label = clean(r[1]);
    const m = label.match(CLASS_LABEL);
    if (m) {
      const name = `${m[1]}th ${m[2]}`;
      const reg = registry.get(name);
      if (!reg) issues.push(`${addr(i, 1)}: class "${label}" → "${name}" is not in CLASSES in registry.js`);
      if (seenClasses.has(name)) issues.push(`${addr(i, 1)}: ${name} appears twice`);
      seenClasses.add(name);
      current = { name, tier: reg ? reg[1] : '', room: reg ? reg[2] : '', dayIdx: 0 };
    }

    const dayLabel = clean(r[2]);
    const day = HINDI_DAYS[dayLabel];
    if (!current || !day || !periodCols) {
      r.forEach((v, c) => {
        const text = clean(v);
        if (text && !TITLE.test(text) && !(c === 1 && m) && !readTimes(text) && !(c === lunchCol && /^[A-Z]$/.test(text))) ignored.push(`${addr(i, c)}: "${text}"`);
      });
      return;
    }
    if (day !== DAYS[current.dayIdx]) issues.push(`${addr(i, 2)}: expected ${DAYS[current.dayIdx] || 'no more days'} for ${current.name}, found ${day}`);
    current.dayIdx++;

    periodCols.forEach((c, p) => {
      const where = `${current.name} ${day} P${p + 1} (${sheetName}!${addr(i, c)})`;
      const raw = String(r[c] == null ? '' : r[c]);
      const { lanes, problems } = parseCell(raw);
      problems.forEach(x => issues.push(`${where}: ${x} in "${clean(raw)}"`));
      const f = P.formatLanes(lanes.map(l => ({ subject: l.subject, teachers: l.teachers, room: '' })));
      const key = clean(raw);
      if (key) {
        if (!cells.has(key)) cells.set(key, { text: key, subject: f.subject, teacher: f.teacher, count: 0, at: [] });
        const e = cells.get(key); e.count++; if (e.at.length < 3) e.at.push(addr(i, c));
      }
      if (!f.subject) blanks.push(`${current.name} ${day} P${p + 1}`);
      rows.push({ Day: day, Period: p + 1, Class: current.name, Tier: current.tier, Subject: f.subject, Teacher: f.subject ? f.teacher : '', Room: current.room, Source: `${sheetName}!${addr(i, c)}` });
    });
    r.forEach((v, c) => {
      const text = clean(v);
      if (!text || c <= 2 || periodCols.includes(c)) return;
      if (c === lunchCol && /^[A-Z]$/.test(text)) return; // the "L U N C H" letters
      ignored.push(`${addr(i, c)} (${current.name} ${day}): "${text}"`);
    });
    if (current.dayIdx === DAYS.length) current = null;
  });

  if (!periodCols) issues.push(`${sheetName}: no header row with the period numbers 1–8`);
  seenClasses.forEach(c => {
    const n = rows.filter(x => x.Class === c).length;
    if (n !== DAYS.length * 8) issues.push(`${c}: found ${n / 8} day rows, expected 6 (Monday–Saturday)`);
  });
  if (lunch) notes.push(`Lunch ${lunch} (between periods 4 and 5).`);
  return { rows, issues, blanks, ignored, notes, bell: bell || [], lunch, cells };
}

module.exports = { parseClassBlocks, readTimes, HINDI_DAYS };
