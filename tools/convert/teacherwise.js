/**
 * teacherwise.js — reads the teacher-wise sheet ("Teacherwise" in Timetable.xlsx)
 * and uses it to (1) fill in teachers the class-wise sheets leave out and
 * (2) cross-check the two sheets against each other.
 *
 * Layout: one row per teacher. "Employee Name (Code)" column, "CLASS TEACHER" column,
 * "SUBJECT" column, then one column per period (numbers 1–8 in the header row).
 * A period cell lists classes, optionally with days, medium and subject hints:
 *
 *   "10TH C"                         every day, 10th C
 *   "11TH A SAT 9TH A"               11th A on the days not named (Mon–Fri), 9th A on Saturday
 *   "MON-TUE-10TH A WED-THU-10TH B+C" Mon/Tue 10th A; Wed/Thu 10th B and 10th C together
 *   "12TH C (ACC) MON TO THU"        days may also follow the class
 *   "11TH C B.STU(E.M)"              subject and medium hints
 *
 * A class list without days applies to every day no other part of the cell names.
 */
'use strict';

const XLSX = require('xlsx');
const { DAYS, CLASSES, TEACHERS } = require('./registry');
const { teacherName, splitMedium } = require('./lib/cell-grammar');
const { HINT_SUBJECTS } = require('./lib/vocabulary');
const { loadGrammar } = require('./lib/gas');

const clean = s => String(s == null ? '' : s).replace(/\s+/g, ' ').trim();
const DAY_WORDS = { MON: 0, TUE: 1, TUES: 1, WED: 2, THU: 3, THUR: 3, THURS: 3, FRI: 4, SAT: 5 };
const HINTS = Object.fromEntries(Object.entries(HINT_SUBJECTS).map(([k, v]) => [k.toUpperCase(), v]));

/**
 * Parses one period cell into claims.
 * @returns {{claims: Array<{day: string, cls: string, subject: string, medium: string}>, problems: string[]}}
 */
function parseClaimCell(text) {
  const t = clean(text).toUpperCase();
  const problems = [];
  const segments = []; // {days: number[]|null, classes: string[], subject, medium}
  let pendingDays = [], rangeFrom = null, last = null;
  let i = 0;
  const flushDaysInto = seg => { if (pendingDays.length) { seg.days = (seg.days || []).concat(pendingDays); pendingDays = []; } };
  while (i < t.length) {
    const rest = t.slice(i);
    let m;
    if ((m = rest.match(/^[\s,\-–./]+/))) { i += m[0].length; continue; }
    if ((m = rest.match(/^TO(?![A-Z])/))) { rangeFrom = pendingDays.length ? pendingDays[pendingDays.length - 1] : null; i += m[0].length; continue; }
    if ((m = rest.match(/^(MON|TUES?|WED|THU(?:RS?)?|FRI|SAT)(?![A-Z])/))) {
      const d = DAY_WORDS[m[1]];
      if (rangeFrom !== null) { for (let k = rangeFrom + 1; k <= d; k++) pendingDays.push(k); rangeFrom = null; } else pendingDays.push(d);
      i += m[0].length; continue;
    }
    if ((m = rest.match(/^(\d{1,2})\s*TH\s*([A-E](?:\s*\+\s*[A-E])*)(?![A-Z])/))) {
      const classes = m[2].split('+').map(x => `${m[1]}th ${x.trim()}`);
      last = { days: null, classes, subject: '', medium: '' };
      flushDaysInto(last);
      segments.push(last);
      i += m[0].length; continue;
    }
    if ((m = rest.match(/^\(([^)]*)\)?/))) {
      const inner = clean(m[1]).replace(/\.$/, '');
      if (!last) problems.push(`"(${inner})" before any class`);
      else if (/^(H|E)\.?M$/.test(inner)) last.medium = inner[0];
      else if (/^H\.?M\s*&\s*E\.?M$/.test(inner)) last.medium = '';
      else if (HINTS[inner]) last.subject = HINTS[inner];
      else problems.push(`unknown hint "(${inner})"`);
      i += m[0].length; continue;
    }
    if ((m = rest.match(/^[A-Z.]+/))) {
      const word = m[0].replace(/\.$/, '');
      const hint = HINTS[word];
      if (hint && last) last.subject = hint;
      else problems.push(hint ? `"${word}" without a class` : `unknown word "${word}"`);
      i += m[0].length; continue;
    }
    problems.push(`cannot read "${rest}"`);
    break;
  }
  if (pendingDays.length) { // trailing days ("12TH C (ACC) MON TO THU") belong to the last class list
    if (last && !last.days) last.days = pendingDays; else problems.push('days without a class');
  }
  const named = new Set(segments.flatMap(s => s.days || []));
  const others = [0, 1, 2, 3, 4, 5].filter(d => !named.has(d));
  const claims = [];
  segments.forEach(s => (s.days || others).forEach(d => s.classes.forEach(cls => claims.push({ day: DAYS[d], cls, subject: s.subject, medium: s.medium }))));
  return { claims, problems };
}

/**
 * @returns {{claims: Object[], teachers: Object[], classTeachers: Map<string, string[]>, issues: string[], notes: string[]}}
 */
function parseTeacherwise(file, sheetName) {
  const wb = XLSX.readFile(file);
  const ws = wb.Sheets[sheetName];
  if (!ws) throw new Error(`Sheet "${sheetName}" not found in ${file} (sheets: ${wb.SheetNames.join(', ')})`);
  const grid = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', blankrows: true, raw: false });
  const addr = (r, c) => XLSX.utils.encode_cell({ r, c });
  const issues = [], notes = [], claims = [], teachers = [];
  const classSet = new Set(CLASSES.map(c => c[0]));

  const head = grid.findIndex(r => r.some(v => /employee name/i.test(clean(v))));
  if (head < 0) throw new Error(`${sheetName}: no "Employee Name" header`);
  const headRows = [grid[head - 1] || [], grid[head]];
  const colOf = re => { for (const r of headRows) { const c = r.findIndex(v => re.test(clean(v))); if (c >= 0) return c; } return -1; };
  const nameCol = colOf(/employee name/i), ctCol = colOf(/class teacher/i), subjCol = colOf(/^subject/i);
  const periodCol = {};
  headRows.forEach(r => r.forEach((v, c) => { if (/^[1-8]$/.test(clean(v))) periodCol[Number(clean(v))] = c; }));
  if (Object.keys(periodCol).length !== 8) issues.push(`${sheetName}: expected period columns 1–8 in the header, found ${Object.keys(periodCol).join(', ')}`);

  const classTeachers = new Map();
  for (let i = head + 1; i < grid.length; i++) {
    const r = grid[i];
    const written = clean(r[nameCol]);
    if (!written) continue;
    const nm = written.match(/^(.*?)\s*-?\s*\(([A-Z0-9]+)\)\s*$/);
    const plainName = clean(nm ? nm[1] : written);
    const teacher = teacherName(plainName);
    const cellsWithText = Object.entries(periodCol).filter(([, c]) => clean(r[c]));
    if (!teacher) {
      if (cellsWithText.length) issues.push(`${addr(i, nameCol)}: unknown teacher "${plainName}" — add the spelling to TEACHER_ALIASES in lib/vocabulary.js`);
      else notes.push(`${addr(i, nameCol)}: "${plainName}" is not in the registry (no periods in classes 6–12)`);
      continue;
    }
    const ct = clean(r[ctCol]);
    const ctClasses = [];
    const ctm = ct.match(/^(\d{1,2})\s*(?:st|nd|rd|th)\s*([A-E](?:\s*\/\s*[A-E])*)$/i);
    if (ctm) ctm[2].split('/').forEach(x => { const c = `${ctm[1]}th ${x.trim().toUpperCase()}`; if (classSet.has(c)) ctClasses.push(c); });
    ctClasses.forEach(c => classTeachers.set(c, (classTeachers.get(c) || []).concat(teacher)));
    teachers.push({ teacher, written: plainName, code: nm ? nm[2] : '', subject: clean(r[subjCol]), classTeacher: ct, row: i + 1 });

    cellsWithText.forEach(([p, c]) => {
      const text = clean(r[c]);
      const cell = `${sheetName}!${addr(i, c)}`;
      const parsed = parseClaimCell(text);
      parsed.problems.forEach(x => issues.push(`${cell}: ${teacher} P${p} "${text}" — ${x}; not used`));
      parsed.claims.forEach(cl => {
        if (!classSet.has(cl.cls)) { notes.push(`${cell}: ${teacher} P${p} names ${cl.cls}, which is not a class in the timetable ("${text}")`); return; }
        claims.push({ teacher, period: Number(p), ...cl, cell, text });
      });
    });
  }
  // the same name twice with different employee codes
  const byName = new Map();
  teachers.forEach(t => byName.set(t.teacher, (byName.get(t.teacher) || []).concat(t)));
  byName.forEach((list, n) => { if (list.length > 1) notes.push(`${n} is listed ${list.length} times (${list.map(t => t.code || 'no code').join(', ')}) — different employees with the same name?`); });
  return { claims: dedupe(claims), teachers, classTeachers, issues: [...new Set(issues)], notes: [...new Set(notes)] };
}

function dedupe(claims) {
  const seen = new Map();
  claims.forEach(c => {
    const k = `${c.teacher}|${c.day}|${c.period}|${c.cls}`;
    if (!seen.has(k)) seen.set(k, c);
    else { const o = seen.get(k); o.subject = o.subject || c.subject; o.medium = o.medium || c.medium; }
  });
  return [...seen.values()];
}

/** Base subject of a lane for matching: "Proficiency English" → "English", "Physics (H.M)" → "Physics". */
const baseOf = subject => splitMedium(subject).base.replace(/^Proficiency\s+/, '').replace(/\s+Practical$/, '').trim().toLowerCase();

/**
 * Fills groups that have no teacher using the teacher-wise sheet, when the evidence
 * points to exactly one teacher. Returns the new rows plus a log of what was filled
 * and what could not be.
 *
 * A candidate is a teacher the teacher-wise sheet puts in this class at this day and
 * period who is not already in the cell and would not clash with what the class-wise
 * sheet already gives them at that time. Score:
 *   +4 the hint names this subject · +2 the hint's medium matches (a different medium excludes)
 *   +3 the teacher teaches this subject to this class elsewhere in the class-wise sheet
 *   +2 otherwise teaches it anywhere, or lists it as their subject
 * The best candidate is used when it scores ≥ 2 and no other candidate scores the same.
 */
function fillFromClaims(rows, tw) {
  const P = loadGrammar().ScheduleParser;
  const subjectsOf = new Map(); // teacher → Set(base subject)
  const classSubjects = new Map(); // teacher|class → Set(base subject)
  rows.forEach(r => P.parseLanes(r).lanes.forEach(l => l.teachers.forEach(t => {
    const b = baseOf(l.subject);
    (subjectsOf.get(t) || subjectsOf.set(t, new Set()).get(t)).add(b);
    const k = `${t}|${r.Class}`;
    (classSubjects.get(k) || classSubjects.set(k, new Set()).get(k)).add(b);
  })));
  const listed = new Map();
  TEACHERS.forEach(([n, spec]) => listed.set(n, String(spec || '').toLowerCase()));
  tw.teachers.forEach(t => listed.set(t.teacher, `${listed.get(t.teacher) || ''} ${t.subject.toLowerCase()}`));
  const listsSubject = (t, b) => (listed.get(t) || '').split(/[\s/,]+/).some(w => w && (w === b || (b.length > 3 && w.startsWith(b.slice(0, 4)))));

  const at = new Map();
  tw.claims.forEach(c => { const k = `${c.day}|${c.period}|${c.cls}`; (at.get(k) || at.set(k, []).get(k)).push(c); });
  // What each teacher already teaches at each day / period in the class-wise sheet: a
  // candidate who would clash there is not used (the clash list and the comparison show it).
  const busy = new Map();
  const addBusy = (t, day, period, cls, subject) => { const k = `${t}|${day}|${period}`; (busy.get(k) || busy.set(k, []).get(k)).push({ cls, subject }); };
  rows.forEach(r => P.parseLanes(r).lanes.forEach(l => l.teachers.forEach(t => addBusy(t, r.Day, r.Period, r.Class, l.subject))));
  const clashesWith = (t, r, subject) => {
    const already = (busy.get(`${t}|${r.Day}|${r.Period}`) || []).filter(x => x.cls !== r.Class);
    return already.length && P.classifyTeacherSlot([...already, { cls: r.Class, subject }]).clash ? already : null;
  };

  const filled = [], unfilled = [];
  const out = rows.map(r => {
    const parsed = P.parseLanes(r);
    const empty = parsed.lanes.filter(l => String(l.subject || '').trim() && !l.teachers.length);
    if (!empty.length) return r;
    const inCell = new Set(parsed.lanes.flatMap(l => l.teachers));
    let pool = (at.get(`${r.Day}|${r.Period}|${r.Class}`) || []).filter(c => !inCell.has(c.teacher));
    let changed = false;
    empty.forEach(lane => {
      const b = baseOf(lane.subject);
      const medium = splitMedium(lane.subject).medium;
      const busyNotes = [];
      const scored = pool.map(c => {
        let s = 0;
        if (c.subject) { if (c.subject.toLowerCase() === b || (c.subject === 'Practical' && /practical/i.test(lane.subject))) s += 4; else return { c, s: -1 }; }
        if (c.medium) { if (medium && c.medium !== medium) return { c, s: -1 }; if (c.medium === medium) s += 2; }
        const other = clashesWith(c.teacher, r, lane.subject);
        if (other) { busyNotes.push(`${c.teacher} (${c.cell.replace(/^.*!/, '')}) — but the class-wise sheet has them in ${other.map(x => `${x.cls} ${x.subject}`).join(', ')}`); return { c, s: -1 }; }
        if ((classSubjects.get(`${c.teacher}|${r.Class}`) || new Set()).has(b)) s += 3;
        else if ((subjectsOf.get(c.teacher) || new Set()).has(b) || listsSubject(c.teacher, b)) s += 2;
        return { c, s };
      }).filter(x => x.s >= 0).sort((a, b2) => b2.s - a.s);
      const where = { cls: r.Class, day: r.Day, period: r.Period, subject: lane.subject, source: r.Source || '' };
      if (scored.length && scored[0].s >= 2 && !(scored[1] && scored[1].s === scored[0].s)) {
        const c = scored[0].c;
        lane.teachers = [c.teacher];
        pool = pool.filter(x => x !== c);
        addBusy(c.teacher, r.Day, r.Period, r.Class, lane.subject);
        filled.push({ ...where, teacher: c.teacher, evidence: `${c.cell} "${c.text}"` });
        changed = true;
      } else {
        unfilled.push({ ...where, candidates: [...scored.map(x => `${x.c.teacher} (${x.c.cell.replace(/^.*!/, '')})`), ...busyNotes] });
      }
    });
    if (!changed) return r;
    const f = P.formatLanes(parsed.lanes);
    return { ...r, Subject: f.subject, Teacher: f.teacher };
  });
  return { rows: out, filled, unfilled };
}

/**
 * Compares the class-wise timetable with the teacher-wise sheet.
 * @returns {{classOnly: Object[], teacherOnly: Object[]}} one entry per teacher / day / period / class
 */
function compare(rows, tw) {
  const P = loadGrammar().ScheduleParser;
  const fromClass = new Map();
  rows.forEach(r => P.parseLanes(r).lanes.forEach(l => l.teachers.forEach(t => {
    fromClass.set(`${t}|${r.Day}|${r.Period}|${r.Class}`, { teacher: t, day: r.Day, period: r.Period, cls: r.Class, subject: l.subject, source: r.Source || '' });
  })));
  const fromTw = new Map(tw.claims.map(c => [`${c.teacher}|${c.day}|${c.period}|${c.cls}`, c]));
  const listedTeachers = new Set(tw.teachers.map(t => t.teacher));
  const classOnly = [...fromClass.entries()].filter(([k, v]) => listedTeachers.has(v.teacher) && !fromTw.has(k)).map(([, v]) => v);
  const teacherOnly = [...fromTw.entries()].filter(([k]) => !fromClass.has(k)).map(([, v]) => v);
  return { classOnly, teacherOnly };
}

module.exports = { parseTeacherwise, parseClaimCell, fillFromClaims, compare, baseOf };
