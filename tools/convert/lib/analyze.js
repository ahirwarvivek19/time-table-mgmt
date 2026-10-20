/**
 * analyze.js — checks a converted timetable the same way the sheet does, and
 * lists what a reviewer needs to see: teacher clashes, combined lectures
 * (one teacher, one subject, several sections at once) and split periods
 * (one section divided into groups with their own subject / teacher / room).
 */
'use strict';

const { loadGrammar } = require('./gas');

const DAYS = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
const norm = s => String(s == null ? '' : s).replace(/\s+/g, ' ').trim().toLowerCase();

/**
 * Puts every row into the canonical cell form and reports rows whose groups do not
 * line up. Returns new row objects; the input is not modified.
 * @param {Array<Object>} rows {Day, Period, Class, Tier, Subject, Teacher, Room, Source?}
 * @returns {{rows: Array<Object>, errors: string[]}}
 */
function canonicalize(rows) {
  const { ScheduleParser: P } = loadGrammar();
  const errors = [];
  const out = rows.map(r => {
    const where = `${r.Class} ${r.Day} P${r.Period}${r.Source ? ' (' + r.Source + ')' : ''}`;
    const parsed = P.parseLanes(r);
    const subjects = P.splitList(r.Subject);
    const teacherLanes = P.splitList(r.Teacher);
    if (parsed.mismatch) errors.push(`${where}: ${subjects.length} subjects but ${teacherLanes.length} teacher groups — "${r.Subject}" ↔ "${r.Teacher}"`);
    if (parsed.roomMismatch) errors.push(`${where}: rooms "${r.Room}" do not line up with subjects "${r.Subject}"`);
    if (teacherLanes.length === 1 && subjects.length > 1) {
      errors.push(`${where}: one teacher for ${subjects.length} subjects "${r.Subject}" — use split(ln(..), ln(..)) or a subject name without "/"`);
    }
    const f = P.formatLanes(parsed.lanes);
    return { ...r, Subject: f.subject, Teacher: f.teacher, Room: f.room };
  });
  return { rows: out, errors };
}

/** Lanes of one row, each {subject, teachers[], room}. */
function lanesOf(row) {
  return loadGrammar().ScheduleParser.parseLanes(row).lanes;
}

/**
 * @param {Array<Object>} rows canonical rows
 * @returns {{clashes: Array, combined: Array, splits: Array, roomClashes: Array}}
 */
function analyze(rows) {
  const teacherSlots = {}; // day|period|teacher → [{cls, subject}]
  const roomSlots = {};
  const splits = [];
  rows.forEach(r => {
    const lanes = lanesOf(r);
    if (lanes.length > 1) splits.push({ cls: r.Class, day: r.Day, period: +r.Period, lanes });
    lanes.forEach(l => {
      l.teachers.forEach(t => {
        const k = `${r.Day}|${r.Period}|${t}`;
        (teacherSlots[k] = teacherSlots[k] || []).push({ cls: r.Class, subject: l.subject, cell: r.Source || '' });
      });
    });
    const laneRooms = lanes.map(l => l.room || '');
    const perLane = lanes.length > 1 && laneRooms.some(x => x !== laneRooms[0]);
    (perLane ? lanes.map(l => ({ room: l.room, sig: norm(l.subject) + '|' + l.teachers.map(norm).sort().join('+') }))
      : [{ room: r.Room, sig: norm(r.Subject) + '|' + norm(r.Teacher) }])
      .forEach(u => {
        if (!u.room || /^(n\/?a|none|-)$/i.test(String(u.room).trim())) return;
        const k = `${r.Day}|${r.Period}|${norm(u.room)}`;
        (roomSlots[k] = roomSlots[k] || []).push({ cls: r.Class, sig: u.sig, room: u.room });
      });
  });

  const clashes = [], combined = [];
  Object.entries(teacherSlots).forEach(([k, list]) => {
    const [day, period, teacher] = k.split('|');
    if (list.length < 2) return;
    const verdict = loadGrammar().ScheduleParser.classifyTeacherSlot(list);
    if (verdict.clash) {
      clashes.push({ day, period: +period, teacher, reason: verdict.reason, detail: list.map(x => `${x.cls} (${x.subject})`).join(', '), cells: [...new Set(list.map(x => x.cell).filter(Boolean))].join(', ') });
    } else {
      combined.push({ day, period: +period, teacher, subject: list[0].subject, classes: [...new Set(list.map(x => x.cls))] });
    }
  });
  const roomClashes = [];
  Object.entries(roomSlots).forEach(([k, list]) => {
    if (new Set(list.map(x => x.sig)).size > 1) {
      const [day, period] = k.split('|');
      roomClashes.push({ day, period: +period, room: list[0].room, detail: list.map(x => x.cls).join(', ') });
    }
  });
  const byTeacher = (a, b) => a.teacher.localeCompare(b.teacher) || DAYS.indexOf(a.day) - DAYS.indexOf(b.day) || a.period - b.period;
  return { clashes: clashes.sort(byTeacher), combined: combined.sort(byTeacher), splits, roomClashes };
}

/** "Mon, Tue, Wed" or "Mon–Sat" for a set of day names. */
function dayList(days) {
  const idx = [...new Set(days)].map(d => DAYS.indexOf(d)).sort((a, b) => a - b);
  if (idx.length === 6) return 'Mon–Sat';
  return idx.map(i => DAYS[i].slice(0, 3)).join(', ');
}

/** Merges entries that differ only by day: key(entry) → one line with a day list. */
function collapseDays(entries, key) {
  const groups = new Map();
  entries.forEach(e => {
    const k = key(e);
    if (!groups.has(k)) groups.set(k, { ...e, days: [] });
    groups.get(k).days.push(e.day);
  });
  return [...groups.values()].map(g => ({ ...g, dayText: dayList(g.days) }));
}

module.exports = { canonicalize, analyze, lanesOf, collapseDays, dayList, DAYS };
