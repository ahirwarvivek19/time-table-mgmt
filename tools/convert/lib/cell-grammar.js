/**
 * cell-grammar.js — reads one period cell of the school's class-wise timetable
 * into groups ("lanes"): who teaches what to which part of the class.
 *
 *   "SANSKRIT NEHA JAMIDAR"                           → Sanskrit — Neha Jamidar
 *   "ENG \nRuchika V."                                → English — Ruchika Vyas
 *   "SANSKRIT NEHA.J IT-DIVYA BEAUTY-KRISHNA"         → split: Sanskrit / IT / Beauty
 *   "PHYSICS GK.DIXIT (H.M) / CHEMISTRY D.BHATI (E.M)" → split by medium
 *   "SO.SCI RUBAL VERMA(H.M)/ I.MATHUR (E.M)"         → one subject, a group per medium
 *   "CCLE INDRA.M/ KRISHNA.P"                         → one group, two teachers (co-teaching)
 *   "MATHS /CCLE NEHA JHALA"                          → one group "Mathematics & CCLE"
 *   "PLOY.SCI (RO.S)/ HINDI (ARCHANA.P)"              → teacher written in brackets
 *   "DAKSHATA HINDI"                                  → Proficiency Hindi, no teacher written
 *
 * Rules, in order of the words in the cell:
 *  - every subject word starts a new group;
 *  - a teacher belongs to the group before it (teachers written before any subject
 *    belong to the first subject: "Ruchika V./UMANG ADDHA PHOOL");
 *  - "(H.M)" / "(E.M)" marks the group's medium. A second teacher after a group that
 *    already has a teacher and a medium opens a new group of the same subject;
 *  - "(PRACTICAL)" turns "Geography" into "Geography Practical";
 *  - "X /CCLE" with no teacher between joins the two into one group "X & CCLE".
 * Words the vocabulary does not know are reported, never guessed.
 */
'use strict';

const { TEACHERS } = require('../registry');
const V = require('./vocabulary');
const { loadGrammar } = require('./gas');

/** Upper case, "-" read as a space, single spaces, no space after "." — the form all words are compared in. */
const norm = s => String(s == null ? '' : s).toUpperCase().replace(/[-–—]/g, ' ').replace(/\s+/g, ' ').replace(/\.\s+/g, '.').trim();
const isWordChar = ch => /[A-Z0-9\u0900-\u097F]/.test(ch || '');
const MEDIUM_RE = /^(H|E)\.?M\.?$/;
const MEDIUM_LABEL = { H: '(H.M)', E: '(E.M)' };

let matcher = null;
/** All known words, longest first, so "SO.SCI" wins over "SCI" and "KRISHNA.P" over "KRISHNA". */
function words() {
  if (matcher) return matcher;
  const list = [];
  Object.entries(V.SUBJECTS).forEach(([w, v]) => list.push({ w: norm(w), kind: 'S', v }));
  Object.entries(V.SUBJECT_PREFIXES).forEach(([w, v]) => list.push({ w: norm(w), kind: 'P', v }));
  Object.entries(V.TEACHER_ALIASES).forEach(([w, v]) => list.push({ w: norm(w), kind: 'T', v }));
  TEACHERS.forEach(([n]) => list.push({ w: norm(n), kind: 'T', v: n }));
  matcher = list.sort((a, b) => b.w.length - a.w.length);
  return matcher;
}

/** The registry name for a written teacher name, or '' when unknown. */
function teacherName(written) {
  const w = norm(written).replace(/\.$/, '');
  const hit = words().find(x => x.kind === 'T' && x.w.replace(/\.$/, '') === w);
  return hit ? hit.v : '';
}

/**
 * Splits a cell into tokens: {kind: 'S'|'P'|'T'|'M'|'PRAC'|'SLASH'|'?', v}.
 * @param {string} raw
 */
function tokenize(raw) {
  const t = norm(raw);
  const out = [];
  let i = 0;
  while (i < t.length) {
    const ch = t[i];
    if (ch === '/') { out.push({ kind: 'SLASH' }); i++; continue; }
    if (/[\s,.&+]/.test(ch)) { i++; continue; }
    if (ch === '(') {
      const j = t.indexOf(')', i);
      const inner = (j < 0 ? t.slice(i + 1) : t.slice(i + 1, j)).trim();
      i = j < 0 ? t.length : j + 1;
      if (MEDIUM_RE.test(inner)) out.push({ kind: 'M', v: inner[0] });
      else if (/^PRAC(TICAL)?$/.test(inner)) out.push({ kind: 'PRAC' });
      else if (teacherName(inner)) out.push({ kind: 'T', v: teacherName(inner) });
      else out.push({ kind: '?', v: '(' + inner + ')' });
      continue;
    }
    const hit = words().find(x => t.startsWith(x.w, i) && !isWordChar(t[i + x.w.length]));
    if (hit) { out.push({ kind: hit.kind, v: hit.v }); i += hit.w.length; continue; }
    let j = i;
    while (j < t.length && !/[\s/()]/.test(t[j])) j++;
    out.push({ kind: '?', v: t.slice(i, j) });
    i = j;
  }
  return out;
}

/**
 * Reads a cell into groups.
 * @param {string} raw cell text
 * @returns {{lanes: Array<{subject: string, teachers: string[], medium: string}>, problems: string[]}}
 *          medium is 'H', 'E' or ''. subject already carries the medium label, e.g. "Physics (H.M)".
 */
function parseCell(raw) {
  const tokens = tokenize(raw);
  const lanes = [], problems = [];
  let cur = null, pendingTeachers = [], prefix = '', slashBefore = false;
  tokens.forEach(tok => {
    if (tok.kind === 'SLASH') { slashBefore = true; return; }
    if (tok.kind === '?') { problems.push(`unknown word "${tok.v}"`); slashBefore = false; return; }
    if (tok.kind === 'P') { prefix = tok.v; slashBefore = false; return; }
    if (tok.kind === 'S') {
      const subject = prefix ? `${prefix} ${tok.v}` : tok.v;
      prefix = '';
      if (cur && slashBefore && V.PAIRS_WITH_PREVIOUS.includes(tok.v) && !cur.teachers.length && !cur.medium) {
        cur.subject = `${cur.subject} & ${subject}`;
      } else {
        cur = { subject, teachers: pendingTeachers, medium: '' };
        pendingTeachers = [];
        lanes.push(cur);
      }
    } else if (tok.kind === 'T') {
      if (!cur) pendingTeachers.push(tok.v);
      else if (cur.medium && cur.teachers.length) {
        cur = { subject: cur.subject, teachers: [tok.v], medium: '' };
        lanes.push(cur);
      } else if (!cur.teachers.includes(tok.v)) cur.teachers.push(tok.v);
    } else if (tok.kind === 'M') {
      if (!cur) problems.push('medium written before any subject');
      else if (cur.medium && cur.medium !== tok.v) problems.push(`"${cur.subject}" marked both (H.M) and (E.M)`);
      else cur.medium = tok.v;
    } else if (tok.kind === 'PRAC') {
      if (cur) cur.subject += ' Practical'; else problems.push('"(PRACTICAL)" written before any subject');
    }
    slashBefore = false;
  });
  if (prefix) problems.push(`"${prefix}" without a subject`);
  if (pendingTeachers.length) problems.push(`teacher ${pendingTeachers.join(', ')} without a subject`);
  lanes.forEach(l => { if (l.medium) l.subject = `${l.subject} ${MEDIUM_LABEL[l.medium]}`; });
  return { lanes, problems };
}

/** "Physics (H.M)" → {base: "Physics", medium: "H"} — the sheet's own rule (ScheduleParser.splitMedium). */
const splitMedium = subject => loadGrammar().ScheduleParser.splitMedium(subject);

module.exports = { parseCell, tokenize, teacherName, splitMedium, norm };
