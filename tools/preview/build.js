#!/usr/bin/env node
/**
 * build.js — see the timetable locally, without pushing anything to Apps Script.
 *
 *   npm run preview          → writes preview/timetable.html (open it in any browser)
 *   npm run preview:serve    → also serves it at http://127.0.0.1:8765 (this machine only)
 *   node tools/preview/build.js --serve --port 9000
 *
 * It runs the real .gs files against the in-memory Sheets stand-in the tests use
 * (tests/harness/gas-mock.js): Setup → Import Timetable (the built-in data in
 * PrefillData.gs) → every Class_View, Teacher_View and Teacher_Day_View is rendered, and
 * the cell text, colours and hover notes the sheet would show are copied into one
 * self-contained page. Fonts, row heights and column widths are not reproduced.
 *
 * Re-run after `npm run convert` to see a new timetable. The page needs no network.
 */
'use strict';

const fs = require('fs');
const path = require('path');
const http = require('http');
const crypto = require('crypto');
const { createEnvironment, loadProject } = require('../../tests/harness/gas-mock');

const ROOT = path.resolve(__dirname, '..', '..');
const OUT = path.join(ROOT, 'preview', 'timetable.html');
const TEMPLATE = path.join(__dirname, 'page.html');
const SCRIPT = path.join(__dirname, 'page.js');
const HOST = '127.0.0.1'; // loopback only: the preview is for this machine, never widen this bind
const DEFAULT_PORT = 8765;

const plain = x => JSON.parse(JSON.stringify(x));
const text = v => String(v == null ? '' : v);

/** Cell text, background and note for a block of a mock sheet. */
function readBlock(sheet, row, col, rows, cols) {
  const values = sheet.getRange(row, col, rows, cols).getValues();
  return values.map((r, i) => r.map((v, j) => {
    const cell = sheet.getRange(row + i, col + j);
    return { v: text(v), bg: cell.getBackground(), note: text(cell.getNote()) };
  }));
}

/**
 * The files the built-in data came from, from the PrefillData.gs header:
 * "Timetable.xlsx (sheets 6-8, 9-12, Teacherwise)". Unrecognised lines are kept as they are.
 */
function sourcesOf() {
  const head = fs.readFileSync(path.join(ROOT, 'PrefillData.gs'), 'utf8').split('*/')[0];
  const lines = head.split('\n').filter(l => /^ \* (Sources:| {9})/.test(l)).map(l => l.replace(/^ \* (Sources:)?\s*/, '').trim());
  const files = new Map(), other = [];
  lines.forEach(l => {
    const m = l.match(/\(([^,()]+), sheet ([^)]+)\)/);
    if (m) (files.get(m[1]) || files.set(m[1], []).get(m[1])).push(m[2]); else other.push(l);
  });
  return [...[...files].map(([f, sheets]) => `${f} (sheet${sheets.length > 1 ? 's' : ''} ${sheets.join(', ')})`), ...other];
}

/** Renders every view in the in-memory sheet and returns what the page shows. */
function collect() {
  const env = createEnvironment();
  const { api } = loadProject(ROOT, env);
  api.setupInitialSpreadsheet();
  api.importExcelData();
  const ss = env.ss;
  const C = plain(api.CONFIG);
  const days = C.DAYS, nP = C.PERIODS.length, nD = days.length;
  const classes = plain(api.getClassNames_());
  const teachers = plain(api.getTeacherNames_());
  const check = plain(api.validateSchedule_()) || {};

  // Class_View: row 5 = period headers, then one subject row and one teacher row per day.
  const classData = classes.map(name => {
    api.ClassViewManager.renderClassView(name);
    const sh = ss.getSheetByName(C.SHEETS.CLASS_VIEW);
    return {
      name,
      days: days.map((d, i) => {
        const subj = readBlock(sh, 6 + 2 * i, 2, 1, nP)[0];
        const tchr = readBlock(sh, 7 + 2 * i, 2, 1, nP)[0];
        return subj.map((s, p) => [s.v, tchr[p].v, s.bg, s.note]);
      }),
    };
  });
  const header = readBlock(ss.getSheetByName(C.SHEETS.CLASS_VIEW), 5, 2, 1, nP)[0]
    .map((h, p) => ({ label: `Period ${p + 1}`, time: h.v.split('\n')[1] || '' }));

  // Teacher_View: rows 6–11 = Monday–Saturday.
  const load = new Map(plain(api.DataAccess.getSheetDataAsObjects(C.SHEETS.TEACHERS))
    .map(t => [text(t['Teacher Name']).trim(), { hours: Number(t['Total Hours Scheduled']) || 0, max: Number(t['Max Hours / Week']) || 0 }]));
  const teacherData = teachers.map(name => {
    api.TeacherViewManager.renderTeacherView(name);
    const cells = readBlock(ss.getSheetByName(C.SHEETS.TEACHER_VIEW), 6, 2, nD, nP).map(r => r.map(c => [c.v, c.bg]));
    return { name, ...(load.get(name) || { hours: 0, max: 0 }), cells };
  });

  // Teacher_Day_View: one row per teacher.
  const dayData = days.map(day => {
    api.TeacherDayViewManager.renderTeacherDayView(day, C.ALL_CLASSES);
    const sh = ss.getSheetByName(C.SHEETS.TEACHER_DAY_VIEW);
    const rows = readBlock(sh, 6, 1, teachers.length, nP + 1);
    return { day, rows: rows.map(r => ({ teacher: r[0].v, cells: r.slice(1).map(c => [c.v, c.bg]) })) };
  });

  // Master_Grid_View: one row per class.
  api.generateMasterGrid(true);
  const grid = ss.getSheetByName(C.SHEETS.MASTER_GRID).getRange(2, 1, classes.length, nP + 1).getValues().map(r => r.map(text));

  // Rows Check For Conflicts flagged.
  const ms = ss.getSheetByName(C.SHEETS.SCHEDULE);
  const conflicts = ms.getRange(2, 1, ms.getLastRow() - 1, C.SCHEDULE_HEADERS.length).getValues()
    .filter(r => text(r[C.COL.CLASH - 1]).trim())
    .map(r => [r[0], r[1], r[2], r[4], r[5], r[C.COL.CLASH - 1]].map(text));

  return {
    generated: new Date().toLocaleString('en-IN', { dateStyle: 'medium', timeStyle: 'short' }),
    sources: sourcesOf(),
    days, periods: header,
    summary: {
      classes: classes.length, teachers: teachers.length, rows: ms.getLastRow() - 1,
      splitSlots: check.splitSlots || 0, combinedLectures: check.combinedLectures || 0,
      rowsFlagged: check.rowsFlagged || 0, counts: check.counts || {}, overloaded: check.overloaded || [],
    },
    classes: classData, teachers: teacherData, dayView: dayData, grid, conflicts,
  };
}

/** Fills the page template: data as a JSON block, script inline with a CSP hash. */
function renderHtml(data) {
  const script = fs.readFileSync(SCRIPT, 'utf8');
  const hash = crypto.createHash('sha256').update(script).digest('base64');
  // "<" is escaped so no value can close the data <script> element.
  const json = JSON.stringify(data).replace(/</g, '\\u003c');
  return fs.readFileSync(TEMPLATE, 'utf8')
    .replace('__SCRIPT_HASH__', () => `sha256-${hash}`)
    .replace('__DATA__', () => json)
    .replace('__SCRIPT__', () => script);
}

/**
 * Builds the preview page.
 * @param {{out?: string}} [opts]
 * @returns {{file: string, data: Object}}
 */
function buildPreview(opts = {}) {
  const file = opts.out || OUT;
  const data = collect();
  fs.mkdirSync(path.dirname(file), { recursive: true });
  fs.writeFileSync(file, renderHtml(data));
  return { file, data };
}

/**
 * Serves the page (read from disk on every request, so a re-run of `npm run preview`
 * shows up on reload) at http://127.0.0.1:<port>/. Nothing else is served.
 * @param {{port?: number, file?: string}} [opts] port 0 picks a free port
 * @returns {Promise<{server: http.Server, url: string}>}
 */
function startServer(opts = {}) {
  const file = opts.file || OUT;
  const server = http.createServer((req, res) => {
    const url = (req.url || '/').split('?')[0];
    if (req.method !== 'GET' && req.method !== 'HEAD') {
      res.writeHead(405, { Allow: 'GET, HEAD', 'Content-Type': 'text/plain; charset=utf-8' });
      return res.end('Method not allowed');
    }
    if (url !== '/' && url !== '/timetable.html') {
      res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
      return res.end('Not found');
    }
    let body;
    try { body = fs.readFileSync(file); } catch (e) {
      res.writeHead(503, { 'Content-Type': 'text/plain; charset=utf-8' });
      return res.end('No preview yet: run npm run preview');
    }
    res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8', 'Cache-Control': 'no-store', 'X-Content-Type-Options': 'nosniff' });
    res.end(req.method === 'HEAD' ? undefined : body);
  });
  return new Promise((resolve, reject) => {
    server.once('error', reject);
    server.listen(opts.port == null ? DEFAULT_PORT : opts.port, HOST, () => {
      resolve({ server, url: `http://${HOST}:${server.address().port}/` });
    });
  });
}

function parseArgs(argv) {
  const opts = {};
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === '--serve') opts.serve = true;
    else if (a === '--port') {
      const p = Number(argv[++i]);
      if (!Number.isInteger(p) || p < 1024 || p > 65535) throw new Error('--port needs a number from 1024 to 65535');
      opts.port = p;
    } else if (a === '--help' || a === '-h') opts.help = true;
    else throw new Error(`Unknown option ${a} (see --help)`);
  }
  return opts;
}

async function main() {
  let opts;
  try { opts = parseArgs(process.argv.slice(2)); } catch (e) { console.error('✗ ' + e.message); process.exit(2); }
  if (opts.help) { console.log(fs.readFileSync(__filename, 'utf8').split('*/')[0].replace(/^#!.*\n\/\*\*\n?/, '').replace(/^ \* ?/gm, '')); return; }
  const { file, data } = buildPreview();
  const s = data.summary;
  console.log(`✓ ${path.relative(ROOT, file)} — ${s.classes} classes, ${s.teachers} teachers, ${s.rows} periods, ${s.rowsFlagged} rows with conflicts`);
  if (!opts.serve) { console.log('  open it in a browser, or run npm run preview:serve'); return; }
  try {
    const { url } = await startServer({ port: opts.port });
    console.log(`  serving at ${url} (this machine only) — Ctrl+C to stop`);
  } catch (e) {
    console.error(e.code === 'EADDRINUSE' ? `✗ port ${opts.port || DEFAULT_PORT} is in use — pass --port <another>` : '✗ ' + e.message);
    process.exit(1);
  }
}

module.exports = { buildPreview, renderHtml, startServer, HOST };
if (require.main === module) main();
