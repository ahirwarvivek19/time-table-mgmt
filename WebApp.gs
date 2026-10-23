/**
 * WebApp.gs
 * Optional READ-ONLY timetable page for teachers (mobile-friendly, printable).
 *
 * Nothing is reachable until you deploy it. Deploy it ONLY like this:
 *   Apps Script editor → Deploy → New deployment → Web app
 *   Execute as:      "User accessing the web app"   ← required
 *   Who has access:  "Anyone within <school domain>" (or "Anyone with Google account")
 * and share the spreadsheet with teachers as **Viewer**.
 *
 * Why "User accessing": a web app page can call ANY public server function via
 * google.script.run (including the Cover Manager and menu functions that write).
 * Running as the visitor means Google's normal sheet permissions apply, so a Viewer
 * can only read. Deploying as "Me" would let every visitor write with your access.
 */

function doGet() {
  return HtmlService.createTemplateFromFile('TimetableWeb')
    .evaluate()
    .setTitle('School Timetable')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Spreadsheet for web-app calls (no "active" spreadsheet outside the Sheets UI). */
function getTimetableSpreadsheet_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('Timetable not configured yet: open the sheet and run Setup once.');
  return SpreadsheetApp.openById(id);
}

/** Remembers the bound spreadsheet id so the web app can open it. Called from setup/refresh. */
function rememberSpreadsheetId_() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());
  } catch (err) {
    // Not fatal: the web app is optional.
  }
}

function readObjects_(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());
  return values.slice(1).map(r => {
    const o = {};
    headers.forEach((h, i) => { if (h) o[h] = r[i]; });
    return o;
  });
}

/** Class and teacher lists for the page's selectors. */
function apiWebGetOptions() {
  const ss = getTimetableSpreadsheet_();
  const S = CONFIG.SHEETS;
  const unique = list => list.map(v => String(v || '').trim()).filter((v, i, a) => v && a.indexOf(v) === i);
  return {
    classes: unique(readObjects_(ss, S.CLASSES).map(c => c['Class Name'])),
    teachers: unique(readObjects_(ss, S.TEACHERS).map(t => t['Teacher Name'])),
    days: CONFIG.DAYS,
    periods: CONFIG.PERIODS.map(p => ({ period: p, time: periodTime_(p) }))
  };
}

/**
 * Weekly grid for one class or one teacher.
 * @param {'class'|'teacher'} kind
 * @param {string} name
 * @returns {{title: string, rows: Array<{day: string, cells: string[]}>}}
 */
function apiWebGetTimetable(kind, name) {
  const ss = getTimetableSpreadsheet_();
  const opts = apiWebGetOptions();
  // Only names that exist in the registries can be queried
  const list = kind === 'teacher' ? opts.teachers : opts.classes;
  const match = list.find(n => ScheduleParser.sameName(n, name));
  if (!match) throw new Error('Unknown ' + (kind === 'teacher' ? 'teacher' : 'class') + ': ' + name);

  const grid = {};
  CONFIG.DAYS.forEach(d => { grid[d] = {}; CONFIG.PERIODS.forEach(p => { grid[d][p] = []; }); });

  readObjects_(ss, CONFIG.SHEETS.SCHEDULE).forEach(row => {
    const p = parseInt(row.Period);
    const day = normalizeDay_(row.Day);
    if (!grid[day] || !grid[day][p]) return;
    if (kind === 'teacher') {
      if (!ScheduleParser.rowIncludesTeacher(row, match)) return;
      grid[day][p].push(String(row.Class) + ' · ' + (ScheduleParser.getSubjectForTeacher(row, match) || row.Subject || ''));
    } else {
      if (!ScheduleParser.rowIncludesClass(row, match)) return;
      if (!row.Subject && !row.Teacher) return;
      grid[day][p].push(String(row.Subject || '') + (row.Teacher ? ' · ' + row.Teacher : ''));
    }
  });

  return {
    title: match,
    rows: CONFIG.DAYS.map(d => ({ day: d, cells: CONFIG.PERIODS.map(p => [...new Set(grid[d][p])].join('\n')) }))
  };
}
