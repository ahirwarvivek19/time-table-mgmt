/**
 * Validation.gs
 * Conflict detection for Master_Schedule. Writes a reason into "Clash Status"
 * (column H) for every row involved in a problem and shows one summary.
 *
 * Checks:
 *  - Teacher clash: one teacher in two places at the same day/period teaching
 *    different subjects. The same teacher + same subject in several sections of
 *    one grade is a combined lecture and is allowed (whole-school activities in
 *    CONFIG.CROSS_GRADE_SUBJECTS may also combine grades).
 *  - Room clash: two different lectures in the same real room at the same time
 *    (blank / "N/A" rooms are ignored).
 *  - Unavailable day: teacher scheduled on a day listed in "Days Unavailable".
 *  - Unregistered teacher / unknown class: names that don't match the Teachers /
 *    Classes sheets (these silently disappear from the views otherwise).
 *  - Overload: periods per week above "Max Hours / Week".
 *  - Lane / room mismatch: a split period whose subjects, teachers or rooms
 *    don't line up ("Hindi / IT / Beauty" with only two teacher groups).
 *
 * A split period ("Hindi / IT" ↔ "A / B") is several groups at once, not a clash.
 */

function runValidation() {
  const result = validateSchedule_();
  if (!result) {
    SpreadsheetApp.getUi().alert('Master Schedule is empty.');
    return;
  }
  SpreadsheetApp.getUi().alert(formatValidationSummary_(result));
}

/**
 * Runs all checks and writes Clash Status. Returns null when there is no data.
 * @returns {{rowsFlagged: number, counts: Object<string, number>, overloaded: string[]}|null}
 */
function validateSchedule_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.SCHEDULE);
  if (!sheet || sheet.getLastRow() < 2) return null;

  const C = CONFIG.COL;
  const values = sheet.getRange(1, 1, sheet.getLastRow(), CONFIG.SCHEDULE_HEADERS.length).getValues();
  const norm = s => ScheduleParser.normalizeName(s);

  // Lookups from the registry sheets
  const teachersData = DataAccess.getSheetDataAsObjects(CONFIG.SHEETS.TEACHERS);
  const registered = {};
  const unavailable = {};
  const maxHours = {};
  teachersData.forEach(t => {
    const key = norm(t['Teacher Name']);
    if (!key) return;
    registered[key] = String(t['Teacher Name']).trim();
    unavailable[key] = String(t['Days Unavailable'] || '').split(/[,\/;]/).map(normalizeDay_).filter(Boolean);
    const max = Number(t['Max Hours / Week']);
    if (max > 0) maxHours[key] = max;
  });
  const hasRegistry = Object.keys(registered).length > 0;
  const classSet = {};
  getClassNames_().forEach(c => { classSet[norm(c)] = true; });
  const hasClasses = Object.keys(classSet).length > 0;

  const issues = values.map(() => []); // per row (index 0 = header, unused)
  const add = (i, msg) => { if (issues[i].indexOf(msg) === -1) issues[i].push(msg); };
  const counts = {};
  const count = k => { counts[k] = (counts[k] || 0) + 1; };

  const teacherSlots = {}; // "day|period|teacher" → [{i, subject}]
  const roomSlots = {};    // "day|period|room"    → [{i, signature}]
  let splitSlots = 0;      // periods where a class is split into groups
  let combinedLectures = 0; // one teacher, one subject, several sections at once

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const day = String(row[C.DAY - 1]).trim();
    const period = String(row[C.PERIOD - 1]).trim();
    const cls = String(row[C.CLASS - 1]).trim();
    if (!day && !period && !cls) continue;

    if (hasClasses && cls) {
      ScheduleParser.splitClasses(cls).forEach(c => {
        if (!classSet[norm(c)]) { add(i, 'Unknown Class (' + c + ')'); count('Unknown class'); }
      });
    }

    const laneRow = { Teacher: row[C.TEACHER - 1], Subject: row[C.SUBJECT - 1], Room: row[C.ROOM - 1] };
    const parsed = ScheduleParser.parseLanes(laneRow);
    if (parsed.mismatch) { add(i, 'Lane Mismatch (subjects and teachers do not line up)'); count('Lane mismatch'); }
    if (parsed.roomMismatch) { add(i, 'Room Mismatch (rooms and subjects do not line up)'); count('Room mismatch'); }
    if (parsed.lanes.length > 1) splitSlots++;

    const assignments = ScheduleParser.parseRowAssignments(laneRow);
    assignments.forEach(a => {
      if (!a.teacher) return;
      const key = norm(a.teacher);
      if (hasRegistry && !registered[key]) { add(i, 'Unregistered Teacher (' + a.teacher + ')'); count('Unregistered teacher'); }
      if ((unavailable[key] || []).indexOf(normalizeDay_(day)) !== -1) { add(i, 'Unavailable Day (' + a.teacher + ')'); count('Unavailable day'); }
      const slot = day + '|' + period + '|' + key;
      (teacherSlots[slot] = teacherSlots[slot] || []).push({ i: i, subject: norm(a.subject), teacher: a.teacher, cls: cls });
    });

    // Rooms: one Room value = the whole slot uses that room; one room per lane = each
    // group uses its own room. A combined lecture (same subject + teachers) may share a room.
    const laneRooms = parsed.lanes.map(l => String(l.room || ''));
    const perLane = parsed.lanes.length > 1 && !parsed.roomMismatch && laneRooms.some(r => r !== laneRooms[0]);
    const usages = perLane
      ? parsed.lanes.map(l => ({ room: l.room, signature: norm(l.subject) + '|' + l.teachers.map(norm).sort().join('+') }))
      : [{ room: laneRow.Room, signature: norm(laneRow.Subject) + '|' + norm(laneRow.Teacher) }];
    usages.forEach(u => {
      if (!isRealRoom_(u.room)) return;
      const slot = day + '|' + period + '|' + norm(u.room);
      (roomSlots[slot] = roomSlots[slot] || []).push({ i: i, signature: u.signature });
    });
  }

  Object.keys(teacherSlots).forEach(k => {
    const list = teacherSlots[k];
    if (list.length < 2) return;
    const verdict = ScheduleParser.classifyTeacherSlot(list.map(x => ({ cls: x.cls, subject: x.subject })));
    if (verdict.clash) {
      count('Teacher clash');
      list.forEach(x => add(x.i, 'Teacher Clash (' + x.teacher + (verdict.reason === 'different subjects' ? '' : ': ' + verdict.reason) + ')'));
    } else {
      combinedLectures++;
    }
  });

  Object.keys(roomSlots).forEach(k => {
    const list = roomSlots[k];
    const sigs = {};
    list.forEach(x => { sigs[x.signature] = true; });
    if (list.length > 1 && Object.keys(sigs).length > 1) {
      count('Room clash');
      list.forEach(x => add(x.i, 'Room Clash'));
    }
  });

  // Overload (periods/week > max) — counts a combined lecture once
  const load = DataAccess.updateTeacherLoad();
  const overloaded = Object.keys(load)
    .filter(k => maxHours[k] && load[k] > maxHours[k])
    .map(k => (registered[k] || k) + ' (' + load[k] + '/' + maxHours[k] + ')');

  // Write Clash Status and highlight flagged rows
  const n = values.length - 1;
  const statuses = [];
  const backgrounds = [];
  let rowsFlagged = 0;
  for (let i = 1; i < values.length; i++) {
    const msg = issues[i].join(' / ');
    if (msg) rowsFlagged++;
    statuses.push([msg]);
    backgrounds.push([msg ? '#FDEDEC' : null]);
  }
  sheet.getRange(2, C.CLASH, n, 1).setValues(statuses);
  sheet.getRange(2, C.CLASH, n, 1).setBackgrounds(backgrounds);

  return { rowsFlagged: rowsFlagged, counts: counts, overloaded: overloaded, splitSlots: splitSlots, combinedLectures: combinedLectures };
}

function formatValidationSummary_(result) {
  const info = 'Allowed, not conflicts: ' + (result.combinedLectures || 0) + ' combined lecture(s) (one teacher, several sections), ' +
    (result.splitSlots || 0) + ' split period(s) (a class split into groups).';
  if (result.rowsFlagged === 0 && result.overloaded.length === 0) {
    return '✅ Validation complete — no conflicts found.\n\n' + info;
  }
  const lines = ['Validation complete — ' + result.rowsFlagged + ' row(s) flagged in the Clash Status column.', '', info, ''];
  Object.keys(result.counts).forEach(k => lines.push('• ' + k + ': ' + result.counts[k]));
  if (result.overloaded.length) {
    lines.push('', 'Over max periods/week:');
    result.overloaded.slice(0, 15).forEach(t => lines.push('• ' + t));
    if (result.overloaded.length > 15) lines.push('• … and ' + (result.overloaded.length - 15) + ' more');
  }
  return lines.join('\n');
}
