/**
 * ScheduleParser.gs
 * Utility for parsing multi-teacher and optional subject fields in Master_Schedule rows.
 */

const ScheduleParser = {

  /** Lower-cases and collapses whitespace so "Neha  Jhala " matches "neha jhala". */
  normalizeName: function(name) {
    return String(name == null ? '' : name).replace(/\s+/g, ' ').trim().toLowerCase();
  },

  /** Case- and whitespace-insensitive name comparison. */
  sameName: function(a, b) {
    const na = this.normalizeName(a);
    return na !== '' && na === this.normalizeName(b);
  },

  /**
   * Capitalizes a teacher name string into proper Title Case / Camel Case.
   * Handles honorifics (Mrs., Smt., Sh., Mr., Dr.), room tags (Rm 206),
   * and multi-teacher slash/comma/ampersand separated strings.
   * @param {string} name
   * @returns {string}
   */
  formatTeacherName: function(name) {
    if (!name) return '';
    return String(name).split(/(\s*[\/\&\,]\s*)/).map(part => {
      if (['/', '&', ','].includes(part.trim())) return part;
      return part.replace(/\b[a-zA-Z0-9\.]+\b/g, word => {
        const raw = word.replace(/\.$/, '');
        const upper = raw.toUpperCase();
        if (['NSS', 'ICT', 'UHV', 'SUPW'].includes(upper)) {
          return word.endsWith('.') ? upper + '.' : upper;
        }
        if (upper === 'RM') {
          return word.endsWith('.') ? 'Rm.' : 'Rm';
        }
        const capitalized = raw.charAt(0).toUpperCase() + raw.slice(1).toLowerCase();
        return word.endsWith('.') ? capitalized + '.' : capitalized;
      });
    }).join('');
  },

  /**
   * Splits a teachers/subjects string by the canonical delimiter "/" or fallback ",", newline, pipe.
   * Canonical write format is always "/" (e.g. "Mrs. X / Mr. Y" or "Hindi / IP").
   * Comma is accepted on input for robustness but is NOT used when writing back.
   * @param {string} str
   * @returns {Array<string>}
   */
  splitList: function(str) {
    if (!str) return [];
    return String(str)
      .split(/[\/,\n\|]/)
      .map(s => s.trim())
      .filter(s => s.length > 0);
  },

  /** A lane whose teacher (or room) is unknown is written as "-" so the lanes stay aligned. */
  EMPTY_LANE: '-',

  _isEmptyMark: function(s) {
    return ['', '-', '—', '–', '?'].indexOf(String(s == null ? '' : s).trim()) !== -1;
  },

  /** "Dinesh Bhati + Lokit Chawda" → ["Dinesh Bhati", "Lokit Chawda"]. "-" → []. */
  splitCoTeachers: function(lane) {
    if (this._isEmptyMark(lane)) return [];
    return String(lane).split(/\s*\+\s*|\s+&\s+/).map(s => s.trim()).filter(s => s && !this._isEmptyMark(s));
  },

  /**
   * Splits a Master_Schedule row into LANES. A lane is one group of students in
   * the slot: its subject, the teacher(s) who teach it and its room.
   *
   *   Subject  "Hindi / IT / Beauty"                     → 3 lanes (split class)
   *   Teacher  "Ekta Gour / Divya Lodval + Reena Shivhare / Krishna Patale"
   *            "/" separates lanes, "+" separates co-teachers inside one lane
   *   Room     "12 / Lab / -"  one room per lane, or one value for the whole slot
   *
   * Legacy rows (written before lanes existed) keep their meaning:
   *   one subject + several "/" teachers  → one co-taught lane;
   *   one teacher + several subjects       → that teacher teaches the whole slot.
   *
   * @param {Object} row  {Subject, Teacher, Room}
   * @returns {{lanes: Array<{subject: string, teachers: string[], room: string}>,
   *            mismatch: boolean, roomMismatch: boolean}}
   */
  parseLanes: function(row) {
    row = row || {};
    const rawSubject = String(row.Subject == null ? '' : row.Subject).trim();
    const subjects = this.splitList(rawSubject).map(s => (this._isEmptyMark(s) ? '' : s));
    const teacherLanes = this.splitList(row.Teacher).map(l => this.splitCoTeachers(l));
    const rawRoom = String(row.Room == null ? '' : row.Room).trim();
    // "N/A" means no room — it must not be read as two lanes "N" and "A"
    const rooms = /^(n\s*\/\s*a|na|none)$/i.test(rawRoom) ? [] : this.splitList(rawRoom).map(r => (this._isEmptyMark(r) ? '' : r));
    let lanes;
    let mismatch = false;

    if (teacherLanes.length === 0) {
      lanes = subjects.length > 1 ? subjects.map(s => ({ subject: s, teachers: [] })) : [{ subject: rawSubject, teachers: [] }];
    } else if (teacherLanes.length === subjects.length) {
      lanes = subjects.map((s, i) => ({ subject: s, teachers: teacherLanes[i] }));
    } else if (subjects.length <= 1) {
      // Legacy co-teaching: "Science" with "Mrs A / Mr B"
      lanes = [{ subject: subjects[0] || '', teachers: [].concat.apply([], teacherLanes) }];
    } else if (teacherLanes.length === 1) {
      // Legacy: one teacher for a multi-subject cell keeps the whole subject text
      lanes = [{ subject: rawSubject, teachers: teacherLanes[0] }];
    } else {
      // 3 subjects / 2 teacher lanes: pairing is unknown. Pair what we can and flag it.
      mismatch = true;
      const n = Math.max(subjects.length, teacherLanes.length);
      lanes = [];
      for (let i = 0; i < n; i++) lanes.push({ subject: i < subjects.length ? subjects[i] : rawSubject, teachers: teacherLanes[i] || [] });
    }

    let roomMismatch = false;
    lanes.forEach((lane, i) => {
      if (rooms.length === 0) lane.room = '';
      else if (rooms.length === 1) lane.room = rooms[0];
      else if (rooms.length === lanes.length) lane.room = rooms[i];
      else { roomMismatch = true; lane.room = String(row.Room).trim(); }
    });
    return { lanes: lanes, mismatch: mismatch, roomMismatch: roomMismatch };
  },

  /**
   * Inverse of parseLanes. A single lane keeps the legacy " / " co-teacher form
   * so existing sheets read exactly as before; several lanes use " + " inside a lane.
   * @returns {{subject: string, teacher: string, room: string}}
   */
  formatLanes: function(lanes) {
    lanes = (lanes || []).filter(l => l && (String(l.subject || '').trim() || (l.teachers || []).length || String(l.room || '').trim()));
    if (lanes.length === 0) return { subject: '', teacher: '', room: '' };
    const E = this.EMPTY_LANE;
    if (lanes.length === 1) {
      return { subject: String(lanes[0].subject || ''), teacher: (lanes[0].teachers || []).join(' / '), room: String(lanes[0].room || '') };
    }
    const anyTeacher = lanes.some(l => (l.teachers || []).length);
    const rooms = lanes.map(l => String(l.room || '').trim());
    const sameRoom = rooms.every(r => r === rooms[0]);
    return {
      subject: lanes.map(l => String(l.subject || '').trim() || E).join(' / '),
      teacher: anyTeacher ? lanes.map(l => ((l.teachers || []).length ? l.teachers.join(' + ') : E)).join(' / ') : '',
      room: sameRoom ? rooms[0] : rooms.map(r => r || E).join(' / ')
    };
  },

  /** True when the slot has more than one lane (a split class). */
  isSplit: function(row) {
    return this.parseLanes(row).lanes.length > 1;
  },

  /**
   * Parses a Master_Schedule row object into an array of { teacher, subject } assignments
   * (one per teacher per lane). Co-teachers of a lane each get that lane's subject.
   *
   * @param {Object} row  Object with keys 'Teacher', 'Subject' and optionally 'Room'
   * @returns {Array<{teacher: string, subject: string, room: string, lane: number}>}
   */
  parseRowAssignments: function(row) {
    if (!row) return [];
    const parsed = this.parseLanes(row);
    const result = [];
    parsed.lanes.forEach((lane, idx) => {
      lane.teachers.forEach(t => result.push({ teacher: t, subject: lane.subject, room: lane.room, lane: idx }));
    });
    if (result.length === 0) return [{ teacher: '', subject: String(row.Subject || ''), room: String(row.Room || ''), lane: 0 }];
    return result;
  },

  /**
   * Checks if a row includes a target teacher (exact or parsed in multi-teacher list).
   * @param {Object} row
   * @param {string} targetTeacher
   * @returns {boolean}
   */
  rowIncludesTeacher: function(row, targetTeacher) {
    if (!row || !targetTeacher) return false;
    const assignments = this.parseRowAssignments(row);
    return assignments.some(a => this.sameName(a.teacher, targetTeacher));
  },

  /**
   * Gets the specific subject assigned to a target teacher in a row.
   * Returns empty string if teacher is not in this row.
   * @param {Object} row
   * @param {string} targetTeacher
   * @returns {string}
   */
  getSubjectForTeacher: function(row, targetTeacher) {
    if (!row || !targetTeacher) return '';
    const assignments = this.parseRowAssignments(row);
    const match = assignments.find(a => this.sameName(a.teacher, targetTeacher));
    return match ? match.subject : '';
  },

  /**
   * Splits a class string by common delimiters (comma, plus, slash, pipe, newline).
   * @param {string} str  e.g. "11th A, 11th B" or "11th A + 11th B"
   * @returns {Array<string>}
   */
  splitClasses: function(str) {
    if (!str) return [];
    return String(str)
      .split(/[\,\+\/\n\|]/)
      .map(s => s.trim())
      .filter(s => s.length > 0);
  },

  /**
   * Gets list of classes represented in a row.
   * @param {Object} row
   * @returns {Array<string>}
   */
  parseRowClasses: function(row) {
    if (!row) return [];
    return this.splitClasses(row.Class || '');
  },

  /**
   * Checks if a row includes a specific class (or one of comma/plus-separated classes).
   * @param {Object} row
   * @param {string} targetClass
   * @returns {boolean}
   */
  rowIncludesClass: function(row, targetClass) {
    if (!row || !targetClass) return false;
    if (targetClass === 'All Classes' || targetClass === 'All') return true;
    const classes = this.parseRowClasses(row);
    const targets = this.splitClasses(targetClass).map(t => this.normalizeName(t));
    return classes.some(c => targets.includes(this.normalizeName(c)));
  },

  /** "11th C (E)" → 11, "6th A" → 6; null when the class name has no grade number. */
  classGrade: function(cls) {
    const m = String(cls || '').match(/^\s*(\d{1,2})/);
    return m ? Number(m[1]) : null;
  },

  /** True for whole-school activities (Sports, Library, SUPW…) that may combine grades. */
  isCrossGradeSubject: function(subject) {
    const s = this.normalizeName(subject);
    const list = (typeof CONFIG !== 'undefined' && CONFIG.CROSS_GRADE_SUBJECTS) || [];
    return list.some(p => s.indexOf(this.normalizeName(p)) === 0);
  },

  /**
   * Merges the rows of one slot into lanes, keeping subject ↔ teacher ↔ room aligned.
   * Identical lanes (a combined lecture shown for several selected classes) appear once.
   * @param {Array<Object>} rows
   * @returns {Array<{subject: string, teachers: string[], room: string}>}
   */
  mergeLanes: function(rows) {
    const out = [];
    (rows || []).forEach(r => {
      this.parseLanes(r).lanes.forEach(l => {
        if (!String(l.subject || '').trim() && !l.teachers.length) return;
        const same = out.some(o => this.sameName(o.subject, l.subject) &&
          o.teachers.length === l.teachers.length && o.teachers.every((t, i) => this.sameName(t, l.teachers[i])));
        if (!same) out.push({ subject: l.subject, teachers: l.teachers.slice(), room: l.room || '' });
      });
    });
    return out;
  },

  /**
   * What a class timetable cell needs to show beyond its text: is the class split into
   * groups, which other sections share the lecture, and which teachers are double-booked.
   * @param {Array<Object>} slotRows  every Master_Schedule row for this day + period
   * @param {string} className        the class (or "11th A, 11th B") being viewed
   * @returns {{lanes: Array, split: boolean, combinedWith: string[], clashes: string[], note: string}}
   */
  describeSlot: function(slotRows, className) {
    const own = (slotRows || []).filter(r => this.rowIncludesClass(r, className));
    const ownClasses = this.splitClasses(className).map(c => this.normalizeName(c));
    const isOwn = c => ownClasses.indexOf(this.normalizeName(c)) !== -1;
    const lanes = this.mergeLanes(own);
    const combinedWith = [], clashes = [];
    // Which other sections share each group. Kept per group, because in a split period
    // the English group may be shared with 11th B while the IT group is shared with 11th C.
    const shared = lanes.map(() => []);
    const seenTeacher = {};
    lanes.forEach((lane, laneIdx) => lane.teachers.forEach(t => {
      const key = this.normalizeName(t);
      if (seenTeacher[key]) return;
      seenTeacher[key] = true;
      const entries = [];
      (slotRows || []).forEach(r => this.parseLanes(r).lanes.forEach(l => {
        if (l.teachers.some(x => this.sameName(x, t))) {
          this.parseRowClasses(r).forEach(c => entries.push({ cls: c, subject: l.subject }));
        }
      }));
      const others = entries.filter(e => !isOwn(e.cls));
      if (!others.length) return;
      if (this.classifyTeacherSlot(entries).clash) {
        clashes.push(t + ' also has ' + others.map(e => e.cls + ' (' + e.subject + ')').join(', '));
      } else {
        others.forEach(e => {
          if (combinedWith.indexOf(e.cls) === -1) combinedWith.push(e.cls);
          if (shared[laneIdx].indexOf(e.cls) === -1) shared[laneIdx].push(e.cls);
        });
      }
    }));
    const order = c => this._classOrderKey(c);
    combinedWith.sort((a, b) => order(a).localeCompare(order(b)));
    shared.forEach(list => list.sort((a, b) => order(a).localeCompare(order(b))));
    const lines = [];
    if (lanes.length > 1) {
      lines.push('Split class — ' + lanes.length + ' groups:');
      lanes.forEach(l => lines.push('• ' + (l.subject || '?') + ' — ' + (l.teachers.join(' + ') || 'no teacher') + (l.room ? ' · Room ' + l.room : '')));
    }
    const sharedLanes = lanes.map((l, i) => ({ subject: l.subject, teachers: l.teachers, classes: shared[i] })).filter(s => s.classes.length);
    if (sharedLanes.length) {
      lines.push('Taught together:');
      sharedLanes.forEach(s => lines.push('• ' + (s.subject || '?') + ' — ' + (s.teachers.join(' + ') || 'no teacher') + ': with ' + s.classes.join(', ')));
    }
    clashes.forEach(c => lines.push('⚠ Clash: ' + c));
    return { lanes: lanes, split: lanes.length > 1, combinedWith: combinedWith, sharedLanes: sharedLanes, clashes: clashes, note: lines.join('\n') };
  },

  /**
   * Multi-select with a normal dropdown: on a split period, a teacher picked from the
   * dropdown fills the NEXT group that has no teacher yet, instead of replacing the cell.
   *   old "Asha / -"  + pick "Bina", 2 groups  → "Asha / Bina"
   *   old ""          + pick "Asha", 3 groups  → "Asha / - / -"
   *   old "Asha / Bina", 2 groups (all filled) → error: clear the cell to start again
   * @param {string} oldTeacherStr  Teacher cell before the pick
   * @param {string} picked         the one name picked
   * @param {number} groups         number of subjects (groups) in the period
   * @returns {{teacher: string, error?: string}}
   */
  fillNextGroup: function(oldTeacherStr, picked, groups) {
    const name = String(picked == null ? '' : picked).trim();
    const lanes = String(oldTeacherStr == null ? '' : oldTeacherStr).split('/').map(s => s.trim());
    const cells = lanes.length === 1 && !lanes[0] ? [] : lanes;
    if (cells.length > groups) return { teacher: String(oldTeacherStr), error: 'more teacher groups than subjects — fix the teacher cell by typing it in full' };
    if (cells.some(l => this.splitCoTeachers(l).some(t => this.sameName(t, name)))) {
      return { teacher: String(oldTeacherStr) }; // already there: picking again changes nothing
    }
    while (cells.length < groups) cells.push(this.EMPTY_LANE);
    const free = cells.findIndex(l => this._isEmptyMark(l));
    if (free === -1) {
      return { teacher: String(oldTeacherStr), error: 'all ' + groups + ' groups already have a teacher — clear the teacher cell and pick them again in subject order, or type the list' };
    }
    cells[free] = name;
    return { teacher: cells.join(' / ') };
  },

  /** Sort key that puts "6th A" before "10th A" (grade number, then section). */
  _classOrderKey: function(cls) {
    const s = String(cls == null ? '' : cls).trim();
    const m = s.match(/^(\d+)/);
    return m ? ('000' + m[1]).slice(-3) + s.slice(m[1].length).toLowerCase() : 'zzz' + s.toLowerCase();
  },

  /**
   * Checks a typed Subject + Teacher pair before it is saved, so subject i still belongs
   * to teacher group i. Returns '' when it can be saved, otherwise the reason.
   *   "Hindi / IT"  + "Asha / Bina + Chetan"  → ok (2 groups, 2 teacher groups)
   *   "Hindi / IT"  + "Deepa"                 → rejected: who teaches which group?
   *   "Hindi / IT / Art" + "Asha / Bina"      → rejected: 3 groups, 2 teacher groups
   *   "Science"     + "Asha / Bina"           → ok (two teachers share one class)
   */
  checkLaneText: function(subjectStr, teacherStr) {
    const subjects = this.splitList(subjectStr);
    const teacherLanes = this.splitList(teacherStr);
    if (subjects.length > 1 && teacherLanes.length && teacherLanes.length !== subjects.length) {
      return 'this period has ' + subjects.length + ' groups (' + subjects.join(' / ') + ') but ' + teacherLanes.length +
        ' teacher group' + (teacherLanes.length === 1 ? '' : 's') + '. Type one teacher group per subject in the same order, ' +
        'separated by " / " (use " + " for two teachers sharing a group), e.g. "' +
        subjects.map((s, i) => 'Teacher ' + (i + 1)).join(' / ') + '"';
    }
    return '';
  },

  /**
   * Separates a medium mark from a subject: "Physics (H.M)" → {base: "Physics", medium: "H"}.
   * The older marks "(H)", "(E)", "(Hindi)", "(English)" are read too. No mark → medium "".
   */
  splitMedium: function(subject) {
    const s = String(subject == null ? '' : subject).trim();
    const m = s.match(/^(.*?)\s*\(\s*(?:(H|E)\s*\.?\s*M\s*\.?|(H|E|Hindi|English))\s*\)$/i);
    if (!m) return { base: s, medium: '' };
    return { base: m[1].trim(), medium: (m[2] || m[3]).charAt(0).toUpperCase() };
  },

  /**
   * Decides whether one teacher's simultaneous assignments are one lecture.
   *   same subject, same grade             → combined lecture (allowed)
   *   same subject, whole-school activity  → combined lecture (allowed)
   *   different subjects                   → clash
   *   same subject but different grades    → clash ("11th D Maths" and "12th D Maths" at once)
   *   same subject, Hindi- and English-medium groups at once → clash
   * A medium mark on one section only ("Hindi" in 11th B, "Hindi (E.M)" in 11th C) is
   * the same lecture: the English-medium group of 11th C joins 11th B.
   * @param {Array<{cls: string, subject: string}>} slots  one teacher, one day/period
   * @returns {{clash: boolean, combined: boolean, reason: string}}
   */
  classifyTeacherSlot: function(slots) {
    const list = (slots || []).filter(Boolean);
    const classes = [...new Set([].concat.apply([], list.map(s => this.splitClasses(s.cls))).map(c => this.normalizeName(c)))];
    if (list.length <= 1 && classes.length <= 1) return { clash: false, combined: false, reason: '' };
    const parts = list.map(s => this.splitMedium(s.subject));
    const subjects = [...new Set(parts.map(p => this.normalizeName(p.base)))];
    if (subjects.length > 1) return { clash: true, combined: false, reason: 'different subjects' };
    const media = [...new Set(parts.map(p => p.medium).filter(Boolean))];
    if (media.length > 1) return { clash: true, combined: false, reason: 'Hindi- and English-medium groups at once' };
    const grades = [...new Set(classes.map(c => this.classGrade(c)).filter(g => g !== null))];
    if (grades.length > 1 && !this.isCrossGradeSubject(list[0].subject)) {
      return { clash: true, combined: false, reason: 'same subject in different grades' };
    }
    return { clash: false, combined: classes.length > 1, reason: '' };
  },

  /**
   * Groups multiple class assignments for a teacher in a single period slot.
   * Merges simultaneous lectures for the same subject into a combined class display (e.g. 11th A, 11th B - IP).
   * Flags actual clashes only when different subjects are taught at the same time.
   *
   * @param {Array<{cls: string, subject: string}>} slots
   * @returns {{display: string, isFree: boolean, isCombined: boolean, isClash: boolean}}
   */
  groupTeacherSlots: function(slots) {
    if (!slots || slots.length === 0) {
      return { display: 'FREE', isFree: true, isCombined: false, isClash: false };
    }

    if (slots.length === 1) {
      const s = slots[0];
      const clsList = this.splitClasses(s.cls);
      const isCombined = clsList.length > 1;
      const displayCls = isCombined ? clsList.join(', ') : s.cls;
      return {
        display: displayCls + (s.subject ? ' - ' + s.subject : ''),
        isFree: false,
        isCombined: isCombined,
        isClash: false
      };
    }

    // Multiple slot rows for the same teacher & period
    // Collect all class names and check if subjects are identical
    const allClassNames = [];
    const subjects = new Set();

    slots.forEach(s => {
      const parsedCls = this.splitClasses(s.cls);
      allClassNames.push(...parsedCls);
      if (s.subject) subjects.add(s.subject.trim().toLowerCase());
    });

    const uniqueClasses = [...new Set(allClassNames)];

    if (!this.classifyTeacherSlot(slots).clash) {
      // Combined optional/joint class with same subject across multiple sections
      const firstSubject = (slots.find(s => s.subject) || {}).subject || '';
      const subjText = firstSubject ? ' - ' + firstSubject.trim() : '';
      return {
        display: uniqueClasses.join(', ') + subjText,
        isFree: false,
        isCombined: true,
        isClash: false
      };
    }

    // Different subjects double-booked -> True Clash
    const clashDesc = slots.map(s => s.cls + ' (' + s.subject + ')').join(' / ');
    return {
      display: clashDesc,
      isFree: false,
      isCombined: false,
      isClash: true
    };
  },

  /**
   * Removes one teacher from a row. Their lane goes with them when other lanes
   * remain ("Hindi / IT" ↔ "A / B" minus A → "IT" ↔ "B"); a co-teacher leaves the
   * lane to the remaining teacher(s).
   *
   * @param {string} teacherStr  Current Teacher cell
   * @param {string} subjectStr  Current Subject cell
   * @param {string} teacher     Teacher to remove
   * @param {string} [roomStr]   Current Room cell
   * @returns {{teacher: string, subject: string, room: string, removed: boolean}}
   */
  removeTeacher: function(teacherStr, subjectStr, teacher, roomStr) {
    const parsed = this.parseLanes({ Teacher: teacherStr, Subject: subjectStr, Room: roomStr });
    const unchanged = { teacher: String(teacherStr == null ? '' : teacherStr), subject: String(subjectStr == null ? '' : subjectStr), room: String(roomStr == null ? '' : roomStr), removed: false };
    const idx = parsed.lanes.findIndex(l => l.teachers.some(t => this.sameName(t, teacher)));
    if (idx === -1 || parsed.mismatch) return unchanged;

    const lane = parsed.lanes[idx];
    lane.teachers = lane.teachers.filter(t => !this.sameName(t, teacher));
    if (lane.teachers.length === 0 && parsed.lanes.length > 1) parsed.lanes.splice(idx, 1);
    const out = this.formatLanes(parsed.lanes);
    if (parsed.lanes.length === 1 && !parsed.lanes[0].teachers.length) out.subject = String(parsed.lanes[0].subject || '');
    return { teacher: out.teacher, subject: out.subject, room: this._keepRoom(roomStr, out.room, parsed), removed: true };
  },

  /**
   * Puts a teacher into a row from a teacher's own view (Teacher_View / Teacher_Day_View).
   * The teacher TAKES the period: whoever taught that group before is replaced and
   * listed in `replaced`, so the edit never silently turns into a co-teacher or a new
   * split group. Co-teachers ("A + B") and new split groups are set up in Class_View.
   *
   *  - nobody assigned yet, one group        → this teacher owns the slot (typed subject wins)
   *  - teacher already there + new subject   → their own group's subject changes
   *  - teacher already there + another group's subject → they move to that group, replacing its teacher
   *  - new teacher, one group                → replaces that group's teacher(s); a typed subject replaces it too
   *  - new teacher, split period             → takes the group whose subject they typed;
   *                                            no subject / unknown subject is refused
   *
   * Rows whose lanes don't line up (3 teachers / 2 subjects) are left unchanged.
   *
   * @returns {{teacher: string, subject: string, room: string, replaced: Array<{teacher: string, subject: string}>, error?: string}}
   */
  assignTeacher: function(teacherStr, subjectStr, teacher, subject, roomStr) {
    subject = String(subject || '').trim();
    const parsed = this.parseLanes({ Teacher: teacherStr, Subject: subjectStr, Room: roomStr });
    const lanes = parsed.lanes;
    const replaced = [];
    const done = () => {
      const out = this.formatLanes(lanes);
      return { teacher: out.teacher, subject: out.subject, room: this._keepRoom(roomStr, out.room, parsed), replaced: replaced };
    };
    const unchanged = (error) => {
      const r = { teacher: String(teacherStr == null ? '' : teacherStr), subject: String(subjectStr == null ? '' : subjectStr), room: String(roomStr == null ? '' : roomStr), replaced: [] };
      if (error) r.error = error;
      return r;
    };
    const take = lane => {
      lane.teachers.filter(t => !this.sameName(t, teacher)).forEach(t => replaced.push({ teacher: t, subject: lane.subject }));
      lane.teachers = [teacher];
    };
    if (parsed.mismatch) return unchanged();
    if (this.splitList(subject).length > 1) {
      return unchanged('type one subject here — a period is split into groups in Class_View');
    }

    const idx = lanes.findIndex(l => l.teachers.some(t => this.sameName(t, teacher)));
    const target = subject ? lanes.findIndex(l => this.sameName(l.subject, subject)) : -1;
    const groups = () => lanes.map(l => l.subject || '?').join(' / ');

    if (idx !== -1) {
      const lane = lanes[idx];
      if (!subject || this.sameName(lane.subject, subject)) return unchanged();
      if (target === -1) {
        if (lane.teachers.length === 1) { lane.subject = subject; return done(); }
        // A co-teacher leaving to start a group of their own
        lane.teachers = lane.teachers.filter(t => !this.sameName(t, teacher));
        lanes.push({ subject: subject, teachers: [teacher], room: lane.room });
        return done();
      }
      lane.teachers = lane.teachers.filter(t => !this.sameName(t, teacher));
      take(lanes[target]);
      if (lane.teachers.length === 0) lanes.splice(idx, 1);
      return done();
    }

    if (lanes.length <= 1) {
      const lane = lanes[0] || { subject: '', teachers: [], room: '' };
      if (!lanes.length) lanes.push(lane);
      take(lane);
      if (subject) lane.subject = subject;
      return done();
    }
    if (!subject) {
      return unchanged('this period is split into ' + lanes.length + ' groups (' + groups() + ') — type the subject too, e.g. "Class - ' + lanes[0].subject + '"');
    }
    if (target === -1) {
      return unchanged('this period is split into ' + lanes.length + ' groups (' + groups() + ') — type one of those subjects to take that group, or change the groups in Class_View');
    }
    take(lanes[target]);
    return done();
  },

  /** Keeps the user's original Room text when the edit did not change the lanes' rooms. */
  _keepRoom: function(roomStr, formatted, parsed) {
    const original = String(roomStr == null ? '' : roomStr);
    if (parsed.roomMismatch) return original;
    const rooms = this.splitList(original);
    return rooms.length <= 1 ? original : formatted;
  }
};
