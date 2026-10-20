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
    const seenTeacher = {};
    lanes.forEach(lane => lane.teachers.forEach(t => {
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
        others.forEach(e => { if (combinedWith.indexOf(e.cls) === -1) combinedWith.push(e.cls); });
      }
    }));
    const lines = [];
    if (lanes.length > 1) {
      lines.push('Split class — ' + lanes.length + ' groups:');
      lanes.forEach(l => lines.push('• ' + (l.subject || '?') + ' — ' + (l.teachers.join(' + ') || 'no teacher') + (l.room ? ' · Room ' + l.room : '')));
    }
    if (combinedWith.length) lines.push('Taught together with ' + combinedWith.join(', '));
    clashes.forEach(c => lines.push('⚠ Clash: ' + c));
    return { lanes: lanes, split: lanes.length > 1, combinedWith: combinedWith, clashes: clashes, note: lines.join('\n') };
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
   * Adds (or updates) a teacher in a row, optionally with the subject they teach,
   * without disturbing the other lanes.
   *
   *  - nobody assigned yet            → this teacher owns the slot
   *  - teacher already there + subject → their lane's subject changes
   *  - new teacher + an existing lane's subject → co-teacher of that lane
   *  - new teacher + a new subject     → a new lane (the class is now split)
   *  - new teacher, no subject         → co-teacher when the slot has one lane
   *
   * Rows whose lanes don't line up (3 teachers / 2 subjects) are left unchanged.
   *
   * @returns {{teacher: string, subject: string, room: string, error?: string}}
   */
  assignTeacher: function(teacherStr, subjectStr, teacher, subject, roomStr) {
    subject = String(subject || '').trim();
    const parsed = this.parseLanes({ Teacher: teacherStr, Subject: subjectStr, Room: roomStr });
    const lanes = parsed.lanes;
    const done = () => {
      const out = this.formatLanes(lanes);
      return { teacher: out.teacher, subject: out.subject, room: this._keepRoom(roomStr, out.room, parsed) };
    };
    const unchanged = (error) => {
      const r = { teacher: String(teacherStr == null ? '' : teacherStr), subject: String(subjectStr == null ? '' : subjectStr), room: String(roomStr == null ? '' : roomStr) };
      if (error) r.error = error;
      return r;
    };
    if (parsed.mismatch) return unchanged();

    const anyTeacher = lanes.some(l => l.teachers.length);
    const idx = lanes.findIndex(l => l.teachers.some(t => this.sameName(t, teacher)));

    // Nobody (else) assigned yet → this teacher owns the slot.
    if (!anyTeacher || (idx !== -1 && lanes.length === 1 && lanes[0].teachers.length === 1)) {
      return { teacher: teacher, subject: subject || String(subjectStr == null ? '' : subjectStr), room: String(roomStr == null ? '' : roomStr) };
    }

    if (idx !== -1) {
      const lane = lanes[idx];
      if (!subject || this.sameName(lane.subject, subject)) return unchanged();
      const target = lanes.findIndex(l => this.sameName(l.subject, subject));
      if (lane.teachers.length === 1 && target === -1) {
        lane.subject = subject;
        return done();
      }
      lane.teachers = lane.teachers.filter(t => !this.sameName(t, teacher));
      if (target !== -1) lanes[target].teachers.push(teacher);
      else lanes.push({ subject: subject, teachers: [teacher], room: lane.room });
      if (lane.teachers.length === 0) lanes.splice(idx, 1);
      return done();
    }

    if (!subject) {
      if (lanes.length === 1) { lanes[0].teachers.push(teacher); return done(); }
      return unchanged('this period is split into ' + lanes.length + ' groups — type the subject too, e.g. "Class - Subject"');
    }
    const target = lanes.findIndex(l => this.sameName(l.subject, subject));
    if (target !== -1) lanes[target].teachers.push(teacher);
    else lanes.push({ subject: subject, teachers: [teacher], room: lanes.length === 1 ? lanes[0].room : '' });
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
