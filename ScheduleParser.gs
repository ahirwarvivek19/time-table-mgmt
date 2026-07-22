/**
 * ScheduleParser.gs
 * Utility for parsing multi-teacher and optional subject fields in Master_Schedule rows.
 */

const ScheduleParser = {

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
   * Splits a string by common delimiters (/, comma, newline, pipe).
   * @param {string} str
   * @returns {Array<string>}
   */
  splitList: function(str) {
    if (!str) return [];
    return String(str)
      .split(/[\/\n\|]/)
      .map(s => s.trim())
      .filter(s => s.length > 0);
  },

  /**
   * Parses a Master_Schedule row object into an array of { teacher, subject } assignments.
   * Handles optional subjects (e.g. Hindi / IP with Mrs. Farhana / Mr. Somesh)
   * as well as co-taught subjects (e.g. Science with Smt. Bharti / Sh. Rajesh).
   *
   * @param {Object} row  Object with keys 'Teacher' and 'Subject'
   * @returns {Array<{teacher: string, subject: string}>}
   */
  parseRowAssignments: function(row) {
    if (!row) return [];

    const rawTeacher = row.Teacher || '';
    const rawSubject = row.Subject || '';

    const teachers = this.splitList(rawTeacher);
    const subjects = this.splitList(rawSubject);

    if (teachers.length === 0) {
      return [{ teacher: '', subject: rawSubject }];
    }

    if (teachers.length === subjects.length) {
      return teachers.map((t, idx) => ({
        teacher: t,
        subject: subjects[idx]
      }));
    }

    if (subjects.length === 1 && teachers.length > 1) {
      return teachers.map(t => ({
        teacher: t,
        subject: subjects[0]
      }));
    }

    if (teachers.length === 1 && subjects.length > 1) {
      return [{
        teacher: teachers[0],
        subject: rawSubject
      }];
    }

    // Edge case: unequal counts > 1
    const minLen = Math.min(teachers.length, subjects.length);
    const result = [];
    for (let i = 0; i < minLen; i++) {
      result.push({ teacher: teachers[i], subject: subjects[i] });
    }
    for (let i = minLen; i < teachers.length; i++) {
      result.push({ teacher: teachers[i], subject: rawSubject });
    }
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
    return assignments.some(a => a.teacher.toLowerCase() === targetTeacher.toLowerCase());
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
    const match = assignments.find(a => a.teacher.toLowerCase() === targetTeacher.toLowerCase());
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
    const targets = this.splitClasses(targetClass).map(t => t.toLowerCase());
    return classes.some(c => targets.includes(c.toLowerCase()));
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
      if (s.subject) subjects.add(s.subject.trim());
    });

    const uniqueClasses = [...new Set(allClassNames)];
    const subjectList = [...subjects];

    if (subjectList.length <= 1) {
      // Combined optional/joint class with same subject across multiple sections
      const subjText = subjectList.length === 1 ? ' - ' + subjectList[0] : '';
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
      display: '⚠️ CLASH: ' + clashDesc,
      isFree: false,
      isCombined: false,
      isClash: true
    };
  }
};
