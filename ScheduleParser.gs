/**
 * ScheduleParser.gs
 * Utility for parsing multi-teacher and optional subject fields in Master_Schedule rows.
 */

const ScheduleParser = {

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
  }
};
