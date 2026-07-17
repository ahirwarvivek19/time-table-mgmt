const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

// ─────────────────────────────────────────────
// PART 1 : Middle School (6th – 8th) — hardcoded
// ─────────────────────────────────────────────

const MIDDLE_CLASSES = ['6th A', '6th B', '7th A', '7th B', '8th A', '8th B'];
const MIDDLE_TEACHERS = [
  "Mrs. Ruchika Vyas",
  "Mrs. Farhana Khan",
  "Mrs. Lalita Tiwari",
  "Mr. Somesh Vyas",
  "Mrs. Roshi Jahan",
  "Mrs. Geeta Chandaniya",
  "Mrs. Premlata Vishwakarma",
  "Mr. Ashok Indriya",
  "Mrs. Neha Jhala",
  "Dr. Mahesh Sharma"
];

const middleDataMap = {
  "6th A": [
    { period: 1, subject: "English",             teacher: "Mrs. Ruchika Vyas" },
    { period: 2, subject: "Hindi",               teacher: "Mrs. Farhana Khan" },
    { period: 3, subject: "Social Science",      teacher: "Mrs. Lalita Tiwari" },
    { period: 4, subject: "Proficiency English", teacher: "Mrs. Ruchika Vyas" },
    { period: 5, subject: "Science",             teacher: "Mr. Somesh Vyas" },
    { period: 6, subject: "Mathematics",         teacher: "Mrs. Roshi Jahan" },
    { period: 7, subject: "Library / Sports",    teacher: "" },
    { period: 8, subject: "Sanskrit",            teacher: "Mrs. Lalita Tiwari" }
  ],
  "6th B": [
    { period: 1, subject: "Mathematics",         teacher: "Mrs. Roshi Jahan" },
    { period: 2, subject: "Sanskrit",            teacher: "Mrs. Lalita Tiwari" },
    { period: 3, subject: "Science",             teacher: "Mr. Somesh Vyas" },
    { period: 4, subject: "Proficiency Hindi",   teacher: "Mrs. Geeta Chandaniya" },
    { period: 5, subject: "Foundation",          teacher: "Mrs. Ruchika Vyas" },
    { period: 6, subject: "Hindi",               teacher: "Mrs. Farhana Khan" },
    { period: 7, subject: "Social Science",      teacher: "Mrs. Premlata Vishwakarma" },
    { period: 8, subject: "English",             teacher: "Mrs. Ruchika Vyas" }
  ],
  "7th A": [
    { period: 1, subject: "Science",             teacher: "Mr. Somesh Vyas" },
    { period: 2, subject: "English",             teacher: "Mr. Ashok Indriya" },
    { period: 3, subject: "Hindi",               teacher: "Mrs. Farhana Khan" },
    { period: 4, subject: "Proficiency Hindi",   teacher: "Mrs. Farhana Khan" },
    { period: 5, subject: "Sanskrit",            teacher: "Mrs. Geeta Chandaniya" },
    { period: 6, subject: "Social Science",      teacher: "Mrs. Premlata Vishwakarma" },
    { period: 7, subject: "Mathematics",         teacher: "Mrs. Roshi Jahan" },
    { period: 8, subject: "Library / Sports",    teacher: "" }
  ],
  "7th B": [
    { period: 1, subject: "Social Science",      teacher: "Mrs. Premlata Vishwakarma" },
    { period: 2, subject: "Foundation English",  teacher: "" },
    { period: 3, subject: "English",             teacher: "Mr. Ashok Indriya" },
    { period: 4, subject: "Proficiency English", teacher: "Mr. Ashok Indriya" },
    { period: 5, subject: "Library / Sports",    teacher: "" },
    { period: 6, subject: "Hindi",               teacher: "Mrs. Geeta Chandaniya" },
    { period: 7, subject: "Sanskrit",            teacher: "Mrs. Lalita Tiwari" },
    { period: 8, subject: "Mathematics",         teacher: "Mrs. Roshi Jahan" }
  ],
  "8th A": [
    { period: 1, subject: "Sanskrit",            teacher: "Mrs. Geeta Chandaniya" },
    { period: 2, subject: "English",             teacher: "Mrs. Ruchika Vyas" },
    { period: 3, subject: "Proficiency",         teacher: "Mrs. Roshi Jahan" },
    { period: 4, subject: "Hindi",               teacher: "Mrs. Farhana Khan" },
    { period: 5, subject: "Social Science",      teacher: "Mrs. Premlata Vishwakarma" },
    { period: 6, subject: "Library / Sports",    teacher: "" },
    { period: 7, subject: "Mathematics",         teacher: "Mrs. Neha Jhala" },
    { period: 8, subject: "Science",             teacher: "Mr. Somesh Vyas" }
  ],
  "8th B": [
    { period: 1, subject: "English",             teacher: "Mr. Ashok Indriya" },
    { period: 2, subject: "Science",             teacher: "Mr. Somesh Vyas" },
    { period: 3, subject: "Mathematics",         teacher: "Mrs. Neha Jhala" },
    { period: 4, subject: "Proficiency Math",    teacher: "Mrs. Neha Jhala" },
    { period: 5, subject: "Hindi",               teacher: "Mrs. Farhana Khan" },
    { period: 6, subject: "Social Science",      teacher: "Mrs. Premlata Vishwakarma" },
    { period: 7, subject: "Sanskrit",            teacher: "Dr. Mahesh Sharma" },
    { period: 8, subject: "Library / Sports",    teacher: "" }
  ]
};

const DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

let middleSchedule = [];
MIDDLE_CLASSES.forEach(className => {
  middleDataMap[className].forEach(p => {
    DAYS.forEach(day => {
      middleSchedule.push({
        Day: day, Period: p.period, Class: className,
        Tier: "Middle", Subject: p.subject, Teacher: p.teacher
      });
    });
  });
});

const middleClassRows = MIDDLE_CLASSES.map(c => [c, 'Middle', '']);
const middleTeacherRows = MIDDLE_TEACHERS.map(t => [t, '', 40, '', '']);

// ─────────────────────────────────────────────
// PART 2 : Secondary / Higher Secondary (9th – 12th) — parsed from Excel
// ─────────────────────────────────────────────

const wb = XLSX.readFile(path.join(__dirname, 'Class_DayWise_Timetable.xlsx'));
const classSheets = wb.SheetNames.filter(s => s !== 'INDEX');

// Period column indices: cols 2,3,4,5 = Periods 1-4; col 6 = BREAK; cols 7,8,9,10 = Periods 5-8
const PERIOD_COLS = [2, 3, 4, 5, 7, 8, 9, 10];

const upperClasses = [];   // { name, tier, room }
const upperTeacherSet = new Set();
let upperSchedule = [];

// Track seen class names to de-duplicate (11th C Hindi + English Medium)
const classNameCount = {};

classSheets.forEach(sheetName => {
  const ws = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

  if (rows.length < 5) return;

  // Row index 1: "Class: 9th A | ... | Room: 207 | ..."
  const classInfoRaw = String(rows[1][0] || '');
  const classMatch  = classInfoRaw.match(/Class:\s*([^|]+)/);
  let className = classMatch ? classMatch[1].trim() : sheetName.split('(')[0].trim();

  // De-duplicate 11th C (Hindi) vs 11th C (English) by appending suffix
  classNameCount[className] = (classNameCount[className] || 0) + 1;
  if (classNameCount[className] > 1) {
    className = `${className} (${classNameCount[className]})`;
  }

  // Extract room number only
  const roomMatch = classInfoRaw.match(/Room:\s*([^\s|,]+)/);
  const room = roomMatch ? roomMatch[1].trim() : '';

  const num = parseInt(className);
  const tier = (num >= 11) ? 'Higher Secondary' : 'Secondary';

  upperClasses.push({ name: className, tier, room });

  // Data rows: row index 4 = Monday, 5 = Tuesday, ..., 9 = Saturday
  for (let r = 4; r < rows.length && r <= 9; r++) {
    const row = rows[r];
    const dayName = String(row[0] || '').trim();
    if (!DAYS.includes(dayName)) continue;

    PERIOD_COLS.forEach((colIdx, periodOffset) => {
      const period = periodOffset + 1;
      const cellVal = String(row[colIdx] || '').trim();
      if (!cellVal || cellVal.toUpperCase().includes('LUNCH')) return;

      const parts = cellVal.split('\n');

      // Subject = first line, strip any inline [bracket] notes
      let subject = parts[0].trim().replace(/\[.*?\]/gs, '').trim();

      // Teacher = first subsequent line that looks like a name
      let teacher = '';
      for (let p = 1; p < parts.length; p++) {
        const part = parts[p].trim().replace(/\[.*?\]/gs, '').trim();
        if (part && part.match(/^(Smt\.|Sh\.|Dr\.|Scout|NSS)/i)) {
          teacher = part;
          break;
        }
      }

      if (!subject) return;
      if (teacher) upperTeacherSet.add(teacher);

      upperSchedule.push({
        Day: dayName, Period: period, Class: className,
        Tier: tier, Subject: subject, Teacher: teacher
      });
    });
  }
});

const upperClassRows   = upperClasses.map(c => [c.name, c.tier, c.room]);
const upperTeacherRows = [...upperTeacherSet].sort().map(t => [t, '', 40, '', '']);

// ─────────────────────────────────────────────
// PART 3 : Combine and generate ImportData.gs
// ─────────────────────────────────────────────

const allClassRows   = [...middleClassRows,   ...upperClassRows];
const allTeacherRows = [...middleTeacherRows, ...upperTeacherRows];
const allSchedule    = [...middleSchedule,    ...upperSchedule];

// Extract unique subjects (sorted) from the full schedule
const subjectSet = new Set();
allSchedule.forEach(r => { if (r.Subject) subjectSet.add(r.Subject); });
const allSubjectRows = [...subjectSet].sort().map(s => [s]);

const scheduleJson  = JSON.stringify(allSchedule);
const classesJson   = JSON.stringify(allClassRows);
const teachersJson  = JSON.stringify(allTeacherRows);
const subjectsJson  = JSON.stringify(allSubjectRows);

const fileContent = `/**
 * ImportData.gs
 * Auto-generated by gen.js — DO NOT EDIT MANUALLY.
 * Contains Middle School (6th–8th) and Secondary/Higher Secondary (9th–12th) data.
 * Classes: ${allClassRows.length}  |  Teachers: ${allTeacherRows.length}  |  Subjects: ${allSubjectRows.length}  |  Schedule rows: ${allSchedule.length}
 */

const RAW_CLASS_ROWS   = ${classesJson};
const RAW_TEACHER_ROWS = ${teachersJson};
const RAW_SUBJECT_ROWS = ${subjectsJson};
const RAW_SCHEDULE     = ${scheduleJson};

/**
 * Imports all class, teacher, subject, and schedule data into the Google Sheet.
 * Safe to run multiple times — clears existing data before writing.
 */
function importExcelData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // 1. Populate Classes
  const classSheet = ss.getSheetByName('Classes');
  if (classSheet && RAW_CLASS_ROWS.length > 0) {
    const lastRow = classSheet.getLastRow();
    if (lastRow > 1) classSheet.getRange(2, 1, lastRow - 1, classSheet.getLastColumn()).clearContent();
    classSheet.getRange(2, 1, RAW_CLASS_ROWS.length, 3).setValues(RAW_CLASS_ROWS);
  }

  // 2. Populate Teachers
  const teacherSheet = ss.getSheetByName('Teachers');
  if (teacherSheet && RAW_TEACHER_ROWS.length > 0) {
    const lastRow = teacherSheet.getLastRow();
    if (lastRow > 1) teacherSheet.getRange(2, 1, lastRow - 1, teacherSheet.getLastColumn()).clearContent();
    teacherSheet.getRange(2, 1, RAW_TEACHER_ROWS.length, 5).setValues(RAW_TEACHER_ROWS);
  }

  // 3. Populate Subjects (master list for Class View dropdowns)
  const subjectSheet = ss.getSheetByName('Subjects');
  if (subjectSheet && RAW_SUBJECT_ROWS.length > 0) {
    const lastRow = subjectSheet.getLastRow();
    if (lastRow > 1) subjectSheet.getRange(2, 1, lastRow - 1, 1).clearContent();
    subjectSheet.getRange(2, 1, RAW_SUBJECT_ROWS.length, 1).setValues(RAW_SUBJECT_ROWS);
  }

  // 4. Populate Master Schedule
  const scheduleSheet = ss.getSheetByName('Master_Schedule');
  if (scheduleSheet && RAW_SCHEDULE.length > 0) {
    const schedData = RAW_SCHEDULE.map(r => [r.Day, r.Period, r.Class, r.Tier, r.Subject, r.Teacher, '', '']);
    const lastRow = scheduleSheet.getLastRow();
    if (lastRow > 1) scheduleSheet.getRange(2, 1, lastRow - 1, scheduleSheet.getLastColumn()).clearContent();
    scheduleSheet.getRange(2, 1, schedData.length, 8).setValues(schedData);
  }

  ui.alert(
    'Import Complete! All school data loaded:\\n' +
    '• Classes  : ${allClassRows.length}\\n' +
    '• Teachers : ${allTeacherRows.length}\\n' +
    '• Subjects : ${allSubjectRows.length}\\n' +
    '• Schedule : ${allSchedule.length} entries'
  );
}
`;

const outPath = path.join(__dirname, 'ImportData.gs');
fs.writeFileSync(outPath, fileContent);
console.log('✅ ImportData.gs written successfully');
console.log('   Classes  : ' + allClassRows.length + ' (' + middleClassRows.length + ' Middle + ' + upperClassRows.length + ' Upper)');
console.log('   Teachers : ' + allTeacherRows.length + ' (' + middleTeacherRows.length + ' Middle + ' + upperTeacherRows.length + ' Upper)');
console.log('   Subjects : ' + allSubjectRows.length);
console.log('   Schedule : ' + allSchedule.length + ' rows (' + middleSchedule.length + ' Middle + ' + upperSchedule.length + ' Upper)');
