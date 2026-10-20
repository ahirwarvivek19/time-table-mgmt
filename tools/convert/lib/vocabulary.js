/**
 * vocabulary.js — the words that appear in the school's timetable cells.
 *
 * SUBJECTS maps every way a subject is written in the Excel to the name used in
 * the sheet. TEACHER_ALIASES maps every short / informal teacher name to the
 * registry name (registry.js). Matching is case-insensitive; a "." may be
 * followed by spaces ("M. Sharma" = "M.Sharma") and "-" or spaces may join words.
 *
 * When the converter reports `unknown word "X"` or `unknown teacher "X"`, add the
 * new spelling here.
 */
'use strict';

/** Written form → subject name. Longer forms win over shorter ones ("SO.SCI" before "SCI"). */
const SUBJECTS = {
  'sanskrit': 'Sanskrit', 'sanksrit': 'Sanskrit',
  'science': 'Science', 'sci': 'Science',
  'hindi': 'Hindi',
  'so.sci': 'Social Science', 'so.science': 'Social Science', 'social science': 'Social Science', 'social': 'Social Science', 'sst': 'Social Science',
  'maths': 'Mathematics', 'math': 'Mathematics',
  'english': 'English', 'eng': 'English',
  'it': 'IT', 'computer': 'Computer',
  'beauty': 'Beauty',
  'uhv': 'UHV', 'ccle': 'CCLE', 'supw': 'SUPW',
  'geography': 'Geography', 'history': 'History',
  'poly.sci': 'Political Science', 'ploy.sci': 'Political Science', 'pol.sci': 'Political Science',
  'carrier': 'Career', 'career': 'Career',
  'music': 'Music', 'library': 'Library', 'sports': 'Sports', 'dance': 'Dance',
  'physics': 'Physics', 'chemistry': 'Chemistry', 'bio': 'Biology', 'biology': 'Biology',
  'b.stud': 'Business Studies', 'accounts': 'Accountancy', 'acc': 'Accountancy', 'eco': 'Economics',
  'umang-addha phool': 'Life Skills Umang (Aadha Phool)',
};

/** "DAKSHATA ENGLISH" → "Proficiency English" (दक्षता = proficiency periods). Misspellings included. */
const SUBJECT_PREFIXES = { 'dakshata': 'Proficiency', 'dakashta': 'Proficiency', 'dakhsata': 'Proficiency' };

/**
 * Extra short forms used only as hints in the teacher-wise sheet, e.g. "7TH A(SANS)",
 * "9TH D (M)", "11TH C (CARR)". Never used for class cells ("M" would match "M. Sharma").
 */
const HINT_SUBJECTS = {
  ...SUBJECTS, 'sans': 'Sanskrit', 'hin': 'Hindi', 'm': 'Mathematics', 'carr': 'Career', 'b.stu': 'Business Studies', 'prac': 'Practical',
};

/** A subject written "X /CCLE Teacher" means the one teacher takes the period for both (Saturday class-teacher slot). */
const PAIRS_WITH_PREVIOUS = ['CCLE'];

/** Subjects whose cells never name a teacher but which only one teacher teaches (used only when nothing better is known). */
const DEFAULT_TEACHER = { Computer: 'Reena Shivhare' };

// Short / informal names → registry name.
const TEACHER_ALIASES = {
  // classes 6–8 sheet
  'ruchika v.': 'Ruchika Vyas', 'ruchika': 'Ruchika Vyas',
  'premlata v.': 'Premlata Vishwakarama', 'premlata vishvkarma': 'Premlata Vishwakarama',
  'purnima s.': 'Poornima Sarkar', 'roshi j.': 'Roshi Jahan', 'manisha lov.': 'Manisha Lovanshi',
  'geeta ch.': 'Geeta Chandniya', 'lalita tiwari': 'Lalita Tiwari', 'lalita': 'Lalita Tiwari',
  'm. sharma': 'Mahesh Chandra Sharma', 'farhana': 'Farhana Khan', 'farhana k.': 'Farhana Khan',
  'subhadra ag.': 'Subhadra Agrawal', 'jitendra d.': 'Jitendra Dhawan', 'varsha m.': 'Varsha Mandloi',
  'seema parmar': 'Seema Parmar', 'sonali s.': 'Sonali Sharma',
  'garima pur.': 'Garima Purohit', 'ashoka i.': 'Ashok Kumar Indria', 'somesh v.': 'Somesh Vyas',
  'pramod v.': 'Pramod Kumar Vyas', 's. brahmane': 'Sayaki Brahamne', 'dipti t.': 'Deepti Tiwari',
  'rahul rathore': 'Rahul Rathore', 'bhavna sharma': 'Bhavana Sharma', 'reena shivhare': 'Reena Shivhare',
  'moh.ali khan': 'Mohammad Ali Khan',
  'rupali': 'Rupali Chuliwal', // "DANCE RUPALI" — the dance teacher (Rupali Dubey teaches primary classes)
  // classes 9–12 sheet
  'neha jamidar': 'Neha Jamidar', 'neha.j': 'Neha Jamidar', 'neha jhala': 'Neha Jhala',
  'p.sarkar': 'Poornima Sarkar', 'archana.p': 'Archana Pandey', 'premlata.v': 'Premlata Vishwakarama',
  'dipti tiwari': 'Deepti Tiwari', 'r.bairagi': 'Rajesh Kumar Bairagi', 'rajesh bairagi': 'Rajesh Kumar Bairagi',
  'urmila.s': 'Urmila Shrimali', 's.vyas': 'Somesh Vyas', 'ritcha ch.': 'Ritcha Choudhary',
  'divya': 'Divya Lodval', 'roshni suman': 'Roshini Suman', 'ro.s': 'Roshini Suman',
  'a.katare': 'Anuradha Katare', 'a.kartare': 'Anuradha Katare', 'anuradha.k': 'Anuradha Katare',
  'indra.m': 'Indira Mathur', 'i.mathur': 'Indira Mathur', 'ekta gour': 'Ekta Gour',
  'krishna.p': 'Krishna Patale', 'krishna': 'Krishna Patale', 'rubal verma': 'Rubal Verma',
  's.bhrahmane': 'Sayaki Brahamne', 'geeta.c': 'Geeta Chandniya', 'preeti bhoj': 'Preeti Bhoj',
  'rk.jaiswal': 'Rakesh Kumar Jaiswal', 'm.marmat': 'Manjula Marmat', 'gk.dixit': 'Gautam Kumar Dixit',
  'gk dixit': 'Gautam Kumar Dixit', 'r.rathore': 'Rajesh Rathore', 'veershree darpe': 'Virshree Darpe Joshi',
  'vs.darpe': 'Virshree Darpe Joshi', 'd.bhati': 'Dinesh Bhati', 'ak.pandit': 'Arun Kumar Pandit',
  'wahid sir': 'Sayyed Wahid Hussain', 'bhavna.s': 'Bhavana Sharma', 'bhavana.s': 'Bhavana Sharma',
  'b.soni': 'Bharti Soni', 'ma.khan': 'Mohammad Ali Khan', 'm.a.khan': 'Mohammad Ali Khan',
  'reena.s': 'Reena Shivhare', 'rahul.r': 'Rahul Rathore',
};

module.exports = { SUBJECTS, SUBJECT_PREFIXES, HINT_SUBJECTS, PAIRS_WITH_PREVIOUS, DEFAULT_TEACHER, TEACHER_ALIASES };
