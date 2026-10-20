/**
 * registry.js — master lists used by the Excel → pre-fill converter.
 *
 * Teacher names here are the canonical spellings written into the sheet.
 * Every name used in a timetable MUST exist in TEACHERS, otherwise the
 * converter fails (that is what keeps Teacher_View / clash checks reliable).
 */
'use strict';

const DAYS = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
const DAY_KEYS = { Mon: 'Monday', Tue: 'Tuesday', Wed: 'Wednesday', Thu: 'Thursday', Fri: 'Friday', Sat: 'Saturday' };

// [name, specialization, maxPeriodsPerWeek, daysUnavailable]
// Order is preserved from the previous registry (views list teachers in this order).
// "Dinesh Kumar Pandya" was listed twice in the old registry — kept once. The 2026-27
// teacher-wise sheet shows two employees with that name (BD7695, BD9714); neither
// teaches classes 6–12, so one entry is enough for the timetable.
const TEACHERS = [
  ['Shiv Singh Ahirwar', 'PR-HSS'], ['Ravi Sahu', 'Vice-Principal'], ['Virshree Darpe Joshi', 'Vocational Education'],
  ['Prem Singh Panwar', 'HM-MS'], ['Santosh Verma', 'HM-MS'], ['Mahesh Chandra Sharma', 'Sanskrit'],
  ['Dinesh Kumar Pandya', 'General'], ['Ekta Gour', 'Hindi'], ['Rajesh Kumar Bairagi', 'English'],
  ['Anuradha Katare', 'English'], ['Rajesh Rathore', 'Biology'], ['Rakesh Kumar Jaiswal', 'Maths'],
  ['Gautam Kumar Dixit', 'Physics'], ['Dinesh Bhati', 'Chemistry'], ['Arun Kumar Pandit', 'Economics'],
  ['Sayyed Wahid Hussain', 'Commerce'], ['Preeti Bhoj', 'History'], ['Roshini Suman', 'Political Science'],
  ['Rubal Verma', 'Geography'], ['Neha Jamidar', 'Sanskrit'], ['Manjula Marmat', 'Craft'],
  ['Sayaki Brahamne', 'Hindi'], ['Urmila Shrimali', 'Hindi'], ['Indira Mathur', 'English'],
  ['Ruchika Vyas', 'English'], ['Ashok Kumar Indria', 'English'], ['Jitendra Dhawan', 'Biology'],
  ['Poornima Sarkar', 'Biology'], ['Somesh Vyas', 'Biology'], ['Ritcha Choudhary', 'Maths'],
  ['Neha Jhala', 'Maths'], ['Roshi Jahan', 'Maths'], ['Deepti Tiwari', 'Maths'], ['Subhadra Agrawal', 'Maths'],
  ['Lalita Tiwari', 'Social Science'], ['Premlata Vishwakarama', 'Social Science'], ['Garima Purohit', 'Social Science'],
  ['Geeta Chandniya', 'Sanskrit'], ['Pramod Kumar Vyas', 'Sanskrit'], ['Sonali Sharma', 'Sanskrit'],
  ['Varsha Mandloi', 'Sanskrit'], ['Farhana Khan', 'Urdu'], ['Bharti Soni', 'Librarian'],
  ['Adityaraj Singh Chauhan', 'Lab'], ['Lokit Chawda', 'Lab'], ['Shailesh Joshi', 'General'],
  ['Chandra Shekhar Jatwa', 'General'], ['Shanti Parmar', 'General'], ['Jagdish Singh Bapu Singh Kelwa', 'General'],
  ['Ramgopal Rathore', 'General'], ['Anju Shrivastava', 'General'], ['Sunita Chaturvedi', 'General'],
  ['Seema Parmar', 'General'], ['Bhagvansingh Limdiya', 'General'], ['Priyanka Singh Na', 'General'],
  ['Rupali Dubey', 'General'], ['Asha Jain', 'General'], ['Ravi Prakash Shukla', 'General'],
  ['Sanyogita Dagar', 'General'], ['Kiran Thakur', 'General'], ['Seema Bairagi', 'General'],
  ['Shalini Kamble', 'General'], ['Alka Dewda', 'General'], ['Madhubala Dhor', 'General'],
  ['Maya Awasthi', 'General'], ['Pragati Bairagi', 'General'], ['Shashikala Dubey', 'General'],
  ['Sunita Sharma', 'General'], ['Mohammad Ali Khan', 'PTI'], ['Rahul Rathore', 'PTI'],
  ['Sumit Chouhan', 'PTI'], ['Rupali Chuliwal', 'Dance'], ['Bhavana Sharma', 'Music'],
  ['Divya Lodval', 'IT'], ['Krishna Patale', 'Beauty'],
  // New in the 2026-27 timetables (not in the previous registry):
  ['Archana Pandey', 'Hindi'],      // "श्रीमती अर्चना पाण्डे" — 9–12 sheet
  ['Reena Shivhare', 'Computer'],   // "श्रीमती रीना शिवहरे" / "REENA SHIVHARE"
  ['Manisha Lovanshi', 'Hindi'],    // "manisha lov." in the class sheet; "Manisha Lovanshi (GTAM0166)" in the 2026-27 teacher-wise sheet
].map(([name, spec]) => [name, spec, 40, '']); // 40 = value used in the previous registry

const NEW_TEACHERS = ['Archana Pandey', 'Reena Shivhare', 'Manisha Lovanshi'];

// Rooms are kept from the previous data where known. "N/A" is dropped: it made
// every class with "N/A" look like it shared one room (false Room Clash).
// 2026-27 (Timetable.xlsx): 6th E is gone, and 11th C / 12th C are one class each —
// their Hindi- and English-medium groups are split periods marked (H.M) / (E.M).
const CLASSES = [
  ['6th A', 'Middle', ''], ['6th B', 'Middle', ''], ['6th C', 'Middle', ''], ['6th D', 'Middle', ''],
  ['7th A', 'Middle', ''], ['7th B', 'Middle', ''], ['7th C', 'Middle', ''], ['7th D', 'Middle', ''],
  ['8th A', 'Middle', ''], ['8th B', 'Middle', ''], ['8th C', 'Middle', ''], ['8th D', 'Middle', ''],
  ['9th A', 'Secondary', '207'], ['9th B', 'Secondary', ''], ['9th C', 'Secondary', ''], ['9th D', 'Secondary', ''], ['9th E', 'Secondary', ''],
  ['10th A', 'Secondary', '201'], ['10th B', 'Secondary', ''], ['10th C', 'Secondary', ''],
  ['11th A', 'Higher Secondary', ''], ['11th B', 'Higher Secondary', ''], ['11th C', 'Higher Secondary', ''], ['11th D', 'Higher Secondary', '208'],
  ['12th A', 'Higher Secondary', '301'], ['12th B', 'Higher Secondary', ''], ['12th C', 'Higher Secondary', ''], ['12th D', 'Higher Secondary', '206'],
];

// Bell schedule — the same for every class. The converter checks it against the
// header rows of the Excel sheets and stops if they differ, so update it here when
// the school changes its timings. Lunch 12:20–12:50 sits between periods 4 and 5.
const PERIODS = [
  [1, '09:30', '10:20'], [2, '10:20', '11:00'], [3, '11:00', '11:40'], [4, '11:40', '12:20'],
  [5, '12:50', '13:30'], [6, '13:30', '14:10'], [7, '14:10', '14:50'], [8, '14:50', '15:30'],
];

module.exports = { DAYS, DAY_KEYS, TEACHERS, NEW_TEACHERS, CLASSES, PERIODS };
