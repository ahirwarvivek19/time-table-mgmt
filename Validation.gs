/**
 * Validation.gs
 * Handles real-time and post-generation conflict detection.
 * Checks for: Teacher double-booking, Room double-booking, and Days Unavailable.
 */

function runValidation() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Master_Schedule');
  if (!sheet) return;

  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  if (values.length <= 1) {
    SpreadsheetApp.getUi().alert('Master Schedule is empty.');
    return;
  }

  // Build a lookup of teacher -> unavailable days
  const teachersData = DataAccess.getSheetDataAsObjects('Teachers');
  const unavailableMap = {}; // { teacherName: ['Monday', 'Friday', ...] }
  teachersData.forEach(t => {
    const name = t['Teacher Name'];
    const daysStr = t['Days Unavailable'];
    if (name && daysStr) {
      unavailableMap[name] = String(daysStr).split(',').map(d => d.trim()).filter(d => d);
    }
  });

  // Clear previous validation status (Column H, index 7)
  sheet.getRange(2, 8, values.length - 1, 1).clearContent();

  const statuses = [];
  const tracking = {};

  for (let i = 1; i < values.length; i++) {
    const day = values[i][0];
    const period = values[i][1];
    const teacher = values[i][5];
    const room = values[i][6];
    
    let clash = false;
    let clashMsg = [];

    const timeKey = `${day}_${period}`;
    
    if (!tracking[timeKey]) {
      tracking[timeKey] = { teachers: {}, rooms: {} };
    }

    const currentSubject = values[i][4] || '';
    const rowObj = { Teacher: teacher, Subject: currentSubject };
    const assignments = ScheduleParser.parseRowAssignments(rowObj);
    const teacherList = [...new Set(assignments.map(a => a.teacher).filter(Boolean))];

    teacherList.forEach(t => {
      if (tracking[timeKey].teachers[t]) {
        const prev = tracking[timeKey].teachers[t];
        // If assigned to different subjects at the same time -> True Clash
        if (prev.subject && currentSubject && prev.subject.trim().toLowerCase() !== currentSubject.trim().toLowerCase()) {
          clash = true;
          clashMsg.push(`Teacher Clash (${t})`);
        }
      } else {
        tracking[timeKey].teachers[t] = { class: values[i][2], subject: currentSubject };
      }

      if (unavailableMap[t] && unavailableMap[t].includes(day)) {
        clash = true;
        clashMsg.push(`Unavailable Day (${t})`);
      }
    });

    if (room) {
      if (tracking[timeKey].rooms[room]) {
        clash = true;
        clashMsg.push('Room Clash');
      } else {
        tracking[timeKey].rooms[room] = true;
      }
    }

    if (clash) {
      statuses.push([clashMsg.join(' / ')]);
    } else {
      statuses.push([""]);
    }
  }

  // Write statuses back
  sheet.getRange(2, 8, statuses.length, 1).setValues(statuses);
  
  SpreadsheetApp.getUi().alert('Validation Complete! Check the Clash Status column.');
}
