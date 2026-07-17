/**
 * DataAccess.gs
 * Handles batch reading and writing from the Google Sheet.
 * (NFR-01: Performance optimization by minimizing API calls)
 */

const DataAccess = {
  
  /**
   * Fetches all data from a specific sheet and returns it as an array of objects.
   * Assumes row 1 contains headers.
   * @param {string} sheetName 
   * @returns {Array<Object>}
   */
  getSheetDataAsObjects: function(sheetName) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) return [];
    
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    if (values.length <= 1) return []; // Only headers or empty
    
    const headers = values[0];
    const data = [];
    
    for (let i = 1; i < values.length; i++) {
      let rowObj = {};
      for (let j = 0; j < headers.length; j++) {
        rowObj[headers[j]] = values[i][j];
      }
      data.push(rowObj);
    }
    return data;
  },

  /**
   * Clears the Master Schedule. 
   * If tierTarget is provided, it only clears rows for that specific tier.
   * @param {string|null} tierTarget (e.g., 'Primary', null for all)
   */
  clearMasterSchedule: function(tierTarget) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Master_Schedule');
    if (!sheet) return;

    if (!tierTarget) {
      // Clear everything except headers
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
      }
    } else {
      // Tiered clearing: Find rows matching the tier and clear specific columns (Subject, Teacher, Room)
      const values = sheet.getDataRange().getValues();
      const tierColIndex = 3; // Academic Tier is D (0-indexed 3)
      const subjectColIndex = 4; // E
      const teacherColIndex = 5; // F
      const roomColIndex = 6;    // G
      
      for (let i = 1; i < values.length; i++) {
        if (values[i][tierColIndex] === tierTarget) {
          // Clear Subject, Teacher, Room. Keep Day, Period, Class, Tier.
          sheet.getRange(i + 1, subjectColIndex + 1, 1, 3).clearContent();
        }
      }
    }
  },

  /**
   * Writes the generated schedule to the Master Schedule tab in a single batch operation.
   * @param {Array<Array>} dataGrid - 2D Array matching the Master_Schedule columns.
   */
  writeSchedule: function(dataGrid) {
    if (!dataGrid || dataGrid.length === 0) return;
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Master_Schedule');
    if (!sheet) return;
    
    // Write at the bottom of existing data
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, dataGrid.length, dataGrid[0].length).setValues(dataGrid);
  }
};
