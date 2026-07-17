# Spreadsheet Schema & Setup Guide

This document defines the physical layout of the Google Sheet required to power the Timetable Management System.

> [!TIP]
> **Automated Setup:** To save you time, the `Code.js` file (which we are generating next) includes a function called `setupInitialSpreadsheet()`. When you run this function in Apps Script, it will automatically create all these tabs and headers for you!
> 
> However, if you prefer to build it manually, follow the structure below.

## 1. Data Entry Tabs (Admin Controlled)

### Tab: `Teachers`
* **Purpose:** Registry of all teachers and their availability.
* **Columns:**
  * `A: Teacher Name` (e.g., John Doe)
  * `B: Subject Specialization` (e.g., Mathematics, Physics)
  * `C: Max Hours / Week` (e.g., 25)
  * `D: Days Unavailable` (e.g., "Monday, Friday" or blank)
  * `E: Total Hours Scheduled` (Formula auto-calculating from the Master Schedule)

### Tab: `Classes`
* **Purpose:** Defines the academic blocks and sections.
* **Columns:**
  * `A: Class Name` (e.g., 10A, 10B)
  * `B: Academic Tier` (Dropdown: Primary, Middle, Secondary, Higher Secondary)
  * `C: Room Assigned` (e.g., Room 101)


### Tab: `Rooms`
* **Purpose:** Tracks physical rooms and specialized equipment.
* **Columns:**
  * `A: Room Name` (e.g., Chemistry Lab)
  * `B: Capacity` (e.g., 40)
  * `C: Specialized Type` (e.g., Standard, Science Lab, IT Lab)

---

## 2. Timetable Ledgers (System Driven)

### Tab: `Master_Schedule` (The Edit Hub)
* **Purpose:** The "flat database" that powers everything. Every single period for every class is represented as a single row. This allows the backend to be lightning fast and makes manual editing straightforward.
* **Columns:**
  * `A: Day` (e.g., Monday)
  * `B: Period` (e.g., 1)
  * `C: Class` (e.g., 10A)
  * `D: Academic Tier` (e.g., Secondary)
  * `E: Subject` (Editable Dropdown)
  * `F: Teacher` (Editable Dropdown)
  * `G: Room` (Editable Dropdown)
  * `H: Clash Status` (Conditional Formatting triggers if a clash is detected)

---

## 3. Dynamic Dashboards (Read Only)

### Tab: `Class_View`
* **Purpose:** A clean, printable 2D grid showing a specific class's week.
* **Structure:**
  * `Cell B2:` Dropdown to select a "Class".
  * `Grid (B5:J10):` Days on the Y-Axis, Periods on the X-Axis.
  * **Core Formula:** Utilizes `=FILTER(Master_Schedule!E:F, Master_Schedule!C:C = B2)` shaped into a grid.

### Tab: `Teacher_View`
* **Purpose:** A clean, printable 2D grid showing a specific teacher's week.
* **Structure:**
  * `Cell B2:` Dropdown to select a "Teacher".
  * `Grid (B5:J10):` Days on the Y-Axis, Periods on the X-Axis.
  * **Core Formula:** Utilizes `=FILTER(Master_Schedule!C:C & " - " & Master_Schedule!G:G, Master_Schedule!F:F = B2)` shaped into a grid.

### Tab: `Cover_Manager`
* **Purpose:** To manage daily absent teachers and find replacements.
* **Columns:**
  * `A: Date`
  * `B: Absent Teacher` (Dropdown)
  * `C: Period` (Auto-populates based on Master Schedule)
  * `D: Class to Cover` (Auto-populates)
  * `E: Suggested Available Teachers` (Formula checking who is free during this specific Day/Period)
  * `F: Assigned Cover Teacher` (Manual Selection)
