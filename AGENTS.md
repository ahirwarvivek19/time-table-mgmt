# AGENTS.md — Time Table Management System

> **Last updated:** 2026-10-10
> **Current data state:** 25 classes (6th–12th), 76 teachers, 96 subjects, ~1 008 schedule rows.

---

## 1 · Project Overview

A **Google Apps Script** project bound to a single Google Spreadsheet.
It manages a full school timetable: data entry, conflict validation, multi-view
dashboards (Class, Teacher, Day-wise), a Master Grid, and a Cover / Substitution manager.

**There is no Node.js runtime in production.** Node/npm is used locally only for the
`@google/clasp` CLI (push/pull to Apps Script) and `xlsx` (XLSX parsing in `gen.js` / `scratch/` scripts).

---

## 2 · Tech Stack

| Layer | Technology |
|---|---|
| Runtime | Google Apps Script (V8 engine) |
| Language | JavaScript (`.gs` ≡ `.js` for Apps Script) |
| UI dialogs | Google Apps Script HTML Service (`.html` files) |
| Data store | Google Sheets (bound spreadsheet) |
| Local tooling | Node.js + `@google/clasp` (push/pull only) |
| Parsing utility | `xlsx` npm package (local-only, used in `gen.js` and `scratch/` scripts) |

### Key Config Files

| File | Purpose |
|---|---|
| `appsscript.json` | Apps Script manifest — timezone `Asia/Kolkata`, V8 runtime, OAuth scopes: `spreadsheets` + `script.container.ui` |
| `.clasp.json` | Links local repo to scriptId `11u-eHdGgsCoU_scC3DfeouSy6NJ8IZEdbBBzdzIS5Z0m7exhlbfapTYp` and parentId (Google Sheet) `1WPSTS9QUf5_uQOWdTRW3yh9Uqm_AUs8gNzCd5xt94WA` |
| `package.json` | devDep: `@google/clasp ^3.3.0`, dep: `xlsx ^0.18.5` |
| `.claspignore` | Excludes: `node_modules/**`, `gen.js`, `package*.json`, `scratch/**`, `*.xlsx`, `.vscode/**`, `MultiSelectManager.gs`, `MultiSelectUI.html` |
| `.gitignore` | Excludes: `node_modules/`, `package-lock.json`, `.clasp.json`, `scratch/`, `.DS_Store`, `Thumbs.db` |

---

## 3 · Repository Layout

```
f:\Coding\Time Table Mgmt\
│
├── appsscript.json              # Apps Script manifest
├── .clasp.json                  # clasp config (git-ignored — contains tokens)
├── .claspignore                 # Files excluded from clasp push
├── .gitignore                   # Git exclusions
├── package.json                 # npm dependencies
├── spreadsheet_schema.md        # Human-readable schema for all Sheet tabs
├── AGENTS.md                    # ← THIS FILE — project reference for AI agents
│
│   ─── Core Runtime Files (deployed to Apps Script) ───
│
├── Code.gs                      # Entry point: onOpen menu, setup, styling, reorderSheets
├── DataAccess.gs                # Batch read/write abstraction (const DataAccess = {...})
├── EventHandlers.gs             # onEdit dispatcher & per-sheet write-back handlers
├── Validation.gs                # Conflict detection (teacher/room clash, unavailable days)
├── ScheduleParser.gs            # Multi-teacher/class string parsing (const ScheduleParser = {...})
│
├── ClassViewManager.gs          # Renders & writes back Class_View dashboard
├── TeacherViewManager.gs        # Renders & writes back Teacher_View dashboard
├── TeacherDayViewManager.gs     # Renders & writes back Teacher_Day_View (day-filtered)
├── MasterGrid.gs                # Renders Master_Grid_View (compact all-class grid)
├── TeacherAvailabilityGrid.gs   # Renders Teacher_Availability_Grid (on-demand)
│
├── CoverManager.gs              # Backend API for Cover/Substitution manager
├── CoverManagerUI.html          # Modal HTML UI for Cover Manager (google.script.run)
│
├── ImportData.gs                # Bulk import: RAW_TEACHER_ROWS, RAW_CLASS_ROWS,
│                                #   RAW_SUBJECT_ROWS, RAW_SCHEDULE constants + import functions
├── GeneratorAlgorithm.gs        # STUB — auto-generation removed; kept to avoid clasp errors
│
│   ─── Stubs / Excluded Files (in .claspignore) ───
│
├── MultiSelectManager.gs        # REMOVED stub — excluded from clasp push
├── MultiSelectUI.html           # REMOVED stub — excluded from clasp push
│
│   ─── Local-Only Files (never deployed) ───
│
├── gen.js                       # Node.js XLSX pre-processor (Kruti Dev decoding, schedule extraction)
├── Copy of time Table 9-12 (2026-27) 9 F.xlsx   # Source Excel for classes 9–12
├── ms time table.xlsx           # Source Excel for Middle School
│
│   ─── Scratch Scripts (git-ignored, clasp-ignored) ───
│
└── scratch/                     # 39 one-off Node.js helper scripts for data extraction,
                                 #   verification, and ImportData.gs code generation
```

---

## 4 · Menu Structure (`onOpen` in `Code.gs`)

```
Timetable System
├── 🔄 Refresh All Views                   → refreshAllViews()
├── ✅ Check For Conflicts                  → runValidation()
├── 📊 Teacher Availability                → generateTeacherAvailabilityGrid()
├── 👥 Cover Manager                       → openCoverManagerUI()
├── ──────────────────────────────────────
├── ⚙️ Setup & Utilities ▶
│   ├── Setup Spreadsheets (Run Once)      → setupInitialSpreadsheet()
│   ├── Generate Master Grid (Full Restyle)→ generateMasterGrid()
│   ├── ─────────────────────────────────
│   ├── Apply Master Schedule Dropdowns    → applyMasterScheduleDropdowns()
│   ├── Apply Global Styling               → styleEntireSheet()
│   ├── Reorder Spreadsheet Tabs           → reorderSheets()
│   └── Format Teacher Names (Title Case)  → formatAllTeacherNames()
└── 📥 Data Import ▶
    ├── Import All Pre-fill Data           → importExcelData()
    ├── Import Static Teacher Registry     → importTeachersData()
    ├── Import Dynamic Schedule Data       → importScheduleData()
    ├── ─────────────────────────────────
    ├── Save Current Sheet as Pre-fill Data→ saveCurrentSheetAsPrefillData()
    └── Reset Pre-fill Data to Default     → resetPrefillDataToDefault()
```

---

## 5 · Spreadsheet Data Model

The bound spreadsheet contains these tabs (created by `setupInitialSpreadsheet()`):

### 5.1 Data Entry Tabs (user-editable)

| Sheet | Headers (columns A→) |
|---|---|
| `Teachers` | Teacher Name, Subject Specialization, Max Hours / Week, Days Unavailable, Total Hours Scheduled |
| `Subjects` | Subject Name |
| `Classes` | Class Name, Academic Tier, Room Assigned |
| `Master_Schedule` | Day, Period, Class, Academic Tier, Subject, Teacher, Room, Clash Status |

### 5.2 Dashboard Tabs

| Sheet | Editable? | Managed by |
|---|---|---|
| `Class_View` | ✅ Yes (grid cells) | `ClassViewManager` — B3 dropdown selects class; grid edits write back to Master_Schedule |
| `Teacher_View` | ✅ Yes (grid cells) | `TeacherViewManager` — B3 dropdown selects teacher; grid edits write back to Master_Schedule |
| `Teacher_Day_View` | ✅ Yes (grid cells) | `TeacherDayViewManager` — B3 selects day, E3 selects class filter; grid edits write back |
| `Master_Grid_View` | ❌ Read-only | `MasterGrid.gs` — lightweight data refresh on edit; full restyle via menu |
| `Teacher_Availability_Grid` | ❌ Read-only | `TeacherAvailabilityGrid.gs` — generated on demand (prompts for day) |
| `Cover_Manager` | ✅ Append-only | `CoverManager.gs` — `apiAssignCover()` appends rows |

### 5.3 Internal Sheets

| Sheet | Purpose |
|---|---|
| `_Prefill_Data_Backup` | Hidden snapshot saved by `saveCurrentSheetAsPrefillData()`. Used as import source if present; otherwise falls back to `RAW_*` constants in `ImportData.gs`. Skipped by `styleEntireSheet()`. |

> All sheets prefixed with `_` are internal and skipped by global styling.

### 5.4 Current Classes in the System (25 sections)

| Tier | Classes |
|---|---|
| Middle | 6th A, 6th B, 6th C, 7th A, 7th B, 8th A, 8th B |
| Secondary | 9th A, 9th B, 9th C, 9th D, 9th E, 10th A, 10th B, 10th C |
| Higher Secondary | 11th A, 11th B, 11th C (H), 11th C (E), 11th D, 12th A, 12th B, 12th C (H), 12th C (E), 12th D |

> `(H)` = Hindi Medium, `(E)` = English Medium

### 5.5 Tab Order (enforced by `reorderSheets()`)

```
Teachers → Subjects → Classes → Master_Schedule → Class_View → Teacher_View → Teacher_Day_View → Master_Grid_View → Teacher_Availability_Grid → Cover_Manager
```

---

## 6 · Architecture & Data Flow

### 6.1 Source of Truth

```
                    ┌─────────────────────┐
                    │   Master_Schedule    │  ← Single source of truth
                    │  (flat database)     │
                    └──────────┬──────────┘
                               │
          ┌────────────────────┼────────────────────┐
          │                    │                     │
    ┌─────▼──────┐     ┌──────▼──────┐     ┌───────▼────────┐
    │ Class_View │     │Teacher_View │     │Teacher_Day_View│
    │ (editable) │     │ (editable)  │     │  (editable)    │
    └─────┬──────┘     └──────┬──────┘     └───────┬────────┘
          │                    │                     │
          │      write-back    │      write-back     │
          └────────────────────┼────────────────────┘
                               │
                    ┌──────────▼──────────┐
                    │   Master_Schedule    │  → then cascade re-render all views
                    └─────────────────────┘
```

### 6.2 Edit Cascade Flow (`onEdit` in `EventHandlers.gs`)

1. User edits a cell in any editable view or Master_Schedule.
2. `onEdit(e)` dispatches to the appropriate handler:
   - `handleClassViewEdit(e)` → dropdown change (B3) re-renders; grid edit calls `ClassViewManager.updateMasterFromClassView()` → writes to Master_Schedule → re-renders Teacher_View, Teacher_Day_View, Master_Grid_View.
   - `handleTeacherViewEdit(e)` → same pattern via `TeacherViewManager.updateMasterFromTeacherView()`.
   - `handleTeacherDayViewEdit(e)` → same pattern via `TeacherDayViewManager.updateMasterFromTeacherDayView()`.
   - `handleMasterScheduleEdit(e)` → cascades refresh to all view dashboards directly.
3. **Critical rule:** Programmatic `setValue()` / `setValues()` calls do **NOT** fire `onEdit`. Every write-back handler must explicitly re-render all other views after writing to Master_Schedule.

### 6.3 Import Data Flow

```
RAW_TEACHER_ROWS ──┐
RAW_CLASS_ROWS   ──┤
RAW_SUBJECT_ROWS ──┤──→ importExcelData() ──→ Teachers, Classes, Subjects, Master_Schedule tabs
RAW_SCHEDULE     ──┘                         then: applyDropdowns → refreshAllViews
```

If `_Prefill_Data_Backup` sheet exists, import functions read from it **instead** of the `RAW_*` constants (this preserves user customizations).

---

## 7 · Coding Conventions

### 7.1 Module Pattern

Stateless utility/manager modules use **object literals assigned to `const`**:

```js
const DataAccess = {
  getSheetDataAsObjects: function(sheetName) { ... },
  writeSchedule:         function(dataGrid)  { ... },
};
```

Top-level functions are used for menu entries and triggers:

```js
function onOpen() { ... }
function runValidation() { ... }
```

> Do **not** use `class` syntax. The object-literal pattern is established across all files.

### 7.2 File Naming

| Extension | Deployed to | Purpose |
|---|---|---|
| `.gs` | Apps Script (script file) | Server-side JavaScript |
| `.html` | Apps Script (HTML template) | UI dialogs served via `HtmlService` |
| `.js` (local) | NOT deployed | Node.js scripts (`gen.js`, `scratch/*.js`) |

### 7.3 Performance Rules (Apps Script Quota)

- **Batch all Sheets API calls.** Read with `getValues()` / `getDataRange()`, never cell-by-cell.
- **Batch all writes.** Use `setValues(2DArray)` / `setBackgrounds(2DArray)`, never per-row.
- Use `DataAccess.getSheetDataAsObjects(sheetName)` as the standard read helper — returns `[{ headerKey: cellValue }]`.
- Apps Script execution limit: **6 minutes**. Long imports should be split.

### 7.4 Multi-Teacher / Multi-Class Entry Convention

Users type multi-value entries **directly into cells**. No sidebar or picker needed.

| Field | Canonical delimiter | Example |
|---|---|---|
| Teachers | `/` | `Mrs. Farhana / Mr. Somesh` |
| Subjects (same lecture) | `/` | `Hindi / IP` |
| Classes (combined) | `,` | `11th A, 11th B` |

`ScheduleParser` is the **single source of truth** for splitting:
- `splitList(str)` — splits on `/` (canonical) or `,`, `\n`, `|` (fallback)
- `splitClasses(str)` — splits on `,`, `+`, `/`, `\n`, `|`
- Always use `ScheduleParser.parseRowAssignments(row)` before teacher-presence checks.
- When **writing back**: join teachers/subjects with `" / "`, classes with `", "`.

### 7.5 Teacher_View Cell Edit Format

| Format | Example | Meaning |
|---|---|---|
| `ClassName - Subject` | `6th A - Mathematics` | Assign teacher to class/subject |
| `ClassName\nSubject` | (rendered format, copy-paste safe) | Same as above |
| `FREE` or empty | — | Remove teacher from slot |
| Multiple classes | `11th A, 11th B - IP` | Combined optional class |

### 7.6 Styling Constants (`styleEntireSheet` in `Code.gs`)

| Token | Value |
|---|---|
| Header background | `#2C3E50` |
| Header font color | `#FFFFFF` |
| Odd body row | `#F8F9FA` |
| Even body row | `#FFFFFF` |
| Font family | `Montserrat` |
| Border color | `#BDC3C7` |

Custom view dashboards manage their own styling and are **skipped** by `styleEntireSheet()`:
`Master_Grid_View`, `Class_View`, `Teacher_View`, `Teacher_Day_View`, `Teacher_Availability_Grid`.

---

## 8 · Complete Function Reference

### 8.1 Code.gs — Entry Point & Utilities

| Function | Line | Purpose |
|---|---|---|
| `onOpen()` | 7 | Builds the "Timetable System" menu with submenus |
| `refreshAllViews(silent)` | 43 | Re-renders all dashboards, dropdowns, styling, tab order in one call |
| `initClassView()` | 94 | Initializes Class_View with first available class |
| `initTeacherView()` | 107 | Initializes Teacher_View with first available teacher |
| `initTeacherDayView()` | 120 | Initializes Teacher_Day_View for Monday |
| `styleEntireSheet(silent)` | 131 | Applies global Montserrat styling to all standard data tabs |
| `setupInitialSpreadsheet()` | 181 | Creates all required tabs, headers, applies initial setup |
| `applyMasterScheduleDropdowns()` | 266 | Public menu entry — applies dropdowns + shows alert |
| `applyMasterScheduleDropdowns_(silent)` | 281 | Internal — applies Subject/Teacher dropdowns to Master_Schedule |
| `reorderSheets()` | 351 | Enforces canonical tab order |

### 8.2 DataAccess.gs — Batch I/O

| Function | Purpose |
|---|---|
| `DataAccess.getSheetDataAsObjects(sheetName)` | Reads all rows as `[{header: value}]` objects |
| `DataAccess.clearMasterSchedule(tierTarget)` | Clears Master_Schedule (all or by tier) |
| `DataAccess.writeSchedule(dataGrid)` | Appends 2D array to bottom of Master_Schedule |

### 8.3 ScheduleParser.gs — Parsing Engine

| Function | Purpose |
|---|---|
| `ScheduleParser.formatTeacherName(name)` | Title Case formatting with honorific/abbreviation awareness |
| `ScheduleParser.splitList(str)` | Splits teachers/subjects on `/` or fallback delimiters |
| `ScheduleParser.parseRowAssignments(row)` | Parses row into `[{teacher, subject}]` pairs |
| `ScheduleParser.rowIncludesTeacher(row, name)` | Checks if teacher appears in multi-teacher row |
| `ScheduleParser.getSubjectForTeacher(row, name)` | Gets specific subject for a teacher in a row |
| `ScheduleParser.splitClasses(str)` | Splits class string on `,`, `+`, `/`, `\n`, `|` |
| `ScheduleParser.parseRowClasses(row)` | Returns array of class names from a row |
| `ScheduleParser.rowIncludesClass(row, cls)` | Checks if class appears in multi-class row |
| `ScheduleParser.groupTeacherSlots(slots)` | Groups slots into combined/clash display for Teacher_Day_View |

### 8.4 EventHandlers.gs — Edit Dispatcher

| Function | Purpose |
|---|---|
| `onEdit(e)` | Main dispatcher — routes to sheet-specific handlers |
| `handleClassViewEdit(e)` | B3 dropdown → re-render; grid edit → write-back |
| `handleTeacherViewEdit(e)` | B3 dropdown → re-render; grid edit → write-back |
| `handleTeacherDayViewEdit(e)` | B3/E3 selectors → re-render; grid edit → write-back |
| `handleMasterScheduleEdit(e)` | Any cell edit → cascade refresh all views |

### 8.5 View Managers

| Module.Function | File | Purpose |
|---|---|---|
| `ClassViewManager.renderClassView(className)` | ClassViewManager.gs:13 | Redraws Class_View for given class |
| `ClassViewManager.updateMasterFromClassView(className, day, period, editedType, newValue)` | ClassViewManager.gs:232 | Writes grid edit back to Master_Schedule |
| `TeacherViewManager.renderTeacherView(teacherName)` | TeacherViewManager.gs:13 | Redraws Teacher_View for given teacher |
| `TeacherViewManager.updateMasterFromTeacherView(teacherName, day, period, newValue)` | TeacherViewManager.gs:203 | Writes grid edit back to Master_Schedule |
| `TeacherDayViewManager.renderTeacherDayView(selectedDay, selectedClass)` | TeacherDayViewManager.gs:14 | Redraws Teacher_Day_View for day + class filter |
| `TeacherDayViewManager.updateMasterFromTeacherDayView(selectedDay, teacherName, period, newValue)` | TeacherDayViewManager.gs:252 | Writes grid edit back to Master_Schedule |

### 8.6 MasterGrid.gs — Compact All-Class Grid

| Function | Purpose |
|---|---|
| `generateMasterGrid()` | Full generation + premium styling of Master_Grid_View |
| `refreshMasterGridData_()` | Lightweight data-only refresh (no restyle); called on every edit |
| `_buildMasterGridOutput_()` | Internal — builds the 2D output array |
| `_applyMasterGridStyling_(gridSheet, numRows, numCols, numClasses)` | Internal — premium formatting |

### 8.7 TeacherAvailabilityGrid.gs

| Function | Purpose |
|---|---|
| `generateTeacherAvailabilityGrid()` | Prompts for day → renders FREE/BUSY grid for all teachers |

### 8.8 Validation.gs — Conflict Detection

| Function | Purpose |
|---|---|
| `runValidation()` | Detects teacher clashes, room clashes, unavailable days → writes to Clash Status column |

### 8.9 CoverManager.gs — Substitution Manager

| Function | Purpose |
|---|---|
| `openCoverManagerUI()` | Launches Cover Manager modal dialog |
| `apiGetTeachers()` | Returns list of all teachers (for UI dropdown) |
| `apiGetTeacherSchedule(teacherName, day)` | Returns teacher's schedule for a given day |
| `apiGetAvailableTeachers(day, period)` | Returns free teachers for a given slot |
| `apiAssignCover(date, absentTeacher, period, classToCover, coverTeacher)` | Appends cover row to Cover_Manager sheet |

### 8.10 ImportData.gs — Data Import & Pre-fill

| Function | Purpose |
|---|---|
| `importTeachersData(silent)` | Imports teacher registry into Teachers tab |
| `importScheduleData(silent)` | Imports classes, subjects, and schedule into respective tabs |
| `importExcelData()` | Combines `importTeachersData` + `importScheduleData` |
| `saveCurrentSheetAsPrefillData()` | Saves current sheet state as `_Prefill_Data_Backup` |
| `resetPrefillDataToDefault()` | Deletes backup; reverts to `RAW_*` factory defaults |
| `formatAllTeacherNames()` | Applies Title Case formatting across Teachers, Master_Schedule, and backup |

### 8.11 Constants in ImportData.gs

| Constant | Size | Description |
|---|---|---|
| `RAW_CLASS_ROWS` | 25 rows × 3 cols | `[className, tier, room]` for all 25 sections |
| `RAW_TEACHER_ROWS` | 76 rows × 5 cols | `[name, specialization, maxHours, daysUnavail, totalScheduled]` |
| `RAW_SUBJECT_ROWS` | 96 rows × 1 col | `[subjectName]` — includes all variants (practicals, bridge courses, etc.) |
| `RAW_SCHEDULE` | ~1 008 objects | `{Day, Period, Class, Tier, Subject, Teacher, Room}` for 6 days × 8 periods × 21 classes |
| `PREFILL_SHEET_NAME` | constant | `'_Prefill_Data_Backup'` |

---

## 9 · HTML UI Pattern (`google.script.run`)

HTML dialogs communicate with the backend via `google.script.run`:

```js
google.script.run
  .withSuccessHandler(function(result) { /* update UI */ })
  .withFailureHandler(function(err)    { /* show error */ })
  .apiGetAvailableTeachers(day, period);
```

- All `api*` functions in `CoverManager.gs` are designed for this pattern.
- HTML template files use `<?= ... ?>` scriptlets for server-side rendering.

---

## 10 · Local Development Workflow

```bash
# Install dependencies (clasp CLI + xlsx)
npm install

# Login to Google account (one-time)
npx clasp login

# Push local changes to Apps Script
npx clasp push

# Pull current remote state
npx clasp pull

# Open the script project in browser
npx clasp open
```

> **Do not run `clasp push` with Node-only code.** The `.claspignore` file excludes
> `gen.js`, `scratch/`, `*.xlsx`, `MultiSelectManager.gs`, `MultiSelectUI.html`, etc.

### Trigger Setup

The `onEdit` trigger must be set as an **installable trigger** in the Apps Script dashboard
(not just a simple trigger) for write-back logic to work reliably.

---

## 11 · Post-Deploy Checklist (After Every `clasp push`)

> **Rule for AI agents:** After every `clasp push`, always tell the user which menu
> items to run from the **Timetable System** dropdown in Google Sheets.

### Step 1 — Reload the Spreadsheet

Close and reopen the Google Sheet (or hard-refresh) so the new script version loads.

### Step 2 — Run the Appropriate Menu Items

| What changed | Menu item(s) to run |
|---|---|
| Any view manager (`ClassViewManager`, `TeacherViewManager`, etc.) | **🔄 Refresh All Views** |
| `Code.gs` menu / `onOpen` | Reload the sheet (menu rebuilds automatically) |
| `setupInitialSpreadsheet` / sheet structure | **⚙️ Setup & Utilities → Setup Spreadsheets (Run Once)** |
| `reorderSheets` / tab order | **⚙️ Setup & Utilities → Reorder Spreadsheet Tabs** |
| `styleEntireSheet` / global styling | **⚙️ Setup & Utilities → Apply Global Styling** |
| `applyMasterScheduleDropdowns` / dropdowns | **⚙️ Setup & Utilities → Apply Master Schedule Dropdowns** |
| `generateMasterGrid` / Master Grid styling | **⚙️ Setup & Utilities → Generate Master Grid (Full Restyle)** |
| `runValidation` / clash detection | **✅ Check For Conflicts** |
| `EventHandlers` / sync logic | No menu needed — takes effect on next edit |
| `ScheduleParser` / parsing logic | No menu needed — takes effect immediately |
| `ImportData.gs` / pre-fill data | **📥 Data Import → Import Dynamic Schedule Data** |

### Step 3 — Verify

Check the affected sheets to confirm changes rendered correctly.

---

## 12 · Important Constraints & Gotchas

1. **No external HTTP calls.** OAuth scopes are limited to `spreadsheets` and `script.container.ui`. Do not add `UrlFetchApp` calls without updating `appsscript.json`.
2. **`@OnlyCurrentDoc`** annotation in `Code.gs` (line 2) restricts the script to its bound spreadsheet. Do not remove it.
3. **6-minute execution limit.** Split long-running operations.
4. **Quota limits.** Always prefer batch operations (see §7.3).
5. **Cascade writes manually.** Programmatic `setValue()` / `setValues()` do NOT fire `onEdit`. Every write-back handler must explicitly re-render all other views.
6. **`ImportData.gs` is ~192 KB** because it embeds all `RAW_SCHEDULE` data inline as JSON. Changes to this file should be generated by scratch scripts, not hand-edited.
7. **Teacher name spelling must match exactly** between `RAW_TEACHER_ROWS` and `RAW_SCHEDULE`. Mismatches cause teachers to appear as "unregistered" and break view rendering.
8. **Subject strings must match exactly** between `RAW_SUBJECT_ROWS` and `RAW_SCHEDULE`. The `Subjects` tab is populated from `RAW_SUBJECT_ROWS` and feeds dropdown validation.
9. **`(H)` / `(E)` suffix** on class names (e.g. `11th C (H)`) denotes Hindi / English medium sections. These are distinct classes with separate schedules.

---

## 13 · Open Areas / Stubs

| File | Status |
|---|---|
| `GeneratorAlgorithm.gs` | Intentional stub — auto-generation was removed. Safe to delete if no clasp references remain. |
| `MultiSelectManager.gs` | REMOVED — stubbed in place. Excluded via `.claspignore`. Delete from Apps Script project editor. |
| `MultiSelectUI.html` | REMOVED — stubbed in place. Excluded via `.claspignore`. Delete from Apps Script project editor. |
| `gen.js` | Local-only XLSX pre-processor. Not deployed. Superseded by individual `scratch/generate_class*.js` scripts. |
| `ms time table.xlsx` | Middle School source Excel. Not yet tracked in git (untracked). |

---

## 14 · Scratch Scripts Reference

The `scratch/` directory (git-ignored, clasp-ignored) contains Node.js helper scripts used for data extraction and verification. Key scripts:

| Script | Purpose |
|---|---|
| `krutidev.js` / `krutidev_converter.js` | Kruti Dev 010 → Unicode Devanagari Hindi conversion |
| `generate_class9_table.js` | Extracts Class 9 timetable from XLSX → markdown table for review |
| `generate_class10_table.js` | Same for Class 10 |
| `generate_class11_table.js` | Same for Class 11 |
| `generate_class12_table.js` | Same for Class 12 |
| `generate_ms_table.js` | Same for Middle School (6th–8th) |
| `update_importdata.js` | Generates `ImportData.gs` code for Class 12 `RAW_SCHEDULE` entries |
| `update_importdata_class9.js` | Same for Class 9 |
| `update_importdata_class10.js` | Same for Class 10 |
| `update_importdata_class11.js` | Same for Class 11 |
| `update_importdata_ms.js` | Same for Middle School |
| `verify_teachers.js` | Validates teacher names against `RAW_TEACHER_ROWS` |
| `verify_entire_system.js` | Full system consistency check (subjects, teachers, classes) |
| `audit_subjects.js` | Audits subject names in `RAW_SCHEDULE` vs. `RAW_SUBJECT_ROWS` |
| `sync_subjects.js` | Generates updated `RAW_SUBJECT_ROWS` from `RAW_SCHEDULE` |
| `export_markdown_table.js` | Exports full timetable as markdown for user review |

---

## 15 · Git History Summary (Key Milestones)

| Commit | Description |
|---|---|
| `6e7ae8a` | Initial commit |
| `7fafd44` | Add `refreshAllViews()` and multi-class support |
| `c76cfed` | Class filter multi-select dropdown in Teacher_Day_View |
| `26c62ea` | Major refactor: sync all views to Master_Schedule, remove MultiSelect |
| `a511411` | Teacher registry (76 teachers), separate static/dynamic imports |
| `4b78b83` | Class 12 timetable data |
| `c3c4174` | Class 11 timetable data |
| `ff824d5` | Class 10 timetable data |
| `c0681eb` | Class 9 timetable data |
| `7f4ee52` | Middle School timetable data (6th–8th) |
| `9668155` | Teacher ordering in views matches Teachers sheet order |
| `f73c5e0` | Clash display: red background only, no text prefix |
| `651aa50` | Synchronize `RAW_SUBJECT_ROWS` with all schedule subjects |
| `d7df6d9` | Remove class identifiers from subject names |
| `8786efd` | Fix teacher name spelling mismatches |
| `5bfac1c` | (HEAD) Suppress intermediate pop-up alerts; single final alert |

---

## 16 · Standard Operating Procedure: Excel to Pre-fill Data Import

Use this exact 6-step workflow whenever importing or updating timetable schedule data from an external Excel file.

### Step 1 — Sheet Inspection & Period Mapping

1. Read the XLSX file using Node.js (`xlsx` npm package).
2. Locate the timetable sheet and map period columns:

| Column Index | Period | Time |
|---|---|---|
| 4 | Period 1 | 10:40–11:30 |
| 5 | Period 2 | 11:30–12:10 |
| 6 | Period 3 | 12:10–12:50 |
| 7 | Period 4 | 12:50–1:30 |
| 8 | **Break** | 1:30–2:00 |
| 9 | Period 5 | 2:00–2:40 |
| 10 | Period 6 | 2:40–3:20 |
| 11 | Period 7 | 3:20–4:00 |
| 12 | Period 8 | 4:00–4:40 |

### Step 2 — Language & Encoding Decoding

1. **Kruti Dev 010 → Unicode Devanagari**: Use the mapping dictionary in `scratch/krutidev.js` to decode ASCII font text to proper Hindi.
2. **Hindi → English Standardization**:

| Hindi | English |
|---|---|
| राजनीति शास्त्र | Political Science |
| इतिहास | History |
| भौतिक शास्त्र | Physics |
| रसायनशास्त्र | Chemistry |
| गणित | Mathematics |
| जीवविज्ञान | Biology |
| लेखा शास्त्र | Accountancy |
| व्यवसायिक अध्ययन | Business Studies |
| अर्थशास्त्र | Economics |
| भूगोल | Geography |
| हिन्दी | Hindi |
| अंग्रेजी | English |
| संस्कृत | Sanskrit |

3. Match raw Hindi teacher names against the 76-teacher registry in `RAW_TEACHER_ROWS`.

### Step 3 — Class & Section Normalization

1. Identify all medium/section splits (e.g., `12th A`, `12th B`, `12th C (H)`, `12th C (E)`, `12th D`).
2. Verify all section names exist in `RAW_CLASS_ROWS` in `ImportData.gs`.

### Step 4 — Master Teacher & Subject Validation

1. Verify every extracted teacher exists in `RAW_TEACHER_ROWS`.
2. Check if any extracted subject is missing from `RAW_SUBJECT_ROWS`; add missing ones.

### Step 5 — User Review & Approval

1. Generate formatted markdown tables (Section × Day × Period 1–8) showing Subject & Teacher.
2. **Present the full table in the chat for user review.**
3. **STOP and wait for explicit user approval before modifying code or pushing.**

### Step 6 — Codebase Update, Git & Clasp Deployment

Once approved:
1. Run scratch script to update `RAW_SCHEDULE`, `RAW_CLASS_ROWS`, and `RAW_SUBJECT_ROWS` in `ImportData.gs`.
2. `git add -A && git commit -m "feat: update class X timetable..."`
3. `git push origin master`
4. `npx clasp push`
5. Tell user to: reload Google Sheet → run **📥 Data Import → Import Dynamic Schedule Data** → run **🔄 Refresh All Views**.

---

## 17 · Common Tasks Quick Reference

### Adding a New Class

1. Add entry to `RAW_CLASS_ROWS` in `ImportData.gs`.
2. Add all schedule rows (6 days × 8 periods) to `RAW_SCHEDULE`.
3. Push and import.

### Adding a New Teacher

1. Add entry to `RAW_TEACHER_ROWS` in `ImportData.gs`.
2. Update any `RAW_SCHEDULE` entries that reference this teacher.
3. Push and run **📥 Import Static Teacher Registry**.

### Adding a New Subject

1. Add entry to `RAW_SUBJECT_ROWS` in `ImportData.gs`.
2. Push and run **📥 Import Dynamic Schedule Data**.

### Fixing a Teacher Name Spelling

1. Update the name in `RAW_TEACHER_ROWS`.
2. Update all occurrences in `RAW_SCHEDULE` (use find & replace).
3. Push and import both static + dynamic data.

### Debugging View Rendering Issues

1. Check `Master_Schedule` sheet — it's the source of truth.
2. Verify the teacher/subject/class names match exactly (case-sensitive comparison is used in some places).
3. Run `✅ Check For Conflicts` to surface clashes.
4. Check Apps Script Execution Log (View → Executions in the script editor) for errors.
