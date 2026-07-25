# AGENTS.md — Time Table Management System

## Project Overview

A **Google Apps Script** project that runs entirely inside a bound Google Spreadsheet.
It provides a complete school timetable management system: data entry, conflict validation,
multi-view dashboards (Class, Teacher, Day-wise), a Master Grid, and a Cover/Substitution manager.

**There is no Node.js runtime in production.** Node/npm is only used locally for the
`@google/clasp` CLI that pushes code to Google Apps Script.

---

## Tech Stack

| Layer | Technology |
|---|---|
| Runtime | Google Apps Script (V8 engine) |
| Language | JavaScript (`.gs` = `.js` for Apps Script) |
| UI dialogs | Google Apps Script HTML Service (`.html` files) |
| Data store | Google Sheets (spreadsheet bound to the script) |
| Local tooling | Node.js + `@google/clasp` (push/pull only) |
| Parsing utility | `xlsx` npm package (used only in `gen.js` locally) |

---

## Repository Layout

```
/
├── appsscript.json          # Apps Script manifest (scopes, timezone, runtimeVersion)
├── .clasp.json              # clasp config: scriptId, parentId (linked Google Sheet ID)
├── package.json             # devDependency: @google/clasp; dependency: xlsx
├── gen.js                   # LOCAL-ONLY Node script to pre-process XLSX import data
├── spreadsheet_schema.md    # Human-readable schema for all required Sheet tabs
│
├── Code.gs                  # Entry point: onOpen menu, top-level orchestration helpers
├── DataAccess.gs            # Batch read/write abstraction (const DataAccess = {...})
├── EventHandlers.gs         # onEdit dispatcher & per-sheet write-back handlers
├── Validation.gs            # Conflict detection (teacher clash, room clash, unavail. days)
├── ScheduleParser.gs        # Multi-teacher/class string parsing (const ScheduleParser = {...})
│
├── ClassViewManager.gs      # Renders & writes back Class_View dashboard
├── TeacherViewManager.gs    # Renders & writes back Teacher_View dashboard
├── TeacherDayViewManager.gs # Renders & writes back Teacher_Day_View (day-filtered, all-teacher grid)
├── MasterGrid.gs            # Renders Master_Grid_View (compact all-class grid)
├── TeacherAvailabilityGrid.gs  # Renders Teacher_Availability_Grid (on-demand)
│
├── CoverManager.gs          # Backend API for Cover/Substitution manager
├── CoverManagerUI.html      # Modal HTML UI for Cover Manager (calls google.script.run)
│
├── ImportData.gs            # Bulk import of pre-fill timetable data from sheet
├── GeneratorAlgorithm.gs    # STUB — auto-generation removed; kept to avoid clasp errors
└── .claspignore / .gitignore

# Excluded from clasp push (.claspignore):
#   MultiSelectManager.gs  — REMOVED; multi-value entries use "/" and "," notation directly
#   MultiSelectUI.html     — REMOVED
#   gen.js                 — LOCAL only
```

---

## Menu Structure (`onOpen` in `Code.gs`)

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
    ├── Save Current Sheet as Pre-fill Data→ saveCurrentSheetAsPrefillData()
    └── Reset Pre-fill Data to Default     → resetPrefillDataToDefault()
```

---

## Spreadsheet Data Model

The bound spreadsheet must contain these tabs (created by `setupInitialSpreadsheet()`):

### Data Entry Tabs (user-editable)
| Sheet | Key Columns |
|---|---|
| `Teachers` | Teacher Name, Subject Specialization, Max Hours / Week, Days Unavailable, Total Hours Scheduled |
| `Subjects` | Subject Name |
| `Classes` | Class Name, Academic Tier, Room Assigned |
| `Master_Schedule` | Day, Period, Class, Academic Tier, Subject, Teacher, Room, Clash Status |

### Dashboard Tabs
| Sheet | Editable? | Populated / written back by |
|---|---|---|
| `Class_View` | ✅ Yes | `ClassViewManager` — reads from & writes back to Master_Schedule |
| `Teacher_View` | ✅ Yes | `TeacherViewManager` — reads from & writes back to Master_Schedule |
| `Teacher_Day_View` | ✅ Yes | `TeacherDayViewManager` — reads from & writes back to Master_Schedule |
| `Master_Grid_View` | ❌ Read-only | `MasterGrid.gs` — lightweight refresh on edit, full restyle via menu |
| `Teacher_Availability_Grid` | ❌ Read-only | `TeacherAvailabilityGrid.gs` — generated on demand |
| `Cover_Manager` | ✅ Append-only | `CoverManager.gs` — `apiAssignCover()` appends rows |

> Internal scratch sheets are prefixed with `_` (e.g. `_PrefillData`) and are skipped
> by `styleEntireSheet()`.

---

## Coding Conventions

### Module pattern
Stateless utility/manager modules are defined as **object literals assigned to `const`**,
which gives them a namespace in Apps Script's flat global scope:

```js
// Correct pattern for shared utilities / managers
const DataAccess = {
  getSheetDataAsObjects: function(sheetName) { ... },
  writeSchedule:         function(dataGrid)  { ... },
};

// Top-level functions used by menus or triggers
function onOpen() { ... }
function runValidation() { ... }
```

> Do **not** use `class` syntax — it works in V8 but the object-literal pattern is
> already established across all manager files.

### File naming
- `.gs` extension → deployed to Apps Script as a script file
- `.html` extension → deployed as an HTML template (served via `HtmlService`)
- `gen.js` → local Node only; excluded from clasp push via `.claspignore`

### Performance rules (Apps Script quota)
- **Batch all Sheets API calls.** Read with a single `getValues()` / `getDataRange()`,
  never cell-by-cell in a loop.
- **Batch all writes.** Use `setValues(2DArray)` / `setBackgrounds(2DArray)` instead of
  per-row `setValue()` / `setBackground()`.
- Use `DataAccess.getSheetDataAsObjects(sheetName)` as the standard read helper — it
  returns an array of `{ headerKey: cellValue }` row objects.

### Multi-teacher / multi-class entry convention

Users type multi-value entries **directly into cells**.
**No sidebar or picker is needed** — `ScheduleParser` handles splitting automatically.

| Field | Canonical delimiter | Example |
|---|---|---|
| Teachers | `/` | `Mrs. Farhana / Mr. Somesh` |
| Subjects (same lecture) | `/` | `Hindi / IP` |
| Classes (combined) | `,` | `11th A, 11th B` |

`ScheduleParser` is the **single source of truth** for splitting these fields:
- `splitList()` — splits teachers/subjects on `/` (canonical) or `,`, `\n`, `|` (fallback)
- `splitClasses()` — splits classes on `,`, `+`, `/`, `\n`, `|`
- Always use `ScheduleParser.parseRowAssignments(row)` before doing teacher-presence checks; never split raw strings directly.
- When **writing back**, always join teachers/subjects with `" / "` and classes with `", "`.

### Source of truth & editable views

**`Master_Schedule` is the single source of truth.** All view dashboards read from it.

| Sheet | Editable? | Write-back path |
|---|---|---|
| `Master_Schedule` | ✅ Yes | Directly editable; cascades to all views on edit |
| `Class_View` | ✅ Yes | Grid edits → `ClassViewManager.updateMasterFromClassView()` → Master_Schedule |
| `Teacher_View` | ✅ Yes | Grid edits → `TeacherViewManager.updateMasterFromTeacherView()` → Master_Schedule |
| `Teacher_Day_View` | ✅ Yes | Grid edits → `TeacherDayViewManager.updateMasterFromTeacherDayView()` → Master_Schedule |
| `Master_Grid_View` | ❌ Read-only | Lightweight data refresh only; full restyle via menu |
| `Teacher_Availability_Grid` | ❌ Read-only | Generated on demand via menu |
| `Cover_Manager` | ✅ Append-only | `apiAssignCover()` appends rows |

**Cascade rule:** After any write-back to Master_Schedule from a view, the handler in
`EventHandlers.gs` must manually re-render all other views — programmatic writes do NOT
fire `onEdit` again.

### Teacher_View cell edit format
When editing a `Teacher_View` grid cell, the value must be one of:

| Format | Example | Meaning |
|---|---|---|
| `ClassName - Subject` | `6th A - Mathematics` | Assign teacher to this class/subject |
| `ClassName\nSubject` | (rendered format, copy-paste safe) | Same as above |
| `FREE` or empty | — | Remove teacher from this slot |

Multiple combined classes: `11th A, 11th B - IP`

### Styling constants (used in `Code.gs` `styleEntireSheet`)
| Token | Value |
|---|---|
| Header background | `#2C3E50` |
| Header font color | `#FFFFFF` |
| Odd body row | `#F8F9FA` |
| Even body row | `#FFFFFF` |
| Font family | `Montserrat` |
| Border color | `#BDC3C7` |

Custom view dashboards (`Class_View`, `Teacher_View`, `Teacher_Day_View`,
`Master_Grid_View`, `Teacher_Availability_Grid`) manage their own styling — they are
**skipped** by `styleEntireSheet()`.

---

## Key Function Reference

| Function | File | Purpose |
|---|---|---|
| `onOpen()` | `Code.gs` | Builds the "Timetable System" menu with submenus |
| `onEdit(e)` | `EventHandlers.gs` | Dispatches edit events; cascades syncs across all views |
| `setupInitialSpreadsheet()` | `Code.gs` | Creates all required tabs and applies initial setup |
| `refreshAllViews()` | `Code.gs` | Re-renders all dashboards in one call |
| `applyMasterScheduleDropdowns_()` | `Code.gs` | Applies Subject/Teacher data-validation to Master_Schedule |
| `styleEntireSheet()` | `Code.gs` | Applies global Montserrat styling to all standard data tabs |
| `reorderSheets()` | `Code.gs` | Enforces canonical tab order |
| `runValidation()` | `Validation.gs` | Detects clashes and writes to Clash Status column |
| `DataAccess.getSheetDataAsObjects()` | `DataAccess.gs` | Standard batch-read helper |
| `ScheduleParser.parseRowAssignments()` | `ScheduleParser.gs` | Parses multi-teacher rows into `{teacher, subject}` pairs |
| `ScheduleParser.splitList()` | `ScheduleParser.gs` | Splits teacher/subject string on `/` or fallback delimiters |
| `ScheduleParser.splitClasses()` | `ScheduleParser.gs` | Splits class string on `,` or fallback delimiters |
| `ScheduleParser.groupTeacherSlots()` | `ScheduleParser.gs` | Detects combined vs. clash slots for Teacher_Day_View |
| `ClassViewManager.renderClassView()` | `ClassViewManager.gs` | Redraws Class_View for a given class |
| `ClassViewManager.updateMasterFromClassView()` | `ClassViewManager.gs` | Writes Class_View grid edit back to Master_Schedule |
| `TeacherViewManager.renderTeacherView()` | `TeacherViewManager.gs` | Redraws Teacher_View for a given teacher |
| `TeacherViewManager.updateMasterFromTeacherView()` | `TeacherViewManager.gs` | Writes Teacher_View grid edit back to Master_Schedule |
| `TeacherDayViewManager.renderTeacherDayView()` | `TeacherDayViewManager.gs` | Redraws Teacher_Day_View for a given day + class filter |
| `TeacherDayViewManager.updateMasterFromTeacherDayView()` | `TeacherDayViewManager.gs` | Writes Teacher_Day_View grid edit back to Master_Schedule |
| `generateMasterGrid()` | `MasterGrid.gs` | Full generation + premium styling of Master_Grid_View |
| `refreshMasterGridData_()` | `MasterGrid.gs` | Lightweight data-only refresh (no restyle); called on every edit |
| `generateTeacherAvailabilityGrid()` | `TeacherAvailabilityGrid.gs` | Prompts for a day and renders FREE/BUSY grid |
| `openCoverManagerUI()` | `CoverManager.gs` | Launches the Cover Manager modal |
| `apiGetAvailableTeachers(day, period)` | `CoverManager.gs` | Returns free teachers for a given slot |
| `apiAssignCover(...)` | `CoverManager.gs` | Appends a cover assignment row to Cover_Manager sheet |
| `importExcelData()` | `ImportData.gs` | Imports all pre-fill data (teachers + schedule) |
| `importTeachersData()` | `ImportData.gs` | Imports static teacher registry into Teachers tab |
| `importScheduleData()` | `ImportData.gs` | Imports dynamic schedule into Master_Schedule tab |

---

## HTML UI Pattern (google.script.run)

HTML dialogs communicate with the backend via `google.script.run`:

```js
// In .html files — async call with success/failure handlers
google.script.run
  .withSuccessHandler(function(result) { /* update UI */ })
  .withFailureHandler(function(err)    { /* show error */ })
  .apiGetAvailableTeachers(day, period);
```

- All `api*` functions in `CoverManager.gs` are designed to be called this way.
- The HTML template files use `<?= ... ?>` scriptlets for server-side rendering.

---

## Local Development Workflow

```bash
# Install dependencies (clasp CLI)
npm install

# Login to Google account
npx clasp login

# Push local changes to Apps Script
npx clasp push

# Pull current remote state
npx clasp pull

# Open the script project in the browser
npx clasp open
```

> **Do not run `clasp push` with files that contain Node-only code.**
> `gen.js`, `MultiSelectManager.gs`, and `MultiSelectUI.html` are excluded via `.claspignore`.

### Deployment (Triggers)
The `onEdit` trigger must be set as an **installable trigger** in the Apps Script dashboard
(not just a simple trigger) for write-back logic to work reliably.

---

## Post-Deploy Checklist (After Every `clasp push`)

> **Rule for AI agents:** After every `clasp push`, always tell the user which menu
> items to run from the **Timetable System** dropdown in Google Sheets to activate the
> changes. New code is not live until the spreadsheet reloads and the relevant functions
> are executed.

### Step 1 — Reload the spreadsheet
Close and reopen the Google Sheet (or hard-refresh the browser tab) so the new script
version is picked up and the menu rebuilds.

### Step 2 — Run the appropriate menu items

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
| `EventHandlers` / sync logic | No menu needed — takes effect immediately on next edit |
| `ScheduleParser` / parsing logic | No menu needed — takes effect immediately |
| Import / pre-fill data | **📥 Data Import → Import Pre-fill Timetable Data** |

### Step 3 — Verify
Check the affected sheets to confirm the changes rendered correctly.

---

## Open Areas / Stubs

| File | Status |
|---|---|
| `GeneratorAlgorithm.gs` | Intentional stub — auto-generation was removed. Safe to delete if no clasp references remain. |
| `MultiSelectManager.gs` | REMOVED — stubbed in place. Excluded via `.claspignore`. Delete from Apps Script project editor. |
| `MultiSelectUI.html` | REMOVED — stubbed in place. Excluded via `.claspignore`. Delete from Apps Script project editor. |
| `gen.js` | Local-only XLSX pre-processor. Not deployed to Apps Script. |

---

## Important Constraints

1. **No external HTTP calls.** OAuth scopes are limited to `spreadsheets` and
   `script.container.ui`. Do not add `UrlFetchApp` calls without updating `appsscript.json`.
2. **`@OnlyCurrentDoc`** annotation in `Code.gs` restricts the script to its bound
   spreadsheet — do not remove it.
3. **Apps Script execution time limit is 6 minutes.** Long-running imports should use
   continuation tokens or be split across multiple function calls.
4. **Quota limits.** Google Sheets API has daily read/write quotas. Always prefer batch
   operations; see the Performance rules section above.
5. **Tab order matters.** `reorderSheets()` enforces a canonical tab order after any
   setup or refresh operation. Maintain this order when adding new sheets.
6. **Cascade writes manually.** Programmatic `setValue()` / `setValues()` calls do NOT
   fire `onEdit`. Every write-back handler must explicitly re-render all other views after
   writing to Master_Schedule.
