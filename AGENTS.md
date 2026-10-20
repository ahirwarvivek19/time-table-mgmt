# AGENTS.md — Time Table Management System

> **Last updated:** 2026-10-20
> **Current data state:** 28 classes (6th A – 12th D), 78 teachers, 64 subject values, 1,344 schedule rows (2026-27, from `Timetable.xlsx`).

User-facing setup and usage: `README.md`. Sheet layout: `spreadsheet_schema.md`.

---

## 1 · Overview

A **Google Apps Script** project bound to one Google Sheet. `Master_Schedule` is a flat table
(one row per class / day / period) and the single source of truth. Class, teacher and day views
are rendered from it; edits in any view are written back to it and the other views re-render.

Node.js is used **locally only**: `@google/clasp` (push/pull), `xlsx` (Excel conversion) and the test suite.

## 2 · Tech stack

| Layer | Technology |
|---|---|
| Runtime | Google Apps Script, V8 |
| UI | Sheets custom menu + HTML Service dialog (`CoverManagerUI.html`, `google.script.run`) |
| Data | The bound spreadsheet |
| Local tooling | Node 18+, `@google/clasp` 3.4.1, SheetJS `xlsx` 0.20.3 (from cdn.sheetjs.com — the npm `xlsx@0.18.5` has known CVEs), `node:test` |

OAuth scopes (`appsscript.json`): `spreadsheets`, `script.container.ui`, `script.scriptapp` (installable edit trigger).

## 3 · Repository layout

```
Config.gs                 CONFIG: sheet names, days, periods, column indexes; shared helpers
                          (periodTime_ / periodHeader_ — bell times from RAW_PERIOD_ROWS, normalizeDay_, isRealRoom_,
                          getClassNames_, getTeacherNames_, withDocumentLock_)
Code.gs                   onOpen menu, refreshAllViews, setup, global styling, dropdowns, reorderSheets
DataAccess.gs             getSheetDataAsObjects, updateTeacherLoad (Total Hours Scheduled), clear/write helpers
ScheduleParser.gs         Splitting/pairing of multi-teacher / multi-subject / multi-class cells;
                          assignTeacher / removeTeacher keep teacher↔subject pairs aligned
EventHandlers.gs          onEdit dispatcher, installable-trigger install, multi-cell edit sync, cascade refresh
ClassViewManager.gs       Class_View render + write-back
TeacherViewManager.gs     Teacher_View render; updateTeacherSlot_ (shared write-back for both teacher views)
TeacherDayViewManager.gs  Teacher_Day_View render + write-back
MasterGrid.gs             Master_Grid_View (full generate / data-only refresh)
TeacherAvailabilityGrid.gs  FREE / BUSY / UNAVAILABLE grid for one day
Validation.gs             Conflict checks → Clash Status + summary
CoverManager.gs / CoverManagerUI.html   Substitution dialog + API
WebApp.gs / TimetableWeb.html   Optional read-only web page (doGet); deploy as "User accessing the web app" only
ImportData.gs             Import (built-in data or saved snapshot), snapshot save/reset, name formatting
PrefillData.gs            GENERATED data: RAW_CLASS_ROWS, RAW_TEACHER_ROWS, RAW_SUBJECT_ROWS, RAW_PERIOD_ROWS, RAW_SCHEDULE

tools/convert/            Excel → PrefillData.gs converter (local only) — guide: docs/IMPORTING.md
  sources.json            which files/sheets to read (Timetable.xlsx: 6-8, 9-12) + the teacher-wise sheet
  registry.js             canonical teachers, classes, rooms, bell schedule (PERIODS)
  lib/vocabulary.js       every written form of subjects and teacher short names
  lib/cell-grammar.js     reads one class-wise cell into groups (subject, teachers, medium)
  class-blocks.js         reads the class-wise sheets (class blocks, Hindi weekdays, header rows)
  teacherwise.js          reads the teacher-wise sheet; fills teacher-less groups; compares the two sheets
  flat.js                 one-row-per-period CSV/Excel format for any other layout
  build.js                validates and writes PrefillData.gs + docs/TIMETABLE_REVIEW_2026-27.md
tools/preview/            npm run preview: renders every view through gas-mock into preview/timetable.html
                          (page.html + page.js; --serve binds 127.0.0.1 only — keep it that way)
tests/                    node:test suites; harness/gas-mock.js = in-memory SpreadsheetApp etc.
docs/TIMETABLE_REVIEW_2026-27.md   generated conversion review (fills, gaps, sheet disagreements, clashes, load, per-class grids)
Timetable.xlsx            source timetable (sheets 6-8, 9-12, Teacherwise)
```

`.claspignore` excludes `tools/`, `tests/`, `docs/`, `preview/`, `*.md`, `*.xlsx`, `package*.json`, `node_modules/`.

## 4 · Menu (`onOpen` in Code.gs)

```
Timetable System
├── 🔄 Refresh All Views            refreshAllViews()
├── ✅ Check For Conflicts           runValidation()
├── 📊 Teacher Availability         generateTeacherAvailabilityGrid()
├── 👥 Cover Manager                openCoverManagerUI()
├── ⚙️ Setup & Utilities ▶  Setup Spreadsheets (Run Once) · Generate Master Grid (Full Restyle) ·
│                          Apply Master Schedule Dropdowns · Apply Global Styling · Reorder Spreadsheet Tabs ·
│                          Format Teacher Names (Title Case) · Install Edit Sync Trigger
└── 📥 Data Import ▶        Import All Pre-fill Data · Import Static Teacher Registry · Import Dynamic Schedule Data ·
                           Save Current Sheet as Pre-fill Data · Reset Pre-fill Data to Default
```

## 5 · Data flow

- **Edit in a view** → `onEdit` → handler → write-back under `withDocumentLock_` → `refreshViewsAfterWrite_`
  re-renders the other views (programmatic writes never fire `onEdit`, so the cascade is explicit).
- **Multi-cell edits** (paste / fill / delete block) are processed cell by cell, up to `CONFIG.MAX_SYNCED_CELLS`.
- **Triggers**: `installEditTrigger_()` (Setup or menu; editors only — it does a write probe) creates an installable
  `onEdit` trigger and stores its unique id in document property `EDIT_TRIGGER_UID`. `onEdit` processes an event only
  if `e.triggerUid` equals that id; the simple-trigger call and any other installable trigger (hand-made, another
  editor's — `getProjectTriggers()` only lists the caller's own) return immediately. With no id recorded the simple
  trigger does the work (original behaviour). The trigger runs as its installer: re-run the menu if they lose access.
- **Concurrency**: handlers read B3 / column A / edited values first, then do write-back + re-render inside one
  `withDocumentLock_` (tryLock 10 s → "please re-enter" error). Views are shared single-selector sheets.
- **Import**: `importExcelData()` → choose source (saved snapshot vs built-in `PrefillData.gs`) → write the four tabs →
  dropdowns → `refreshAllViews(true)` → generate Master Grid if missing → validation summary.

## 6 · Conventions

- Object-literal modules (`const X = { fn: function() {} }`) for managers; top-level functions for menu entries.
  No `class` syntax.
- Use `CONFIG.*` for sheet names, days, periods and columns — **never** at file top level (load order is not defined).
- Batch reads/writes (`getValues` / `setValues` / `setBackgrounds`). 6-minute execution limit.
- Multi-value cells: teachers and subjects joined with `" / "` in matching order; classes with `", "`.
  Always go through `ScheduleParser` (`parseRowAssignments`, `rowIncludesTeacher`, `assignTeacher`, `removeTeacher`).
- Name comparisons are case/whitespace-insensitive (`ScheduleParser.sameName`).
- Single-activity names that list several teachers must not contain `/` (e.g. `Bal Sabha & Life Skills`, `SUPW & Clubs`),
  otherwise subjects get paired one-to-one with teachers.
- Clash semantics: same teacher + same day/period + **different** subject = clash; same subject = combined lecture
  (same grade, unless a whole-school activity). Medium marks `(H.M)` / `(E.M)` are compared with
  `ScheduleParser.splitMedium`: a mark on one section only is the same lecture, H.M + E.M at once is a clash.

## 7 · Validation (`validateSchedule_`)

Teacher clash · Room clash (real rooms only — blank/`N/A` ignored; identical lecture sharing a room is fine) ·
Unavailable day · Unregistered teacher · Unknown class · Over `Max Hours / Week`. Every row involved is flagged.

## 8 · Excel → pre-fill SOP

1. Replace `Timetable.xlsx` in the repo root (sheet names `6-8`, `9-12`, `Teacherwise`; otherwise update `sources.json`).
2. `npm run convert:check` and fix what it reports (docs/IMPORTING.md has the table): new spellings →
   `lib/vocabulary.js`; new teachers → `registry.js` (`TEACHERS`, `NEW_TEACHERS`); new/removed classes → `CLASSES`;
   new timings → `PERIODS`. Never guess an unknown word — ask the user.
3. `npm run convert` → read `docs/TIMETABLE_REVIEW_2026-27.md` (teachers filled in, groups without a teacher,
   sheet disagreements, clashes); **show it to the user and wait for approval**.
4. `npm test`.
5. Only with the user's explicit permission: commit, `clasp push`, then tell the user to reload the sheet and run
   **Import All Pre-fill Data** (answer **NO** to load the new built-in data) → **Check For Conflicts**.

## 9 · Testing

`npm test` loads every deployable `.gs` file into a VM with `tests/harness/gas-mock.js` and drives real menu functions and
`onEdit` events. The mock stores values, backgrounds and validation lists; other formatting calls are no-ops, and Apps Script
quotas/permissions are not simulated — check those manually (README → "Manual checks after deploying").

## 10 · Post-deploy checklist (after `clasp push`)

| What changed | Run |
|---|---|
| Any `.gs` file | Reload the sheet |
| View managers | 🔄 Refresh All Views |
| `PrefillData.gs` | 📥 Import All Pre-fill Data (answer NO to use the built-in data) |
| Trigger logic / first deploy of this version | ⚙️ Install Edit Sync Trigger |
| Validation | ✅ Check For Conflicts |

## 11 · Gotchas

1. `@OnlyCurrentDoc` is ignored when `appsscript.json` lists explicit scopes; scope changes require re-authorisation.
2. Don't add a manual `onEdit` trigger — use the menu item (it de-duplicates).
3. An old `_Prefill_Data_Backup` snapshot is offered on import; answer NO to load new built-in data.
4. 11th C and 12th C are single classes; their Hindi- and English-medium groups are split periods marked `(H.M)` / `(E.M)`.
   (The older data had separate `11th C (H)` / `(E)` sections; `splitMedium` still reads those marks.)
5. `PrefillData.gs` is generated — edit the converter, not the file. `RAW_PERIOD_ROWS` drives the bell times in every
   view header; `CONFIG.PERIOD_TIMES` is only a fallback and a test keeps it in step.
