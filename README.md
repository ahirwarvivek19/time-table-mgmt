# Time Table Management System

School timetable on a Google Sheet, driven by Google Apps Script. One flat
`Master_Schedule` tab is the source of truth; class, teacher and day views are
generated from it and can be edited directly — edits sync back automatically.

Current data: **2026-27** timetable from `Timetable.xlsx`, 28 sections (6th A – 12th D), 78 teachers,
1,344 periods (6 days × 8 periods, 9:30–3:30 with lunch 12:20–12:50).

## Features

- **One source of truth, many editable views.** `Master_Schedule` holds every period; `Class_View`, `Teacher_View` and
  `Teacher_Day_View` are generated from it, can be edited directly (including pasted blocks), and sync back automatically.
- **Real-school periods.** Split classes (groups with their own subject, teacher and room), co-teaching, Hindi- / English-medium
  groups (`(H.M)` / `(E.M)`), and combined sections (one teacher, several sections) — combined lectures are not clashes.
- **Conflict check.** Teacher clashes, room clashes, unavailable days, unknown teachers or classes, and teachers over their weekly
  maximum, with the reason written next to every affected row. Teacher load (`Total Hours Scheduled`) is kept up to date.
- **Overviews.** `Master_Grid_View` (every class at a glance), `Teacher_Availability_Grid` (free / busy / unavailable for a day),
  bell times in every period header, and colour marks with hover notes for split, combined and clashing periods.
- **Cover Manager.** Pick a date and an absent teacher; it offers only teachers who are free, available and not already covering,
  and logs each substitution.
- **Excel import.** `npm run convert` reads the school's `Timetable.xlsx` automatically, fills in teachers the class sheets leave
  out using the teacher-wise sheet, cross-checks the two sheets, and writes a review report. Any other layout can be imported
  through a simple flat CSV. Snapshots of the sheet can be saved and re-imported.
- **Safe for several editors.** Edits are locked so two people cannot corrupt the schedule, and only one recorded trigger
  processes edits.
- **Optional read-only web page** for teachers (mobile-friendly, printable).
- **Tested locally.** `npm test` runs the real `.gs` code against an in-memory Sheets mock.

## What you get

| Tab | Use |
|---|---|
| `Teachers`, `Subjects`, `Classes` | Registries. `Total Hours Scheduled` is filled in automatically. |
| `Master_Schedule` | One row per class / day / period. Edit here or in any view. |
| `Class_View` | Pick a class in **B3** → weekly grid (subject row + teacher row per day). Editable. |
| `Teacher_View` | Pick a teacher in **B3** → weekly grid. Type `6th A - Mathematics`, or `FREE` to clear. |
| `Teacher_Day_View` | Pick a day (**B3**) and optional class filter (**E3**) → all teachers × periods. Editable. |
| `Master_Grid_View` | Read-only overview of every class. |
| `Teacher_Availability_Grid` | FREE / BUSY / UNAVAILABLE for one day (menu → 📊 Teacher Availability). |
| `Cover_Manager` | Log of substitutions made with 👥 Cover Manager. |

Colours: green = free, blue = class, red = clash (`Teacher_Day_View` also shows combined sections in purple).

## First-time setup (new Google Sheet)

1. Install Node.js 18+ and run `npm install` in this folder.
2. `npx clasp login`, then create `.clasp.json` (git-ignored) pointing at your script:
   `{"scriptId": "<your script id>", "rootDir": "."}` — or `npx clasp create --type sheets` for a new sheet.
3. `npx clasp push` (uploads the `.gs` / `.html` files listed in the project; local tooling is excluded by `.claspignore`).
4. Open the sheet, reload it, then run **Timetable System → ⚙️ Setup & Utilities → Setup Spreadsheets (Run Once)**.
   Approve the permissions prompt (spreadsheet access, dialogs, and triggers — the last one installs the edit-sync trigger).
5. Run **📥 Data Import → Import All Pre-fill Data**. If a saved snapshot exists you are asked which one to load.
6. Run **✅ Check For Conflicts** and review the `Clash Status` column.

## Updating an existing sheet to this version

1. `npx clasp push` → reload the sheet.
2. **⚙️ Setup & Utilities → Install Edit Sync Trigger** (one time; needed for reliable syncing of edits).
   If you previously added an `onEdit` trigger by hand, it is ignored from now on (you can delete it under Extensions → Apps Script → Triggers).
3. **📥 Data Import → Import All Pre-fill Data** → answer **NO** to load the new 2026-27 timetable
   (**YES** reloads your old saved snapshot). This replaces the Teachers, Classes, Subjects and Master_Schedule tabs —
   use **Save Current Sheet as Pre-fill Data** first if you want to keep a copy of the current state.
   The new timetable has 28 classes: 6th E is gone, and 11th C / 12th C are single classes whose medium groups are
   split periods, so old `11th C (H)` / `(E)` rows are replaced.
4. **✅ Check For Conflicts**.

## Day-to-day use

- **Change a period**: edit `Class_View` (subject / teacher rows), `Teacher_View` / `Teacher_Day_View`
  (`Class - Subject`), or `Master_Schedule` directly. The other views refresh automatically.
  Pasting a block of cells works too.
- **Several teachers or optional subjects in one period**: separate with ` / ` in the same order —
  subject `Hindi / IT` with teacher `Ekta Gour / Divya Lodval`. Mark medium groups as in the Excel:
  `Physics (H.M) / Chemistry (E.M)`. Two teachers sharing one group: `Divya Lodval + Reena Shivhare`.
- **Combined sections**: give each section's row the same subject and teacher; it shows as one combined lecture, not a clash.
- **Teacher absent**: 👥 Cover Manager → pick the date (the timetable day is worked out for you) and the absent teacher →
  assign a free teacher per period. Teachers who are busy, marked unavailable that day, absent, or already covering are not offered.
- **Teacher on leave for a weekday**: put the day in `Teachers → Days Unavailable` (e.g. `Monday, Friday`).

## See the timetable locally (no push)

```bash
npm run preview          # writes preview/timetable.html — open it in any browser
npm run preview:serve    # also serves it at http://127.0.0.1:8765 (this machine only; Ctrl+C to stop)
```

The page shows the Class, Teacher and Day views, the all-classes grid and every conflict, with the same text,
colours and notes the sheet would show after **Import All Pre-fill Data → NO**. It is rendered by the real `.gs` code
against the in-memory Sheets mock the tests use, so fonts and column widths differ. Re-run after `npm run convert`.
`preview/` is ignored by git and clasp.

## Optional: read-only timetable page for teachers

`WebApp.gs` + `TimetableWeb.html` serve a mobile-friendly, printable class/teacher timetable. Nothing is reachable until you deploy it:

1. Apps Script editor → **Deploy → New deployment → Web app**.
2. **Execute as: User accessing the web app** (required — with "Me", any visitor could call the write functions with your access).
3. **Who has access:** your school domain or "Anyone with Google account".
4. Share the spreadsheet with teachers as **Viewer**, then send them the web-app URL.

## Loading a new Excel timetable

Full guide: [docs/IMPORTING.md](docs/IMPORTING.md). In short:

```bash
npm run convert:check    # validate Timetable.xlsx (the sources listed in tools/convert/sources.json)
npm run convert          # writes PrefillData.gs and docs/TIMETABLE_REVIEW_<label>.md
node tools/convert/build.js --check --add flat:new.csv   # try any new file in the flat format
```

- Replace `Timetable.xlsx` (sheets `6-8`, `9-12`, `Teacherwise`) with the new version and run the two commands.
  The class-wise sheets are read automatically; teachers they leave out are filled in from the teacher-wise sheet
  when it names exactly one fitting teacher.
- Any other layout: convert it to the flat format (`templates/timetable-template.csv`).
- New teacher or new spelling? Add them to `tools/convert/registry.js` / `tools/convert/lib/vocabulary.js` —
  the converter refuses unknown names and words instead of guessing.
- Read the review file (teachers filled in, groups without a teacher, sheet disagreements, clashes) before importing.

## Tests

```bash
npm test
```

Runs the real `.gs` code against an in-memory Sheets mock (`tests/harness/gas-mock.js`): import, validation,
view edits and write-back, triggers, cover manager, and the converter. The mock does not check visual
formatting or Apps Script quotas — see "Manual checks" below.

## Manual checks after deploying

1. Import → `Class_View` shows **6th A**; `Teacher_View` shows the first teacher; `Master_Grid_View` exists.
2. In `Teacher_View`, type `7th B - Mathematics` into a FREE cell → `Class_View` for 7th B shows it, and the teacher
   disappears from any other class in that period.
3. Paste three subjects across a `Class_View` row → all three appear in `Master_Schedule`.
4. 👥 Cover Manager → assign a cover → a row appears in `Cover_Manager`; assigning the same class/period again is refused.
5. **✅ Check For Conflicts** → summary dialog; flagged rows show the reason in `Clash Status`.

## Troubleshooting

| Problem | Fix |
|---|---|
| Edits in a view don't reach `Master_Schedule` | An editor runs **Install Edit Sync Trigger** (the trigger runs under the account of whoever installed it last — re-run it if that person lost access). Extra `onEdit` triggers are harmless: only the last installed one does the work. |
| A teacher is missing from `Teacher_Day_View` | The name in `Master_Schedule` doesn't match `Teachers`. **Check For Conflicts** lists it as *Unregistered Teacher*. |
| "unknown class" toast after typing in `Teacher_View` | Use the exact class name, e.g. `6th A - Mathematics`. |
| Import loaded old data | Import again and answer **NO** (built-in timetable), or **Reset Pre-fill Data to Default**. |
| "Another timetable edit is still being processed" | Two edits overlapped; re-enter yours. The views are shared, so two people using the same view switch each other's selection (B3). |
| Script times out on very large pastes | Paste into `Master_Schedule`, then **🔄 Refresh All Views**. |
