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
| `Teacher_View` | Pick a teacher in **B3** → weekly grid. Type `6th A - Mathematics` to take that period, or `FREE` to clear. |
| `Teacher_Day_View` | Pick a day (**B3**) and optional class filter (**E3**) → all teachers × periods. Editable. |
| `Master_Grid_View` | Read-only overview of every class. |
| `Teacher_Availability_Grid` | FREE / BUSY / UNAVAILABLE for one day (menu → 📊 Teacher Availability). |
| `Cover_Manager` | Log of substitutions made with 👥 Cover Manager. |
| `Change_Log` | One row per period changed from a view or `Master_Schedule`: time, editor, where, before, after. Created on the first edit. |

Colours: green = free, blue = class, red = clash (`Teacher_Day_View` also shows combined sections in purple).

## First-time setup (new Google Sheet)

1. Install Node.js 18+ and run `npm install` in this folder.
2. `npx clasp login`, then create `.clasp.json` (git-ignored) pointing at your script:
   `{"scriptId": "<your script id>", "rootDir": "."}` — or `npx clasp create --type sheets` for a new sheet.
3. `npx clasp push` (uploads the `.gs` / `.html` files listed in the project; local tooling is excluded by `.claspignore`).
4. Open the sheet, reload it, then run **Timetable System → ⚙️ Setup / Repair**.
   Approve the permissions prompt (spreadsheet access, dialogs, and triggers — the last one installs the edit-sync trigger).
5. Run **📥 Import Timetable**. If a saved snapshot exists you are asked which one to load.
6. Run **✅ Check For Conflicts** and review the `Clash Status` column.

## Updating an existing sheet to this version

1. `npx clasp push` → reload the sheet.
2. **⚙️ Setup / Repair** (one time; installs the edit-sync trigger, and the change trigger that refreshes the views
   after rows are deleted). Approve the new permission prompt: this version also reads the editor's email for `Change_Log`.
   If you previously added an `onEdit` trigger by hand, it is ignored from now on (you can delete it under Extensions → Apps Script → Triggers).
3. **Keeping the data already in the sheet** (you have been editing it): stop here — do **not** run Import Timetable.
   The new code reads the existing tabs as they are; Setup / Repair and Refresh only rewrite the header rows and the
   two computed columns (`Total Hours Scheduled`, `Clash Status`). Run **🔄 Refresh All Views**, then go to step 4.
   *Only to replace the data with the 2026-27 Excel*: **📥 Import Timetable** → answer **NO** to load the new 2026-27 timetable
   (**YES** reloads your old saved snapshot). This replaces the Teachers, Classes, Subjects and Master_Schedule tabs —
   use **💾 Save Snapshot** first if you want to keep a copy of the current state.
   The new timetable has 28 classes: 6th E is gone, and 11th C / 12th C are single classes whose medium groups are
   split periods, so old `11th C (H)` / `(E)` rows are replaced.
4. **✅ Check For Conflicts**. Rows you added by hand for a period that already had a row are now flagged
   *Duplicate Period*; editing that period in `Class_View` merges them into one row.

## Day-to-day use

- **Change a period**: edit `Class_View` (subject / teacher rows), `Teacher_View` / `Teacher_Day_View`
  (`Class - Subject`), or `Master_Schedule` directly. The other views, `Clash Status` and the teachers'
  period counts refresh automatically. Pasting a block of cells works too.
- **Split period (several subjects and teachers at once)**: edit it in `Class_View` or `Master_Schedule`.
  Type every group, in the same order, in both cells — subject `Hindi / IT / Art`, teacher
  `Asha / Bina + Chetan / Deepa` (` / ` between groups, ` + ` for two teachers sharing a group, `-` for a
  group with no teacher yet). A single name on a split period fills the next empty group (see below), so groups never lose their teachers.
- **Several teachers from the dropdown (split periods)**: in `Class_View` or the `Master_Schedule` Teacher column, pick names one at a time from the teacher
  dropdown. Each pick fills the next group that has no teacher, in subject order — subject `Hindi / IT / Art`, pick
  Asha, then Bina, then Deepa → `Asha / Bina / Deepa`. To start again, clear the teacher cell and pick again.
  Two teachers sharing one group still need typing: `Bina + Chetan`. Teacher dropdowns list names A→Z.
- **Teacher_View / Teacher_Day_View**: typing `6th A - Science` makes that teacher take the period; whoever taught it
  is replaced and a "Saved" message names them. On a split period, type one of its subjects to take that group.
  Co-teachers and new split groups are set up in `Class_View`.
- **What gets refused** (in `Class_View` and in single Subject / Teacher cells of `Master_Schedule`; a "Not saved"
  message explains why, and the old value is put back): an unknown teacher or class
  name, a teacher who already teaches another class at that time (a combined lecture — same subject, same grade —
  is fine), subjects and teacher groups that don't line up, a teacher name typed in the subject row, and clearing
  only the subject. Clear both cells to free a period.
- **Combined sections**: give each section's row the same subject and teacher; it shows as one combined lecture, not a clash.
  The cell note in `Class_View` says which group is shared with which sections.
- **Teacher absent**: 👥 Cover Manager → pick the date (the timetable day is worked out for you) and the absent teacher →
  assign a free teacher per period. Teachers who are busy, marked unavailable that day, absent, or already covering are not offered.
- **Teacher on leave for a weekday**: put the day in `Teachers → Days Unavailable` (e.g. `Monday, Friday`).

## See the timetable locally (no push)

```bash
npm run preview          # writes preview/timetable.html — open it in any browser
npm run preview:serve    # also serves it at http://127.0.0.1:8765 (this machine only; Ctrl+C to stop)
```

The page shows the Class, Teacher and Day views, the all-classes grid and every conflict, with the same text,
colours and notes the sheet would show after **📥 Import Timetable → NO**. It is rendered by the real `.gs` code
against the in-memory Sheets mock the tests use, so fonts and column widths differ. Re-run after `npm run convert`.
`preview/` is ignored by git and clasp.

## Optional: read-only timetable page for teachers

`WebApp.gs` + `TimetableWeb.html` serve a mobile-friendly, printable class/teacher timetable. Nothing is reachable until you deploy it:

1. Apps Script editor → **Deploy → New deployment → Web app**.
2. **Execute as: User accessing the web app** (required — with "Me", any visitor could call the write functions with your access).
3. **Who has access:** your school domain or "Anyone with Google account".
4. Share the spreadsheet with teachers as **Viewer**, then send them the web-app URL.

## Automatic push from GitHub

`.github/workflows/clasp-push.yml` runs the tests and then `clasp push --force` on every push to `master`.
It only goes GitHub → Apps Script: edits made in the Apps Script editor are overwritten by the next push.

One-time setup:

1. Turn on the Apps Script API at https://script.google.com/home/usersettings.
2. Run `npx clasp login` locally, then add the whole of `~/.clasprc.json` as a repository secret named `CLASPRC_JSON`
   (Settings → Secrets and variables → Actions). It grants access to your Google account's Apps Script and Drive — keep it secret.
3. Add a repository variable `SCRIPT_ID` with the script ID from your `.clasp.json`.

If the run fails with an auth error, the login has expired: run `npx clasp login` again and update the secret.
A versioned web-app deployment still needs **Deploy → Manage deployments** (or `clasp deploy`) after a push.

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
2. In `Teacher_View`, type `7th B - Mathematics` into a FREE cell → `Class_View` for 7th B shows it, the teacher
   disappears from any other class in that period, and a "Saved" message names anyone replaced.
3. Paste three subjects across a `Class_View` row → all three appear in `Master_Schedule`.
4. In `Class_View`, type one teacher on a split period → "Not saved" message, the cell goes back.
   (Also check the message reaches the person editing when the trigger was installed by someone else.)
5. Delete a row in `Master_Schedule` → `Class_View` drops that period within a few seconds.
6. Type in `Master_Grid_View` → Sheets warns that the tab is generated.
7. 👥 Cover Manager → assign a cover → a row appears in `Cover_Manager`; assigning the same class/period again is refused.
8. **✅ Check For Conflicts** → summary dialog; flagged rows show the reason in `Clash Status`.

## Troubleshooting

| Problem | Fix |
|---|---|
| Edits in a view don't reach `Master_Schedule` | An editor runs **⚙️ Setup / Repair** (the trigger runs under the account of whoever installed it last — re-run it if that person lost access). Extra `onEdit` triggers are harmless: only the last installed one does the work. |
| A teacher is missing from `Teacher_Day_View` | The name in `Master_Schedule` doesn't match `Teachers`. **Check For Conflicts** lists it as *Unregistered Teacher*. |
| "unknown class" toast after typing in `Teacher_View` | Use the exact class name, e.g. `6th A - Mathematics`. |
| Import loaded old data | Import again and answer **NO** (built-in timetable). |
| "Not saved — another timetable edit is still being processed" | Two edits overlapped. What you typed is still on screen but was not saved; re-enter it. The views are shared, so two people using the same view switch each other's selection (B3). |
| Ctrl+Z after a view edit | Undo only reverts the view cell, not `Master_Schedule`. Type the old value back instead; `Change_Log` shows what it was. |
| Script times out on very large pastes | Paste into `Master_Schedule`, then **🔄 Refresh All Views**. |
