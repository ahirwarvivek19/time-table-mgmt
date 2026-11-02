# Timetable Sheet — Operating SOP

For school office staff and the timetable in-charge. All menu items are under **Timetable System** in the sheet's top menu bar.

## 1. First-time setup after an update

1. Reload the sheet (browser refresh) so the **Timetable System** menu appears.
2. Run **Timetable System → ⚙️ Setup / Repair**.
3. When Google asks for permission, choose your school account and click **Allow**.
4. Run **🔄 Refresh All Views**.
5. Run **✅ Check For Conflicts** and look at the `Clash Status` column in `Master_Schedule`.
6. Do **not** run **📥 Import Timetable**. Use it only when the in-charge has decided to replace all data with the new Excel timetable.

## 2. The tabs

| Tab | What it is for |
|---|---|
| `Master_Schedule` | The full timetable: one row per class, day and period. Everything else comes from here. |
| `Class_View` | One class's week. Pick the class in **B3**. Editable. |
| `Teacher_View` | One teacher's week. Pick the teacher in **B3**. Editable. |
| `Teacher_Day_View` | All teachers for one day. Pick the day in **B3**, optional class in **E3**. Editable. |
| `Master_Grid_View` | Every class at a glance. Read only. |
| `Teachers` | Teacher list: name, subject, `Max Hours / Week`, `Days Unavailable`. |
| `Classes` | Class list and rooms. |
| `Subjects` | Subject list used in the dropdowns. |
| `Cover_Manager` | Record of substitutions made with 👥 Cover Manager. |
| `Change_Log` | Who changed which period, and the old and new value. |

Colours: green = free, blue = class, red = clash, purple = combined sections (in `Teacher_Day_View`), amber = DAY OFF. In `Class_View`, amber marks a split period.

The grey **PRAYER** (9:10–10:00) and **LUNCH** (12:50–1:20) columns in every view are for display only. Anything typed there is not saved.

## 3. Daily tasks

### Change one period (Class_View)
1. Open `Class_View` and pick the class in **B3**.
2. Find the day. Each day has two rows: subject (top) and teacher (below).
3. Type or pick the subject in the subject row.
4. Pick the teacher in the teacher row, under that period.
5. To make the period free, clear the **teacher** cell first, then the subject cell (or select both and press Delete). The row stays in `Master_Schedule` with its room, so the next edit fills it again.
6. A normal edit that works shows no message; a message appears only when something is refused or worth knowing.

### Change a split period (several subjects and teachers)
1. In `Class_View` (or `Master_Schedule`), type all subjects in the subject cell, separated by ` / `: `Hindi / IT / Art`.
   If the period already had a teacher, they keep the first group (`Asha / - / -`).
2. Type the teachers in the **same order** in the teacher cell: `Asha / Bina / Deepa`.
3. Two teachers sharing one group: join them with ` + `: `Asha / Bina + Chetan / Deepa`.
4. A group with no teacher yet: type `-` in its place: `Asha / - / Deepa`.
5. Or pick teachers one at a time from the teacher dropdown: each pick fills the next empty group in subject order.
6. To start again, clear the teacher cell and pick again.
7. Two teachers in one group (`Bina + Chetan`) must be typed, not picked.

### Give a teacher a period, or free them (Teacher_View / Teacher_Day_View)
1. In `Teacher_View` pick the teacher in **B3** (or in `Teacher_Day_View` pick the day in **B3** and find the teacher's row).
2. To give a period: type `Class - Subject` in the cell, e.g. `6th A - Mathematics`.
3. The teacher takes that period; whoever taught it before is replaced, and a **Saved** message names them.
4. To free the teacher: type `FREE` (or clear the cell).
5. On a split period, type just one of its subjects to take that group.
   `FREE` on a split period keeps the group with `-` as its teacher, so someone else can take it.
6. Co-teachers and new split groups are set up in `Class_View`, not here.

### Combined sections (one teacher, several sections together)
1. Give each section the same subject and the same teacher for that period (in `Class_View` for each section).
2. Or in `Teacher_View` type the sections separated by `, `: `11th A, 11th B - Physics`.
3. This shows as one combined lecture, not a clash.

### A teacher is absent today
1. Run **👥 Cover Manager**.
2. Check the **Date** (today by default); the timetable day is filled in for you.
3. Pick the **Absent teacher**. Their periods for that day are listed.
4. For each period, pick a cover teacher from the list and click **Assign**.
5. Only free, available teachers who are not already covering are offered.
   If both teachers of a split period are absent, each gets their own cover. One cover teacher can take all sections of a combined class.
6. Each cover is recorded in the `Cover_Manager` tab.

### A teacher is on leave on a fixed weekday
1. Open the `Teachers` tab and find the teacher.
2. In `Days Unavailable` type the day(s): `Monday, Friday`, `Mon & Fri` or `Mon-Wed`.
3. The views update on their own; if one still looks old, run **🔄 Refresh All Views**.
4. In `Teacher_Day_View` that teacher now shows DAY OFF on those days.
5. Run **✅ Check For Conflicts**: any period still booked on that day shows *Unavailable Day*; move it to another teacher.

### A new teacher joins
1. Add a row in the `Teachers` tab: name, subject, `Max Hours / Week`, `Days Unavailable` (if any).
2. The name appears in the dropdowns on its own (run **🔄 Refresh All Views** if it does not).
3. Give them periods in `Teacher_View` or `Class_View`.

### A teacher leaves
1. Open `Teacher_View` and pick the teacher.
2. Give each of their periods to another teacher (in that teacher's `Teacher_View`, type `Class - Subject`), or type `FREE`.
   If the leaver is in two classes in the same period (a clash), give each class to a different teacher: one teacher can take only one of them.
3. When their week is empty, delete their row from the `Teachers` tab.
4. Run **🔄 Refresh All Views**, then **✅ Check For Conflicts**.

### Check for conflicts
1. Run **✅ Check For Conflicts**. A summary appears.
2. Open `Master_Schedule` and read the `Clash Status` column (filter out blanks to see only problems).
3. Fix each flagged row, then run **✅ Check For Conflicts** again until it is clean.

## 4. Rules to keep the data safe

- Never run **📥 Import Timetable** for day-to-day work: it replaces the Teachers, Classes, Subjects and Master_Schedule tabs.
- Run **💾 Save Snapshot** before any big change.
- Only one person edits a given view tab (`Class_View`, `Teacher_View`, `Teacher_Day_View`) at a time. Others use a different tab or `Master_Schedule`.
- Renaming a teacher in the `Teachers` tab does not rename them in the timetable. A **Name changed** message says how many periods still use the old name; change those in `Master_Schedule` (Edit → Find and replace).
- Do not type in `Master_Grid_View`; it is rebuilt automatically.
- Pasting into `Master_Schedule` can show **Check this cell**: if the paste was wrong, press Ctrl+Z.
- Ctrl+Z in a view tab does not undo the change; type the old value back (see `Change_Log`).
- A period with two rows for the same class (*Duplicate Period*) must be fixed in `Master_Schedule`: delete or correct the extra row.

## 5. Messages you may see

| Message | What to do |
|---|---|
| **Not saved** | The edit was refused and the old value put back. Read the reason in the message and correct it. |
| **Check this cell** | A paste in `Master_Schedule` was kept but has a problem. Fix it, or press Ctrl+Z. |
| **Saved** | Done. It may name a teacher who was replaced, or a group that now has no teacher. |
| **Not saved — another timetable edit is still being processed** (may appear as **⚠️ Timetable sync failed**) | Two edits overlapped. Wait a few seconds and type it again. |
| **⚠️ Sync not running** | The edit may not have been saved. Run **⚙️ Setup / Repair**, then make the edit again. |
| **this period is split into … groups** | In `Class_View`: type all subjects in order, e.g. `Sanskrit / Music`. In `Teacher_View`: type `Class - Subject` with one of the subjects listed in the message. |
| **this period has … groups … teacher groups** | The subjects and teachers don't match in number. Clear the teacher cell, then pick the teachers again in subject order. |
| **one subject with … teacher groups** (with **Check this cell**) | Fine if those teachers share the class. If it was a split period, type every subject in order, or press Ctrl+Z. |
| **… is marked unavailable on …** (with **Saved**) | The teacher was booked on their day off. Check it is intended, or give the period to someone else. |
| **Name changed** | Some periods still use the old name. Change them in `Master_Schedule`. |
| **Prayer and Lunch columns are not part of the timetable** | You typed in a grey PRAYER or LUNCH column. Type in a period column instead. |
| *Duplicate Period* | The class has two rows for that period. Delete or fix the extra row in `Master_Schedule`. |
| *Teacher Clash* | The teacher is in two classes at once. Give one of the periods to someone else. |
| *Unavailable Day* | The teacher has a period on a day in their `Days Unavailable`. Move it, or correct `Days Unavailable`. |
| **unknown teacher** | The name is not in the `Teachers` tab. Pick from the list, or add the teacher first. |
| **unknown class** | Use the exact class name, e.g. `6th A - Mathematics`. |
| **subjects and teacher groups don't line up** | Count the subjects and teacher groups: they must match, in the same order (use `-` for a group with no teacher). |
| **already teaches … at this time** | Free the teacher from the other class first, or move them in `Teacher_View`. |

## 6. Troubleshooting

| Problem | Fix |
|---|---|
| Edits in a view don't reach the other tabs | An editor runs **⚙️ Setup / Repair**. |
| A view looks old or out of date | Run **🔄 Refresh All Views**. |
| The **Timetable System** menu is missing | Reload the sheet and wait a few seconds. |
| A teacher is missing from `Teacher_Day_View` | Their name in `Master_Schedule` differs from `Teachers`. **✅ Check For Conflicts** lists it as *Unregistered Teacher*. |
| "Edit too large to sync" after a big paste | Paste into `Master_Schedule` instead, then run **🔄 Refresh All Views**. |
| Import loaded old data | Ask the in-charge: **File → Version history → See version history**, open the version from before the import and click **Restore this version**. Running Import again does not bring the old data back. |
