# Spreadsheet Schema

All tabs are created by **⚙️ Setup / Repair** (`setupInitialSpreadsheet()`).
Names and column positions are defined once in `Config.gs` (`CONFIG.SHEETS`, `CONFIG.COL`).

## Data tabs (admin-edited)

### `Teachers`
| Col | Header | Notes |
|---|---|---|
| A | Teacher Name | Must match the names used in `Master_Schedule` exactly (case and spaces are ignored). |
| B | Subject Specialization | Informational. |
| C | Max Hours / Week | Periods per week; **Check For Conflicts** reports teachers above it. |
| D | Days Unavailable | e.g. `Monday, Friday`. Used by validation, availability grid and cover manager. |
| E | Total Hours Scheduled | Filled automatically on refresh/validation (a combined lecture counts once). |

### `Subjects`
Column A `Subject Name` feeds the subject dropdowns. Multi-subject cells (`Hindi / IT`) are listed as one value.

### `Classes`
| Col | Header | Notes |
|---|---|---|
| A | Class Name | e.g. `6th A`, `11th C`. Row order = order in all dropdowns/views. |
| B | Academic Tier | `Middle`, `Secondary`, `Higher Secondary`. |
| C | Room Assigned | Blank or `N/A` = no room (ignored by the room-clash check). |

### `Master_Schedule` (source of truth)
| Col | Header | Notes |
|---|---|---|
| A | Day | `Monday` … `Saturday` |
| B | Period | `1` … `8` |
| C | Class | One class per row (shared lectures are repeated on each section's row). |
| D | Academic Tier | Copied from `Classes`. |
| E | Subject | One subject, or the groups of a split period: `Sanskrit / IT / Beauty`. Medium groups: `Physics (H.M) / Chemistry (E.M)`. |
| F | Teacher | Same order as the subjects: `Neha Jamidar / Divya Lodval / Krishna Patale`. `+` joins two teachers of one group; `-` = no teacher. |
| G | Room | Optional. |
| H | Clash Status | Written by **Check For Conflicts**. |

## Generated tabs

| Tab | Editable | Layout |
|---|---|---|
| `Class_View` | Yes | B3 = class selector; row 5 = period headers; rows 6–17 = subject/teacher row pair per day. |
| `Teacher_View` | Yes | B3 = teacher selector; rows 6–11 = Mon–Sat; cell = `Class` + newline + `Subject`. |
| `Teacher_Day_View` | Yes | B3 = day, E3 = class filter; one row per teacher. |
| `Master_Grid_View` | No | One row per class; each cell lists entries with the days they run (`[Mon,Tue]`). |
| `Teacher_Availability_Grid` | No | Generated on demand for one day. |
| `Cover_Manager` | Append-only | Date, Absent Teacher, Period, Class to Cover, Suggested Available Teachers, Assigned Cover Teacher. |
| `Change_Log` | Append-only, written by the script | Time, Edited By, Edited In, Day, Period, Class, Before, After. One row per changed period. Warning-only protection. |
| `_Prefill_Data_Backup` | Hidden | Snapshot from **💾 Save Snapshot** (A:C classes, E:I teachers, K subjects, M:T schedule). |
