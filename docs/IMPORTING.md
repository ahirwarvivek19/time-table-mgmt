# Importing a timetable

Turns the school's Excel timetable into `PrefillData.gs`, which the sheet loads with **📥 Import All Pre-fill Data**.

## Quick start

```bash
npm install
npm run convert:check    # validate only — fix every error it prints
npm run convert          # writes PrefillData.gs + docs/TIMETABLE_REVIEW_<label>.md
npm test
npx clasp push           # then in the sheet: 📥 Import All Pre-fill Data → NO (built-in data)
```

Read the review file before importing. It lists the teachers filled in from the teacher-wise sheet, groups that still have no teacher, where the class-wise and teacher-wise sheets disagree, clashes, teacher load, split periods and combined lectures, and every distinct cell with how it was read.

## A new version of `Timetable.xlsx`

Replace `Timetable.xlsx` in the repo root (same sheet names: `6-8`, `9-12`, `Teacherwise`) and run `npm run convert:check`. Then fix what it reports:

| Message | Do this |
|---|---|
| `unknown word "X"` in a cell | A new spelling of a subject or teacher: add it to `SUBJECTS` or `TEACHER_ALIASES` in `tools/convert/lib/vocabulary.js` |
| `Unregistered teacher "X"` | A new teacher: add them to `TEACHERS` in `tools/convert/registry.js` (and `NEW_TEACHERS`), plus any short form to `TEACHER_ALIASES` |
| `class "कक्षा 6 E" → "6th E" is not in CLASSES` / `no source gives a timetable for this class` | Add or remove the class in `CLASSES` in `registry.js` |
| `bell times … differ from PERIODS` | The school changed its timings: update `PERIODS` in `registry.js` (the view headers follow automatically) |
| `found N day rows, expected 6` | A class block is missing a weekday row, or has an extra one |
| A warning about the teacher-wise sheet | That cell is not used; the build still succeeds |

## Sources

`tools/convert/sources.json`:

```json
{ "label": "2026-27",
  "sources": [
    { "format": "class-blocks", "file": "Timetable.xlsx", "sheet": "6-8" },
    { "format": "class-blocks", "file": "Timetable.xlsx", "sheet": "9-12" } ],
  "teacherwise": { "file": "Timetable.xlsx", "sheet": "Teacherwise" } }
```

- Paths are relative to the repo root (or to the JSON file when you pass `--sources other.json`).
- **A class given by a later source replaces that class from earlier sources.** The build prints which classes were replaced.
- Every class in `CLASSES` (`registry.js`) must come from some source.
- `teacherwise` is optional. Without it nothing is filled in and nothing is cross-checked.

| Format | Reads |
|---|---|
| `class-blocks` | The class-wise sheets: a `कक्षा 9 A` block, Hindi weekday in column C, one cell per period. The period columns and bell times are read from the header rows (`1 … 8`, `09:30-10:20 …`). |
| `flat` | CSV or Excel, one row per class / day / period — for **any other layout**. |

## How a class-wise cell is read

`tools/convert/lib/cell-grammar.js`. Each subject word starts a group; teachers belong to the subject before them.

| Cell | Read as |
|---|---|
| `SANSKRIT NEHA JAMIDAR` | Sanskrit — Neha Jamidar |
| `SANSKRIT NEHA.J IT-DIVYA BEAUTY-KRISHNA` | **split**: Sanskrit / IT / Beauty, one teacher each |
| `PHYSICS GK.DIXIT (H.M) / CHEMISTRY D.BHATI (E.M)` | split by medium: `Physics (H.M) / Chemistry (E.M)` |
| `SO.SCI RUBAL VERMA(H.M)/ I.MATHUR (E.M)` | one subject, a group per medium |
| `CCLE INDRA.M/ KRISHNA.P`, `SST Lalita / Farhana` | one group, two teachers (co-teaching) |
| `MATHS /CCLE NEHA JHALA` | one group "Mathematics & CCLE" |
| `PLOY.SCI (RO.S)/ HINDI (ARCHANA.P)` | teacher in brackets |
| `GEOGRAPHY (PRACTICAL)` | Geography Practical |
| `DAKSHATA HINDI` | Proficiency Hindi |
| `GEOGRAPHY`, `COMPUTER` (no name) | no teacher written — see below |

## Teachers the class-wise sheets leave out

`tools/convert/teacherwise.js` reads the teacher-wise sheet (one row per teacher, one cell per period, e.g. `MON-TUE-10TH A WED-THU-10TH B+C`, `11TH C B.STU(E.M)`, `12TH C (ACC) MON TO THU`; classes listed without days apply to the days no other part of the cell names).

A group with no teacher gets one only when the teacher-wise sheet puts **exactly one fitting teacher** in that class at that time: the hint (subject, medium) must not contradict the group, the teacher must not be teaching something else then in the class-wise sheet, and they must teach that subject (to this class, elsewhere, or as their listed subject). Everything filled is listed in the review with the teacher-wise cell. `COMPUTER` with no name and no teacher-wise entry goes to Reena Shivhare, the only computer teacher. Anything else stays empty and is listed under "Groups still without a teacher".

The review also lists every teacher / period where the two sheets put a teacher in different classes. The class-wise sheets are what gets imported.

## Flat format (any other file)

Header row with `Class, Day, Period, Subject, Teacher, Room` (any order, any case; Room optional). Example: `templates/timetable-template.csv`.

| Case | How to write it |
|---|---|
| Normal period | `9th A, Monday, 1, Sanskrit, Neha Jamidar` |
| **Split class** (groups at once) | Subjects, teachers and rooms in the same order with `/`: `Hindi / IT / Beauty` · `Ekta Gour / Divya Lodval / Krishna Patale` · `12 / Lab / 14` |
| Medium groups | `Physics (H.M) / Chemistry (E.M)` |
| Co-teachers in one group | `+`: `Divya Lodval + Reena Shivhare` |
| Group with no teacher | `-`: `English / Library` · `Anuradha Katare / -` |
| **Combined sections** | One row per section, same subject and teacher. Not a clash. |
| Free period | Empty Subject (missing rows are also filled as free and reported) |

- Day: `Monday`, `Mon`, or `सोमवार`. Period: 1–8.
- Teacher names must be in `TEACHERS` (`registry.js`); case/spacing and the short forms in `TEACHER_ALIASES` are accepted.
- Save CSVs as **CSV UTF-8**.

```bash
node tools/convert/build.js --check --add flat:new-9th.csv          # try a file on top of the current sources
node tools/convert/build.js --add flat:new-9th.xlsx#Sheet1          # Excel: give the sheet after #
npm run convert:export-flat                                         # current data → timetable-flat.csv (edit, then re-import)
```

To switch to flat for good, add `{ "format": "flat", "file": "..." }` to `sources.json` (last, so it wins) and remove the old entries.

## Clash rules

- Same teacher, same period, **same subject, same grade** → combined lecture (OK). A medium mark on only one section (`Hindi` in 11th B, `Hindi (E.M)` in 11th C) is still the same lecture.
- Same teacher, same period, different subject, different grade, or `(H.M)` and `(E.M)` groups at once → **clash**. Activities such as Sports, Library or Music may span grades.
- Rooms: same room, same period, different lecture → room clash.

The converter and **✅ Check For Conflicts** in the sheet use the same rules (`ScheduleParser.classifyTeacherSlot`).
