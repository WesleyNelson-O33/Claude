# Utilisation report — complete project

Everything is in this one zip. No folder swapping.

## Setup

1. **Delete the old `Utilisationpbip` folder** in `C:\Utilisation\` (and the old zip).
2. **Extract this zip** into `C:\Utilisation\`. It does not matter if Windows creates a
   subfolder — the project reads your data by absolute path, so it works either way.
3. **Double-click `Utilisation.pbip`.**
4. **Refresh.** First load reads 47,000 timesheet rows out of the workbook — give it a
   minute or two.

**No parameters to set.** They are already pointed at the paths that work on your machine:

| Parameter | Value |
|---|---|
| `UtilWorkbookPath` | `C:\Utilisation\August 2026 Utilisation Report.xlsx` |
| `MappingFolderPath` | `C:\Utilisation` |
| `FY_Start_Year` | 2024 |
| `FY_End_Year` | 2027 |

Your workbook and the three CSVs stay exactly where they are in `C:\Utilisation\`. A copy
of the CSVs is in the `mapping` folder here only in case you ever lose the originals —
you do not need to move them.

If you later tidy the CSVs into `C:\Utilisation\mapping\`, change `MappingFolderPath`
to match. Nothing else changes.

## What is in it

7 tables, 36 measures, 6 relationships, 5 pages, 39 visuals.

**Summary** — Month Year and Person Group slicers, five cards (Utilisation %, variance in
points, chargeable hours, available hours, coverage %), rolling 13-week trend against
target, utilisation by team, hours mix, chargeable by job group, overtime, casual share,
and the data quality flag.

**By person** — one table, worst first. Coverage % sits beside utilisation on purpose:
answer "did they fill in a timesheet" before drawing any conclusion.

**Where the time went** — non-chargeable hours by job, overtime by person, and a
category × month matrix.

**WIP hours** — hours added, hours to date, unbilled, and a job-level table keyed on job
number so it drops beside your dollar schedule.

**Data quality** — the four zero-target cards, the working-days pair that catches a
partial month, coverage by person, and the pay type mapping check.

Month Year is synced across all five pages. Pick August once and it applies everywhere.

## Read the Data quality page first

Four cards, all should read zero:

- `Hours With No Job Number` — expect about 188.25 for August
- `Unmatched Employee Hours` — **this is the one that matters.** It tests whether the
  employee-ID-plus-name resolution held. Above zero means someone's hours are unattributed.
- `Unmapped Job Hours`
- `Unmapped Pay Type Hours`

And the pair that would have stopped the 50.2% going out:

- `Working Days With Data` vs `Working Days In Period` — 15 against 21 for August 2026

## Two things that look wrong but are not

- **`Utilisation vs Target (pp)`** reads oddly until you set real targets in
  `map_employees.csv`. Everyone currently has a group default (Support 0.88,
  Production 0.65, Consulting 0.50, Overhead 0).
- **`Unbilled WIP Hours`** equals `WIP Hours to Date` until `Last Fully Billed Date` is
  filled in. By design — see the baseline assessment.

## Still needs a human

- **`map_employees.csv`** — `Standard Weekly Hours` is 38 for everyone. Fix the
  part-timers. 12 rows carry `NEEDS REVIEW`, three of them employee IDs with two
  different people booking time on overlapping dates.
- **`map_jobs.csv`** — `Last Fully Billed Date` is blank.

## If a visual comes up blank

Valid JSON does not guarantee a rendered visual. Click the empty one, check the Fields
well, and send me its name from the Selection pane. Each visual is its own file, so I fix
that one and nothing else moves.
