# Utilisation — Power BI Project (PBIP)

## Open it

1. Unzip this folder somewhere sensible, e.g. `C:\Utilisation\`.
2. Put your files where the parameters expect them:

```
C:\Utilisation\
    August_2026_Utilisation_Report.xlsx
    mapping\
        map_employees.csv
        map_jobs.csv
        map_paytypes.csv
    Utilisation.pbip              <- double-click this
    Utilisation.SemanticModel\
    Utilisation.Report\
```

3. **Double-click `Utilisation.pbip`.** Power BI Desktop opens the project.
   (Requires Desktop from roughly 2024 onward. If the .pbip option is greyed out,
   turn on File > Options and settings > Options > Preview features >
   "Power BI Project (.pbip) save option".)

4. It will ask for the four parameters, or you set them via
   Transform data > Manage Parameters:

| Parameter | Value |
|---|---|
| `UtilWorkbookPath` | `C:\Utilisation\August_2026_Utilisation_Report.xlsx` |
| `MappingFolderPath` | `C:\Utilisation\mapping` (no trailing backslash) |
| `FY_Start_Year` | `2024` |
| `FY_End_Year` | `2027` |

5. Refresh. It reads 47,000 timesheet rows out of the Excel workbook — give it a minute.

6. **File > Save As > `Utilisation.pbix`** if you want a single-file version.

## What is in it

7 tables, 36 measures, 6 relationships, 4 parameters.

```
Fact_Timesheet  ──  Dim_Date, Dim_Employee, Dim_Job, Dim_PayType
Fact_Capacity   ──  Dim_Date, Dim_Employee
```

Five empty report pages are set up — Summary, By person, Where the time went,
WIP hours, Data quality. Build the visuals in the UI following
`powerbi/docs/report-pages.md`; visual layout is not reliably hand-authorable.

## Do this before you trust a number

Put these four measures on a page as cards. All should read zero:

- `Hours With No Job Number` — 188.25 for August as things stand
- `Unmatched Employee Hours`
- `Unmapped Job Hours`
- `Unmapped Pay Type Hours`

And this pair tells you whether the month is complete:

- `Working Days With Data` vs `Working Days In Period`

15 against 21 is what August 2026 looks like. That is the check that would have
stopped the 50.2% going out.

## Still needs a human

- **`map_employees.csv`** — `Standard Weekly Hours` defaults to 38 for everyone.
  Fix the part-timers. `NEEDS REVIEW` is set on 12 rows, three of which are employee
  IDs with two different people booking time on overlapping dates.
- **`map_jobs.csv`** — `Last Fully Billed Date` is blank. Fill it from the month each
  job's dollar balance on the WIP schedule returns to zero, and `Unbilled WIP Hours`
  becomes a true balance instead of hours booked to date.

## If it will not open

The manual path is in `SETUP.md` one folder up: six M queries to paste, then
`measures.csx` in Tabular Editor for the measures and relationships. About 45 minutes
and it is certain to work.

Tell me the exact error text either way — "unable to open" alone does not tell me
which part it rejected.
