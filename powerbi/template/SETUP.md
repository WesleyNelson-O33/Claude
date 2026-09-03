# Building the model — the working path

## What went wrong with the .pbit

I built one and it does not open. Three faults, one of them fatal:

1. **No `DataMashup` part.** This is the fatal one. Power BI stores all Power Query M
   in a `DataMashup` part — a proprietary serialised binary stream, not the plain JSON
   I assumed. Putting the M in the model partitions alone is not enough for Desktop to
   accept the file.
2. `[Content_Types].xml` was written last in the zip. Open Packaging Conventions
   requires it first.
3. The content-type values were empty strings, which is invalid OPC.

I could fix 2 and 3 blind. I cannot reliably hand-write a `DataMashup` stream, and I
have no Power BI Desktop here to test against, so I am not going to keep guessing at a
binary format with your time. The `.pbit` has been removed rather than left sitting
there looking usable.

`build_pbit.py` is kept — it is the single source of truth that generates everything
below, so the model definition lives in one place.

## The path that works

About 45 minutes, and you end up with a real `.pbix` you own.

### 1. Put the files where the queries expect them

```
C:\Utilisation\
    August_2026_Utilisation_Report.xlsx
    mapping\
        map_employees.csv
        map_jobs.csv
        map_paytypes.csv
```

(Any path works — you set it in step 2. OneDrive or SharePoint is better if you want
scheduled refresh later without a gateway.)

### 2. New Power BI file, create four parameters

Home > Transform data > Manage Parameters > New:

| Name | Type | Current value |
|---|---|---|
| `UtilWorkbookPath` | Text | `C:\Utilisation\August_2026_Utilisation_Report.xlsx` |
| `MappingFolderPath` | Text | `C:\Utilisation\mapping` (no trailing backslash) |
| `FY_Start_Year` | Whole Number | `2024` |
| `FY_End_Year` | Whole Number | `2027` |

Names must match exactly — the queries reference them.

### 3. Paste the six queries, in this order

New Source > Blank Query > Advanced Editor > paste > name it exactly as the filename
says. Order matters because the queries reference each other.

1. `queries/1-Dim_Date.m`
2. `queries/2-Dim_Employee.m`
3. `queries/3-Dim_Job.m`
4. `queries/4-Dim_PayType.m`
5. `queries/5-Fact_Timesheet.m` — reads 47,000 rows, give it a minute
6. `queries/6-Fact_Capacity.m`

Close & Apply.

### 4. Add the measures and relationships

**With Tabular Editor 2** (free, tabulareditor.com — takes 2 minutes to install):
External Tools > Tabular Editor > "C# Script" tab > paste `measures.csx` > F5 > Ctrl+S.

That creates all 36 measures with their format strings and descriptions, all 6
relationships, the sort-by columns, marks `Dim_Date` as the date table, and hides the
plumbing columns. This is by far the faster route.

**Without Tabular Editor:** the relationships are in `powerbi/docs/data-model.md` and
the measures in `powerbi/dax/measures.dax`, to add by hand. Budget an hour, and do the
relationships first or the measures will error as you paste them.

### 5. Check the data quality measures before anything else

Put these on a page as cards. All four should read zero:

- `Hours With No Job Number` (188.25 for August as things stand)
- `Unmatched Employee Hours`
- `Unmapped Job Hours`
- `Unmapped Pay Type Hours`

And this pair tells you whether the month is complete:

- `Working Days With Data` vs `Working Days In Period`

15 against 21 is what August 2026 looks like. That is the check the spreadsheet never had.

### 6. Build the pages

`powerbi/docs/report-pages.md` — Summary, By person, Where the time went, WIP hours,
Data quality. About an hour in the UI. Build Data Quality first.

## Still to do in the mapping files

- **`map_employees.csv`** — `Standard Weekly Hours` defaults to 38 for everyone. Fix the
  part-timers. `NEEDS REVIEW` is set on 12 rows, including three employee IDs with two
  different people booking time on overlapping dates.
- **`map_jobs.csv`** — `Last Fully Billed Date` is blank. Populate it from the month each
  job's dollar balance on the WIP schedule returns to zero, and `Unbilled WIP Hours`
  becomes a true balance rather than hours booked to date.
