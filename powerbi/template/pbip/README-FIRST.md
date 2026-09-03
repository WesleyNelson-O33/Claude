# Utilisation — semantic model for Power BI

## Read this first

The last attempt failed with `Cannot find file 'version.json'`. That error is worth
understanding, because it is good news: **Power BI parsed the semantic model fine.** It
fell over reading the *report* definition — the PBIR scaffolding, which is the one part
I cannot verify from here.

So there are two routes below. **Route A is reliable** — it lets Power BI Desktop
generate the report scaffolding itself, which is the only thing it knows how to write
correctly, and supplies only the semantic model, which is already proven to parse.

## Route A — the one that will work

1. Open Power BI Desktop. **File > New.**
2. **File > Save as** > change "Save as type" to **Power BI project files (.pbip)**.
   Save it as **`Utilisation`** into `C:\Utilisation\`.
   (Name it exactly `Utilisation` — it has to match the model name.)

   Desktop creates:
   ```
   C:\Utilisation\
       Utilisation.pbip
       Utilisation.Report\           <- Desktop's own, valid, leave it alone
       Utilisation.SemanticModel\
   ```

3. **Close Power BI Desktop.**

4. In `C:\Utilisation\Utilisation.SemanticModel\`, **delete the `definition` folder**
   Desktop just made, and copy in the `definition` folder from this zip
   (`Utilisation.SemanticModel\definition`) in its place.

   Leave Desktop's `definition.pbism` and `.platform` exactly as they are — only the
   `definition` folder gets swapped.

5. Put your data alongside:
   ```
   C:\Utilisation\
       August_2026_Utilisation_Report.xlsx
       mapping\                      <- the three CSVs from this zip
   ```

6. **Double-click `Utilisation.pbip`.** Set the four parameters, refresh, and
   File > Save As > `.pbix` if you want a single file.

| Parameter | Value |
|---|---|
| `UtilWorkbookPath` | `C:\Utilisation\August_2026_Utilisation_Report.xlsx` |
| `MappingFolderPath` | `C:\Utilisation\mapping` (no trailing backslash) |
| `FY_Start_Year` | `2024` |
| `FY_End_Year` | `2027` |

## Route B — try the project as shipped

I have added a `version.json` to the report definition, which is the file the error
named. I could not verify its exact contents — Microsoft's schema host and every blog
documenting it are blocked from where I am working — so this is a best guess.

Unzip to `C:\Utilisation\` and double-click `Utilisation.pbip`. If it opens, you saved
yourself five minutes. If it throws the same class of error, use Route A and do not
spend a third attempt on it.

## What you are getting either way

7 tables, 36 measures, 6 relationships, 4 parameters.

```
Fact_Timesheet  ──  Dim_Date, Dim_Employee, Dim_Job, Dim_PayType
Fact_Capacity   ──  Dim_Date, Dim_Employee
```

Report pages are yours to build in the UI — `powerbi/docs/report-pages.md` has the
spec. Visual layout is the part that is not reliably hand-authorable, which is exactly
what this error demonstrated.

## Before you trust a number

Four cards, all should read zero:

- `Hours With No Job Number` — 188.25 for August as things stand
- `Unmatched Employee Hours`
- `Unmapped Job Hours`
- `Unmapped Pay Type Hours`

And this pair tells you the month is complete:

- `Working Days With Data` vs `Working Days In Period` — 15 against 21 is August 2026

## Still needs a human

- **`map_employees.csv`** — `Standard Weekly Hours` is 38 for everyone. Fix the
  part-timers. 12 rows carry `NEEDS REVIEW`, three of them employee IDs with two
  different people on overlapping dates.
- **`map_jobs.csv`** — `Last Fully Billed Date` is blank. Fill it from the month each
  job's WIP dollar balance returns to zero.

## If Route A also fails

`SETUP.md` one folder up: six M queries to paste, then `measures.csx` in Tabular
Editor. 45 minutes, no format guessing, certain to work.
