# Utilisation.pbit — setup

## Why a .pbit and not a .pbix

A `.pbix` cannot be hand-authored. Its data model is a compiled Analysis Services
partition — proprietary, undocumented, and written only by Power BI Desktop itself.
Anything carrying that extension that I generated would not open.

A `.pbit` (Power BI Template) can, because its model is plain TMSL JSON. You open the
template, Desktop builds the model, and you save it as `.pbix` from there. Same
destination, one extra step.

**I could not test this file in Power BI Desktop** — there is no Desktop in this
environment. The model has been validated for JSON structure, encoding, M syntax
balance, and every measure / column / relationship reference resolving. If it fails to
open, the fallback is section 5.

## 1. Put the files somewhere

```
C:\Utilisation\
    August_2026_Utilisation_Report.xlsx      <- your existing workbook, unchanged
    mapping\
        map_employees.csv
        map_jobs.csv
        map_paytypes.csv
```

Use OneDrive or SharePoint rather than a local drive if you want scheduled refresh in
the Service without a gateway.

## 2. Open the template

Double-click `Utilisation.pbit`. Power BI Desktop prompts for four parameters:

| Parameter | Value |
|---|---|
| `UtilWorkbookPath` | `C:\Utilisation\August_2026_Utilisation_Report.xlsx` |
| `MappingFolderPath` | `C:\Utilisation\mapping` |
| `FY_Start_Year` | `2024` |
| `FY_End_Year` | `2027` |

Then **File > Save As > Utilisation.pbix**.

The first refresh reads 47,000 timesheet rows out of the Excel workbook and will take a
minute or two.

## 3. Check the Data Quality measures before trusting anything

Drop these on a page as cards. Every one should read zero:

- `Hours With No Job Number` — currently 188.25 for August
- `Unmatched Employee Hours`
- `Unmapped Job Hours`
- `Unmapped Pay Type Hours`

And these two together tell you whether you are looking at a complete month:

- `Working Days With Data` vs `Working Days In Period`

That pair is what would have caught the August 2026 problem: 15 against 21.

## 4. The mapping files — what still needs a human

Everything is pre-populated from your actual data. Three things need your eyes:

**`map_employees.csv`** — 168 rows, 162 employee IDs.
- `NEEDS REVIEW` is set on **12 rows**. Three IDs have two different people booking time
  under them on **overlapping dates** (`7118495`, `9569711`, `9771951`); the rest look
  like recycled IDs or renames. Each has been given a distinct `Employee Key`
  (`7118495-1`, `7118495-2`) so the model is correct either way — but confirm which are
  genuinely two people.
- `Utilisation Scope` defaults to `Delivery` for Support / Production / Consulting /
  Video / Onsite / Integration, `Overhead` for Management / Finance / Admin. **147
  Delivery, 21 Overhead.** Move anyone whose chargeable share runs above ~25%.
- `Standard Weekly Hours` defaults to 38 and `FTE` to 1 for everyone. **Fix the
  part-timers** — this is the column that carried the reverse-engineered plugs
  (`=23*5.130434`, `=23*2.52173913`) in the spreadsheet.
- `Target Utilisation %` defaults by group: Support 0.88, Production 0.65,
  Consulting 0.50, Overhead 0.

**`map_jobs.csv`** — 611 jobs, classified from your job numbering:
- `Recurring` (122) — five-digit legacy numbers, resident onsite contracts billed
  monthly. No WIP.
- `Project` (436) — date-coded numbers. These accrue WIP hours.
- `Non-chargeable` (53) — the 9000-series overhead codes.
- `Override Billing Basis` wins over the derived value where the rule gets one wrong.
- `Last Fully Billed Date` is blank. Populate it from the month each job's dollar
  balance on the WIP schedule returns to zero, and `Unbilled WIP Hours` becomes a true
  balance instead of hours booked to date.

**`map_paytypes.csv`** — 52 pay types, classified Ordinary / Leave / Public Holiday /
Overtime / Other Payment. The whitespace variants (` - Public Holiday` vs
`  - Public Holiday` vs ` - Public Holiday `) are normalised on a trimmed key, so the
three-different-formulas problem cannot recur.

## 5. If the template will not open

The model is also available as plain text in this repo, to paste in by hand:

- `powerbi/queries/*.m` — Power Query, one per table
- `powerbi/dax/measures.dax` — the measure library
- `powerbi/docs/data-model.md` — relationships and model housekeeping

Or regenerate the template after editing: `python3 build_pbit.py`

## 6. Report pages

The template ships five empty pages (Summary, By person, Where the time went, WIP hours,
Data quality). Visual layout is not reliably hand-authorable, so build them in the UI —
about an hour, following `powerbi/docs/report-pages.md`.

**Build the Data Quality page first** and get it to zero before showing anyone the
Summary.

## What is in the model

7 tables, 36 measures, 6 relationships, 4 parameters.

```
Fact_Timesheet  ── Dim_Date, Dim_Employee, Dim_Job, Dim_PayType
Fact_Capacity   ── Dim_Date, Dim_Employee
```

Key definitions, all settled in the baseline assessment:

- `Utilisation %` = chargeable ordinary hours ÷ available hours, Delivery scope only.
- `Available Hours` = capacity − leave. Public holidays are already excluded from
  capacity and are **not** subtracted twice.
- Casual capacity = hours engaged. Permanent capacity = contracted hours ÷ 5 per working
  day, limited to days they were employed.
- Overtime is outside both the numerator and the denominator, reported as
  `Overtime Hours`.
- `WIP Hours Added` = chargeable hours on `Project` jobs.
