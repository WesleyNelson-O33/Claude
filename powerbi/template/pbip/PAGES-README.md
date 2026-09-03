# Dropping in the report pages

Your semantic model and parameters live in `Utilisation.SemanticModel` and are **not
touched** by this. Only the report pages change, so your paths and refresh stay as they are.

## Steps

1. **Close Power BI Desktop.**
2. **Back up first:** copy `Utilisation.Report\definition\pages` somewhere safe, or copy
   the whole `Utilisation.Report` folder. Takes five seconds and saves an argument later.
3. Delete the existing `Utilisation.Report\definition\pages` folder — the one holding
   `Page01` … `Page05`.
4. Copy the `pages` folder from this zip into `Utilisation.Report\definition\` in its place.
5. Reopen `Utilisation.pbip`.

## What you get

Five pages, 39 visuals. Every field reference has been checked against the model, so
nothing should come up blank because of a bad binding.

**Summary** — two slicers (Month Year, Person Group), five cards (Utilisation %,
variance in points, chargeable hours, available hours, coverage %), the rolling 13-week
trend against target, utilisation by team, hours mix, chargeable by job group, plus
overtime, casual share and the data quality flag.

**By person** — one table, sorted worst first. Coverage % sits beside utilisation on
purpose: answer "did they fill in a timesheet" before drawing any conclusion.

**Where the time went** — non-chargeable hours by job (your list of things to kill),
overtime by person, and a category × month matrix.

**WIP hours** — cards for hours added, hours to date and unbilled, plus a job-level
table that drops beside your dollar schedule on the same job key.

**Data quality** — the four zero-target cards, the working-days pair that catches a
partial month, coverage by person, and the pay type mapping check.

## The slicers are synced

Month Year is synced across all five pages, so picking August once sets it everywhere.

## Two things to expect

- **`Utilisation vs Target (pp)`** will read oddly until you fill in
  `Target Utilisation %` properly in `map_employees.csv`. Right now it is a group default.
- **`Unbilled WIP Hours`** reads the same as `WIP Hours to Date` until
  `Last Fully Billed Date` is populated. That is by design, not a bug.

## If a visual comes up blank

Valid JSON does not guarantee a rendered visual. If one is empty, click it and check the
Fields well — it will show which binding it did not like. Send me the visual name
(shown in the Selection pane) and I will fix that one file.
