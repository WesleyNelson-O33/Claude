# Report pages

Four pages. Resist adding a fifth until someone actually asks for it.

Shared slicers on every page (sync them: View > Sync slicers):
`FY Label`, `Month Year`, `Team`, `Employment Type`, `Location`.

---

## Page 1 - Summary

For the person who has 30 seconds.

**Top row of cards**
- `Utilisation %` (big number)
- `Utilisation vs Target (pp)` - conditional font colour from `Utilisation Status Colour`
- `Billable Hours`
- `Available Hours`
- `Active Headcount`
- `Data Quality Flag` - small, top right, red when not "OK"

**Main visual: line chart**
- X axis: `Dim_Date[Week Ending]`
- Y axis: `Utilisation % R13W` and `Target Utilisation %`
- Do **not** plot raw weekly `Utilisation %` here. A single week swings 15
  points on one person taking leave and it makes the chart useless.

**Right panel: bar chart, utilisation by team**
- Y axis: `Dim_Employee[Team]`
- X axis: `Utilisation %`
- Add `Target Utilisation %` as a constant-per-category line if you have the
  Analytics pane available, otherwise show `Utilisation vs Target (pp)` as a
  second bar.

**Bottom: donut or stacked bar, hours mix**
- Legend: `Dim_WorkType[Utilisation Category]`
- Values: `Total Hours`

---

## Page 2 - By person

The accountability page. Expect this one to get screenshotted into meetings,
so make it defensible.

**Table** (not a matrix - you want sortable columns):

| Column | Measure/field |
|---|---|
| Employee | `Dim_Employee[Employee Name]` |
| Team | `Dim_Employee[Team]` |
| Available hrs | `Available Hours` |
| Billable hrs | `Billable Hours` |
| Non-billable hrs | `Non-Billable Hours` |
| Leave hrs | `Leave Hours` |
| Utilisation % | `Utilisation %` |
| Target % | `Target Utilisation %` |
| Variance pp | `Utilisation vs Target (pp)` |
| Coverage % | `Timesheet Coverage %` |

Conditional formatting: data bars on `Utilisation %`, font colour on
`Variance pp` via `Utilisation Status Colour`.

**Put `Timesheet Coverage %` on this page and leave it there.** The first
question anyone asks about a low utilisation figure is "did they even fill
in their timesheet". Answer it in the same view.

---

## Page 3 - Where the time went

Diagnostic. Explains *why* the number is what it is.

- **Matrix**: rows `Dim_WorkType[Utilisation Category]` then
  `Dim_WorkType[Work Type]`; columns `Dim_Date[Month Year]`;
  values `Total Hours` and `% of column total`.
- **Bar chart**: non-billable hours by work type, descending. This is your
  list of things to kill.
- **Bar chart**: `Above-Capacity Hours` by employee. Overtime concentrated on
  two or three names is a resourcing problem wearing a utilisation costume.
- **Bar chart**: hours by `Location`.

---

## Page 4 - Data quality

Boring, and the reason anyone will believe pages 1 to 3.

**Cards**
- `Unmapped Work Type Hours` (target: 0)
- `Unmatched Employee Hours` (target: 0)
- `Employees With Capacity But No Timesheets` (target: 0)
- `Timesheet Coverage %` (target: ~100%)
- `Data Age (days)`
- `Last Timesheet Date`

**Table: unmapped work types**
- Filter `Dim_WorkType[Utilisation Category] = "UNMAPPED"`
- Columns: `Work Type`, `Total Hours`, distinct employees
- This is your worklist. Every row here is a line to add to
  `worktype-utilisation-map.csv`.

**Table: people with capacity but no timesheets**
- `Dim_Employee[Employee Name]`, `Capacity Hours`, `Total Hours`
- Filter to `Total Hours` is blank.

**Table: source files loaded**
- `Fact_Timesheet[Source File]`, min and max `Work Date`, row count.
- Catches the "we forgot to export the last fortnight" failure, which
  otherwise shows up as a mysterious utilisation cliff.
