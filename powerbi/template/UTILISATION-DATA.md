# Utilisation schedule — the numbers

Two CSVs, computed from `August 2026 Utilisation Report.xlsx` using the definitions we
settled: casual capacity as hours engaged, available hours as capacity less leave,
overtime outside the ratio, delivery scope only for the headline.

## Loading into Power BI

Get data > Text/CSV > pick the file > Load. That is the whole job — no model, no
parameters, no relationships. Use it to see numbers today while the full model's report
pages get sorted.

## `utilisation_schedule.csv` — 1,523 rows

One row per person per month, 123 people, Jul 2024 to Aug 2026.

| Column | Meaning |
|---|---|
| Month | YYYY-MM |
| Employee, Employee Id, Person Group, Employment Type | who |
| Utilisation Scope | Delivery or Overhead |
| Working Days | weekdays in the month they were employed for |
| Capacity Hours | permanents: working days x contracted/5 x FTE. Casuals: hours engaged |
| Leave Hours, Public Holiday Hours | |
| Available Hours | Capacity less leave. Public holidays are not subtracted twice |
| Chargeable Hours, Non-Chargeable Hours | ordinary time only |
| Overtime Hours | reported beside the ratio, never inside it |
| Utilisation % | Chargeable / Available |
| Target %, Variance pp | variance in percentage POINTS |
| Coverage % | (worked + leave) / capacity. Below 100% means missing timesheets |

## `utilisation_by_month.csv` — 26 rows

Delivery staff rolled up by month. Chart `Utilisation %` and `Target %` against Month
and you have the trend in one visual.

## What it says

```
Month     ppl     avail    charge    util  target  var pp
2025-12    51    3922.5    2777.5   70.8%   81.4%  -10.6
2026-01    41    3956.8    2582.8   65.3%   80.6%  -15.4
2026-02    52    4494.9    3632.0   80.8%   79.8%   +1.0
2026-03    51    4927.1    4056.2   82.3%   80.8%   +1.6
2026-04    49    4793.8    3478.8   72.6%   81.0%   -8.5
2026-05    51    4559.2    3717.2   81.5%   79.8%   +1.7
2026-06    50    4520.1    3479.0   77.0%   79.9%   -3.0
2026-07    51    4449.7    3682.6   82.8%   80.2%   +2.6
2026-08    50    3189.9    2764.5   86.7%   79.0%   +7.7
```

The business runs at or near target most months. December and January are genuinely
low (holiday period), April dips on Easter. Those are real seasonal effects, not errors.

**August 2026 reads 86.7%, not 50.2%.** August capacity is lower than other months
because the data stops on 21 August and capacity is only counted up to each person's
last timesheet — so this is a like-for-like partial-month figure, not a full month.

## Two things to know

- **Zero unmatched hours.** Every timesheet row resolved to a person, including the
  employee IDs shared by two people. The key resolution held.
- **Targets are group defaults** (Support 0.88, Production 0.65, Consulting 0.50) until
  you set real per-person targets in `map_employees.csv`. The Target % column moves when
  you do.
