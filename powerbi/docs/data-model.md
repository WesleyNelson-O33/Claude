# Data model

A plain star schema. Two fact tables at different grains, joined only
through shared dimensions - never to each other.

```
                   +----------------+
                   |    Dim_Date    |
                   | (Mark as Date) |
                   +--------+-------+
                            | 1
              +-------------+-------------+
              | *                         | *
    +---------+--------+        +---------+--------+
    | Fact_Timesheet   |        |  Fact_Capacity   |
    | grain: employee  |        | grain: employee  |
    |   x day x work   |        |   x working day  |
    |   type x location|        |                  |
    +---------+--------+        +---------+--------+
              | *                         | *
              +-------------+-------------+
                            | 1
                   +--------+-------+
                   |  Dim_Employee  |
                   +----------------+

    Fact_Timesheet * ----- 1 Dim_WorkType
    Fact_Capacity  * ----- 1 Dim_Date (via Date Key)
    Dim_PublicHoliday is consumed inside the Fact_Capacity query only;
    it does not need a relationship in the model.
```

## Relationships to create

| From (many)                    | To (one)                  | Cardinality | Cross-filter | Active |
|--------------------------------|---------------------------|-------------|--------------|--------|
| Fact_Timesheet[Date Key]       | Dim_Date[Date Key]        | Many-to-one | Single       | Yes    |
| Fact_Timesheet[Employee Key]   | Dim_Employee[Employee Key]| Many-to-one | Single       | Yes    |
| Fact_Timesheet[Work Type Key]  | Dim_WorkType[Work Type Key]| Many-to-one| Single       | Yes    |
| Fact_Capacity[Date Key]        | Dim_Date[Date Key]        | Many-to-one | Single       | Yes    |
| Fact_Capacity[Employee Key]    | Dim_Employee[Employee Key]| Many-to-one | Single       | Yes    |

Leave every cross-filter direction on **Single**. Bi-directional filtering
between two fact tables and a shared dimension produces ambiguous paths and
wrong totals, and it is very hard to spot once the report looks plausible.

## Model housekeeping that actually matters

1. **Mark Dim_Date as the date table.** Table tools > Mark as date table >
   Date column = `[Date]`. Without this, `DATESINPERIOD` and
   `SAMEPERIODLASTYEAR` in the measure library return wrong results rather
   than errors.
2. **Sort columns.** Set `Month` sorted by `Month Sort`, `Day Name` sorted by
   `Day of Week No`. Otherwise your axis reads Apr, Aug, Dec.
3. **Hide the keys.** Hide `Date Key`, `Employee Key`, `Work Type Key` on
   every table. Nobody should ever drag them onto a visual.
4. **Hide the raw hour columns** on `Fact_Timesheet` and `Fact_Capacity` so
   people use the measures instead of implicit sums.
5. **Put every measure in a `_Measures` table.** Create it with
   Home > Enter data, one dummy column, load, then hide the column.

## Grain warnings

- `Fact_Timesheet` is **not** unique per employee per day. Someone can log
  three work types on one day. Never assume one row equals one day.
- `Fact_Capacity` **is** unique per employee per working day. If it is not,
  your `employee-targets.csv` has duplicate `Employee External ID` values and
  the capacity denominator is multiplied. The `Table.Distinct` step in
  `04-dim-employee.m` guards against this, but it silently keeps the first
  row - check the file rather than relying on the guard.
