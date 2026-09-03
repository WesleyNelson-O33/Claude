# Utilisation report - Employment Hero timesheets to Power BI

Build guide plus a working starter kit: Power Query scripts, a DAX measure
library, mapping templates and a page-by-page report spec.

---

## 1. The short version

Do this, in this order:

1. **Decide what "billable" means and write it down.** Employment Hero has no
   billable flag. You create one by classifying every Work Type in
   `mapping/worktype-utilisation-map.csv`. This is the entire report.
2. **Export timesheets to CSV** from Employment Hero into one folder. Do not
   start with the API.
3. **Build the capacity table yourself** from contracted hours, employment
   dates and public holidays. Employment Hero will not give you a denominator.
4. Load the two facts and four dimensions, wire up the star schema
   (`docs/data-model.md`), paste in the measures (`dax/measures.dax`).
5. Build the four pages in `docs/report-pages.md`.
6. Move to the API or a staged database **only** when the manual export
   genuinely becomes the bottleneck.

Realistic effort: a day to a first working version if your Work Types are
already clean, a week if they are not. The modelling is not the hard part.

---

## 2. The thing that will actually sink this project

**Employment Hero does not know which hours are billable.** A timesheet line
has an employee, a date, a duration, and optionally a Work Type, a Location
and a Pay Category. None of those mean "chargeable to a client" unless you
have deliberately set them up that way.

So before any Power BI work, answer these:

- **Is there a Work Type on every timesheet line?** If most lines are blank or
  everything is "Ordinary Hours", you cannot split billable from non-billable
  and no amount of DAX will fix it. Fix the Employment Hero configuration
  first: create the Work Types you need and make them mandatory.
- **Do you need client or project level detail?** Work Type is a flat list. If
  you want utilisation by client, you are pushing Employment Hero past what it
  is built for, and you should be asking whether your practice management or
  job costing system is the better source.
- **What is the denominator?** Pick one and stick to it:

  | Denominator | Definition | Use when |
  |---|---|---|
  | **Available hours** | Contracted hours less leave and public holidays | Professional services. This is the default in the kit. |
  | **Contracted hours** | Contracted hours, leave included | Capacity planning and forecasting |
  | **Rostered hours** | Hours the person was actually rostered on | Rostered/shift operations - see section 6 |

The kit ships all three as separate measures (`Utilisation %`,
`Utilisation % (Gross)`, and the roster variant in section 6) so you are not
locked in. But publish **one** as the headline. A report with three
utilisation numbers on it turns every meeting into an argument about
definitions.

---

## 3. Getting the data out - three options, ranked

### Option A: CSV export into a folder (start here)

Employment Hero exports timesheets from **Time & Attendance > Reports**
(and Payroll has its own timesheet reports). Drop each export into one folder
and point the `TimesheetFolderPath` parameter at it. `01-fact-timesheet-csv.m`
unions every CSV in the folder, normalises the headers, filters to approved
lines and de-duplicates re-exported periods.

- **Cost:** nothing.
- **Effort:** an hour to set up, then a few minutes each pay period.
- **Refresh:** put the folder on OneDrive or SharePoint and Power BI Service
  can refresh on a schedule against it. On a local drive you need a gateway.
- **Honest downside:** someone has to remember to do the export. The
  `Source File` table on the Data Quality page exists to catch it when they
  do not.

### Option B: Employment Hero Payroll API

Base URL is `https://api.yourpayroll.com.au/api/v2` and the endpoints are
documented at [api.keypay.com.au](https://api.keypay.com.au/australia/guides/Home)
(Employment Hero Payroll is the former KeyPay platform, acquired in 2022).
API access requires a **Platinum plan or above** - check your subscription
before you spend time here, per
[Employment Hero's own help article on API access](https://help.employmenthero.com/hc/en-au/articles/17426618690063-Access-the-Employment-Hero-API).

`02-fact-timesheet-api.m` gives you a paged, refresh-safe starting point, but
**verify the endpoint path and query parameters against the current docs
before you trust it** - I have written the shape the API uses, not a
guarantee. The two things that will bite you:

- **Dynamic data sources.** If you build the URL by string concatenation,
  the query works in Desktop and then fails on the Service. Always use
  `Web.Contents(base, [RelativePath=..., Query=...])`. The starter file does.
- **Cross-endpoint joins.** Timesheets, employees, work types and locations
  are separate endpoints. Joining them in Power Query means re-querying per
  row, which is slow enough to time out. This is the main reason people who
  try the API route end up back at option A or moving to option C.

### Option C: Stage into a database first

Pull the API into SQL (Azure SQL, Fabric, or a small managed Postgres) on a
schedule, then point Power BI at the database. This is the right answer at
scale and the wrong answer for a first build. Third-party middleware such as
SyncHub sells this as a product if you would rather not build it.

**Recommendation:** Option A now. It gets you a working, trusted report in
days. Revisit only when the manual export is the actual constraint - and it
usually is not, because the constraint is normally Work Type hygiene.

---

## 4. Setting it up

### Files in this kit

```
powerbi/
  queries/
    00-parameters.m              parameters to create (read this first)
    01-fact-timesheet-csv.m      Fact_Timesheet from a folder of CSVs
    02-fact-timesheet-api.m      Fact_Timesheet from the API (verify first)
    03-dim-date.m                Australian FY calendar
    04-dim-employee.m            from employee-targets.csv
    05-dim-worktype.m            the billable/non-billable classification
    06-dim-publicholiday.m       from public-holidays.csv
    07-fact-capacity.m           the denominator
  dax/
    measures.dax                 the full measure library
  mapping/
    worktype-utilisation-map.csv edit this - it defines the report
    employee-targets.csv         replace the example rows
    public-holidays.csv          replace with your state's dates
  docs/
    data-model.md                schema, relationships, gotchas
    report-pages.md              page-by-page build spec
```

### Steps

1. Copy the three files from `mapping/` to a folder on OneDrive or SharePoint.
   Fill them in. The examples in `employee-targets.csv` and
   `public-holidays.csv` are placeholders - delete them.
2. Create a folder for timesheet exports. Put at least one real export in it.
3. New Power BI file. Create the parameters listed in `00-parameters.m`.
4. For each query file: Home > Transform data > New Source > Blank Query >
   Advanced Editor > paste the file contents > name the query exactly as the
   comment header says (`Fact_Timesheet`, `Dim_Date`, and so on). Names
   matter - the queries reference each other.
   Load them in this order, because of the dependencies:
   `Dim_Date` > `Dim_Employee` > `Dim_PublicHoliday` > `Fact_Timesheet` >
   `Dim_WorkType` > `Fact_Capacity`.
5. **Open one of your real CSV exports and check the header row against the
   `RenameMap` in `01-fact-timesheet-csv.m`.** Employment Hero's column names
   differ between Payroll classic, current Payroll, and Time & Attendance, so
   there is no single correct list and I have not pretended otherwise. Add
   your actual headers to the map.
6. Build the relationships in `docs/data-model.md`. Mark `Dim_Date` as the
   date table.
7. Paste the measures from `dax/measures.dax`.
8. Build the Data Quality page **first**. Get every number on it to zero
   before you show anyone page 1.
9. Build pages 1 to 3.

---

## 5. How the maths works

Worked example. One consultant, one fortnight (10 working days), 38-hour week,
one public holiday, one day of annual leave, 75% target.

**Capacity**

```
Daily capacity        = 38 / 5                    = 7.6 hrs
Gross capacity        = 10 days x 7.6             = 76.0 hrs
Less public holiday   = 76.0 - 7.6                = 68.4 hrs   <- [Capacity Hours]
```

The public holiday is removed inside `07-fact-capacity.m`, so `Capacity Hours`
is already net of it.

**Available hours**

```
Timesheet shows 7.6 hrs of annual leave.
Available Hours = Capacity Hours - Leave Hours
                = 68.4 - 7.6                      = 60.8 hrs
```

Public holidays are **not** subtracted again here. They are already zero in
capacity, and subtracting twice is the single most common error in a
hand-built utilisation model - it quietly inflates utilisation by 10%.

**Utilisation**

```
Timesheet: 44.0 billable, 12.0 non-billable, 7.6 leave, 4.0 overtime.

Utilisation %        = 44.0 / 60.8                = 72.4%
Utilisation % (Gross)= 44.0 / 68.4                = 64.3%
Productive %         = (44.0 + 12.0) / 60.8       = 92.1%
Billable Mix %       = 44.0 / (44.0 + 12.0)       = 78.6%
Coverage %           = (44.0 + 12.0 + 7.6) / 68.4 = 93.0%
```

Note the 4.0 hours of overtime does not appear in any of those. It is mapped
with `Counts Toward Capacity = No`, so it sits in `Above-Capacity Hours` and
is reported separately. If you fold overtime into the numerator without
adding it to the denominator, a person working unpaid overtime looks like a
star performer, and you will make bad decisions off that.

**Against target**

```
Target billable hours = 60.8 x 75%                = 45.6 hrs
Variance              = 44.0 - 45.6               = -1.6 hrs
Variance in points    = 72.4% - 75.0%             = -2.6 pp
```

Report the variance in percentage **points**, never as a percentage change.
"Utilisation is down 3%" is ambiguous; "down 2.6 points" is not.

---

## 6. If you run a rostered operation, not a services firm

The repo this kit sits in is payroll operations - awards, casuals, rosters,
overtime matrices. If your question is "are we using our rostered labour
efficiently" rather than "are our consultants billing enough", the numerator
and denominator both change:

- **Denominator** becomes rostered hours, not contracted hours. Replace
  `07-fact-capacity.m` with a load of the Employment Hero roster export
  (Rostering > Reports), at grain employee x date x rostered hours.
- **Numerator** becomes worked hours, not billable hours.
- The measure becomes:

```dax
Roster Adherence % =
DIVIDE ( [Worked Hours], SUM ( Fact_Roster[Rostered Hours] ) )
```

and the interesting numbers are the two failure directions, not the ratio:

```dax
Under-Worked Hours = MAX ( SUM ( Fact_Roster[Rostered Hours] ) - [Worked Hours], 0 )
Over-Worked Hours  = MAX ( [Worked Hours] - SUM ( Fact_Roster[Rostered Hours] ), 0 )
```

Everything else in the kit - the date table, employee dimension, work type
classification, data quality measures, star schema - carries across unchanged.

---

## 7. Things that will go wrong, and what they look like

| Symptom | Cause | Fix |
|---|---|---|
| Utilisation over 100% | Overtime mapped as `Counts Toward Capacity = Yes` | Set it to No in the work type map |
| Utilisation ~10% too high | Public holidays subtracted twice | Do not subtract PH in `Available Hours`; it is already zero in capacity |
| Utilisation drifts down over months | New Work Types added in Employment Hero, unmapped | Watch `Unmapped Work Type Hours` on the Data Quality page |
| One team looks terrible | They are not filling in timesheets | Check `Timesheet Coverage %` before drawing any conclusion |
| Capacity is a suspiciously round multiple | Duplicate `Employee External ID` in `employee-targets.csv` | De-duplicate the file |
| Numbers change when you add a slicer | Bi-directional relationship somewhere | Set every cross-filter direction to Single |
| Works in Desktop, fails in the Service | Dynamic data source in a `Web.Contents` call | Use `RelativePath` and `Query`, static base URL |
| A fortnight vanished | Nobody ran the export | The Source File table on the Data Quality page |

## 8. Refresh limits

Power BI Pro allows 8 scheduled refreshes a day; Premium, PPU and Fabric
allow 48. Timesheet data updates once a pay period, so one refresh a day is
already more than the data warrants. Do not build a real-time pipeline for a
fortnightly number.

---

## Sources

- [Employment Hero Payroll API reference (api.keypay.com.au)](https://api.keypay.com.au/australia/guides/Home)
- [Access the Employment Hero API - Employment Hero Help Centre (AU)](https://help.employmenthero.com/hc/en-au/articles/17426618690063-Access-the-Employment-Hero-API)
- [Building your integration with Employment Hero Payroll - Employment Hero Partners](https://partners.employmenthero.com/resources/building-your-integration-with-employment-hero-payroll)
- [Employment Hero API references (developer.employmenthero.com)](https://developer.employmenthero.com/api-references)
- [Set up your Timesheet Import (CSV) - Employment Hero Help Centre (AU)](https://help.employmenthero.com/hc/en-au/articles/11373882391183-Set-up-your-Timesheet-Import-CSV)
- [How to connect Employment Hero Payroll to Power BI (DataSights)](https://datasights.co/how-to-connect-employment-hero-payroll-to-power-bi/)
- [Connect Employment Hero HR to Power BI (SyncHub)](https://www.synchub.io/employmenthero-to-powerbi-connector)
