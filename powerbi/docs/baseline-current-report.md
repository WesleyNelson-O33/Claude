# Baseline: the current utilisation report

Assessment of `August_2026_Utilisation_Report.xlsx` as supplied, 3 September 2026.
Full write-up with evidence: https://claude.ai/code/artifact/31999da1-cf04-4fe5-a579-b82e009b0356

## Verdict

The workbook cannot currently produce a correct utilisation number. It publishes
50.2% against an 81.2% target. On a like-for-like basis the same people and the
same hours give roughly 80-88%. The gap is four compounding denominator errors
plus a filter that drops 29% of chargeable work - not performance.

## The correction bridge (August 2026, same 23 people, same 1,980.25 chargeable hrs)

| Basis | Utilisation |
|---|---|
| As published (23 business days, leave not deducted) | 50.2% |
| Corrected to August 2026's actual 21 weekdays | 54.9% |
| Leave deducted from capacity | 60.2% |
| Like-for-like: only the 15 weekdays that have data | 87.7% |

Excluding the three people with broken capacity factors: 45.1% -> 54.6% -> 79.7%.

## Defects

**Critical**

- **D1** Report excludes 29% of chargeable work. 58 people logged chargeable hours
  in Aug-26; 23 appear on the report. 826.50 of 2,806.75 hrs excluded. Two of the
  excluded are full-time.
- **D2** Partial month against a full-month denominator. Data runs 3-21 Aug =
  15 working days. Nothing on the report says so.
- **D3** Business-day count wrong and in places impossible. August 2026 has 21
  weekdays; the report uses 23. The `Days` sheet claims more business days than
  the calendar has weekdays in 4 of 49 months (Aug-26, Nov-26, May-27, Jul-27),
  and carries three `#REF!` cells.
- **D4** One employee's hours split across two name spellings under one employee
  ID (`7118495F`). "John Rizvi" shows 40.08 hrs / 21.8% against an 88% target;
  "John Abbas Rizvi" adds 80.00 hrs the report never sees. Real total 120.08.
- **D5** Leave never deducted from capacity, so approved leave reads as poor
  performance. Sebastien Hathaway: 66.25 hrs leave, capacity still 184, published
  at 18.2%.

**High**

- **D6** Capacity hand-typed per person with reverse-engineered plugs
  (`=23*8+0.05`, `=23*8-0.05`, `=23*5.130434`, `=23*2.52173913`). Layla Phillips
  is published at 206.9% chargeable.
- **D7** Grand total is broken. Nine error cells on the August tab including
  I31/I32/N31/N32 `#REF!` and P31/P32 `#N/A`. Workbook-wide: 114,031 error cells.
- **D8** "Non-chargeable %" is `1 - chargeable %`, i.e. unaccounted time, not
  non-chargeable work. Contradicts column E on every Support row. Also returns
  the text `"0"` rather than zero.
- **D9** Three different public-holiday formulas down column S of `Raw data`
  (31,061 / 13,732 / 2,554 rows), each matching a different whitespace variant of
  the same pay type. Three leave types missed entirely.

**Medium**

- **D10** `GETPIVOTDATA` with employee names as string literals; spelling already
  inconsistent between adjacent columns ("Milo Rankin" vs "Milo RANKIN").
- **D11** `VLOOKUP` ranges slide one row per row (`I2:K410`, `I3:K411`, ...);
  P28 reads C29's name.
- **D12** Grand total target `=AVERAGE(L14:L30,L6:L10,L11:L12)` averages the
  Production subtotal in with the five individual Production rows.
- **D13** 48 sheets (27 hidden), 28.4 MB, ~470k live lookup formulas, duplicate
  sheet names, `Full-Time`/`Full-time`, `Non-chargeable`/`Non-Chargeable`, and a
  `PH` chargeable category (592 rows) that matches neither hour formula.

## What carries across unchanged

1. Chargeable is driven by job number via `Job List` - objective and already
   maintained. Much better than a work-type-based classification.
2. `Raw data` is at the right grain: employee x day x job x hours.
3. Targets are already differentiated per person (0.88 Support, 0.65-0.70
   Production, 0.50 Consulting).
4. The `Days` sheet is the right concept; it needs generating, not typing.

## Model changes from the generic kit in this folder

- `Dim_WorkType` is replaced by `Dim_Job` (from `Job List`), keyed on Job No.
  The chargeable flag lives there.
- Add `Dim_PayType` to classify leave, public holiday and overtime from the
  payroll `Pay Type` string, with trimming - see D9.
- Every join is on `Employee Id`, never on `Name` - see D4.
- Capacity is generated per employee per working day from contracted hours and
  FTE, never typed - see D3 and D6.

## Open decisions

1. Are casuals in or out of the company utilisation figure? They are 29% of
   chargeable hours.
2. Is overtime above capacity or inside it? 10+ overtime pay types at loadings
   1.25-2.0. Recommendation: hold it out of the ratio and report it beside.

---

# Decisions settled

## 1. Casuals are IN the company utilisation figure

Casual capacity = **hours engaged** (chargeable + non-chargeable), not 8h x days
attended. You carry no idle cost for a casual you did not call in.

August 2026 on 15 working days of data:

| Scope | People | Chargeable | Capacity | Utilisation |
|---|---|---|---|---|
| As published today | 23 | 1,980 | 3,946 | 50.2% |
| Delivery staff, permanents only | 25 | 2,104 | 2,858 | 73.6% |
| **Delivery staff, casuals in** | **50** | **2,796** | **3,598** | **77.7%** |
| Everyone incl Management/Finance/Admin | 58 | 2,807 | 4,398 | 63.8% |

**Correction to the earlier estimate.** The 80-88% quoted before this decision used
the workbook's own per-person capacity factors, three of which are reverse-engineered
plugs (D6). On a consistent 8h-per-working-day basis the figure is 73.6% permanent,
77.7% with casuals.

**Be clear-eyed about what this measures.** Casuals sit at 98.9% under this definition
(100.0% Support, 98.8% Production) because 26 casuals logged 692.25 chargeable and only
57.25 non-chargeable hours. Their capacity and their chargeable work are nearly the same
number by construction. Consequences:

- Including casuals lifts the company figure ~4 points and will keep doing so as the
  casual share grows, with no change in anyone's performance.
- A utilisation target for casuals is meaningless. Do not set one.
- Report two casual-specific measures beside the headline instead:
  `Casual Share of Delivered Hours` and `Casual Non-Chargeable %`.

## 2. Overtime sits OUTSIDE the ratio

Excluded from both numerator and denominator, reported as its own measure. Already
implemented in the kit as `Above-Capacity Hours` via
`Dim_WorkType[Counts Toward Capacity Flag] = FALSE()` - move this to `Dim_PayType`
for this business, since overtime is a pay type here, not a job attribute.

---

# 3. WIP extension: hours not charged to customers

Source reviewed: `202608_WIP.xlsx` - a deferred revenue schedule, GL 11300
"Prepaid Revenue", per job, month by month, in dollars. 452 job rows, 191 distinct
job numbers, 155 columns.

## The join works

132 job numbers appear in both the timesheet data and the WIP schedule. `Job No` is a
genuine shared key.

## "Hours not charged" is three different numbers

| | Treatment |
|---|---|
| **a.** Chargeable hours delivered, not yet invoiced | **The WIP number.** Contract asset / accrued revenue. |
| **b.** Chargeable hours beyond what the job will recover | Write-off. Not an asset. Margin story, belongs on the utilisation report. |
| **c.** Non-chargeable hours (9000-series codes) | Overhead, expensed. Never WIP. |

The current report only separates (c). Splitting (a) from (b) is the actual work.

## August 2026 join result

```
                  on jobs in WIP    on jobs NOT in WIP
chargeable              708.83              2,097.92
non-chargeable           38.75              1,164.33
public holiday            0.00                  8.00
```

Most of the 2,097.92 is resident onsite / managed-service work (DTTL Melb Tech,
DTTL SYD Technician, CBA Brisbane Onsite) billed monthly on contract, which correctly
generates no project WIP. Nothing in the data distinguishes those from a project job
that has been missed.

**188.25 hours in August carry no job number at all** - 4.7% of the month.

## Three blockers

1. **No charge-out rate exists in either workbook.** `Base Hourly` in the timesheet
   extract is a *pay type*, not a rate. Decide whether WIP is carried at cost or at
   charge-out rate - an accounting policy question that changes both the number and
   the disclosure.
2. **No billing basis per job.** Recurring / fixed price / time and materials. Only
   the latter two generate WIP. The `Job Type` column in `Job List` exists and is empty.
3. **Invoices are hand-maintained** across 155 columns. Should come from Xero on the
   same job key.

## Measures once those exist

```dax
Delivered Value =
    SUMX( Fact_Timesheet,
          Fact_Timesheet[Chargeable Hours] * RELATED( Dim_Rate[Charge Rate] ) )

Invoiced to Date =
    CALCULATE( SUM( Fact_Invoice[Amount] ), Dim_Date[Date] <= MAX( Dim_Date[Date] ) )

Unbilled WIP =                                  -- contract asset
    CALCULATE( MAX( [Delivered Value] - [Invoiced to Date], 0 ),
               Dim_Job[Billing Basis] IN { "Fixed Price", "Time and Materials" } )

Overbilled =                                    -- contract liability, = GL 11300
    CALCULATE( MAX( [Invoiced to Date] - [Delivered Value], 0 ),
               Dim_Job[Billing Basis] IN { "Fixed Price", "Time and Materials" } )
```

The schedule supplied is the **liability** side (most balances negative - invoiced
ahead of delivery). The hours question is the **asset** side. Report both together or
the movement will never reconcile to the GL.

## Sequencing

Build the utilisation model first. The WIP extension is a separate project: it needs a
rate card, a billing basis on 745 jobs, and an invoice feed - the first two are
decisions, not engineering. `Fact_Invoice` hangs off the same `Dim_Job` and `Dim_Date`,
so there is no rework in doing utilisation first.

## Still open

- Are Management, Finance and Admin inside the chargeable ratio? (77.7% -> 63.8%)
- WIP at cost or at charge-out rate?
- Who assigns billing basis to 745 jobs, and who maintains it for new jobs?
