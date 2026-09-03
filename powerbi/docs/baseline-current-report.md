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
