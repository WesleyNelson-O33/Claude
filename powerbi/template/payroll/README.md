# Cost per employee from Employment Hero

## What this does

`extract_cost.py` reads Employment Hero **Pay Run Audit Report** exports and produces
actual employment cost per employee — and, because the audit report carries a job number
on every earnings line, **cost per employee per job**.

You already download these every fortnight. Your own manual, Part 6 Step 6:

> `Pay run › Reports › Audit report; tick everything on the left, Run report, download as Excel`

There are roughly 26 of them sitting in your fortnight folders. No new process needed.

## Running it

```
python3 extract_cost.py <folder of PayRunAudit*.xlsx> <output folder>
```

Point it at a folder holding every fortnight's audit file and it processes the lot.

**Set `ON_COST_UPLIFT` at the top of the script first.** It defaults to 0. Payroll tax
and workers comp are not in the audit report because they are levied on your total wage
bill rather than per person, so they have to be added as an uplift. Your accountant will
have the figure — it is typically 5-8% combined, but use the real one. Payroll tax only
applies above the state threshold, so if you sit near it the uplift is not linear.

## Output

**`cost_by_person_job.csv`** — one row per earnings line: pay period, employee, **job
number**, job name, pay category, hours, rate, gross, super. This is the file that makes
job-level profitability possible.

**`cost_rate_by_employee.csv`** — one row per employee per pay period: hours paid, gross,
super, employer liabilities, total cost, and a derived `Cost Rate $/hr`.

## Cost basis

```
Total cost = Gross Earnings + SG Super + Employer Liabilities, x (1 + on-cost uplift)
```

Actual cost, not a rate. It picks up overtime loadings, casual loading, leave loading and
allowances automatically — which matters here, because you run 20 overtime pay types at
loadings up to 2.0. A contract hourly rate would miss all of it.

## The reconciliation check

Every audit sheet ends with a **Total** row. That row carries the period total in the
dollar columns but no Units — so summing it doubles every dollar figure while hours still
look correct. It is a silent, plausible-looking doubling.

The script skips it and prints a reconciliation of Earnings Details against Pay Run
Totals on every run. **If that line does not say OK, do not use the output.**

## What two fortnights say

Period 25 Jul - 21 Aug 2026, matched to timesheet hours over the same dates:

| Scope | Group | Paid hrs | Cost | Chargeable hrs | Cost per chargeable hr |
|---|---|---|---|---|---|
| Delivery | Support | 3,120.7 | $123,777 | 2,711.9 | **$45.64** |
| Delivery | Production | 1,376.2 | $71,828 | 882.8 | **$81.37** |
| Overhead | Management | 640.0 | $44,272 | 16.0 | n/a |
| Overhead | Finance | 134.2 | $5,735 | 0.0 | n/a |
| Overhead | Admin | 113.6 | $4,514 | 0.0 | n/a |

**Blended delivery cost per chargeable hour: $53.02.**
Total employment cost in the period $250,125 — delivery $195,604, overhead $54,520.
Overhead is 21.8% of the wage bill.

Those figures carry no on-cost uplift yet. At 6.5% they become roughly $48.60 (Support),
$86.66 (Production), $56.47 blended.

## Two things to check

**Consulting shows 94.4 chargeable hours but zero cost.** None of the three consultants
appear in these two fortnightly audit files. Most likely they sit on a different pay
schedule — monthly, or a separate pay run. Find that pay run and include its audit
exports, or consulting margin will read as pure profit.

**Employer Liabilities was zero in both files.** If you carry anything there in other
periods the script picks it up automatically, but worth knowing it is currently empty.

## What is still missing

Cost is now solved. Revenue is not, for 97% of hours:

- **Consulting** — rate card exists ($165 / $150 / $145 by grade). Needs each consultant
  mapped to a grade.
- **Support (81% of hours)** — resident onsite contracts. Revenue is the monthly contract
  value per head, not hours x rate.
- **Production (17%)** — quoted per job. Revenue is the invoiced job value.

For the last two, actual invoiced revenue from Xero, joined on job number, is the right
source.
