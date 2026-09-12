# Profitability per person — what exists and what does not

## Short version

- **Job numbers: yes.** `utilisation_by_job.csv` has every timesheet hour allocated to a
  job, per person per month.
- **Clients: partly.** 73% of hours carry a client. Not 100%, and the gap is fixable.
- **Cost per person: does not exist anywhere.** Not in the utilisation workbook, not in
  the WIP schedule. This is the blocker.
- **Charge rate: also does not exist.** Same problem, other side of the equation.

You cannot calculate profitability per person from the files supplied. Both the revenue
rate and the cost rate have to come from somewhere else. Everything up to that point is
built and waiting.

## What you have now

### `utilisation_by_job.csv` — 5,544 rows

Employee x month x job. Every chargeable, non-chargeable and overtime hour, allocated.

| Column | Notes |
|---|---|
| Month, Employee, Employee Id, Person Group, Employment Type, Utilisation Scope | who |
| Job No, Job Name, Job Group, Billing Basis | what |
| Client, Client Group, Sector | who for |
| Chargeable / Non-Chargeable / Overtime / Total Hours | how much |

### `cost_and_rate_template.csv` — 94 people

One row per person over the last 12 months (Sep 2025 - Aug 2026), with their chargeable,
non-chargeable and overtime hours already totalled. Two blank columns:

- **`Cost Rate $/hr`** — fully loaded cost: base pay + super + leave loading + payroll
  tax + workers comp. Not just the hourly rate on their contract.
- **`Charge Rate $/hr`** — what a client is billed for an hour of their time.

Fill those two columns and profitability calculates. Nothing else is needed.

## The maths, once the rates are in

```
Revenue        = Chargeable Hours x Charge Rate
Direct cost    = (Chargeable + Non-Chargeable + Overtime) Hours x Cost Rate
Gross margin $ = Revenue - Direct cost
Gross margin % = Gross margin $ / Revenue
```

Note the asymmetry, because it is where this usually goes wrong: **you earn on chargeable
hours only, but you pay for every hour.** Non-chargeable time is pure cost. That is what
makes the utilisation number worth measuring in the first place.

Overtime needs a decision: if you pay a 1.5x or 2x loading but bill at the standard rate,
overtime destroys margin and the cost side must carry the loading. There are 20 overtime
pay types in your data at loadings from 1.25 to 2.0.

## Where the rates come from

**Cost rate — payroll.** Employment Hero holds pay rates. Fully loaded means adding
roughly 25-35% on top of base for super, leave, payroll tax and workers comp, but use
your own figures rather than a rule of thumb.

**Charge rate — your rate card or quotes.** For the resident onsite contracts (DTTL, CBA,
Bankwest) it is whatever the contract says per head. For project work it is the quoted
rate or derived from the job value.

**Or from Xero.** You have Xero connected to this session. Invoiced revenue by job or
contact would give actual revenue rather than a rate card assumption, which is better.
Say the word and I will check what is actually there before you spend time on rate cards.

## The client gap

73% of hours carry a client. The missing 27% is because client is maintained in three
different places and none of them is complete:

| Source | Jobs it filled in |
|---|---|
| `Job List` (Client column) | 700 |
| WIP schedule (Client column) | 180 |
| `Lookup` (Client column) | 53 |

I merged all three, which is why coverage went from 6% to 73%. To close the rest, client
needs to live in one place - realistically `Lookup`, since that is the table your raw data
already points at and the only one containing your current date-coded jobs.

Profitability by client will be understated until that is fixed, because a quarter of the
hours will not attribute.
