# CTS Monthly Business Review deck

`build_mbr.js` builds `CTS Monthly Business Review - <month>.pptx` on the CTS
slide masters (`../cts_masters.js`). Every figure comes from `mbr_data.json`.

    node presentation/mbr/build_mbr.js

## Monthly routine

1. Open `mbr_data.json` and replace the numbers with the month's actuals from
   the Financial Controller Pack (`pack/CTS Financial Controller Pack.xlsx`):
   Summary for the P&L, P&L FY27 for the departments, Clients for the top 10,
   Utilisation for hours and FTE. Set `month`, `monthShort` and `ytdLabel`.
2. Rebuild. The file name follows `month`.
3. Present, typing the commentary into the panel on each slide, then carry the
   agreed actions into the final table.

The data file ships with ILLUSTRATIVE figures. Nothing in it is a real CTS
result. Client names are placeholders (`Client A` to `Client J`).

## Slides

| # | Slide | Data used |
|---|---|---|
| 1 | Cover | month, presenter |
| 2 | Agenda | fixed |
| 3 | Section: Financial performance | |
| 4 | P&L summary | `plSummary` rows and six-month trend |
| 5 | P&L per department (table) | `departments` |
| 6 | Revenue and contribution by department (chart) | `departments` |
| 7 | Risk revenue profile | `departments` contract / non-contract split, `riskFlags` |
| 8 | Section: Customers and cash | |
| 9 | Top 10 customers | `topCustomers` |
| 10 | Receivables and DSO | `dso` |
| 11 | Top customers year on year | `topCustomers` ytd vs lastYtd |
| 12 | Section: People and utilisation | |
| 13 | Utilisation by department | `utilisation`, `workingDays`, `hoursPerDay` |
| 14 | Profit based on utilisation | `labourProfit` rates with the utilisation hours |
| 15 | Section: Summary and actions | |
| 16 | Commentary and actions | `actions` |

## Definitions baked into the deck

- Variances are actual less budget. Costs are stored as negatives, so a
  positive variance is always favourable, shown in teal; unfavourable in
  burgundy.
- Contribution = gross profit less the department's direct overheads. Admin
  carries shared overheads, so the CTS total is EBITDA.
- Risk flags follow the pack's Setup sheet: under $5k = Low, 10% or more =
  Medium, 25% or more or $50k or more = High.
- DSO = receivables / (last three months' revenue / 92 days).
- Utilisation % = chargeable hours / (FTE x working days x hours per day).
- Labour revenue = chargeable hours x charge rate. Labour cost = FTE x
  working hours x loaded cost per hour. "+5 pts" = what five more points of
  utilisation would add at the same charge rate, before any extra cost.

## Commentary

Each content slide has a mint panel on the right with prompts; type over the
prompts. The same prompts are in the speaker notes. The last slide collects
the commentary and actions with an owner and a date.
