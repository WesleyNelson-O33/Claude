# CTS Business Intelligence Portal

One HTML file, one script and a folder of data. It runs entirely in the browser:
no backend, no build step, no network calls at runtime. Every figure is
re-aggregated from the ledger each time the page loads, the same way the
Controller Pack re-aggregates a GL paste.

## Running it

```
cd portal
python3 -m http.server 3000
```

then open <http://localhost:3000/CTS%20Business%20Intelligence%20Portal.html>.

Opening the HTML file directly by double-clicking mostly works, but some
browsers refuse to load sibling files from a `file://` path, which stops the
data files loading. If the portal shows a data load error, serve the folder as
above. Nothing here needs installing.

## What is in it

21 pages across seven sections.

| Section | Pages |
|---|---|
| Dashboard | Dashboard, Business Context |
| Finance | P&L, P&L by Department, P&L Spread, Overhead Allocation, P&L Control |
| Budget | Budget vs Actual, Actions |
| Revenue | Revenue Summary, Revenue Schedule, Top Clients, Clients by Department |
| Utilisation | Utilisation, Profitability per FTE |
| Ledger | Detail Records, Charts |
| Admin | Setup, Data Loaders, Config & Variables, About |

## The conventions it is built on

These are the CTS rules, taken from the Controller Pack and the Monthly
Reporting manual rather than invented:

- **Amount is credit less debit.** Income is positive, costs are negative, so
  gross profit is income *plus* cost of sales.
- **Department comes from the transaction, not the account.** The Cost Centres
  column first, then the bracket tag on the job number. Three quarters of the
  P&L sits in accounts with no department in the name.
- **The P&L is actual to the reporting month and budget after it**, so the year
  always reads as twelve months with the shortfall ahead visible.
- **The risk flag tests the materiality floor before the percentage**, so a
  large percentage on a small dollar variance is still Low.
- **Admin overheads are pushed out on three bases** (3 Way, Staff, Office
  Dept), each line labelled with the one that applies. A department's own
  overhead and its share of the pool are kept as two separate columns, because
  the documented way this goes wrong is one being counted twice.
- **Two utilisation figures exist and are not meant to agree.** The pack nets
  public holidays off the month and counts worked hours; the graphs file counts
  every logged hour against a plain weekday year, which is why that denominator
  lands on 2,080 or 2,088.

## Loading real data

Data Loaders takes a paste, not a file. Copy a range in Excel and paste it in:
the clipboard carries it tab separated, which is what the boxes read. Three
loaders: the Xero account transactions report, the Xero P&L, and utilisation
hours. Everything is parsed in the browser and nothing is uploaded. A load lives
in that browser tab only; the save button writes a replacement data file you can
drop into `portal/data`.

## Rebuilding the data files

```
python3 portal/build/build_data.py
```

Deterministic: the same inputs give the same output every time. It reads
`reporting/data/accounts.json` and `reporting/data/validation.json` and writes
every file in `portal/data` except `CTS_config_data.js`, which is hand
maintained and is the one file meant to be edited.

## What is real and what is not

The shipped data is a **seed**. These figures are the real Xero ones and the
portal ties to them to the cent:

- FY26 control totals (income 6,113,186.03, cost of sales 3,497,093.83, gross
  profit 2,616,092.20, other income 13,413.68, expenses 2,380,879.16, net profit
  248,626.72), pulled from Xero on 3 September 2026
- the July 2025 month
- the August 2026 category totals, from the Controller Pack's PL_Check worked
  example
- the 190 account chart, the department tags, the risk thresholds, the 3 Way
  split at 34/33/33, the production and integration sub splits, and the margin
  and seasonality norms

Everything between those anchors is modelled from the documented norms. Client
names in the seed are fictional. Three things are stand-ins and the portal
flags them on screen wherever they are used:

- the **Staff** and **Office Dept** split percentages, which live in the budget
  workbook and are not in this repository
- **utilisation targets** other than consulting, which are not written down
- **which basis each overhead line uses**, which is labelled per row on the live
  allocation tab and is derived from the account here

## Colour

Six departments, six categorical slots in fixed order, never cycled. The palette
is validated in both light and dark mode: all checks pass, with three light-mode
slots below 3:1 contrast, which is why every chart ships with its data table
beside it rather than on its own.
