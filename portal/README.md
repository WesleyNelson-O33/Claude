# CTS Business Intelligence Portal

One HTML file, one script and a folder of data. It runs entirely in the browser:
no backend, no build step, no network calls at runtime. Every figure is
re-aggregated from the ledger each time the page loads, the same way the
Controller Pack re-aggregates a GL paste.

## Running it

Double-click `CTS Business Intelligence Portal.html`. Nothing needs installing
and nothing needs a server. Keep the folder together: the HTML needs
`CTS_bi_bundle.js`, `CTS_build.js`, `CTS_email.js`, `vendor` and `data` beside it.

For the team, host the folder in a SharePoint library that everyone syncs and
opens from the synced copy. `docs/HOSTING.md` has the setup; `docs/MONTHLY.md`
the routine; `docs/AUTOMATION.md` the email flow.

## The monthly cycle in one line

Paste each export into its template in `templates`, open the portal, Admin,
Build, and click through Choose folder, Read templates, Write data files. The
portal rewrites its own data files from the templates, keeps last month's in
`data/_previous`, and OneDrive carries the result to everyone.

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

## The templates

Eight workbooks in `templates`, each shaped to the export it receives so the
paste is header for header. Built by `portal/build/build_templates.py`.

| Template | From | Status |
|---|---|---|
| 01 Xero P&L | Xero, Profit and Loss | Matched |
| 02 Xero GL Transactions | Xero, Account transactions for P&L analysis | Matched, the thirteen columns from the Controller Pack |
| 03 Employment Hero Earnings | Employment Hero, Earnings Details | Matched, the seventeen columns, plus Locations, Pay categories and Adjustments tabs |
| 04 Zoho Deals | Zoho CRM, Deals | Placeholder on Zoho's standard export until a real one is matched |
| 05 OnRent Orders | OnRent Events | Placeholder until a real export is matched |
| 06 Qwilr Quotes | Qwilr | Placeholder until a real export is matched |
| 07 Budget | Once a year | Same layout as the P&L |
| 08 Config | Monthly reporting month; otherwise rarely | Departments, split bases, holidays, who sees what, who gets the email |

The Build page (`CTS_build.js`, SheetJS vendored in `vendor`) reads them with
the File System Access API, which Edge and Chrome have, checks them, and writes
`data`. Data Loaders keeps the paste boxes as the ad hoc path.

## The monthly email

`CTS_email.js` renders three tiers, executive, department head and finance, and
the Distribution page writes them into `outbox/YYYY-MM` as JSON and HTML. The
portal never sends: a Power Automate flow watches that folder, or you copy the
HTML into Outlook until it does. `docs/AUTOMATION.md` has the flow.

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
