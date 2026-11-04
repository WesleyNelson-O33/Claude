# CTS Business Intelligence Portal: handover

Written 4 November 2026 for whoever picks this up next, including a Claude
session in another account. Everything here comes from the repository and
from the work done with Danica Nelson between September and November 2026.
Where something is a placeholder or unverified it says so.

Read this first, then `README.md`, then `docs/MONTHLY.md`. The older
`CTS Technical Handover.md` at the repository root covers the manuals and
workbooks that came before the portal; it is background, not the portal.

---

## 1. What this is, in one page

A static business intelligence portal for Corporate Technology Services Pty
Ltd (CTS), an Australian audio-visual company. One HTML file, four scripts and
a folder of data files. It runs entirely in the browser from a synced folder:
no server, no backend, no network calls at runtime, nothing to install. It
replaces the manual Monthly Reporting process (the P&L analysis file, the
graphs file and the management report) that the Monthly Reporting manuals
describe.

Finance pastes Xero, Employment Hero and Zoho exports into fourteen Excel
templates, opens the portal, clicks Build, and the portal rewrites its own
data files. Everyone else opens the portal and reads. It also writes the
monthly emails, a fortnightly update and the Monthly Business Review deck as
PowerPoint, all from the same data.

**Where it is up to**

| Item | State |
|---|---|
| Code | Complete and working. 33 pages, Build for all 14 templates, emails, fortnightly update, deck, controller checks, revenue risk, project P&L |
| Repository | `WesleyNelson-O33/Claude`, branch `claude/cts-replica-n0d209`, folder `portal/`. Latest commit `0ef956f` (Project P&L page). No pull request has been opened |
| Data in the repository | Seed only. Fictional client names, modelled months between real anchors. Reporting month in the seed config is August 2026 |
| Live data | September 2026 is closed and built. It lives only in two places: the private Claude artifact at https://claude.ai/artifact/VonqRbokE2PDjfaJ51LZ4V (version 25) and the zip `CTS Business Portal (September live).zip` sent to Danica on 29 October 2026. It must not be committed to the repository |
| September numbers | Revenue $561,945, gross profit $187,814, net loss $41,464. Ledger ties to the closed Xero P&L in every month. Controller checks 0 failed, 12 look |

---

## 2. Where everything lives

```
portal/
  CTS Business Intelligence Portal.html   the page: loader, splash, CSS
  CTS_bi_bundle.js     engine (E), UI (U), pages (P), commentary generators (GEN), router, store
  CTS_build.js         the Build page: template readers, converters (C.*), checks, data writers
  CTS_email.js         the monthly and fortnightly emails, three tiers
  CTS_deck.js          the Monthly Business Review deck (PptxGenJS, 16:9)
  vendor/              xlsx.full.min.js (SheetJS), pptxgen.bundle.js, licences
  data/                the data files the portal loads (seed in the repo)
  templates/           01 to 14, the Excel templates finance pastes into
  build/
    build_templates.py   generates every template from scratch (openpyxl)
    build_data.py        generates the seed data files
    fill_test_templates.py   fills templates with seed data for testing
    schedule_lines.json  the 147 revenue schedule lines from the FY27 budget
    budget_groups.json   FY27 budget by department group
  docs/
    MONTHLY.md     the monthly routine, template by template
    HOSTING.md     SharePoint and OneDrive hosting
    AUTOMATION.md  Power Automate flow for the outbox
    FORECASTING.md the forecast design and what is built
    HANDOVER.md    this file
```

Nothing outside `portal/` is needed to run or build the portal.

---

## 3. How it is put together

### Loading

The HTML loads data files one by one with their own error handlers.
Mandatory: config, cal, accounts, gl, fin. A missing one stops the portal with
a message naming the file. Optional: util, clients, users, staff, brand,
pipeline, forecast, cash, commentary, distribution, aged, checks. A missing
optional file only disables what depends on it. Then the vendor libraries,
then `CTS_build.js`, `CTS_email.js`, `CTS_deck.js`, then the bundle, which
calls `CTS.boot()`.

### The data files

Every file is `window.CTS_<NAME> = {...}` with a `meta` block saying when it
was built and from what, and `seed: true` on seed data. The portal shows a
seed banner on every page while any loaded file is seed.

| File | Holds | Built from |
|---|---|---|
| CTS_config_data | Departments, split bases, risk thresholds, utilisation settings, seasons, distribution, reporting month | 08 Config |
| CTS_cal_data | Months, financial years, working days per state, public holidays | 08 Config Holidays tab |
| CTS_accounts_data | Chart of accounts with category, sub group, department | The ledger and the P&L |
| CTS_gl_data | Every ledger line: date, account, contact, debit, credit, source, job number, invoice, cost centre, description, GST, tax name | 02 Xero GL Transactions |
| CTS_fin_data | P&L by account by month, actual and budget | 01 Xero P&L and 07 Budget |
| CTS_util_data | Hours by department by month (chargeable, non chargeable, leave, public holiday) and hours by job | 03 Employment Hero Earnings |
| CTS_staff_data | Hours and gross pay by employee by month | 03 Employment Hero Earnings |
| CTS_clients_data | Client display names, aliases, department | 13 Revenue Schedule Clients tab and the ledger |
| CTS_pipeline_data | Zoho deals, OnRent orders, Qwilr quotes, revenue schedule lines | 04, 05, 06, 13 |
| CTS_forecast_data | Forecast methods and overrides | 09 Forecast |
| CTS_cash_data | Bank and card balances, commitments, timing settings | 10 Cash and Commitments |
| CTS_aged_data | Aged receivables and payables, one row per open invoice or bill | 11 and 12 |
| CTS_checks_data | The controller check results for the month | Computed at Build from everything above plus 14 |
| CTS_users_data | Pages, roles, who sees what | 08 Config |
| CTS_brand_data | Logos and marks as data URIs | Fixed |
| CTS_commentary_data | Published commentary | Commentary page, Publish |
| CTS_distribution_data | Published recipient list | Distribution page, Publish |

### Conventions the engine is built on

- **Money is held in cents** internally (integers). Formatters divide by 100.
- **Amount is credit less debit.** Income positive, costs negative, so gross
  profit is income plus cost of sales.
- **Department comes from the transaction**: the Cost Centres column first,
  then the bracket tag on the Job Numbers column (`[ONS]`, `[PRD]`, `[INT]`,
  `[CONS]`, `[VID]`, `[CTS]`). Unmatched lines are UNALLOCATED.
- **Actual to the reporting month, budget (or forecast) after it**, so the
  year always reads as twelve months.
- **Admin overheads are pushed out** on three bases, 3 Way, Staff and Office
  Dept. The 3 Way split is confirmed at 34/33/33. Staff and Office Dept are
  placeholders until Danica types the live percentages on the Config Split
  bases tab.
- **Two utilisation figures exist** and are not meant to agree: the portal's
  (chargeable over chargeable plus non chargeable, leave excluded) and the
  graphs file's rolling twelve months against 2,080 or 2,088 hours.
- **Every page follows the department picker** in the period bar, except Cash.
  The emails and the deck ignore the picker on purpose; a department head's
  role fixes what they get.
- **Visibility is not security.** Roles decide which tabs are shown. Anyone
  with the folder can read the data files. A department head who must not
  see the ledger needs a separate restricted build (see HOSTING.md).

### Pages

33 pages. Pages marked sensitive are hidden from the Viewer and Department
head roles.

| Section | Pages |
|---|---|
| Dashboard | Dashboard, Business Context, Commentary, The Month |
| Finance | P&L, P&L by Department, P&L Spread, Overhead Allocation, P&L Control, Controller Checks, Trends |
| Budget | Budget vs Actual, Actions, Forecast |
| Cash | Cash Flow |
| Revenue | Revenue Summary, Revenue Schedule, Revenue Forecast, Revenue Risk, Project P&L, Pipeline, Top Clients, Clients by Department |
| Utilisation | Utilisation, Profitability per FTE, Profitability by Employee |
| Ledger | Detail Records |
| Admin | Setup, Build, Distribution, Config & Variables, Access Control, About |

Roles: admin (Finance head, 33 pages, can change access), finance (32),
exec (24), deptHead (15, own department only), viewer (4). Users and roles
are on the Users tab of 08 Config. The rows there are still template
examples (Danica, Duncan, Graham, Jordan, Production manager); real names
and emails are awaited.

### Adding a page

Three places: `P.<id>` in `CTS_bi_bundle.js` (section, title, sub, render
returning an array of nodes), the PAGES and ROLES lists in
`build/build_templates.py` so 08 Config carries it, and `data/CTS_users_data.js`
for the seed. Optionally `GEN.<id>` for proposed commentary. The Project P&L
commit `0ef956f` is a complete example.

---

## 4. The monthly build

`docs/MONTHLY.md` is the routine. The short version:

1. Paste each export into its template. Change the reporting month on the
   Reporting tab of 08 Config.
2. Portal, Admin, Build: Choose folder, Read templates, Write data files.
   Needs Edge or Chrome (File System Access API).
3. Read the checks on the Build page, then open Controller Checks.
4. Write and publish commentary, then on Distribution download the deck and
   write the emails to `outbox/`.

What Build does that the manuals list as manual: codes Stripe fee rows to
Admin, splits legal fees by contact, spreads the department group budget
across accounts on last year's mix when the Budget tab has no account rows,
folds client aliases, flags consulting schedule lines with budget and no
revenue as lost or deferred.

### The two ledger layouts Build accepts

- **Preferred:** the saved Xero custom report "Account transactions for P&L
  analysis" with the thirteen columns (Account Code, Account, Source, Date,
  Contact, Debit, Credit, Job Numbers, Invoice Number, Reference, Description,
  Cost Centres, plus GST and tax columns if present). Several exports can be
  pasted one under another. This carries cost centres and job numbers, which
  the department P&L and the Project P&L need.
- **Fallback:** Xero's standard General Ledger Detail report pasted whole.
  Build reads the account from the heading rows. It carries no cost centres or
  job numbers, so nothing is coded to a department.

### Pay runs

03 Employment Hero Earnings takes the Earnings Details tab of each Pay Run
Audit report, pasted one run under the last with its title row. Build splits
the runs on the title rows, works out each run's fortnight, spreads undated
lines over the run's weekdays by month, reads the department from the
Locations tab or the bracket tag, and keeps hours by job from the Location
External Id. The Employees tab maps external id to department, employment
type and utilisation target.

### What the Xero connector can and cannot do

A Claude session with the Xero MCP connector attached (organisation id
`3993749e-5d5f-4997-bb08-96616f256085`) can pull: profit and loss by period,
aged receivables and payables, financial position totals, invoices, bills,
contacts, chart of accounts, tracking categories, repeating templates. It
cannot pull the general ledger or journals, so 02 Xero GL Transactions is
always a paste from a Xero export. Pulled P&L and aged data were written into
templates 01, 11 and 12 by a script; the templates then went through Build
like any paste.

---

## 5. The September 2026 build

Inputs used: the Xero P&L Jul-25 to Sep-26 from the connector, the custom
Account Transactions export 1 July 2025 to 30 September 2026 (two files,
FY26 and FY27), eight Pay Run Audit reports (to the run ending 18 September),
the Zoho All Deals by Stage export at end of September, aged receivables and
payables at 30 September, the FY27 budget by department group, and the
revenue schedule lines from the FY27 budget workbook.

| | September 2026 |
|---|---|
| Revenue | $561,945 |
| Gross profit | $187,814 |
| Net result | ($41,464) |
| By department | Onsite $242,948, Integration $234,858, Production $73,304, Video $7,303, Consulting $3,531 |
| Utilisation (pay runs to 18 Sep) | Onsite 98.9%, Production 47.4%, Consulting 19.1% |
| Danica's own utilisation report | Onsite 90.3%, Production 38.3%, Consulting 21.8% |
| Jobs with activity | 155 in September, 201 year to date |
| Income with no job number | $26,722 in September |

The utilisation difference is explained: the pay run data stopped at 18
September and the two methods count differently (section 3). It is not a bug.

**Still placeholder in the September build**

- Users and Distribution rows on 08 Config are template examples.
- Staff and Office Dept split percentages are placeholders.
- Cash balances, cards and commitments are empty (10 not filled).
- 14 Balance Sheet and Trial Balance was not supplied, so the balance sheet
  checks read not loaded.
- Hours per job are blank on Project P&L because the September utilisation
  file was built before the converter kept hours by job. The next Build with
  the pay runs fills the column.
- The deck button does not work on the Claude artifact because the PptxGenJS
  vendor file is not published there. It works from the folder or zip.

**Awaited from Danica**

- Pay Run Audit reports ending 2 October and 16 October.
- Balance Sheet and Trial Balance at 30 September for template 14.
- Staff and Office Dept split percentages from the budget workbook.
- Real Users and Distribution rows for 08 Config, and the SharePoint link.
- Whether the August work in progress jobs (Barco E3, PwC screens, Deloitte
  microphones) were invoiced in September.
- OnRent and Qwilr exports so templates 05 and 06 can be matched.
- Bank and card balances and commitments for 10.
- An account level FY27 budget if one exists.

---

## 6. Accounting findings from the September review

An 18-finding review workbook (`CTS accounts review - Sep 2026 close.xlsx`)
was given to Danica on the September close. These are hers to action in
Xero, not portal faults, but the next person should know them because the
controller checks keep flagging them until they are fixed:

- Work in progress of about $252k parked in August with no matching release.
- No depreciation posted; an instant asset write off used as a plug.
- Equity movement does not equal profit.
- A $500k transfer to CTS Holdings with no narration.
- Entertainment claimed with GST.
- Leave provisions posted in lumps rather than monthly.
- Superannuation split across accounts inconsistently.
- An insurance premium expensed in one month rather than prepaid.
- Overdue receivables and payables beyond the thresholds.
- Chart hygiene: duplicate and unused accounts, every repeating template in
  Xero is DRAFT or stale.

One correction from Danica to carry forward: RJW INV-0048 is one $12k
invoice entered as two $6,000 bills (August and September), not a duplicate.
The duplicate check now ignores deposit, balance, progress, claim, stage and
instalment references for that reason.

---

## 7. Publishing the full site

Three ways, from most to least suitable for real use.

### A. SharePoint library (the intended home)

`docs/HOSTING.md` has the steps. Copy the whole `portal` folder with the
live `data` folder into a document library, set Finance to Edit and
everyone else to Read, everyone syncs it and opens the HTML from the synced
copy. This is the only option where the deck, Build and the outbox all work,
and the only one with real access control.

### B. A private Claude artifact

This is what the current artifact is. Steps for a Claude session with the
Artifact tool:

1. Publish `CTS Business Intelligence Portal.html` as the page.
2. Pass the rest as `files`: `CTS_bi_bundle.js`, `CTS_build.js`,
   `CTS_email.js`, `CTS_deck.js`, every `data/*.js`, `vendor/xlsx.full.min.js`,
   `README.md`, `docs/*.md`, `build/*.json`. The Python file must be sent with
   `contentType: "text/plain"`; `.xlsx`, `.pptx` and `.zip` cannot be attached
   at all.
3. `vendor/pptxgen.bundle.js` is refused by the artifact service, so the
   deck button on the artifact says the module is missing. Everything else
   works. Build on the artifact can read templates but cannot write into the
   artifact, so Build is for checking only there.
4. The ledger file is about 4.9MB for fifteen months, well under the 16MB
   page limit. Watch it as the years grow.
5. The artifact is private until shared from its Share menu. Anyone it is
   shared with can read every data file, including the ledger, whatever
   their portal role.

To move the site to a work account, publish a new artifact there from the
same files (an artifact cannot be transferred between accounts). The
September live data files are in the zip sent to Danica and can be read
back from the current artifact with the Artifact read action while it
exists.

### C. Any static host

Any web server that serves the folder works. Do not use a public host
(GitHub Pages included) with real data.

**In every case:** never commit `data/` files built from real exports to the
repository. The repository holds the seed only. Real client names, wages and
ledger lines stay in the synced folder, the artifact or the zip.

---

## 8. Testing without a browser window

The portal has no test suite. The pattern that has been used is a harness
page served from a copy of the folder and run in headless Chromium:

```
cp -r portal /some/scratch/portal
cd /some/scratch/portal && python3 -m http.server 8765 &
chrome --headless=new --no-sandbox --disable-gpu --virtual-time-budget=60000 \
  --dump-dom http://localhost:8765/harness.html
```

The harness page loads the data files and the bundle in the same order as
the real page, sets an identity in localStorage (`cts.bi.identity`), calls
`CTS.engine.init()`, then renders a page with `CTS.pages.<id>.render()` into
a div and writes results into `<pre id="out">`. For Build, it fetches the
templates as ArrayBuffers and calls `B.runFromBuffers(buffers, { reportingMonth })`,
which returns the checks and the data file contents without touching disk.

Checks worth keeping: ledger ties to the P&L in every month; job income
plus no-job income equals P&L income; derived account budget ties to the
group budget; every page renders without an error; the deck builds and
passes the pptx skill's `validate.py` (needs `lxml` and `defusedxml`).

Syntax check before any commit:

```
node -e "new Function(require('fs').readFileSync('CTS_bi_bundle.js','utf8'))"
node -e "new Function(require('fs').readFileSync('CTS_build.js','utf8'))"
```

Python Playwright is not installed in the cloud container; the Chromium
binary path under `/opt/pw-browsers` changes between container versions, so
find it with `find /opt/pw-browsers -name chrome -type f`.

---

## 9. Working rules

These were set by Danica or by the repository and should carry over.

- Australian spelling and context throughout: financial year 1 July to 30
  June, GST, superannuation, state public holidays.
- No em dashes or en dashes in repository files. Use a comma, a full stop or
  the word "to".
- Never commit real data. Column shapes and derived names only.
- Never open a pull request unless asked. Develop on the designated branch
  and push there only.
- Do not put model names or identifiers in commit messages, code or
  comments. End commits with whatever attribution lines the session requires.
- The emails and the deck must never leak the in-portal department filter.
- The data files are generated; the one hand edited file is
  `data/CTS_config_data.js`, and even that is normally rebuilt from 08 Config.
- Templates are generated by `build/build_templates.py`; edit the script,
  not the workbook, or the next regeneration discards the edit.
- Talk to Danica in plain English, answer first, workings underneath, and
  say clearly when something is a placeholder or unverified.

---

## 10. Backlog

In rough priority order.

1. Rebuild September with the two October pay runs and the balance sheet and
   trial balance, so utilisation covers the whole month, hours per job fill,
   and the balance sheet checks run.
2. Real Users and Distribution rows, the SharePoint link, and the Staff and
   Office Dept split percentages, then set Confirmed on Config.
3. Fill 10 Cash and Commitments so the Cash Flow page and the cash tile have
   something to show.
4. Match templates 05 OnRent Orders and 06 Qwilr Quotes to real exports so
   the confirmed work layer of the revenue forecast is live.
5. October close: the same routine with the reporting month moved on.
6. Forecast accuracy (FORECASTING.md step 5) once three months of archive
   exist.
7. Optional: the Power Automate flow (AUTOMATION.md stage 2) so the outbox
   sends itself.

---

## 11. A starting prompt for the next session

Paste this, with the repository attached, into the new session:

> This is the CTS Business Intelligence Portal in `portal/` on branch
> `claude/cts-replica-n0d209`. Read `portal/docs/HANDOVER.md` first, then
> `portal/README.md` and `portal/docs/MONTHLY.md`. The data in the
> repository is seed; the live September 2026 data is in the zip I will
> upload. Do not commit real data. Australian spelling, no em dashes in
> repository files, no pull requests unless I ask. First job: [publish the
> full site as a private artifact from the portal folder with the live data
> / rebuild September with the October pay runs / close October].
