# CTS Business Intelligence Portal: full handover

Written 8 November 2026 for the person or Claude agent who takes this over on
a local PC without access to the original repository or chat. Everything in
it comes from the code in this folder and from the work done with Danica
Nelson (finance head, Corporate Technology Services) between September and
November 2026. Where something is a placeholder, unverified or lost, it says
so.

Reading order: this file, then `CLAUDE.md` at the repository root, then
`portal/README.md`, `portal/docs/MONTHLY.md`, `portal/docs/HOSTING.md`,
`portal/docs/AUTOMATION.md`, `portal/docs/FORECASTING.md`. The older
`CTS Technical Handover.md` at the root covers the payroll and reporting
manuals and workbooks that came before the portal; background only.

Contents

1. What this is and where it is up to
2. Setting up on a local PC
3. Architecture
4. Data model
5. The Build pipeline
6. The templates, tab by tab
7. Controller checks
8. Deck, emails and the fortnightly update
9. Forecast and cash
10. The September 2026 build
11. Accounting findings to carry forward
12. Publishing the site
13. Testing
14. Decisions and traps learned the hard way
15. Working rules and how Danica likes to work
16. Backlog
17. Starter prompt for the work agent

---

## 1. What this is and where it is up to

A static business intelligence portal for Corporate Technology Services Pty
Ltd (CTS), an Australian audio-visual and technology services company based
in Victoria. Financial year 1 July to 30 June, reporting in AUD. Six
departments: Onsite (managed services at client sites), Production (events),
Video, Integration (systems installs), Consulting, Admin. Accounting is Xero,
payroll is Employment Hero (fortnightly), CRM is Zoho, hire orders are
OnRent, quotes are Qwilr.

The portal is one HTML file, four scripts, two vendored libraries and a
folder of data files. It runs entirely in the browser from a synced folder.
No server, no backend, no network calls at runtime, no accounts, nothing to
install. It replaces the manual Monthly Reporting process (the P&L analysis
file, the graphs file and the management report) described in the Monthly
Reporting manuals in `manuals/`, and it adds things the manual process never
had: a revenue risk profile, a project P&L, an auditor style controller
check schedule, a forecast, a cash flow, a Monthly Business Review deck and
the monthly emails.

Finance pastes exports into fourteen Excel templates, opens the portal,
clicks Build, and the portal rewrites its own data files. Everyone else
opens the same folder and reads. OneDrive or SharePoint sync carries the
files.

### State at handover

| Item | State |
|---|---|
| Code | Working. 33 pages, Build for all 14 templates, three email tiers, fortnightly update, deck, 28 controller checks, revenue risk, project P&L, forecast, cash flow |
| Repository | GitHub `WesleyNelson-O33/Claude`, branch `claude/cts-replica-n0d209`, folder `portal/`. Latest commits: `0ef956f` Project P&L page, `b77649e` handover. The work agent will not have access to this remote; the zip carries the full clone with history |
| Data in the repository | Seed only. Fictional client names, modelled months between real anchors. The seed config's reporting month is August 2026 |
| Live data | September 2026 is closed and built. The data files are in `portal/_live_september/data/` in the handover zip (ignored by git on purpose) and on a private Claude artifact in Danica's personal account. They must never be committed |
| September result | Revenue $561,945, gross profit $187,814, net loss $41,464. Ledger ties to the closed Xero P&L in every month. Controller checks: 0 failed, 12 look |
| Hosting | Not yet hosted for the team. Danica has the folder as a zip. The intended home is a SharePoint library (section 12) |

---

## 2. Setting up on a local PC

### 2.1 What to install (Windows)

| Tool | Why | Needed for |
|---|---|---|
| Microsoft Edge or Google Chrome | The Build page uses the File System Access API, which only these have | Build, Distribution writes |
| Git for Windows | Version control, pushing to the work GitHub | Development |
| Node.js 20 or later | Syntax checks of the scripts, no npm packages needed | Development |
| Python 3.11 or later with `openpyxl` | Regenerating the templates (`build/build_templates.py`), serving the folder for headless tests | Development, template changes |
| Claude Code (desktop or CLI) | The work agent | Development |

Nothing is needed to run the portal itself. Double-clicking the HTML file
works from a local folder. If a browser blocks sibling file loads from a
`file://` path, serve the folder instead:

```
cd portal
python -m http.server 3000
```

and open `http://localhost:3000/CTS%20Business%20Intelligence%20Portal.html`.

### 2.2 Unpacking the handover zip

```
CTS Portal Handover/
  START HERE.md                 what is in the zip and the order to read
  CLAUDE.md                     instructions for the Claude agent (also at repo root)
  HANDOVER.md                   this file (also at portal/docs/HANDOVER.md)
  repo/                         the full repository clone, .git included
    CLAUDE.md
    CTS Technical Handover.md   the older manuals and workbooks handover
    portal/                     the portal (see 3.1)
      _live_september/data/     the real September 2026 data files, git-ignored
    manuals/ pack/ reporting/ tools/ checklists/   the earlier deliverables
```

Put `repo/` somewhere sensible, for example `C:\Work\CTS\Claude`. Open that
folder in Claude Code. `CLAUDE.md` at its root is read automatically.

### 2.3 Pointing the clone at the work GitHub

The clone still points at the personal remote, which the work account
cannot reach. Change it:

```
cd repo
git remote remove origin
git remote add origin https://github.com/<work-org>/<work-repo>.git
git push -u origin claude/cts-replica-n0d209
```

Or push the branch as `main` if the work repository is new:

```
git push -u origin claude/cts-replica-n0d209:main
```

Before the first push, confirm no real data is staged: `git status` must
not list anything under `portal/_live_september/` (it is covered by the
`portal/_*` rule in `.gitignore`) and `portal/data/` must still hold the
seed files (their `meta.seed` is `true`).

### 2.4 Making the portal show the real September data locally

```
cd portal
xcopy /E /Y _live_september\data data
```

(or copy the files in Explorer). Then open the HTML. The seed banner
disappears and the reporting month reads September 2026. To go back to the
seed, `git checkout -- portal/data`.

Never commit after that copy without restoring the seed first. The simplest
habit: keep two folders, the repository clone with seed data for development,
and a separate hosted copy with live data that is never a git repository.

### 2.5 Regenerating the templates

```
pip install openpyxl
python portal/build/build_templates.py
```

Writes all fourteen workbooks into `portal/templates/`. Run this after any
change to the page list (08 Config carries pages and roles), the revenue
schedule lines (`build/schedule_lines.json`) or the budget groups
(`build/budget_groups.json`). Templates are generated; never hand edit the
workbooks in the repository.

### 2.6 Rebuilding the seed data

```
python portal/build/build_data.py
```

Deterministic. Reads `reporting/data/accounts.json` and
`reporting/data/validation.json` and writes every seed file in
`portal/data` except `CTS_config_data.js`. Rarely needed.

---

## 3. Architecture

### 3.1 Files

```
portal/
  CTS Business Intelligence Portal.html   page shell, CSS, splash, the loader
  CTS_bi_bundle.js     engine E, UI U, formatters F, pages P, commentary GEN, router, store
  CTS_build.js         Build page: template registry, readers, converters C.*, checks, writers B.*
  CTS_email.js         monthly and fortnightly emails, module M
  CTS_deck.js          Monthly Business Review deck, module D (PptxGenJS)
  vendor/xlsx.full.min.js      SheetJS, reads the templates
  vendor/pptxgen.bundle.js     PptxGenJS, writes the deck
  data/                the data files the page loads
  templates/           01 to 14
  build/               build_templates.py, build_data.py, fill_test_templates.py, schedule_lines.json, budget_groups.json
  docs/                MONTHLY, HOSTING, AUTOMATION, FORECASTING, HANDOVER
```

Sizes: bundle about 426 KB, build 130 KB, deck 38 KB, email 35 KB, HTML 36 KB.
The seed ledger file is about 2 MB; the live September ledger (fifteen
months) is about 4.9 MB.

### 3.2 Loading order

The HTML's inline loader appends one script tag at a time.

1. Mandatory data: `CTS_config_data`, `CTS_cal_data`, `CTS_accounts_data`,
   `CTS_gl_data`, `CTS_fin_data`. A missing one replaces the splash with a
   message naming the file.
2. Optional data: `CTS_util_data`, `CTS_clients_data`, `CTS_users_data`,
   `CTS_staff_data`, `CTS_brand_data`, `CTS_pipeline_data`,
   `CTS_forecast_data`, `CTS_cash_data`, `CTS_commentary_data`,
   `CTS_distribution_data`, `CTS_aged_data`, `CTS_checks_data`. A missing one
   is logged and the pages that need it say so.
3. `vendor/xlsx.full.min.js`, `vendor/pptxgen.bundle.js`, `CTS_build.js`,
   `CTS_email.js`, `CTS_deck.js`. Each optional; a missing module disables
   its page or button.
4. `CTS_bi_bundle.js`, then `CTS.boot()`.

Everything hangs off `window.CTS`. The modules attach to it in whatever
order they load, and the bundle resolves them lazily, so load order between
the modules does not matter.

### 3.3 The bundle, `CTS_bi_bundle.js`

One IIFE per concern, in this order.

**Formatters `F` (`CTS.fmt`)**: `money` (cents to `$1,234` with brackets for
negatives in tables), `dollars` (whole dollars), `k` (thousands), `pct`,
`hours`, `num`, `date`.

**Store `CTS.store`**: `get(key, default)` and `set(key, value)` over
localStorage with the prefix `cts.bi.`. Everything a user changes in the
browser lives here: `identity`, `period`, `theme`, `previewRole`,
`focusDept`, section collapsed states, `pnlDeptExOI`, `pnlDeptStaff`,
`proj.search`, `proj.lifetime`, `proj.overheads`, `proj.sel`,
`checks.<month>` (reviewed ticks and notes), commentary drafts, recipient
edits. Nothing in the store is shared between people; publishing writes a
data file instead.

**Engine `E` (`CTS.engine`)**, the part that matters most:

- Loading. `E.init()` runs `loadGL`, `loadFin`, `loadUtil`, `loadUsers`,
  `loadPipeline`, `loadStaff`, `loadForecast`, `loadCash`, `loadCommentary`.
  `loadGL` walks every ledger row once and builds every index the pages
  read: by department, category and month; by account and month; by
  contact; by sub group; by job (`E.jobs`, `E.idx.jobCatMonth`,
  `jobAcctMonth`, `jobSubMonth`). Client aliases from the Clients tab are
  folded here through `E.aliasOf`.
- Departments. `E.resolveDept(costCentre, jobNo)` applies the rule: Cost
  Centres column first, then the bracket tag on the job number. `E.deptOf`
  is the department map from config. `E.viewDept()` is the department picker
  value; `E.withoutView(fn)` runs a function with the picker cleared, which
  is how the emails and deck stay company wide.
- Calendar and periods. `E.monthIdx`, `E.reportingMonth()`,
  `E.setReportingMonth`, `E.currentFY()`, `E.monthsOfFY`, `E.ytd`,
  `E.quarterMonths`, `E.priorYearMonths`, `E.hasActual(month)`.
  `CTS.period()` returns the selected period `{id, label, keys, fy}` and
  `CTS.periodBar()` renders the picker with the department picker beside it.
- P&L. `E.acctActual`, `E.acctBudget`, `E.acctBlend` (actual to the
  reporting month, budget or forecast after), `E.buildPnl`, `E.pnlActual`,
  `E.pnlBudget`, `E.pnlBlend`, `E.pnlForecast`. `E.CAT_KEYS` is
  `income, cos, expenses, other_income, other_expenses`; `E.CAT_LABEL` the
  labels.
- Allocation. `E.splitExact`, `E.basis`, `E.basisWeights`, `E.subSplits`,
  `E.basisFor(sub)`, `E.deptRaw`, `E.allocate`, `E.allocationIsSourced`,
  `E.deptShareOfAccount`, `E.budgetByDept`. The three bases are 3 Way (34/33/33,
  confirmed), Staff and Office Dept (placeholders). Each overhead sub group is
  assigned a basis on the Config Split assignment tab.
- Risk. `E.risk(actual, budget)` tests the materiality floor ($5,000) before
  the percentage (10% medium, 25% or $50,000 high). `E.riskRank`.
- Utilisation. `E.utilFor(dept, keys)`, `E.utilRolling`, `E.workingDays`,
  `E.businessDays`, `E.hoursPerDay`, `E.utilState`. Chargeable over chargeable
  plus non chargeable; leave and public holidays excluded. FTE is worked hours
  over working days for the state times hours per day.
- Revenue. `E.clientRows`, `E.schedule` (budget, actual, last year per
  schedule line, lost or deferred flag), `E.revenueRisk(keys)`,
  `E.arByContact`, `E.revenueMix`, `E.mixOf`, `E.gmScenario`,
  `E.subByContact` (legal fees by contact), `E.plCheck`.
- Projects. `E.jobKey(jobNo)` returns the leading number of a job string
  such as `2604704 - Aware Super Seminar Room Audio [INT]`; `E.jobName`
  strips the number and the tag. `E.projectRows(keys, {overheads})` and
  `E.projectDetail(id, keys)`. Overhead jobs are the 9000 series.
- Forecast. `E.FC_METHODS`, `E.forecastRule`, `E.forecastMonths`,
  `E.forecast`, `E.forecastLayers`, `E.acctForecast`, `E.forecastByDept`,
  `E.forecastLines`, `E.pipelineSettings`, `E.pipelineTargets`.
- Cash. `E.cashSettings`, `E.cashBalances`, `E.cashCards`, `E.cashflow`
  (twelve months, indirect method).
- People and access. `E.identity`, `E.setIdentity`, `E.currentUser`,
  `E.currentRole`, `E.role`, `E.rolePages`, `E.setRolePages`, `E.userList`,
  `E.can(pageId)`, `E.isAdmin`, `E.isRealAdmin`, `E.previewRole`,
  `E.deptScope`, `E.visibleDepts`, `E.focusDept`.
- Commentary. `E.commentaryFor(page, month)`, `E.addCommentary`,
  `E.updateCommentary`, `E.removeCommentary`, `E.commentaryExport`,
  `E.canComment`, `E.canPublishCommentary`, `E.proposeCommentary(pageId)`
  which calls `GEN[pageId]`.
- Fortnight. `E.progressMonth`, `E.snapshotDate`, `E.fortnightFor`.
- Snapshots. `E.snapshot`, `E.snapshots`, `E.restoreSnapshots` keep the
  last loaded data in IndexedDB so a load survives a refresh.
- `E.CHECKLIST` is the list of controller checks for the page.

**Commentary generators `GEN`**: one function per page that drafts a
paragraph from the page's numbers with square brackets where the why is
needed. Aliases: `profit-fte` uses `util`, `pipeline` uses `rev-forecast`,
`client-dept` uses `clients`, `story` uses `home`, `pnl-spread` and
`actions` use `pnl`. `GEN.projects` exists.

**UI `U`**: `tile`, `section(title, kids, note, opts)` (collapsible card,
remembers state, `opts.badge`, `opts.collapsed`), `table(cols, rows, opts)`
(a column is `{key, label, align, width, fmt(v, r), value(r), cell(r)}`
where `cell` returns a class name), `figure`, `columns`, `lines`, `bars`,
`waterfall`, `spark` (all hand written SVG), `legend`, `note`, `source`,
`flag`, `cover`, `icon`, `chev`, `commentary`, `money`, `moneyCell`,
`SERIES` and `colourOf` (six fixed department colours, never cycled).

**Pages `P`**: an object per page `{section, title, sub, render()}`.
`render` returns an array of DOM nodes built with `h(tag.class, attrs,
children)`. The 33 pages are listed in 3.5. `P.build`, `P.distribution`,
`P.access`, `P.config`, `P.setup`, `P.about` are the admin pages.

**Router and boot**: hash routing `#/<pageId>`. `CTS.router.render(id)`,
`CTS.router.reload()`, `CTS.router.current`. `CTS.boot()` initialises the
engine, builds the sidebar from the role's pages, restores theme and
collapsed state, and renders the current page. `CTS.seedBanner()` and
`CTS.allocBanner()` are the warnings shown while data is seed or an
allocation rests on a placeholder.

### 3.4 Conventions the engine is built on

- Money is held in cents as integers. Formatters divide by 100. Hours are
  decimal hours.
- Amount is credit less debit. Income positive, costs negative. Gross profit
  is income plus cost of sales.
- Department comes from the transaction, not the account: Cost Centres
  column first, then the bracket tag on the job number (`[ONS]`, `[PRD]`,
  `[VID]`, `[INT]`, `[CONS]`, `[CTS]`). Unmatched lines read UNALLOCATED.
- Actual to the reporting month, then budget (or forecast when enabled), so
  the year always reads as twelve months.
- The risk flag tests the materiality floor before the percentage.
- A department's own overhead and its share of the Admin pool are two
  separate columns, because the documented failure is counting one twice.
- Two utilisation figures exist and are not meant to agree (3.3 Utilisation
  versus the graphs file's rolling twelve months against 2,080 or 2,088).
- Every page follows the department picker except Cash. The emails and the
  deck ignore the picker; a department head's role fixes what they get.
- Visibility is not security. Anyone with the folder can read the data
  files. A department head who must not see the ledger needs a separate
  restricted build.

### 3.5 Pages and roles

Pages marked * are sensitive and hidden from Viewer and Department head.

| Section | Pages |
|---|---|
| Dashboard | Dashboard, Business Context, Commentary, The Month |
| Finance | P&L*, P&L by Department*, P&L Spread*, Overhead Allocation*, P&L Control*, Controller Checks*, Trends |
| Budget | Budget vs Actual*, Actions*, Forecast* |
| Cash | Cash Flow* |
| Revenue | Revenue Summary, Revenue Schedule, Revenue Forecast, Revenue Risk, Project P&L*, Pipeline, Top Clients, Clients by Department |
| Utilisation | Utilisation, Profitability per FTE*, Profitability by Employee* |
| Ledger | Detail Records* |
| Admin | Setup, Build*, Distribution*, Config & Variables, Access Control*, About |

| Role | Pages | Notes |
|---|---|---|
| admin (Finance head) | 33 | Can change who sees what. One person |
| finance | 32 | Runs the month |
| exec | 24 | Company result, no ledger or workings |
| deptHead | 15 | Own department only |
| viewer | 4 | Headline and context |

Users and roles come from the Users and Tabs by role tabs of 08 Config. The
rows in the templates and the seed are examples (Danica Nelson admin, Duncan
exec, Graham exec, Jordan deptHead Consulting, Production manager deptHead).
Real names and addresses are awaited.

### 3.6 Adding a page

1. `P.<id> = { section, title, sub, render }` in the bundle.
2. The PAGES and ROLES lists in `build/build_templates.py`, so 08 Config
   carries it, then regenerate the templates.
3. `data/CTS_users_data.js` for the seed.
4. Optionally `GEN.<id>` for proposed commentary.

Commit `0ef956f` (Project P&L) touches all four and is the example.

---

## 4. Data model

Every file is `window.CTS_<NAME> = {...}` with a `meta` block: `built`,
`source`, `seed` (true on seed data), sometimes a `note`.

| File | Shape | Built from |
|---|---|---|
| CTS_config_data | `ORG` (name, fyStartMonth, reportingMonth, priorFY, currentFY, portalUrl), `DEPARTMENTS` (code, short, long, order, isRevenue, utilTarget, gmNorm, npNorm, tags, suffix, normNote), `DEPT_RESOLUTION`, `SPLIT_BASES` (3WAY, STAFF, OFFICE with groups PRD/VID, ONS, CONS/INT and `source` confirmed or placeholder), `SUB_SPLITS` (production 0.79 of PRD/VID, integration 0.72 of CONS/INT), `SPLIT_ASSIGNMENT` (sub group to basis), `MANUAL_LINES`, `RISK`, `UTIL`, `SEASON`, `COMMENTARY`, `CATEGORY_ORDER`, `DISTRIBUTION`, `VERSION`, `META` | 08 Config |
| CTS_cal_data | `months[]` (key, y, m, fy, label, long, period, quarter, wd per state, bd, days), `holidays[]`, `quarters[]` | 08 Config Holidays tab, three years across |
| CTS_accounts_data | Account name to `{code, cat, sub, dept}` | The P&L and ledger |
| CTS_gl_data | `cols` then `rows`, columns `date, account, contact, debit, credit, source, jobNo, invoiceNo, costCentre, description, gst, taxName, narration`. Cents. `meta.layout` says custom or detail | 02 |
| CTS_fin_data | Per account per month actual and budget, cents | 01 and 07 |
| CTS_util_data | `rows` with `month, dept, chargeable, nonChargeable, leave, publicHoliday, fte`; `jobs` rows `[month, jobExternalId, hours]` | 03 |
| CTS_staff_data | `month, empId, name, dept, chargeable, nonChargeable, leave, publicHoliday, grossCents, superCents, employment, target` | 03 |
| CTS_clients_data | Display name, aliases, department, schedule lines | 13 and the ledger |
| CTS_pipeline_data | `deals` (Zoho: name, stage, amount cents, closing date, department, job), `orders`, `quotes`, settings | 04, 05, 06 |
| CTS_forecast_data | Methods per category, growth, pipeline weights, overrides | 09 |
| CTS_cash_data | `settings` (debtorDays, creditorDays, basFrequency, superTiming, paygInstalment, facility, openingAR, openingAP, cardsPaidInFull), `balances`, `cards`, `commitments` | 10 |
| CTS_aged_data | `receivables` and `payables`: `contact, date, due, number, ref, current, lt1, m1, m2, m3, older, total` cents | 11 and 12 |
| CTS_checks_data | The check results for the month: id, area, title, status, items, cols, detail, guide | Computed at Build |
| CTS_users_data | `PAGES`, `ROLES`, `USERS`, `DEFAULT_ROLE` | 08 Config |
| CTS_brand_data | Logos and marks as data URIs, palette | Fixed |
| CTS_commentary_data | Published notes `{id, page, month, dept, text, author, at}` | Commentary page |
| CTS_distribution_data | Published recipient list | Distribution page |

Only `CTS_config_data.js` is meant to be edited by hand, and even that is
normally rebuilt from 08 Config.

---

## 5. The Build pipeline

`CTS_build.js`. The Build page runs in Edge or Chrome and needs the File
System Access API to read the templates and write the data files.

### 5.1 Template registry

```
key       file                                   sheet          writes
pnl       01 Xero P&L.xlsx                       P&L            CTS_fin_data.js        required
gl        02 Xero GL Transactions.xlsx           GL             CTS_gl_data.js         required
earnings  03 Employment Hero Earnings.xlsx       Earnings       CTS_util_data.js, CTS_staff_data.js
deals     04 Zoho Deals.xlsx                     Deals          CTS_pipeline_data.js
orders    05 OnRent Orders.xlsx                  Orders         CTS_pipeline_data.js   placeholder shape
quotes    06 Qwilr Quotes.xlsx                   Quotes         CTS_pipeline_data.js   placeholder shape
budget    07 Budget.xlsx                         Budget         CTS_fin_data.js        required
config    08 Config.xlsx                         all tabs       CTS_config_data.js, CTS_cal_data.js, CTS_users_data.js   required
forecast  09 Forecast.xlsx                       Methods        CTS_forecast_data.js
cash      10 Cash and Commitments.xlsx           Bank balances  CTS_cash_data.js
ar        11 Xero Aged Receivables.xlsx          Receivables    CTS_aged_data.js
ap        12 Xero Aged Payables.xlsx             Payables       CTS_aged_data.js
schedule  13 Revenue Schedule.xlsx               Schedule       CTS_clients_data.js
bstb      14 Balance Sheet and Trial Balance     Balance Sheet  CTS_checks_data.js
```

### 5.2 Steps

1. `B.pickFolder()` asks for the portal folder. `B.readTemplates()` reads
   each workbook with SheetJS into array buffers.
2. `B.runFromBuffers(buffers, opts)` does all the work without touching
   disk, which is what the tests call. `opts.reportingMonth` overrides the
   month from 08 Config.
3. `B.table(wb, sheet, {minFilled, findHead})` finds the header row in a
   sheet. `findHead` is a predicate on the row, needed because pasted Xero
   reports carry title rows above the header.
4. Converters run in order: config (departments, calendar across three
   years, users), pnl, budget and `C.budgetDept`, gl (layout detection:
   General Ledger Detail heading rows mean `C.glDetail`, else the custom
   report through `C.gl`), earnings, deals, orders, quotes, forecast, cash,
   aged, schedule, balance sheet and trial balance, then `C.checks`.
5. Checks are collected per step and shown on the page: ledger against P&L
   by month and category, rows read, accounts not in the chart, locations
   with no department, back pays reallocated, placeholder split bases,
   schedule lines with no client match, aliases applied, Stripe rows coded,
   revenue rows with no contact.
6. `B.write()` copies the current `data/` to `data/_previous`, writes the
   new files, writes `archive/YYYY-MM/` with the data files and the
   templates as pasted, then reloads.
7. `B.writeCommentary`, `B.writeRecipients`, `B.writeOutbox`,
   `B.writeBinary` serve the Commentary and Distribution pages.

### 5.3 The ledger converters

`C.gl` reads the custom Xero report "Account transactions for P&L analysis"
(Account Code, Account, Source, Date, Contact, Debit, Credit, Job Numbers,
Invoice Number, Reference, Description, Cost Centres, optionally GST and Tax
Rate). Several exports can be pasted one under another; the header row is
found by `findHead` looking for account name, date and debit together.
Stripe fee rows with no cost centre are coded to `9000 - OFFICE / ADMIN
[CTS]`. Values must be numbers, not text, before pasting.

`C.glDetail` reads Xero's standard General Ledger Detail report pasted whole
from A1: the account comes from the heading rows above each block, the
contact is parsed from the description, balance sheet lines are kept
separately in `bs` for the checks, and `detail: true` is set on the result
so the pages can say cost centres are missing.

### 5.4 The earnings converter

`C.earnings(table, locations, payCategories, adjustments, fallbackMonth, employees)`.
Pay runs are pasted one under another with their title rows. The converter
splits on "Earnings Details" title rows; each block's fortnight is the
fourteen day window holding most of its dated lines; undated lines are
spread over the block's weekdays by month. Location decides department:
`C.guessLocation` tries a bracket tag, then the tag anywhere in the name,
then non chargeable wording. The Employees tab maps external id or name to
department, employment type and utilisation target; with it, hours go to the
person's department and chargeable follows the job. Hours by job are kept by
Location External Id and emitted as `util.jobs` rows. Pay categories decide
chargeable, non chargeable, leave, public holiday; adjustments reallocate
back pays.

### 5.5 Budget

`C.budget` reads the Budget tab (same layout as the P&L). `C.budgetDept`
reads the Departments tab: Consulting (CONSULTING, INTEGRATION), Production
(PRODUCTION, VIDEO), Support (ONSITE), CTS Overheads (ADMIN). When the Budget
tab has no account rows, the group budget is spread across accounts on the
prior financial year's ledger mix (the prior year is computed arithmetically
because the calendar is built after this step). The derived budget ties to
the group budget to the dollar.

### 5.6 Schedule and aliases

`C.schedule` reads the Schedule tab (one line per onsite role or client per
month with a type) and the Clients tab (display name, aliases, department).
Aliases are folded in `E.loadGL` so a client with two spellings in Xero is
one client everywhere.

### 5.7 Aged, balance sheet, trial balance

`C.aged` reads the Xero Aged Receivables Detail and Aged Payables Detail
layouts as they come. `C.balanceSheet` and `C.trialBalance` read the whole
exports pasted from A1 and feed the checks only.

---

## 6. The templates, tab by tab

Every workbook has a Read me tab that says exactly what to run in the source
system and where to paste. Generated by `build/build_templates.py`.

| Template | Tabs | Notes |
|---|---|---|
| 01 Xero P&L | Read me, P&L | Xero Profit and Loss, year to date, one column per month, three years across. Paste the account rows |
| 02 Xero GL Transactions | Read me, GL | The thirteen column custom report from 1 July. Fallback: General Ledger Detail pasted whole |
| 03 Employment Hero Earnings | Read me, Earnings, Locations, Pay categories, Adjustments, Employees | Earnings Details of each Pay Run Audit report, one under the last with title rows. Locations maps location to department and chargeable; Employees maps person to department, employment, target |
| 04 Zoho Deals | Read me, Deals, Departments | All Deals by Stage (FY26-FY27) exported to Excel, pasted whole from A1. Departments maps Zoho's department field |
| 05 OnRent Orders | Read me, Orders | Placeholder shape until a real export is matched |
| 06 Qwilr Quotes | Read me, Quotes | Placeholder shape until a real export is matched |
| 07 Budget | Read me, Budget, Departments | Budget by account by month (optional) and by department group (prefilled FY27 from the budget workbook) |
| 08 Config | Read me, Reporting, Departments, Split bases, Split assignment, Holidays, Users, Tabs by role, Distribution | Reporting month is the one monthly change. Portal link goes on Reporting |
| 09 Forecast | Read me, Assumptions, Pipeline, Methods, Overrides | Rarely changes |
| 10 Cash and Commitments | Read me, Settings, Bank balances, Credit cards, Commitments | Month end balances, one row per account per month |
| 11 Xero Aged Receivables | Read me, Receivables | Aged Receivables Detail as at month end |
| 12 Xero Aged Payables | Read me, Payables | Aged Payables Detail as at month end |
| 13 Revenue Schedule | Read me, Schedule, Clients | 147 lines prefilled from the FY27 budget workbook, amounts to type; Clients with display names and aliases |
| 14 Balance Sheet and Trial Balance | Read me, Balance Sheet, Trial Balance, Checklist | Whole exports from A1 |

---

## 7. Controller checks

`C.checks(ctx)` in `CTS_build.js`, thresholds in the `CK` object, results
written to `CTS_checks_data.js` and shown on the Controller Checks page with
a Reviewed tick and note per check per month (stored locally) and an Add to
commentary button. Statuses: fail, look, pass, info, na (file not loaded).

| Area | Check | What it does |
|---|---|---|
| Integrity | What was posted | Which files were loaded for the month |
| Integrity | Ledger agrees with the P&L by category | Per category, cents tolerance $1 |
| Integrity | Trial balance balances | Debits equal credits |
| Integrity | Trial balance agrees with the P&L year to date | Per P&L account |
| Integrity | Current year earnings equal the P&L year to date | Balance sheet line against P&L |
| Revenue | Revenue posted without a contact | Manual journals to income with no contact |
| Revenue | Sales invoiced without GST | Tax rate other than GST on Income |
| Revenue | Credit notes raised this month | $500 and over |
| Revenue | Clients billed last month and not this month | Regular clients missing or down by half |
| Revenue | Work in progress | WIP posted this month and earlier postings with no reversal |
| Costs | Credits in cost accounts | Outside journals |
| Costs | Manual journals on P&L accounts | Round numbers of $5,000 or more highlighted |
| Costs | Possible duplicate bills | Same supplier and amount within 45 days, $100 and over; same date and reference excluded; recurring words excluded (monthly, contract, subscription, rent, retainer, marketing, SEO, deposit, balance, progress, claim n, stage n, instalment, and others) |
| Costs | Cost accounts well above run rate | More than 3 times the median prior month, floor $2,000 |
| Costs | Recurring costs with nothing posted | Posted each of the last three months, nothing this month, $3,000 and over |
| Costs | GST on expenses | Entertainment with GST, bank fees with GST, overseas suppliers with GST ($500 and over), mixed treatment on one account |
| Costs | Expensed items that look like assets | Equipment words, $1,000 and over, outside a client job |
| Costs | Largest lines this month | Fifteen largest, informational |
| Payroll | Payroll ratios and movements | Super 11.2% to 12.8% of direct wages, indirect super at most 13.5%, payroll tax 2.5% to 7%, wages within 25% of last month |
| Payroll | Payroll accrual posted and reversed | |
| Payroll | Pay runs against the wage accounts | Gross from the pay runs against the ledger |
| Balance sheet | Receivables ageing and tie | Overdue share over 30%, over 60 days listed, tie to trade debtors |
| Balance sheet | Payables ageing and tie | |
| Balance sheet | Accruals, prepayments, provisions and payroll liabilities moved | Silence on an account that moves monthly |
| Balance sheet | Related party and shareholder balances and postings | Holdings, director loans, drawings, dividends |
| Balance sheet | Credit card postings | |
| Balance sheet | Clearing and suspense accounts | Nil at month end, lines through them listed |
| Balance sheet | Depreciation booked this month | |
| Cut-off | Large invoices in the last days and references naming another month | $10,000 and over in the last three days; bills dated after month end fail |

Each check carries a `guide` sentence saying what to look for, written for
CTS (contract sites bill monthly, production season spikes are expected,
Employsure retainer arrives monthly, and so on).

---

## 8. Deck, emails and the fortnightly update

### 8.1 The deck, `CTS_deck.js`

Module `D`. `D.supported()`, `D.build()` (runs inside `E.withoutView` so the
department picker never leaks), `D.bytes()`, `D.download()`,
`D.fileName()` ("CTS Monthly Business Review - Sep 2026.pptx"),
`D.messages()` for the covering notes. PptxGenJS, 16:9 (10 by 5.625
inches), Calibri, the CTS palette.

Seventeen slides in the September 2026 Monthly Business Review layout
Danica supplied: cover; agenda; divider 01 Financial performance; P&L
summary with EBITDA (add backs: Depreciation, Depreciation ROU, Instant
Asset Write Off, Interest Expense, Income Tax Expense, Under/(Over)
provision for tax) and the last six months chart; P&L per department with
contribution; revenue and contribution chart; risk revenue profile;
divider 02 Customers and cash; top 10 customers; receivables and DSO (AR
over the last three months' revenue per day); top customers year on year;
divider 03 People and utilisation; utilisation by department against prior
month and target; profit based on utilisation (plus five points
sensitivity); divider 04 Summary and actions; summary from commentary;
thank you. Every content slide has a commentary panel that carries the
published commentary for the matching pages, or the proposed commentary
marked as proposed.

Validated structurally with the pptx skill's `validate.py`. LibreOffice
cannot render these files, so visual checks are done in PowerPoint.

### 8.2 The emails, `CTS_email.js`

Module `M`. `M.recipients()` reads the published list or the Config
Distribution tab. `M.render(recipient)` builds the monthly message for the
recipient's tier; `M.renderAll()`. `M.renderFortnightly`,
`M.renderAllFortnightly`. `M.mailto(msg)` for the plain text Outlook route.
All rendering runs inside `E.withoutView`.

| Tier | Who | Content |
|---|---|---|
| 1 | Executives | Company result, departments after the split, top clients under four headings (increased, decreased, new, dropped), utilisation, cash line, full year outlook, what is worth a comment |
| 2 | Department heads | Own department in depth, one line on the rest |
| 3 | Finance | The controls: ledger tie, uncoded lines, placeholders, what the build did |

Messages are written to `outbox/YYYY-MM/<name>.json` and `.html`, never
sent. `docs/AUTOMATION.md` has the Power Automate flow that sends them.

### 8.3 The fortnightly update

Mid month: paste the year to date ledger (it carries the month in progress)
and the latest pay runs, Build, then Write the fortnightly outbox. Fortnight
1 is the 1st to the 15th, fortnight 2 the 16th to month end, picked from the
ledger's latest date. Department heads get revenue by fortnight, hours, sell
rate, utilisation, team, key calls, the other departments in one line, and
the forward view. Written to `outbox/YYYY-MM-F1` or `-F2`.

---

## 9. Forecast and cash

`docs/FORECASTING.md` is the design. Built: step 1 (P&L forecast on run rate
and seasonality with overrides, 09 Forecast), step 2 (twelve month indirect
cash flow, 10 Cash and Commitments), step 3 (pipeline driven revenue from
Zoho with stage weights). Designed only: thirteen week direct cash, and
forecast accuracy once three months of archive exist. OnRent and Qwilr are
placeholder shapes, so the confirmed work layer is empty until an export of
each is matched.

The forecast is a third series beside actual and budget. When it is on, the
months after the reporting month come from the forecast and budget stays as
the comparison. Every forecast figure carries its method, and run rate based
figures are flagged the way placeholder allocations are.

---

## 10. The September 2026 build

### 10.1 Inputs

- Xero Profit and Loss Jul-25 to Sep-26, pulled through the Xero connector
  in Danica's personal Claude and written into 01.
- Account transactions for P&L analysis, 1 July 2025 to 30 September 2026,
  two exports (FY26 and FY27), pasted one under the other into 02.
- Eight Pay Run Audit reports into 03, the last ending 18 September. Two ad
  hoc runs among them.
- Zoho All Deals by Stage at end of September into 04.
- Aged Receivables and Payables at 30 September from the connector into 11
  and 12.
- FY27 budget by department group into 07 Departments; the account level
  budget derived on the FY26 mix.
- Revenue schedule lines from the FY27 budget workbook into 13.
- Three labour hour reports (Consulting, Onsite, Production) and Danica's
  September utilisation report, used to confirm the non chargeable locations
  and compare utilisation.

### 10.2 Results

| | September 2026 |
|---|---|
| Revenue | $561,945 |
| Gross profit | $187,814 |
| Net result | ($41,464) |
| Revenue by department | Onsite $242,948, Integration $234,858, Production $73,304, Video $7,303, Consulting $3,531 |
| Utilisation, portal (pay runs to 18 Sep) | Onsite 98.9%, Production 47.4%, Consulting 19.1% |
| Utilisation, Danica's report | Onsite 90.3%, Production 38.3%, Consulting 21.8% |
| Ledger to P&L | Ties in every month |
| Derived account budget to group budget | Ties to the dollar |
| Controller checks | 0 failed, 12 look, rest pass or info |
| Jobs with activity | 155 in September, 201 financial year to date |
| Income with no job number | $26,722 in September |
| Largest job | 2604704 Aware Super Seminar Room Audio, $200,767 income, 19% GM, quoted $214,690 in Zoho |

The utilisation gap is explained by the pay run cut off at 18 September and
by the two methods counting differently. It is not a defect.

### 10.3 Still placeholder in the September build

- Users and Distribution rows on 08 Config are the template examples.
- Staff and Office Dept split percentages are placeholders; the allocation
  banner shows.
- Cash balances, cards and commitments are empty (10 not filled), so Cash
  Flow shows settings only.
- 14 Balance Sheet and Trial Balance was not supplied; those checks read not
  loaded.
- Hours per job on Project P&L are blank because the September utilisation
  file was built before the converter kept hours by job. The next Build with
  the pay runs fills the column.

### 10.4 Awaited from Danica

- Pay Run Audit reports ending 2 October and 16 October (and onward).
- Balance Sheet and Trial Balance at 30 September for 14.
- Staff and Office Dept split percentages from the budget workbook.
- Real Users and Distribution rows for 08 Config, and the SharePoint link.
- Whether the August work in progress jobs (Barco E3, PwC screens, Deloitte
  microphones) were invoiced in September.
- OnRent and Qwilr exports to match templates 05 and 06.
- Bank and card balances and commitments for 10.
- An account level FY27 budget if one exists.

---

## 11. Accounting findings to carry forward

An eighteen finding review was given to Danica at the September close as a
workbook (`CTS accounts review - Sep 2026 close.xlsx`, in her files, not in
the repository). The controller checks keep flagging these until they are
fixed in Xero. They are hers to action, not portal faults.

- Work in progress of about $252k parked in August with no matching release
  or reclass.
- No depreciation posted during the year; an instant asset write off used as
  a plug.
- Equity movement does not equal profit.
- A $500k transfer to CTS Holdings with no narration or documented terms.
- Entertainment claimed with GST.
- Leave provisions posted in lumps rather than monthly.
- Superannuation split inconsistently between direct and indirect accounts.
- An insurance premium expensed in one month rather than prepaid.
- Receivables and payables overdue beyond the thresholds.
- Chart hygiene: duplicate and unused accounts; every repeating invoice and
  bill template in Xero is DRAFT or stale.

Corrections from Danica already applied: RJW INV-0048 is one $12k invoice
entered as two $6,000 bills (August and September), not a duplicate; the
duplicate check now ignores deposit, balance, progress, claim, stage and
instalment references.

---

## 12. Publishing the site

### 12.1 SharePoint library (intended home)

`docs/HOSTING.md` has the steps: a document library `CTS Business Portal`,
copy the whole `portal` folder with the live `data` into it, Finance gets
Edit and everyone else Read, everyone syncs it and marks it Always keep on
this device, and opens the HTML from the synced folder. This is the only
option where Build, the deck, the outbox and the archive all work, and the
only one with real access control. Clicking the file on the SharePoint web
view downloads it rather than running it; tell people up front.

### 12.2 A private Claude artifact (what Danica's personal account used)

For a Claude session with the Artifact tool. Publish
`CTS Business Intelligence Portal.html` as the page and pass the rest as
files with `root` set to the portal folder: `CTS_bi_bundle.js`,
`CTS_build.js`, `CTS_email.js`, `CTS_deck.js`, every `data/*.js`,
`vendor/xlsx.full.min.js`, `README.md`, `docs/*.md`, `build/*.json`.
Limits learned:

- `vendor/pptxgen.bundle.js` is refused by the artifact service, so the deck
  button on an artifact says the module is missing.
- `.py` files must be sent with `contentType: "text/plain"`. `.xlsx`,
  `.pptx` and `.zip` cannot be attached at all.
- Build can read templates on an artifact but cannot write into it; it is
  for checking only there.
- The ledger file is 4.9 MB for fifteen months, under the 16 MB page limit,
  but it will grow.
- Anyone the artifact is shared with can read every data file, including
  wages and the ledger, whatever their portal role.
- An artifact cannot move between accounts; publish a fresh one from the
  work account.

### 12.3 Local or any static host

Any web server that serves the folder works. Never a public host with real
data.

### 12.4 Hosting on Danica's PC for herself

Copy `portal` with the live data to a folder, double-click the HTML. Edge
is fine. For Build to write, the folder must be on a local or synced drive,
not a network share that blocks the File System Access API.

---

## 13. Testing

There is no automated test suite. The pattern used throughout is a harness
HTML page served from a copy of the portal and run in headless Chromium,
with results dumped into a `<pre id="out">` and read from `--dump-dom`.

```
cp -r portal /scratch/portal          # never test in the repo folder
cd /scratch/portal && python3 -m http.server 8765 &
chrome --headless=new --no-sandbox --disable-gpu --virtual-time-budget=60000 \
  --dump-dom http://localhost:8765/harness.html
```

The harness loads the data files and the bundle in the same order as the
real page, sets `localStorage["cts.bi.identity"]` to a user name, calls
`CTS.engine.init()`, then renders pages with `CTS.pages.<id>.render()` into
a div. For Build it fetches the templates as ArrayBuffers and calls
`B.runFromBuffers(buffers, { reportingMonth: "2026-09" })`, which returns
the checks and the data file contents without touching disk.

Checks that have been kept green and should stay so:

- Ledger ties to the P&L in every month and category.
- Job income plus no-job income equals P&L income (Project P&L identity).
- Derived account budget ties to the group budget.
- Every page renders without a thrown error on seed and on live data.
- The deck builds and passes `validate.py` from the pptx skill (needs
  `lxml` and `defusedxml`).
- Both scripts parse:

```
node -e "new Function(require('fs').readFileSync('CTS_bi_bundle.js','utf8'))"
node -e "new Function(require('fs').readFileSync('CTS_build.js','utf8'))"
```

On Windows, Edge's own headless mode works the same way:
`msedge --headless=new --dump-dom http://localhost:8765/harness.html`.

---

## 14. Decisions and traps learned the hard way

- **Header rows in pasted Xero reports.** `B.table` took the title row as
  the header. Fixed with a `findHead` predicate. Any new Xero paste needs
  one.
- **Pay run undated lines** all landed in the reporting month. Fixed by
  splitting runs on their title rows and spreading by weekday. If a pay run
  is pasted without its title row the split fails silently; the Build check
  "pay runs read" shows the count.
- **Money units in the checks.** Several checks mixed dollars and cents.
  Everything in the checks works in cents now; `money(cents)` formats.
- **Financial year as two digits.** The calendar's `fy` is 27 for FY27, not
  2027. Use `E.currentFY()`.
- **Derived budget produced nothing** because the calendar is built after
  the budget step. The prior year is now computed arithmetically in that
  step. Keep that in mind when reordering steps.
- **Duplicate bill noise.** Same date and reference, monthly contracts,
  marketing retainers and progress claims were all flagged. Excluded by
  keyword and by same date and reference.
- **GL layout detection.** General Ledger Detail has account heading rows
  and no cost centres; the custom report has them. Build detects by the
  header. A GL Detail paste makes every department read UNALLOCATED, which
  looks like a bug but is the data.
- **Stripe fees** have no cost centre in Xero. Build codes them to Admin and
  logs how many.
- **Cents and the Zoho export.** Deal amounts are converted to cents at
  Build like everything else; a quote of 21469000 is $214,690.
- **Table `cell` callbacks** return a class name string and receive the row,
  not the value. Returning a DOM node there throws.
- **LibreOffice cannot open the deck**, so visual checks are PowerPoint only.
- **The scratch folder in a cloud Claude session is wiped on container
  restart.** Filled templates, the review workbook and the test harnesses
  were lost that way once. Anything worth keeping goes in the repository
  (if not real data) or back to Danica as a file.
- **Seed versus live confusion.** Danica once saw "Sep as budget" and
  thought the live build was wrong; she was looking at the seed copy whose
  reporting month is August. Check the seed banner and the reporting month
  before debugging.

---

## 15. Working rules and how Danica likes to work

Set by Danica or by the repository. Carry them over.

- Australian spelling and context: financial year 1 July to 30 June, GST,
  superannuation guarantee 12%, state public holidays, Victoria by default.
- No em dashes or en dashes in repository files. Use a comma, a full stop or
  "to".
- Never commit real data. Column shapes and derived names only. The seed
  stays in `portal/data`; live data lives in `portal/_live_september` or a
  hosted copy, both outside git.
- Never open a pull request unless asked. Work on the one branch.
- No model names or identifiers in commit messages, code or comments.
- The emails and the deck must never leak the in-portal department filter.
- Data files are generated; templates are generated. Edit the generator,
  not the output.
- Talk to Danica plainly: answer first, workings underneath, say clearly
  what is a placeholder or unverified, do not pad. She is a finance head,
  not a developer; describe changes in terms of what she sees on the page
  and what she has to do, not in code. She prefers dollars against dollars
  over percentages. When she is closing a month, give the fastest practical
  path.
- When she uploads files, treat them as the source of truth over anything
  modelled. When she says a flagged item is fine (as with RJW INV-0048),
  change the check so it stops flagging that pattern.
- Deliver files she can use: a zip of the hosted folder, a workbook, a
  deck. She works in Excel, PowerPoint, Outlook and Edge.

---

## 16. Backlog

In rough priority order.

1. Rebuild September with the 2 and 16 October pay runs and the balance
   sheet and trial balance: whole month utilisation, hours per job, balance
   sheet checks.
2. Real Users and Distribution rows, the SharePoint link on Config
   Reporting, Staff and Office Dept split percentages, then Confirmed.
3. Host the folder in the SharePoint library and walk the executives
   through opening it from the synced folder.
4. Fill 10 Cash and Commitments so Cash Flow and the cash tile have data.
5. Match templates 05 OnRent Orders and 06 Qwilr Quotes to real exports so
   the confirmed work layer of the forecast is live.
6. October close with the same routine.
7. Forecast accuracy (FORECASTING step 5) after three months of archive.
8. Power Automate flow (AUTOMATION stage 2) so the outbox sends itself.
9. Nice to have: a restricted build for department heads (data set without
   the ledger), a thirteen week direct cash flow from the aged reports.

---

## 17. Starter prompt for the work agent

Open the `repo` folder in Claude Code and paste:

> This folder is the CTS Business Intelligence Portal, a static finance
> portal for Corporate Technology Services. Read CLAUDE.md, then
> portal/docs/HANDOVER.md, portal/README.md and portal/docs/MONTHLY.md
> before touching anything. The data in portal/data is seed. The real
> September 2026 data is in portal/_live_september/data and must never be
> committed. Australian spelling, no em dashes in repository files, no pull
> requests unless I ask, never put model names in commits. First job:
> [point the git remote at my work GitHub and push / host the portal with
> the live data / rebuild September with the October pay runs I will upload].
