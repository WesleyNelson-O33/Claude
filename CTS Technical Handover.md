# CTS Finance and Payroll Systems - Technical Handover

Two source layers went into this: (1) direct reading of every script, spec and data file in this repository, and (2) the generated deliverables themselves (the .docx manuals and the three .xlsx workbooks). Where the two agree I have merged them. Where a build script and the text it writes into its own README disagree, I have said so rather than picking one. Where the live business process has clearly moved on since a script was written, I have flagged that explicitly rather than presenting the script as current truth.

**What is still not fully verified**: the twelve .docx manuals (about 22MB in total) are generated artefacts. I have read every JSON spec that produces them, which is the actual source of truth, but I have not opened the built Word files to confirm they match. The two large workbooks are the bigger gap. `CTS Financial Controller Pack.xlsx` (3.58MB) and `CTS Budget FY27 Final - 3 Way fixed.xlsx` (500KB) could not be opened in this session because openpyxl is not installed here; their sheet names I confirmed by reading the raw XML, but their cell contents below come from the build scripts and from the two check scripts, not from opening the files. The upstream sources are not in this repository at all: the Xero exports, the Employment Hero data, the Y drive, the budget file, the bonus leave tracker, the rate increase spreadsheet and the seven screen recordings all sit outside it. Everything said about them below comes from the manuals and the scripts that consume them.

---

## 1. What this actually is

A **documentation and tooling repository**, not an application. Nothing here runs in production and nothing here talks to Xero or Employment Hero. It builds three kinds of deliverable, and the deliverables are what people actually use:

1. **Training manuals** (twelve .docx files in `manuals/`) - built by Node scripts from JSON specs. Two series: an eight-part Payroll Processing manual and a three-part Monthly Reporting manual, each also published as a single bound volume.
2. **Working workbooks** (three .xlsx files in `reporting/`, `pack/` and `checklists/`) - built by Python scripts using openpyxl, or in one case written straight into an existing workbook's XML.
3. **Handover documents** (two .docx files) - the unanswered questions and their sources, written for a payroll handover where the outgoing officer left before the incoming one had run a pay run unsupervised.

The single design rule running through all of it: **the generated file is never the source of truth**. The spec JSON is the source for a manual; the Python script is the source for a workbook. Editing a built .docx or .xlsx by hand and then re-running the build silently discards the edit. This matters more than it sounds, because two of the three workbooks cannot currently be rebuilt at all (section 10).

**Business context** (confirmed from `reporting/data/validation.json` and the manual specs): Corporate Technology Services Pty Ltd, an Australian audio-visual and technology services business. Financial year 1 July to 30 June, reporting in AUD. Six operating departments: Onsite, Production, Video, Integration, Consulting and Admin. Payroll runs fortnightly on Employment Hero; accounting runs on Xero; banking runs through CommBiz (Commonwealth Bank). Casual staff are covered by the Live Performance Award; salaried and full-time staff are paid above award under an internal CTS Overtime Matrix.

---

## 2. Folder architecture (confirmed from the build scripts)

```
manuals/       specs/*.json          the source of every manual
               screenshots/*/        38 redacted PNGs, grouped by part
               redaction-*.json      which pixels get blurred in which frame
               *-narration-transcript.txt   machine transcripts of the recordings
               *.docx                 the built manuals (generated, do not edit)

reporting/     build_accounts.py      writes data/accounts.json
               build_pnl.py           writes CTS P&L Reporting.xlsx
               data/accounts.json     190 accounts, grouped and divisioned
               data/validation.json   Xero FY26 control totals
               source/                the Video 7 transcript

pack/          build_pack.py          writes CTS Financial Controller Pack.xlsx
               build_bridge.py        adds two tabs to the FY27 budget workbook
               loadgl.py              one-off: loads a GL export into the pack
               check.py, check2.py    ad hoc verification of the built pack
               validate.py            generic xlsx package validator

checklists/    Payroll Processing Checklist.xlsx   the fortnightly tick sheet

tools/         build_manual.js        one spec  -> one .docx
               build_master.js        several specs -> one bound .docx
               manual_lib.js          shared docx rendering
               extract_frames.py      video -> candidate frames
               redact.py              frames -> blurred screenshots
               transcribe.py          audio -> timestamped transcript
```

There is no CLAUDE.md, no README.md and no build runner. The build commands live only in the docstring at the top of each script.

**The video-to-manual pipeline**, confirmed from `tools/`:

1. `extract_frames.py` pulls two frame sets from a screen recording: `scene/` at detected scene changes, and `grid/` at a fixed interval as a fallback for slow-moving footage. It shells out to the ffmpeg binary bundled with imageio-ffmpeg, and parses ffmpeg's own stderr banner for duration and resolution because imageio-ffmpeg ships ffmpeg but not ffprobe.
2. `transcribe.py` runs a local sherpa-onnx Whisper model over 16kHz mono WAV, in 28-second chunks because Whisper reads a fixed 30-second window, timestamping each chunk from its offset in the file. Nothing is sent to a cloud service.
3. `redact.py` blurs the regions named in a per-video redaction JSON. It does not simply Gaussian-blur: it downsamples each region by a factor of twelve, resamples back with NEAREST, and only then blurs, because a Gaussian blur alone leaves large text faintly legible. An optional `crop` trims the frame afterwards, which is how the webcam column is removed from a call recording rather than blurred.
4. The redacted PNGs land in `manuals/screenshots/<part>/` and are referenced by filename from the spec.
5. `build_manual.js` or `build_master.js` renders the spec to .docx.

Frames and source video are not committed. Only the final redacted PNGs are. Re-running redaction therefore needs the original recordings, which are not in this repository.

---

## 3. Build sequence per deliverable

**A single manual**: `node tools/build_manual.js manuals/specs/part2.json`. The spec carries `title`, `subtitle`, `runningHead`, `screenshotDir`, `output`, an optional `toc` flag, and a `blocks` array. `manual_lib.js` renders nine block types: `h1`, `h2`, `step`, `p`, `bullet`, `shot`, `callout`, `table`, `questions`, `sourced` and `next`. Figures are scaled to 600px wide against the 722/1280 aspect ratio of the extracted frames and captioned "Figure N: ...".

**A bound manual**: `node tools/build_master.js manuals/specs/master.json`. The master spec lists the part specs in order and renders each one nested, so every heading drops a level and the contents page reads as a hierarchy. A single shared figure counter runs across the whole book. Because the parts are read from the same specs as the standalone guides, the bound volume can never drift from them.

**The P&L workbook**: `python reporting/build_accounts.py` then `python reporting/build_pnl.py`. Fully reproducible from the repository; `data/accounts.json` and `data/validation.json` are both committed.

**The Controller Pack**: `python pack/build_pack.py`. **Not currently reproducible.** `main()` reads its account master from `/tmp/claude-0/.../scratchpad/lists.json`, a scratchpad path that no longer exists. Without that file the script will not run. See section 10.

**The budget bridge**: `python pack/build_bridge.py`. **Not currently reproducible either.** It reads the original budget workbook from a `/root/.claude/uploads/` path and a row skeleton from the same scratchpad.

**Node dependencies are undeclared.** `.gitignore` excludes `node_modules/`, `package.json` and `package-lock.json`, so the `docx` package the manual builders require is nowhere recorded. Anyone rebuilding a manual has to know to `npm install docx` first, and there is no pinned version.

---

## 4. Access, people and approvals

There is no software access control anywhere in this repository. The control model is human, and it is documented in the manuals rather than enforced by anything.

**Systems a payroll officer needs** (confirmed from `part1.json` and the handover spec): the CTS Accounts mailbox in Outlook, the Y drive (`CTS NAS Accounts (Y:) Backup`, via OneDrive), Employment Hero Payroll, Employment Hero HR, the bonus leave tracker, the rate increase spreadsheet, and CommBiz. The handover document records CommBiz as the one piece of access that was outstanding at the time it was written.

**Named people and what they do**:

| Person | Role in the process |
|---|---|
| Graham | Approves the pay run inside Employment Hero. Finalisation cannot complete without a second approver. |
| Duncan | Approves the payment in CommBiz, and approves the collated payroll notes. Also the person who asks for old charts to be reproduced. Advises where non-standard legal spend is coded. |
| Danica | Copied on the payroll email each fortnight. Body shading in the Controller Pack is matched to a layout marked up by Danica. |
| Jordan | The person to ask whether missing consulting revenue was lost or simply deferred to a later month. |
| The outgoing payroll officer | The source of all seven recordings. |

**The single biggest documented gap**: neither approver has a documented backup. `handover.json` names three questions as the ones not to leave the handover meeting without, and two of them are this: who approves the pay run and the payment if Graham or Duncan is unavailable, and who the escalation is if something goes wrong on pay day. The third is who to contact at Employment Hero when something breaks mid pay run. As committed, those questions are recorded as unanswered.

---

## 5. Data sources

### 5.1 Upstream systems

| Source | What comes out of it | Used by |
|---|---|---|
| Xero, Profit and Loss report | Monthly P&L by account, year to date | P&L workbook Actuals sheet; the P&L Analysis file; the Controller Pack's PL_Check |
| Xero, saved custom report "Account transactions for P&L analysis" | One row per GL transaction, 13 columns | Controller Pack GL_Paste; the P&L Analysis file's GL transactions tab |
| Xero, repeating bills | The Employsure retainer counter | Monthly reporting, legal fees split |
| Employment Hero Payroll | Pay runs, timesheets, rates, leave balances, the ABA file, five fortnightly reports | The whole payroll cycle |
| Employment Hero HR | Leave without pay requests, leave balances for cross-checking | Parts 3c and 7 |
| The utilisation report (kept as a separate saved file because the full report is slow) | Hours per department split chargeable and non-chargeable, plus a Days tab of working days by state | Monthly reporting Part 2 |
| Budgets and Forecast FY27 | Budget by month by department, and the overhead split percentages | Monthly reporting Part 1; the budget bridge |
| Bonus leave tracker | Service length, eligibility dates, one row per month per employee | Payroll Part 5 |
| Rate increase spreadsheet | Which rates changed this period | Payroll Part 5 |
| CommBiz | Payment import and approval | Payroll Part 6 |
| Live Performance Award | Casual entitlements, clauses 62 and 63 | Payroll Part 3a |
| CTS Overtime Matrix (HR Resource Library on SharePoint) | Overtime multipliers for salaried staff | Payroll Part 3b |

None of these are stored in this repository. The repository holds the tools that consume them and the manuals that describe them.

### 5.2 The chart of accounts (confirmed, `reporting/data/accounts.json`)

190 accounts, read from the Xero chart of accounts on 3 September 2026 so the names match Xero exactly and an export pastes straight in. Split by group:

| Group | Accounts |
|---|---|
| Income | 35 |
| Cost of Sales | 66 |
| Other Income | 9 |
| Expenses | 80 |

**Division is derived from the account name suffix**, by a regex anchored to the end of the name that accepts either `" - CODE"` or `"-CODE"`: ONS is Onsite, PRD is Production, VID is Video, CONS is Consulting, plus INTEGRATION and ADMIN. 88 accounts carry a division; the other 102 are shared overheads and fall to "Unallocated". Each of the five main divisions carries exactly 15 accounts, and Admin carries 13, which is worth knowing as a quick sanity check.

Note the two spelling inconsistencies that are real and handled deliberately: `Credit Card Collected-INTEGRATION` and `Direct Superannuation-INTEGRATION` have no spaces around the hyphen, which is exactly why the suffix regex allows `\s*-\s*`.

### 5.3 The GL paste - the core mechanic of the Controller Pack

This is the one thing pasted into the pack each month, and everything else follows from it.

**Expected columns**, in order, matching the Xero export (`GL_HEADERS` in `build_pack.py`): Index, Account Code, Account Name, Source, Date, Contact, Debit, Credit, Job Numbers, Invoice Number, Reference, Description, Cost Centres.

**The Cleanup sheet** adds six derived fields the export does not carry, one row per GL line, all by formula:

| Field | Derivation |
|---|---|
| Amount | `Credit - Debit`. Income is therefore **positive** and costs **negative**. |
| Month | `DATE(YEAR(date), MONTH(date), 1)`, so every line collapses to a month start |
| Department | `INDEX/MATCH` on the Cost Centres column against the department tag table. If Cost Centres is blank, it falls back to the bracket tag inside the Job Numbers column, extracted with `MID(..., FIND("["), FIND("]") - FIND("[") - 1)`. Unmatched lines read "UNALLOCATED". |
| Category | `INDEX/MATCH` on Account Code against the account master. Unmatched reads "UNKNOWN". |
| Subcategory | Same lookup, subcategory column |
| Contact | `TRIM` of the Contact column, for the client charts |

**The department tag table** accepts twelve tags mapping to six departments: ONSITE/ONS to ONSITE, PRODUCTION/PRD to PRODUCTION, VIDEO/VID to VIDEO, INTEGRATION/INT to INTEGRATION, CONSULTING/CONS to CONSULTING, and CTS/ADMIN to ADMIN. Note that `CTS` maps to ADMIN, which is how the Stripe fee rows described in Monthly Reporting Part 1 end up in the right place.

**Why per transaction and not per account** (the pack's own README says this outright): three quarters of the P&L sits in accounts with no department in the name, Contract Support Staff and Equipment Hires among them. Only the transaction knows which department earned or spent it.

**Setup reports data quality as calculated status lines**: GL lines pasted, GL capacity, lines with no department tag, lines with an unknown account, earliest and latest GL date. On the August 2026 export all 3,771 lines were tagged, per the README.

### 5.4 Two different department mechanisms, and they do not agree

This is worth being explicit about because the two workbooks look similar and are not.

- **`CTS P&L Reporting.xlsx`** derives division from the **account name suffix**. An account with no suffix is Unallocated, permanently. 102 of 190 accounts fall there. Its By Division sheet is therefore a partial view.
- **`CTS Financial Controller Pack.xlsx`** derives department from the **transaction's cost centre**, falling back to a bracket tag on the job number. Nearly every line gets a department.

The pack is the better method and the pack's README says why. The P&L workbook was built first, is simpler, and was designed around a straight paste of the Xero P&L export, which carries no transaction detail at all. They are not two versions of the same thing and reconciling them line for line will not work.

### 5.5 Overhead allocation - the three split bases

Confirmed from Monthly Reporting Part 1 and from `build_bridge.py`, which reads the percentages live off the budget workbook.

Anything sitting in Admin is pushed out across the departments using preset percentages pulled in with XLOOKUP. There is not one split, there are three, and each overhead row is labelled with the one that applies to it:

| Basis | What it means |
|---|---|
| 3 Way | Split roughly evenly three ways, near enough a third each |
| Staff | Split by headcount, so it follows where the people actually are |
| Office Dept | Weighted heavily to one department, for costs that mostly belong there |

The percentages live in the budget file, set against three groups: production and video, onsite, and consulting and integration. `build_bridge.py` reads them from `'FY27 Budget - detailed'` rows 17, 19 and 21 for the 3 Way, Staff and Office Dept bases respectively, and then splits the grouped percentages into six departments using two sub-splits it holds as editable inputs: production takes 0.79 of PRD/VID and integration takes 0.72 of CONS/INT, with video and consulting taking the remainder. Those two defaults are the business's own FY23 actuals off the PRD and CONS tabs.

**Two standing warnings from the manual, both of them live risks**:

- The staff split follows headcount, and consulting is nearly empty. The day someone is hired into consulting, that percentage has to be changed by hand. Nothing prompts it, and until it is changed consulting carries almost no overhead.
- The split must not change mid-year. Changing it in December means the six months already reported were built on a different split and no longer compare. If it is revisited, it takes effect at a new financial year.
- Whatever is changed in the actuals split has to be changed in the budget file as well, or the two will never reconcile.

### 5.6 Reference data maintained entirely by hand

Each of these has no system behind it and nothing prompts you to update it:

- Public holidays, one row per holiday per state, on the pack's Lists sheet. 120 rows provided, with a state dropdown over VIC/NSW/QLD/WA/SA/TAS/ACT/NT. The working-days grid subtracts them per state per month with COUNTIFS rather than NETWORKDAYS, because NETWORKDAYS cannot take a filtered holiday range.
- The client display names on the top clients chart, which are stock exchange listing names rather than trading names and have to be matched by hand in the yellow cells.
- The client industry and engagement type mapping, which lives in the previous financial year's June file on a client tab and is already partly out of date.
- The bonus leave tracker, including extending it as people pass each service milestone.
- The Employsure legal retainer at $800 a month, entered cumulatively (so $1,600 on the two-month year to date tab, $2,400 on the quarter tab).

---

## 6. Calculation logic confirmed in the code

This section is read straight from the build scripts, so treat it as ground truth about the workbooks rather than inference.

### 6.1 Sign conventions - there are two, and they differ

- **`CTS P&L Reporting.xlsx`**: figures are pasted as Xero presents them in the P&L report. Gross Profit is `Total Income - Total Cost of Sales`. Net Profit is `Gross Profit + Total Other Income - Total Expenses`.
- **`CTS Financial Controller Pack.xlsx`**: Amount is `Credit - Debit`, so **costs are negative**. Gross profit is therefore Income **plus** Cost of Sales, not minus. The pack's README calls this out as the first thing that will catch you out, and the Charts sheet relies on it: net profit is computed as a straight sum of all five category totals.

Move a formula from one workbook to the other and it will be wrong by the sign of every cost line.

### 6.2 CTS P&L Reporting.xlsx

Ten sheets: Read Me, Summary, Actuals, Budget, This Year, Last Year, Year on Year, By Division, Budget P&L, Budget vs Actual.

The design rule, stated in the build script's own docstring: **the only sheets anyone types into are Actuals and Budget**. Everything else is formulas pointed at those two, so appending a month updates the whole pack at once.

- **Actuals and Budget**: 190 accounts down, 24 months across (FY26 in columns D to O, FY27 in P to AA), yellow input cells, frozen at D4, autofilter on the header.
- **This Year / Last Year / Budget P&L**: full P&L, accounts down, twelve months across, FY total in column N. Every cell is `SUMIFS` against the matching data sheet keyed on the account name in column A, which means an account renamed in one place and not the other silently returns zero rather than erroring.
- **Gross Profit % and Net Profit %** are computed per column, wrapped in IFERROR against a zero-income month.
- **Year on Year and Budget vs Actual**: row-for-row comparisons that work because all three P&L sheets share a layout, so rows align exactly. Variance % is left blank on percentage rows, since the difference between two margins is already in the variance column.
- **Summary**: monthly, quarterly and year on year on one page. The quarterly margin rows are rebuilt from the quarter's own numerator and denominator rather than summed, with the explicit comment that a margin is not the sum of three margins.
- **Reconciliation Check**: hardcodes the FY26 control totals from `validation.json` (Total Income 6,113,186.03; Cost of Sales 3,497,093.83; Gross Profit 2,616,092.20; Other Income 13,413.68; Expenses 2,380,879.16; Net Profit 248,626.72) in blue type, by convention, against what the workbook itself computes. Once FY26 is pasted in, the difference column should read zero.

### 6.3 CTS Financial Controller Pack.xlsx

Fifteen sheets, ordered README, Setup, Summary, Charts, Clients, Engine, PL_Check, P&L FY27, Actual FY26, Budget FY27, Budget FY26, GL_Paste, Utilisation, Lists, Cleanup. Tabs you type into are coloured yellow (Setup, GL_Paste, PL_Check, both Budgets, Utilisation, Clients); everything else is navy.

**Setup is the single control panel.** Cell B5 is the reporting month, picked from a dropdown bound to the Months named range. The financial year, the period number in the year, the prior year and the comparative month last year are all derived from it by formula. Change B5 and the entire pack repoints.

**The Engine** is one row per department per subcategory, six departments across every subcategory in category order, with 24 month columns (FY26 then FY27) plus a total. Each cell is a three-condition `SUMIFS` over Cleanup on department, subcategory and month. Everything downstream reads the Engine; nothing downstream reads Cleanup directly except PL_Check and Clients.

**One row skeleton is shared** by P&L FY27, both budget sheets, Actual FY26 and Summary, so a paste into any of them lines up with the rest. The skeleton is department, then category, then indented subcategory, with grouped outline levels so the subcategory rows collapse. `summaryBelow` is set false, so the totals sit above their detail rather than below.

**The P&L is half actual and half forecast, by design.** Each month cell reads `IF(month <= Setup!B5, Engine value, Budget FY27 value)`. Months up to the reporting month come from the GL; months after it come from the budget, so you always see a full twelve and the shortfall ahead.

**Year to date** is a `SUMIF` across the header row on `"<="&Setup!$B$5` rather than a fixed column range, so it moves with the reporting month automatically. The quarter is a `SUMIFS` bounded below by `FLOOR(MONTH-1,3)+1`, the first month of the quarter the reporting month falls in.

**PL_Check** is the control that proves the GL paste agrees with Xero. You type the five Xero P&L category totals into the yellow cells and each line reads OK if the absolute difference is under $1, CHECK otherwise. `loadgl.py` left the August 2026 figures in as a worked example (Income 603,466.54; Cost of Sales -335,516.00; Expenses -208,889.82; Other Income 861.35; Other Expenses 0.00).

### 6.4 Risk flags

Four states, evaluated in this order, off thresholds that live on Setup:

1. `"No Activity"` if both actual and budget are zero
2. `"Low"` if the dollar variance is below the materiality floor (default $5,000) - **this is checked before the percentage test**, so a 90% variance on a $200 line is still Low
3. `"High"` if the dollar variance is at or above $50,000 **or** the percentage variance is at or above 25%
4. `"Medium"` if the percentage variance is at or above 10%, otherwise `"Low"`

### 6.5 Utilisation and FTE

Two different definitions exist in the same business, and they are not the same number.

**In the Controller Pack**: you type chargeable, non-chargeable and leave hours per department per month. Utilisation is `Chargeable / (Chargeable + Non-chargeable)`, with leave excluded from both sides. FTE is `(Chargeable + Non-chargeable) / working hours in the month`, where working hours is working days for the selected state multiplied by the hours-per-day on Setup (default 8). The state is a single cell, defaulting to VIC, which means the whole utilisation sheet is computed on one state's calendar.

**In the monthly reporting graphs file** (Part 2): FTE is `(department hours + Other CC hours) / total working hours in a rolling twelve months`, where the denominator comes from the Days tab of the utilisation report and should land on either 2,080 or 2,088 hours. The manual is explicit that this must be checked rather than assumed, because the Days tab is also split by state.

Also note that the Charts sheet's utilisation series does not use the Utilisation sheet's own percentage rows. It recomputes utilisation as `Chargeable / SUMIFS(..., "<>Leave")`, which is the same arithmetic reached a different way.

**The profitability per employee figure is explicitly flagged as weak** in Part 2: it is departmental profit spread across a calculated headcount, it is a recent idea, and the person who built it has said openly that a better method would be welcome. Do not present it as more precise than it is.

### 6.6 Clients and charts

The Clients sheet is thirty typed client names in yellow cells; everything to the right calculates. For each name it sums Cleanup on Contact and Category "Income" over the month, the prior month, the quarter, the prior quarter, the year to date, and the same year to date a year earlier. Rank is `RANK` over the year to date column. Below it, a second block of thirty rows mirrors the same names and splits their year to date revenue across the six departments.

The name in column B has to match the Xero contact string exactly. This is the same fragility the monthly reporting manual warns about from the other end: every Commonwealth Bank variant has to become one spelling or the client appears several times and is understated in each.

The Charts sheet holds eight chart blocks, each a labelled table on the left with its chart beside it, stacked down the page. Every series is a formula off the P&L or the Clients sheet, so nothing is copied in from another workbook. The exception is days sales outstanding: trade receivables is not in a P&L export, so the closing balance is typed each month into a yellow column and DSO calculates as `receivables / revenue * days in month`.

### 6.7 The budget bridge

`build_bridge.py` adds two tabs to the FY27 budget workbook so it can feed the Controller Pack, and it writes them straight into the workbook's XML rather than going through openpyxl. The reason is stated in the docstring: the budget workbook carries pivot tables, charts, threaded comments and external links, and openpyxl would drop them.

- **Pack Map** is the mapping engine: 59 rows, one per budget line, each carrying the split basis read live off the budget's own matrix, a weight per pack department, and the pack category and subcategory it maps to. A sign column flips costs negative to match the pack's convention. Four superannuation lines are netted off their sibling lines rather than taken gross.
- **Pack Budget FY27** is the pack's own 227-row Budget FY27 skeleton, filled by `SUMPRODUCT` off Pack Map. It is loaded by copying B7:M232 and pasting values into cell B7 of Budget FY27 in the pack. It is a paste, not a link.

The script also rewires five overhead lines (rows 26, 28, 49, 50 and 53 of the CTS Overheads sheet) from the Staff or Office Dept basis onto the 3 Way basis at 34/33/33, and bands those five rows bright orange so the change is visible. It drops `calcChain.xml` and sets `fullCalcOnLoad`, so Excel rebuilds the formula order on open rather than being handed a stale cache.

---

## 7. The deliverables

**Payroll Processing manual**, eight parts, published standalone and bound (`Payroll Processing Manual - Complete.docx`, 11.9MB):

| Part | Covers | Source |
|---|---|---|
| 1 | Payroll Prep: collecting notes, the timesheet template, timesheet approval gate | Video 1 |
| 2 | Creating a New Pay Run | Video 2 |
| 3a | Reconciling Casual Timesheets (Live Performance Award) | Video 3 |
| 3b | Reconciling Salaried and Full-Time Timesheets (CTS Overtime Matrix), with a worked correction | Videos 3 and 4 |
| 3c | Employees Needing Manual Handling | Videos 3 and 5 |
| 5 | Reconciling and Balancing the Pay Run | Video 5 |
| 6 | Finalisation and Bank Upload | Video 6 |
| 7 | Terminations and Out-of-Cycle Pay Runs | Video 7 |

**There is deliberately no Part 4.** Video 4 worked through a correction to a salaried employee, so its material sits at the end of Part 3b rather than in a guide of its own. The part numbers otherwise follow the source recordings so a step can be traced back to the video it came from. This was a later decision: git history shows Part 4 built as a standalone guide first and folded into 3b afterwards.

**Monthly Reporting manual**, three parts, published standalone and bound (`Monthly Reporting Manual - Complete.docx`, 2.0MB): Part 1 the P&L Analysis file, Part 2 the Graphs file, Part 3 the Management Report. All three were recorded live in Video 7.

**Also built**: `Payroll Step by Step - Listening Version.docx` (a spoken-form walkthrough), `Payroll Handover - Questions and Scenarios.docx`, `Payroll Handover - Meeting Prep.docx` (which cites the recording, timestamp and quoted words behind each handover question), and `Payroll Processing Checklist.xlsx` (a 60-plus step tick sheet mirroring the manual, with a Done column that turns the row green and an N/A option).

---

## 8. The two operating cycles

### 8.1 The fortnightly payroll cycle

Start on the **Friday before** the Monday you process. Timesheets are due from staff by **10am Monday**. The pay date is always the **Tuesday** of the processing week.

Order of work: collect and collate notes (Part 1) -> copy the timesheet template into the fortnight folder and populate it -> confirm nothing sits under Submitted in Timesheet Approval -> create the pay run on the `CTS - Fortnight` schedule, importing **Timesheets for this pay period** and never All unpaid timesheets -> reconcile alphabetically, picking up 3a, 3b or 3c per employee by employment type -> apply notes, per diem and bonus leave (Part 5) -> balance -> finalise, upload, pay (Part 6).

**The rules most likely to cost money if they are missed**, all confirmed in the specs:

- Casual overtime is 1.75 on hours above 38 **in a week**, not a fortnight (clause 63). Employment Hero catches this only sometimes.
- Daily overtime above 8 hours and weekly overtime above 38 must not both be paid for the same hours. Pay one, delete the other.
- More than 4 hours entitles a 30 minute meal break (clause 62.2). More than 5 continuous hours with no suitable break is paid at 200% of the minimum hourly rate (clause 62.1c). If no break is recorded, Employment Hero **silently deducts 30 minutes**. The fix is to set the work type to Working During Meal Break on the timesheet, not to add half an hour to the pay run.
- A salaried fortnight is exactly 80 hours unless overtime was pre-approved. A back-end rule left over from the ADP migration **caps captured hours at eight per day and also rounds short days up to eight**, so the pay run total and the real timesheet total can differ when nothing is wrong. You cannot reconcile from the pay run alone.
- Salaried staff are not eligible for the award break condition that triggers when a shift starts less than ten hours after the previous one ended. Those rows have to be returned to Permanent Ordinary Hours.
- **Changing a category wipes that row's Notes field, with no undo.** Copy the note out, change the category, paste it back.
- Time in lieu taken goes in the In Lieu Taken column; ordinary hours are left to the formula, which takes 80 and subtracts columns F to S.
- Salaried hourly rate = annual salary / 1.12 / 52 / 40. The 1.12 removes superannuation, because the spreadsheet holds rates excluding super and excluding casual loading.
- Per diem is $75 per night at the CTS rate, entered as units multiplied by nights. The award rate is substantially lower and is not used.
- Bonus leave accrues in the **last pay run of the month only**, always a month ahead, and only where the employee's letter has been issued and signed.
- Balancing is two checksums: total hours and total dollar value. Hours almost always balance first. Hours right and dollars wrong means a rate is wrong, and it is usually an overtime multiplier on a stale base rate.
- Finalisation needs a second approver (Graham). **Once super has been paid against a pay run, Employment Hero will not unlock it.**
- Transfer enough to cover net payroll **plus super** and still leave roughly a $20,000 buffer. Gross earnings is not what leaves the account.
- Pay slips are published the **following day**, after Duncan has approved the payment, because publishing notifies employees immediately.

**Known unresolved fault**: the timesheet spreadsheet's total hours column deliberately excludes the leave without pay column, so total hours are wrong for anyone with leave without pay in the fortnight. This was identified during training and left unresolved. Until it is fixed, verify those totals by hand.

### 8.2 The monthly reporting cycle

Three stages, strictly in order, because each feeds the next: the P&L Analysis file, then the Graphs file, then the Management Report. Going back to correct an earlier stage means redoing everything after it.

The rules that matter most:

- **Re-paste every prior month, not just the new one.** Earlier months can move; an Aware Super adjustment retrospectively changed a July figure, and anyone working off the old paste would never have seen it.
- Run the GL transactions custom report from 1 July, sorted by account code, and **convert the values column to number before pasting** or the totals will not add up.
- Stripe fee rows arrive with no cost centre because Xero splits a card payment three ways. Set them all to Admin and CTS. Do not try to allocate them properly; the amounts are cents.
- Only revenue rows (the 4 codes) need contacts filled in, and each client must be spelled one consistent way.
- Tabs 2 and 3 reconcile the data dump back to the P&L. A FALSE is usually rounding; check the cell reference first, then round.
- The tabs that **exclude** other income are the reporting ones and are highlighted differently. Every graph should reference those.
- Hard code last month's budget figures **before** re-pointing the budget links, or you will lose them. Then move each reference along one column (main tab H to I; Onsite F to G; Consulting G to H; Production H to I), leaving the grey total cells alone. The check cell should read zero.
- **Hard code a quarter as soon as it is finished**, before pasting the next one, and leave a note on the tab saying so.
- The running total chart is cumulative. Entering the month on its own makes the line go up and down.
- Filter the utilisation data to a **rolling twelve months**, not the month and not the year to date.

**What normal looks like** (from Part 3, and this is the part hardest to pick up from a spreadsheet):

| Department | Normal |
|---|---|
| Support and managed services | Always profitable. Gross margin around 45%, net profit above 10%. |
| Production | Very seasonal. Gross margin 55% or better; in the 40s is worth investigating. |
| Integration | Gross margin at least 15%. Net profit skewed by how few people work in it. |
| Consulting | Frequently below budget. A loss in most months is normal, not a red flag. |

Production calendar: August and September are results season; October to December is the peak, driven by AGM season; December and January are quiet; February picks up; March is quiet; May and June are the peak for integration and consulting as the financial year closes.

---

## 9. Version history

The repository runs from 26 August 2026 to 17 September 2026, in three distinct phases.

**26 August, the payroll manuals.** Started with the video pipeline (ffmpeg frame extraction, then local speech to text), then Part 1, then the builder was made spec-driven so later parts were data rather than code. Parts 2, 4 and 6 followed, then master binding, then Parts 3a to 3c and 5. The handover questions document was added the same day and then revised four times in a row: resolved access questions trimmed, answered questions recorded, the new starter question reframed, and the whole thing rewritten around the fact that the first unsupervised pay run would happen with the outgoing officer already on leave.

**3 to 6 September, consolidation.** Every em dash replaced with ordinary punctuation (a house style rule that still applies). The contents page made navigable and the part count corrected. Part 4 folded into Part 3b and a structure map added explaining why there is no Part 4. The fortnightly checklist workbook and the P&L reporting workbook added. Part 7 added from the seventh recording. The Monthly Reporting manual added on 4 September, and a listening version of the payroll walkthrough on 6 September.

**17 September, the Controller Pack.** Built, then rebuilt twice in one day: first the Summary reorganised around rolling months, then the whole pack rebuilt to layouts marked up on the yellow tabs. Two-year GL, per-state holidays, clients and charts followed, then shading matched to Danica's and the chart tables moved back beside their charts. The FY27 budget work came last: five overhead lines switched to the 3 Way split, the paste cells made white instead of yellow, the five 3 Way rows highlighted bright orange, then the Pack Map and Pack Budget FY27 tabs added and finally a corrupt budget workbook fixed and a package validator added.

That last one is worth knowing about. `pack/validate.py` checks an xlsx the way Excel does before it offers to repair: every XML part parses, every part has a content type, every relationship target resolves, sheet names and ids are unique, style counts match their children, every `s=` index is in range, and worksheet child elements follow schema order. If you hand-edit workbook XML again, run it.

---

## 10. Risks and open questions

1. **Two of the three workbooks cannot be rebuilt from this repository.** `build_pack.py` reads its account master from a scratchpad `lists.json` that is gone, and `build_bridge.py` reads both the source budget workbook and a row skeleton from paths that no longer exist. The P&L workbook is the only one that is fully reproducible. If either of the other two needs a change, the first job is reconstructing those inputs, not editing the script. Committing `lists.json` and `skeleton.json` would fix this permanently and they are small.
2. **The pack's README understates its own capacity by more than half.** `GL_ROWS = 45000`, but the README line says "GL_Paste holds 20,000 lines, about a full year at your volume", and `loadgl.py` reads only to row 20001. The formulas are built to 45,000 so the capacity is real; the documentation is stale. Anyone pasting a second year of GL will be told by the README that they cannot.
3. **The pack's Summary sheet computes the same thing twice.** Column F is labelled "YTD vs Rolling Forecast" and column G "YTD vs Budget", but both are `SUMIF` over Budget FY27 for the year to date. Either a rolling forecast sheet was intended and never built, or the label is wrong. Confirm which before anyone reads F as a forecast comparison.
4. **The Node build has no declared dependencies.** `package.json` is gitignored, so the `docx` version that produced the committed manuals is unrecorded. A different major version could change the output silently. Committing a `package.json` with a pinned version is a ten minute job and removes a real reproducibility risk.
5. **The frames the redaction specs refer to are not committed**, only the redacted output. Re-running redaction needs the original recordings. If those recordings are ever lost, the screenshots cannot be regenerated or re-cropped.
6. **Two department models coexist** (section 5.4), and the weaker one is the one in the simpler workbook. Be clear which workbook someone is reading before answering a question about departmental revenue.
7. **The utilisation sheet in the pack runs on one state's calendar**, defaulting to VIC, while the monthly reporting FTE calculation uses a rolling twelve months off a Days tab that is split by state. Two utilisation numbers for the same business are therefore expected to differ, and neither is wrong.
8. **The overhead split has a live trigger nobody watches**: hire one person into consulting and the staff-basis percentages become wrong in both the actuals and the budget, with nothing to prompt the change.
9. **The payroll approval chain has no documented backup**, and the handover document records this as unanswered. Two people, Graham and Duncan, each sit on a single point in the chain that cannot proceed without them.
10. **The leave without pay total hours fault in the timesheet spreadsheet is open**, known, and has to be worked around by hand every fortnight it applies.
11. **Several monthly reporting steps are still manual and are named as the first automation candidates**: the legal fees split, the budget column links, the labour hours, the client name matching and the sales forecast. The manual says plainly that these are where the mistakes happen.
12. **House style**: no em dashes anywhere in this repository. There is a commit whose only purpose was removing them. Keep to it.

---

## 11. Suggested next steps

- If the Controller Pack needs any change at all, reconstruct and commit `lists.json` first. Everything else is blocked behind it, and the reconstruction gets harder the longer it is left.
- Fix the three documentation defects in one pass: the 20,000 line capacity note, the Summary column F label, and a committed `package.json`. None of them changes a number, all of them will mislead someone.
- If you want the built workbooks read cell by cell rather than inferred from their build scripts, that needs openpyxl installed. Say which sheets matter and it is a short job.
- The handover document is only useful filled in. Its own advice is that three answers matter more than the rest: who approves when Graham or Duncan is unavailable, who to call at Employment Hero mid pay run, and who the pay day escalation is. If those are still blank, they are the highest value thing in this repository to close.
- The P&L workbook holds FY26 and FY27 only. When FY28 starts it needs rolling forward, and the script that does it is the one to edit, not the workbook.
