# CLAUDE.md

Instructions for any Claude agent working in this repository. Read
`portal/docs/HANDOVER.md` before the first change.

## What this repository is

Corporate Technology Services Pty Ltd (CTS), an Australian audio-visual
company. The live project is the **CTS Business Intelligence Portal** in
`portal/`: a static, browser only finance portal (one HTML file, four
scripts, vendored SheetJS and PptxGenJS, a folder of generated data files).
The other folders (`manuals/`, `pack/`, `reporting/`, `tools/`,
`checklists/`) are earlier deliverables: payroll and monthly reporting
manuals and workbooks. They are background and rarely change.

The owner is Danica Nelson, finance head at CTS. She is not a developer.

## Hard rules

1. **Live data never goes into this folder's `portal/data/`.** It holds
   seed data only (`meta.seed: true`, fictional clients). Real exports,
   built live data, wages, ledgers and client names live in
   `portal/_live_september/` or in the separate hosted copy of the portal.
   `portal/_*` folders are ignored by git for this reason. If git is ever
   used, real data is never committed.
2. **No em dashes or en dashes** in any file, code or docs. Use a
   comma, a full stop or the word "to".
3. **Australian spelling and context** everywhere: colour, organisation,
   utilisation, financial year 1 July to 30 June, GST, superannuation.
4. **This is a local build with no remote.** Before editing the HTML or
   any of the four scripts, copy them into
   `portal/_versions/YYYY-MM-DD-<short note>/` so the change can be undone.
   The `.git` folder here is history and an optional local undo log; never
   push anywhere or open a pull request unless Danica asks. Never put a
   model name or identifier in code, comments, commit messages or notes.
5. **Generated files are not edited by hand.** Templates come from
   `portal/build/build_templates.py`; seed data from
   `portal/build/build_data.py`; live data from the Build page. Change the
   generator, regenerate, keep both.
6. **The emails and the deck ignore the department picker.** Anything that
   renders for a recipient must run inside `E.withoutView`.
7. **Money is cents (integers) inside the engine.** Amount is credit less
   debit: income positive, costs negative, gross profit is income plus cost
   of sales.
8. **Department comes from the transaction**: Cost Centres column first,
   then the bracket tag on the job number. Never from the account name.

## Layout

```
portal/CTS Business Intelligence Portal.html   loader and CSS
portal/CTS_bi_bundle.js   engine E, UI U, formatters F, pages P, GEN, router, store
portal/CTS_build.js       Build page: TEMPLATES registry, converters C.*, checks C.checks, writers B.*
portal/CTS_email.js       module M, monthly and fortnightly emails
portal/CTS_deck.js        module D, Monthly Business Review deck
portal/data/              seed data files (generated)
portal/templates/         01 to 14 Excel templates (generated)
portal/build/             generators and the JSON they read
portal/docs/              MONTHLY, HOSTING, AUTOMATION, FORECASTING, HANDOVER, COWORK
```

## How to work

- **Before editing**, read the relevant section of `portal/docs/HANDOVER.md`
  (section 3 architecture, 5 build, 7 checks, 8 deck and emails).
- **Syntax check both scripts** before handing any change over:
  `node -e "new Function(require('fs').readFileSync('portal/CTS_bi_bundle.js','utf8'))"`
  and the same for `portal/CTS_build.js`.
- **Test in a copy, never in the repo folder.** Copy `portal` to a scratch
  folder, overlay live data if needed, serve it with `python -m http.server`
  and drive it with headless Edge or Chrome (`--headless=new --dump-dom`)
  through a harness page that calls `CTS.engine.init()` and
  `CTS.pages.<id>.render()`, or `B.runFromBuffers()` for Build. HANDOVER
  section 13 has the pattern and the identities that must hold (ledger ties
  to P&L; job income plus no-job income equals P&L income; derived budget
  ties to group budget; every page renders).
- **Finishing a change** means: syntax check passed, the harness identities
  hold, a dated copy is in `portal/_versions`, and a list for Danica of the
  files to copy into the hosted portal (`portal/docs/COWORK.md` section 5).
- **Adding a page** means four places: `P.<id>` in the bundle, the PAGES and
  ROLES lists in `build/build_templates.py` (then regenerate templates),
  `data/CTS_users_data.js`, and optionally `GEN.<id>`.
- **Adding a template or column** means the registry in `CTS_build.js`, a
  converter, the generator in `build_templates.py`, and the table in
  `docs/MONTHLY.md`.
- **Adding a controller check** means an `add({...})` block in `C.checks`
  with id, area, title, status, items, cols, detail and a `guide` sentence
  written for CTS, plus a row in HANDOVER section 7. Thresholds go in `CK`.
- **UI table columns**: `fmt(value, row)` formats; `cell(row)` returns a
  class name string, not a node; `value(row)` may return a node.
- Keep each change minimal and in the style of the surrounding code: plain
  ES5 style functions, no build step, no frameworks, no new dependencies.
  Charts are hand written SVG.

## How to talk to Danica

- Answer first, then the workings. Plain English. Say what she will see on
  the page and what she has to do, not what the code does.
- State clearly what is a placeholder, unverified or an assumption. Do not
  guess with confidence. Dollars against dollars rather than percentages.
- When she is closing a month, give the fastest practical path.
- Her uploads are the source of truth over anything modelled or remembered.
  When she says a flagged item is fine, fix the check so the pattern stops
  flagging.
- Deliver things she can open: a zip of the hosted folder, an Excel
  workbook, a PowerPoint. She works in Excel, PowerPoint, Outlook and Edge.
- Do not ask a string of clarifying questions. Make a reasonable assumption,
  say what it was, and proceed.

## Two folders

The repository clone is for building and holds seed data. The team's
portal is a separate hosted copy with live data, run from a Cowork project.
`portal/docs/COWORK.md` section 5 says what to copy across after a change:
the HTML, the four scripts, `vendor` and `docs`, never `data`.

## Current state (November 2026)

September 2026 is closed and built; see HANDOVER section 10 for the numbers
and what is still placeholder (users and distribution rows, Staff and
Office Dept split percentages, cash balances, balance sheet and trial
balance, hours per job). Next: rebuild September with the October pay runs
and the balance sheet, host the folder in SharePoint, then the October
close. HANDOVER section 16 is the backlog.
