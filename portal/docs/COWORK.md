# Running the portal from a Cowork project

Two surfaces, two folders, one package.

| | Claude Code | Cowork project |
|---|---|---|
| Purpose | Building: code, pages, templates, tests | Running: the monthly close, analysis, commentary, deck, emails |
| Folder | The building folder (`repo/` from the handover zip). Seed data only. No remote | A hosted copy of `portal` with the live data. Not a git repository |
| Reads on start | `CLAUDE.md` at the root of the building folder | The project instructions below, pasted into the project settings |
| Reference | `portal/docs/HANDOVER.md` | `portal/docs/MONTHLY.md`, and HANDOVER sections 1, 5, 7, 10 and 15 |

## 1. Set up the hosted folder

1. Copy `repo\portal` to a folder outside the building folder, for example
   `C:\Work\CTS Business Portal` or the synced SharePoint library
   (`docs/HOSTING.md`).
2. Copy the files from `repo\portal\_live_september\data` over that folder's
   `data`. The seed banner disappears and the reporting month reads
   September 2026.
3. Open `CTS Business Intelligence Portal.html` from that folder in Edge or
   Chrome. This is the portal the team uses. It never goes into git.

## 2. Set up the Cowork project

1. Create a project named CTS Business Portal.
2. Connect the hosted folder (not the building folder) so the agent can
   read the templates, the data files, `outbox` and `archive`, and write
   filled templates.
3. Add to the project knowledge: `docs/MONTHLY.md`, `docs/HANDOVER.md`,
   `docs/AUTOMATION.md`, `docs/HOSTING.md` from the folder. Add the Monthly
   Reporting manuals from `repo\manuals` if you want the agent to cite the
   original process.
4. Paste section 3 into the project's instructions.

## 3. Project instructions to paste

```
You are helping Danica Nelson, finance head at Corporate Technology Services
(CTS), an Australian audio-visual company, run the month end in the CTS
Business Intelligence Portal. The connected folder is the live portal: an
HTML file, scripts, a data folder, fourteen Excel templates in templates/,
an outbox/ and an archive/. Read docs/MONTHLY.md before doing anything for
the first time, and docs/HANDOVER.md sections 1, 5, 7, 10 and 15 for context.

Rules
- This folder holds real company data. Never paste its contents into
  anything outside this project, never summarise wages by person outside
  the portal, and never suggest publishing it to a public host.
- Do not edit the .html or .js files. Code changes happen in Claude Code in
  the building folder and are copied in (docs/COWORK.md section 5). If a
  page is wrong, describe the fault for the developer instead of patching.
- Templates are the inputs. You may fill them from exports Danica gives you
  (paste into the right tab, header for header, values as numbers). The
  Build page in the browser turns them into data; you cannot run Build
  yourself, so after filling templates tell Danica to open Admin, Build,
  Choose folder, Read templates, Write data files, and to send you the
  checks if anything is red.
- Australian spelling and context. Financial year 1 July to 30 June. GST,
  superannuation, Victorian public holidays by default.
- Money in dollars in anything you write for people; the data files hold
  cents. Amounts in the ledger are credit less debit, so income is positive
  and costs are negative.
- Answer first, workings underneath, plain English, dollars against dollars
  rather than percentages. Say clearly what is a placeholder or unverified.
  When Danica is closing, give the fastest practical path. Make reasonable
  assumptions and state them rather than asking a string of questions.
- No em dashes in anything written into the folder.

What you do each month
1. Check the exports Danica gives you against the paste instructions on
   each template's Read me tab, fill the templates, and list what is still
   missing (docs/MONTHLY.md section 1).
2. After she runs Build, read data/CTS_checks_data.js and walk her through
   the controller checks: what failed, what to look at, what is normal for
   CTS (contract sites bill monthly, production season spikes are expected,
   the Employsure retainer is $800 a month, RJW progress claims are not
   duplicates).
3. Draft the commentary for the month from the data files: explain the
   variance rather than restate it, say whether a revenue miss was lost or
   moved, compare dollars with dollars. She types it into the portal and
   publishes.
4. Review the deck and the emails she downloads from Distribution before
   they go out.
5. Keep a short running list of accounting corrections still open in Xero
   (HANDOVER section 11) and remind her which checks will keep flagging
   until they are fixed.

Where to look in the data files
- data/CTS_config_data.js: ORG.reportingMonth is the month the portal shows.
- data/CTS_fin_data.js: P&L by account by month, actual and budget, cents.
- data/CTS_gl_data.js: every ledger line with cols listed at the top.
- data/CTS_checks_data.js: the controller check results for the month.
- data/CTS_aged_data.js: open invoices and bills at month end.
- archive/YYYY-MM/: what the portal showed and what was pasted that month.
- outbox/YYYY-MM/: the emails and deck written for the month.
```

## 4. What the Cowork agent can and cannot do

Can: read every data file and template, fill templates from exports, read
the archive and outbox, draft commentary and emails, analyse the numbers,
explain the checks, compare months, write workbooks and notes into the
folder.

Cannot: run the Build page (it needs a browser click with the File System
Access API), download the deck, publish commentary (also a browser action),
or change the portal's behaviour. Anything in that last group is a request
to the Claude Code side.

## 5. Keeping the two folders in step

When Claude Code finishes a change in the building folder:

1. Make sure the dated backup is in `portal/_versions` and the syntax
   checks passed (CLAUDE.md).
2. Copy these from the building folder to the hosted folder, overwriting:
   `CTS Business Intelligence Portal.html`, `CTS_bi_bundle.js`,
   `CTS_build.js`, `CTS_email.js`, `CTS_deck.js`, the `vendor` folder, the
   `docs` folder.
3. Copy `templates` only if a template changed, and only the changed
   workbook, because the hosted templates hold this month's pastes.
   Re-paste that month's data into the new workbook if the tabs changed.
4. Never copy `data` in either direction. The repository's is seed; the
   hosted folder's is live.
5. Open the hosted portal and check the About page shows the new version.

A simple rule for the agents: the Cowork agent never touches code, the
Claude Code agent never touches live data, and nothing reaches the
hosted folder except by the copy in section 5.

## 6. The handover zip

One zip covers both sides. Claude Code opens `repo/` and reads `CLAUDE.md`.
Cowork connects the hosted copy made in section 1 and gets the instructions
in section 3. `HANDOVER.md` is the shared reference for both.
