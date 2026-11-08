# Handover: CTS finance work for Danica

Written 8 November 2026 for the next agent. Read this file and CLAUDE.md before you do anything.

## Read this first

CLAUDE.md is the rulebook. It is not optional. The rules that matter most:

- Danica is dyslexic and listens to replies with text to speech. Short sentences, plain words, answer first.
- No em dashes anywhere. Hard rule. Use a comma, a full stop or a new sentence.
- No nested bullets. One level only.
- Give the answer, not the working. Only show figures she has to type, and say where they go.
- Do the work exactly as the recordings and last month's file show it. Never add your own method. If the source does not cover something, say so and ask.
- Never show a correction as done unless she has told you it is done.
- Say plainly when something comes from a search, a guess or a gap, not a source document.
- She quotes the numbering she sees in chat, not the numbering inside a document. Resolve by what she describes.
- When she says "speechify format", give flowing spoken paragraphs under headings. No tables, no bullets.

## Who and what

Danica works in finance at CTS, Corporate Technology Services. She is taking over payroll, payroll tax, month-end journals and monthly reporting from colleagues who recorded handover sessions.

People named in the recordings:

- Em (Emelda Waiya) ran the month-end journals and WIP session.
- Gi ran the payroll tax session.
- The reporting recording (video 7) walks her through the P&L, graphs and management report.

Systems:

- Xero for the ledger, invoices and manual journals. A Xero connector is set up in this account.
- Employment Hero (EH) for payroll, timesheets and leave.
- CommBiz for the bank upload of pay runs.

Departments and their tags:

- ONS is Onsite, also called managed services.
- PRD is Production. VID is Video, which sits with Production on some invoices.
- INT is Integration. CONS is Consulting.
- CTS is overheads and admin.

Department tags appear in square brackets on GL lines and tracking options, for example [ONS].

## Big warning: input files are gone

Every script in this repo read Danica's uploaded files from /root/.claude/uploads. That folder belonged to the old session and no longer exists.

The output folders (wip/out, leave/out, ptax/out, accruals/out, payexp/out, wages/out) are in .gitignore. They held employee and payroll data, so they were never committed. They are gone too.

What this means:

- The scripts are a record of exactly what was done and how. They will not run again without her source files.
- To rerun or roll a script forward, ask Danica to upload the files again. Then change the path constants at the top of the script.
- Most scripts take paths as command line arguments. The older ones (leave, pack, ptax/fix_sept, payexp) have hard-coded upload paths at the top.
- Some scripts also read saved Xero tool results (chart of accounts, tracking categories) from the old session folder. Pull those fresh from the Xero connector.

## Git

- Repo: wesleynelson-o33/claude.
- Working branch: claude/funny-euler-87s5em. It has all 50 commits.
- The other branch, claude/video-training-manual-playwright-yi3j2w, has nothing this branch does not already have.
- No pull request has been opened. Do not open one unless she asks.

## How the scripts edit Excel files

This is the main technical lesson. Her workbooks carry pivot tables, threaded comments, tables, drawings, protection and external links. openpyxl drops many of these when it saves.

So most scripts edit the sheet XML inside the xlsx zip directly, and copy every other part across untouched. Keep doing it this way for any file of hers that has pivots or comments.

- pack/validate.py and ptax/validate.py check a finished xlsx the way Excel does before it offers to repair. Run one on every file before you hand it back.
- Use openpyxl only to read values, or for workbooks that were built from scratch here.

## Folder by folder

### manuals

Training manuals built from the recordings, as Word documents. Built from the JSON specs in manuals/specs by tools/build_manual.js and tools/build_master.js.

- Payroll Processing Parts 1, 2, 3a, 3b, 3c, 5, 6 and 7, plus the complete bound manual.
- There is a Part 4 transcript and redaction file but no Part 4 Word document in the repo. I did not find out why. Ask before assuming.
- Payroll Step by Step, Listening Version, is the speechify style version.
- Payroll Handover, Meeting Prep and Questions and Scenarios.
- Monthly Reporting Parts 1 to 3 and the complete manual.
- Payroll Tax, The Monthly Process.
- Month-End Journals, WIP.
- Screenshots were pulled from the recordings with tools/extract_frames.py and blurred with tools/redact.py, using the redaction JSON files.

### Transcripts (the source of truth for method)

- manuals/part1 to part6 narration transcripts: the payroll processing recordings.
- reporting/source/video7-reporting-narration-transcript.txt: monthly reporting, P&L analysis, graphs, management report, sales forecast update.
- ptaxvid/transcripts: Gi's payroll tax session, 25 August 2026. Use payroll-tax-full-ordered.txt.
- mejvid/transcripts: Em's month-end journals session, 9 September 2026. Use month-end-journals-full.txt. The screen share ends about two hours twenty in.

All are machine transcribed. Names and figures may be misheard. The videos themselves are on GitHub releases, not in the tree. tools/transcribe.py rebuilds a transcript with a local Whisper model.

### ptax (payroll tax)

- fix_sept.py: September payroll tax fixes on the payroll tax workbook and split file.
- complete_sept.py: completes the September payroll tax workbook the way Gi does it.
- build_split.py and fill_split.py: a simplified rebuild of the EH Payroll Split file. Three pastes, all formulas, no pivots.
- build_annual_recon.py and fill_annual_recon.py: annual reconciliation workbook for the other states (ACT, QLD, SA, VIC, WA).
- Employees are put in a state by residential state, the way Gi does it.

### leave

- build_sept.py: rolled the AL-LSL Provision workbook from August to September 2026, the same way it was rolled from July to August.
- sept_calc.py: the September annual leave, long service leave and time in lieu journals.
- xero_sept.py: Xero manual journal import files for those journals.
- Later commits added cost centres and one description to the leave journals.

### accruals

- sept_accruals.py: updated the 2-1210 expense accruals reconciliation for September and wrote the Xero journal. Amounts are ex GST.

### payexp (staff expense reimbursements)

- build_staff_exp.py: one Xero journal per pay run, copying August's journal 69541 layout.
- Pay runs 37005 and 37006 dated 30/09/2026, and 37007 dated 02/10/2026, as Danica set them.
- Final state: GST is on the expense lines. ASX confetti claims were left out because they go through bills.

### wages

- check_wage_split.py: checks each pay run reclass journal against Employment Hero.
- Danica's rule: client job is Direct, internal code is Indirect. Casual is Wages, full or part time is Salaries. Department comes from the job tag.

### wip (revenue tracker and WIP, September 2026)

This was the most recent work, finished around 15 October.

- fill_fy27_sept.py: brought the FY27 Revenue Tracker up to date for September and corrected the Finance cells to what is in Xero.
- fill_fy27_sept_wip.py: brought the WIP side of the tracker into line with Xero account 11300, journals 70539, 71186 and 71188.
- xero_wip_sept.py: September WIP journals for Xero import, in the same layout as August's journal in the recording.
- change_schedule_sept.py: schedule of every tracker change, with WIP Schedule row references.
- wip_edit_schedule_sept.py: the September WIP Adjustment Schedule, built on her August file. This is the layout she approved as the model for all adjustment journal reviews. CLAUDE.md describes it in full.
- wip_edit_schedule_status.py: marked that schedule as not yet corrected. Done is unticked and the status lines say the corrections are still to do.
- add_tracker_fix_sheet.py: added a "Fix on tracker" sheet listing tracker lines that still need to catch up.
- fix_tracker_sept_v2.py: applied those fixes to her own 15/10 tracker copy.
- fix_tracker_sept_v3.py: cleared a 12,249 Month-End revenue variance. Added two missing September rows: Aware Super INV-10652 for 11,000 on Onsite, and Talking Philanthropy INV-10669 for 1,249 on Production Video. Set the CN-10666 credit note line to minus 1,186.15.
- pl_cc_alloc_sept.py and fill_pl_summary_sept.py: split September trading income by cost centre for her PL and Revenue summary, the same way August's was built.
- fix_utilisation.py: September utilisation report. Overtime only, taken from paid overtime in the raw data.

### pack (Financial Controller Pack)

Built in September 2026, before the month-end work.

- build_pack.py builds CTS Financial Controller Pack.xlsx. One paste of the Xero account transactions export drives the department P&L, budget variance and summary.
- The Risk column flags No Activity, Low, Medium or High. Thresholds on the Setup tab are 10 percent, 25 percent and 50,000 dollars.
- Future months on the P&L come from Budget FY27, so it always shows a full twelve months.
- build_bridge.py added Pack Map and Pack Budget FY27 tabs to the FY27 budget so it can feed the pack.
- Five overhead lines were switched to the 3 Way split and highlighted orange.
- The budget file here is "CTS Budget FY27 Final - 3 Way fixed.xlsx".

### reporting

- build_pnl.py and build_accounts.py build CTS P&L Reporting.xlsx. Only the Actuals and Budget tabs are typed into. Everything else is formulas.
- data/accounts.json holds the chart of accounts with each account's division.

### checklists

- Payroll Processing Checklist.xlsx.

### tools

- Split video for chat (.ps1 and .bat): a Windows script Danica uses to cut a recording into pieces small enough to upload to chat.

## Where things stand

What I know from the commits, as of 15 October:

- September WIP corrections in Xero were not done when the schedule was written. The schedule says so. I do not know if she has done them since. Ask her.
- September payroll tax, leave, accruals, staff expenses and the wage split check all have scripts. I cannot confirm from the repo whether each was posted in Xero. Ask before treating any as final.
- No October month-end work is in the repo. It is now November, so October is likely the next job. Roll each process forward from September's file and the recordings, exactly as before.

## Open conversation: revenue risk forecast

On 18 October she asked how best to produce a revenue risk forecast for CTS. Nothing has been built yet.

What I told her:

- No recording shows a revenue risk forecast. Any method is my suggestion, not the CTS process. Check who wants it and in what format before building.
- The best base is the three department Sales Forecast tabs in the budget file: CONS Sales Forecast, ONS Sales Forecast and PRD Sales Forecast. Each has a row per client and job, with monthly forecast, actual and variance.
- Update those tabs to actual first, the way video 7 shows it from about 2:17:40.
- Tag each future line as secured (contract, repeating invoice, licence), likely (quoted or in progress) or at risk (unsigned pipeline).
- Total secured, secured plus likely, and full forecast per department, against budget. Flag the gap with the pack's existing risk thresholds.
- Read it against the seasonality described in video 7: Production peaks in Q2 with AGM season, December, January and March are quiet, Integration and Consulting peak in May and June, managed services is steady.
- Note that Production beat its budget at 110 percent because it understated its pipeline.
- Do not invent probability percentages per job.
- The CONS Sales Forecast tab in the repo copy still shows 2023 dates in its header row. Her live file may be newer.
- I offered to build it from the Sales Forecast tabs and Xero repeating invoices. She has not answered yet.

## Connectors

- Xero, Google Drive, Google Calendar and GitHub are connected.
- Finance plugin connectors for BigQuery and Slack needed sign-in at last check. They are not needed for any work so far.

## How to start the next session

1. Read CLAUDE.md, then this file.
2. Ask Danica what she wants done first. Do not assume October month-end.
3. Ask her to upload the source files for that job: last month's workbook and the new month's exports.
4. Find the matching recording transcript and last month's script. Copy the method exactly.
5. Edit her files through the XML when they have pivots or comments. Run validate.py before you hand anything back.
6. Commit scripts to claude/funny-euler-87s5em. Keep employee data out of git.
