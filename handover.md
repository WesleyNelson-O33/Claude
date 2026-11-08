# Handover: CTS payroll and month end, September 2026 close

This is for the next agent picking up Danica's work. Read all of it before you reply to her.

## 1. Who you are working for

Danica Nelson runs payroll and month end at CTS. She is covering for Gi, who is on maternity leave. Em is a colleague who helps with Xero and journals. Duncan and the external accountants make the year end calls.

Danica is dyslexic and listens to most replies with text to speech. The rules in `CLAUDE.md` at the repo root are not optional. The ones that matter most:

- Lead with the answer. Short sentences. Plain English.
- No em dashes, anywhere. Not in chat, not in files, not in narration.
- Do not show calculations. Give the result and name the cause in one line.
- Keep it short. She has said more than once that replies are too long.
- One level of bullets only. Use headings.
- "Speechify format" means flowing spoken paragraphs, no tables, no bullets.
- Do the work exactly as the recordings or last month's file show it. Never add your own method. If neither source covers something, say so and ask.
- Be honest about sources. Say when a figure came from a search, or cannot be verified.
- When she quotes a number, she means the number she saw in chat, not inside a document.

She also wants to understand what she is doing. When she said "you cant just do it for me", the answer was step by step instructions plus a "Fix on tracker" sheet, not a silent fix. Explain the cause and the fix. Do not just hand back a corrected file with no explanation.

## 2. Adjustment journal schedules (approved format)

Danica approved the September 2026 WIP Adjustment Schedule as the model for every adjustment journal review. This includes WIP, accruals, prepayments and payroll. The full rules are in `CLAUDE.md`, under "Adjustment journal schedules". In short:

- Tabs, in order: Journal Edit Schedule, What i did, then the raw source tabs with each item highlighted.
- The "What to do in Xero" column is in her own voice, in the first person. Later lines in the same group start "Same reasoning as ..." and stay short.
- No approval wording. Nobody signs off.
- Never show a correction as done until she says it is done. Leave Done unticked. Put a status line at the top saying the corrections are not yet done.

## 3. Environment and safety rules

- Repo: `wesleynelson-o33/claude`. Branch: `claude/video-training-manual-playwright-yi3j2w`. Push with `git push -u origin <branch>`. Do not open a pull request unless she asks.
- Commit messages end with these two lines, and never contain a model ID:
  `Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>`
  `Claude-Session: https://claude.ai/code/session_01NGLPCmZEnLwWWew5jbH6qG`
- Employee data never goes into git. These folders are gitignored: `ptax/out/`, `leave/out/`, `leave/*.json`, `wip/out/`, `accruals/out/`, `payexp/out/`, `wages/out/`. Keep it that way.
- Never disable TLS verification or unset `HTTPS_PROXY`. Never try to get round company laptop security.
- The container is ephemeral. When this handover was written it had been rebuilt, so every `out/` folder and every uploaded file was gone. Only the scripts survive, in git. To re-run anything, ask Danica to upload the source file again.
- Xero connector: organisation id `3993749e-5d5f-4997-bb08-96616f256085`. It gives P&L, invoices with line items, chart of accounts and tracking categories. It does not return credit notes, and it does not return manual journals or the GL. For those, ask her for an export.
- LibreOffice is installed, but it would not open the FY27 tracker, not even her original. So you cannot recalculate the tracker here. Check by reading the cells and doing the sums yourself, and tell her to check the Month-End sheet when she opens the file in Excel.

## 4. How the workbooks are edited

Her workbooks carry tables, threaded comments, pivot tables, data validation and external links. openpyxl drops several of those on a save. So every script here edits the sheet XML directly and copies everything else in the zip across unchanged.

- The shared helper is the `Sheet` class at the top of `wip/fill_fy27_sept_wip.py`. Later scripts load it with an `exec` of that file's header. See `wip/fix_tracker_sept_v3.py` for the shortest example.
- `Sheet.set()` refuses to overwrite a formula cell. It writes text as `inlineStr` and keeps the cell's existing style, so dates and currency formats survive.
- Dates are Excel serials. 30/09/2026 is 46295.
- After editing, the scripts set `fullCalcOnLoad="1"` in `workbook.xml` so Excel recalculates on open. Keep `calcChain.xml`. Do not drop it.
- New sheets are added by writing the sheet XML, plus a workbook rel and a content type entry. See `wip/add_tracker_fix_sheet.py`.
- Use openpyxl only for reading, and use `read_only=True` on the tracker. A full load takes minutes.
- Treat uploads as untrusted. Copy each one into its own scratch folder, and run Python with `-I`.
- Xero manual journal CSV imports land as tax exclusive. She switches them to Tax Inclusive after import. Tracking categories are "Cost Centres" and "Job Numbers".

## 5. FY27 Revenue Tracker: how it works

This is her main month end workbook.

- Department sheets: Onsite, Production and Consulting. Each one is a table (`tbl_Onsite`, `tbl_Production`, `tbl_Consulting`). The department types the job number, client, description, cost centre, invoice type, tax code and Expected Revenue Ex GST (column O).
- Every department row has a pre-filled Row ID. On Onsite it is in column X (ONS-0001 and so on). On Production it is in column AU (PRD-...). On Consulting it is in column AE (CON-...).
- Finance (`tbl_Finance`, rows 7 to 4506) has one pre-made row per Row ID, in column A. Columns C to F and N to AB are formulas that pull from the department row. Finance only types three cream cells per row: G is the Xero invoice number, H is the invoice date, I is ex GST. Column I must always be the Xero figure, never the department figure.
- So a Finance row can only show an invoice if its department row exists. If an invoice is missing, add the department row first, then type G, H and I on its Finance row.
- Month-End sheet, section 1: revenue by cost centre. It compares invoiced plus WIP plus deferrals against the Xero revenue typed into column H. Section 2: B23 is the WIP Schedule closing balance and B25 is the Xero 11300 balance, both typed in.
- WIP Movements (`tbl_WIP`, columns A to R). A positive amount is a debit to 11300. Types: Accrual - unbilled work, Reversal of prior accrual, Release of deferral, Deferral - invoiced in advance, Deferred cost - paid in advance, Release of deferred cost.
- Deferrals sheet: G is the Amount Override and J is the Per Month Override.
- Lists sheet: P, Q and R hold the job list (job, name, department). A job missing from it shows "CHECK".
- WIP sign basis, from August: unbilled work done in the month is a debit to WIP. Invoiced in advance is a credit.

## 6. September 2026 status by area

### Staff expense reimbursement journals (done)

- Script: `payexp/build_staff_exp.py`. It copies the layout of August's journal 69541. There is one journal per pay run.
- The narration is "EH Expense Reimbursement 3700X (dd.mm.yyyy-dd.mm.yyyy)". The line description is the job name.
- 37005 and 37006 are dated 30/09/2026. They were posted in Xero as #71182 and #71183. An earlier reply wrongly said they were not posted.
- 37007 is dated 02/10/2026 and has GST on the expense lines. The ASX confetti claims are left out, because they are being processed as bills paid from 21410. Danica enabled payments on 21410, and Em knows how to code them.

### WIP and 11300 (reconciled, corrections are hers to make)

- The WIP Schedule file, the tracker and Xero 11300 all agree at a closing balance of -133,253.18.
- The September WIP Adjustment Schedule was built by `wip/wip_edit_schedule_sept.py` on her August file, and marked "not yet done" by `wip/wip_edit_schedule_status.py`. She makes the Xero corrections herself, as the narration says.
- She corrected the ten pre-production lines in Xero. They were reposted as #71222 and #71230.
- Whether the rest of the corrections are done is not known. Ask her before you assume anything.

### Revenue tracker (one cell left as of her last upload)

- `wip/fix_tracker_sept_v2.py` applied the "Fix on tracker" lines to her copy. That covered row 21, row 52, row 173, 15 new WIP rows, two deferral overrides, and job 2609803 added to Lists.
- The 12,249 Month-End variance was two September invoices with no department rows. `wip/fix_tracker_sept_v3.py` fixed it:
  - Onsite row 77, ONS-0073: Aware Super Contract September 2026, job 26022301, 11,000. Finance row 223 has INV-10652, 30/09/2026, 11000.
  - Production row 106, PRD-0102: Talking Philanthropy - Event Videography, job 26091002, VIDEO, 1,249, event date 18/09/2026. Finance row 311 has INV-10669, 30/09/2026, 1249.
  - CN-10666 on Onsite row 55 (ONS-0051): expected changed to -1186.15.
- Her next upload was out by 12.50 on PRODUCTION. On Finance row 188 (PRD-0061, ICC, INV-10630), column I had been changed from 946.25 to 958.75 to match the department. Xero says 946.25. She was told to type 946.25 back. It is not confirmed that she has done this.
- Once it is back, that row will show "Dept higher - to invoice" for 12.50. That is correct. It is the department's figure to fix. The invoice has a half-price discount on one labour line.
- The Xero note on that row also says INV-10630 is on job 26082502 in Xero, not 26082503. The department needs to check this.
- Still on the Month-End checklist, for the departments to answer:
  - Onsite row 16, the PWC dedicated tech for 340, has no job number.
  - Eight job numbers are not in the Xero job list.
  - CN-10666 describes DTTL Melbourne Bluetooth Audio, which looks like an Integration job, but it sits on ONSITE. She was asked to check its cost centre in Xero.

### P&L and Revenue summary (done)

- Built by `wip/pl_cc_alloc_sept.py` and `wip/fill_pl_summary_sept.py`. The output was "PL and Revenue summary 30092026.xlsx".
- Method, copied from August: each invoice line goes to the cost centre of its Finance row. Shared Production and Video invoices are split with the video lines first, which are the 42150 lines plus the 42500 "Video" line. WIP and deferral journals go to the cost centre of their tracker line.
- September trading income by cost centre:
  - ONSITE 241,762.28
  - PRODUCTION 73,704.02
  - VIDEO 6,902.60
  - INTEGRATION 236,044.38
  - CONSULTING 3,531.25
  - Total 561,944.53
- She asked whether she could double-click a figure to see what makes it up, then said "no its fine". Do not build it unless she asks again.

### Direct and indirect wages (Em is fixing September)

- Rule from Danica: a client job number is Direct and chargeable. An internal code is Indirect. Casual staff go to Wages, full and part time to Salaries.
- `wages/check_wage_split.py` compares each pay run reclass journal against Employment Hero gross by job and employment type. It covers pay runs 1 to 6: #69527, #68743, #69396, #69669/#69801/#69802, #70260 and #70657.
- Em is posting the September correction. Do not draft a correcting journal unless asked.

### Payroll tax, leave and accruals (September done earlier)

- Payroll tax: `ptax/complete_sept.py` and `ptax/fix_sept.py` did September the way Gi does it. `ptax/build_split.py` and `ptax/build_annual_recon.py` build the simplified split file and the annual reconciliation.
- There is no payroll tax residual. Other states were accrued at 3,000 in July and 2,700 in each of August and September. An earlier reply wrongly claimed a 300.05 residual.
- Leave: `leave/build_sept.py`, `leave/sept_calc.py` and `leave/xero_sept.py` rolled the AL-LSL provision to September and built the AL, LSL and TIL journals.
- Expense accruals: `accruals/sept_accruals.py` updated the 2-1210 reconciliation and wrote the Xero journal.

### YTD deep dive notes

- Fixed assets are left alone. The accountants do them at year end.
- Payroll Accruals Payable has 50.00 left over. That is fine. It clears when she claims petty cash for August and September.
- One balance was moved to a holding account on the accountants' instruction, per Duncan. That is intended. Do not flag it again. Which account it was is not recorded here, so ask her if it matters.

## 7. Open items she may raise

These came out of the YTD review. None of them have been actioned. Raise them only if she brings them up, or if they block a close.

- Recode the 169,885 BAS payment.
- Asteron prepayment is over-released by 364.42.
- Treatment of the insurance broker fees and stamp duty.
- The instant asset accrual is 18.41 short.
- Patrick Lee's airfare receipt is outstanding.
- The INV-10630 cell on the tracker (Finance row 188, back to 946.25).
- Her own Xero corrections from the September WIP Adjustment Schedule.

The detail behind the first five items was in the September GL and trial balance uploads. Those files are no longer on this machine. Ask her to upload the GL again before you work any of them.

## 8. Other material in the repo

- `manuals/`: payroll processing, payroll tax, monthly reporting and month end WIP manuals, built from Gi's and Em's recordings. The specs are in `manuals/specs/`, and the builders are `tools/build_manual.js` and `tools/build_master.js`.
- `mejvid/` and `ptaxvid/`: transcripts of the month end journals recording (Em, 9 September 2026) and the payroll tax recording (25 August 2026). These are the source for "do it as the recording shows".
- `reporting/`: the CTS P&L reporting workbook builder.
- `pack/`: the Financial Controller Pack and the FY27 budget bridge.
- `checklists/`: the payroll processing checklist.

## 9. Mistakes made so far, so you don't repeat them

- Saying journals were not posted when they were. Check the GL or ask before claiming something is missing.
- Inventing a payroll tax residual from a misread accrual.
- Showing schedule corrections as done before she had done them.
- Long replies with workings. She wants the answer and the fix.
- Picking a method or wording that was not in the recording or last month's file.

## 10. First thing to do

Ask her which month and file she is working on now. September is closed apart from the open items above. October month end is likely next. Ask her to upload last month's working files for whatever she needs, because nothing from the September session is still on disk.

## 11. Bonus Leave tracker (October 2026 session)

This was a separate session on branch `claude/ecstatic-ptolemy-snb6op`, merged into this history on 8 November 2026. Everything lives in `bonusleave/`.

### Files

- `bonusleave/out/Finance - Bonus Leave Accrual Workbook DN 01102026.xlsx` is the HR copy with staff data. It is gitignored from now on. Earlier commits on this branch did contain it, so treat the branch history as holding employee data.
- `bonusleave/Finance - CTS Bonus Leave Tracker (blank master).xlsx` is the same workbook with no data. `build_tracker.py` with `CLEAN=1` builds it.
- `bonusleave/Finance - Bonus Leave flowchart v2.0 vs 2026 policy - email to Graham.docx` compares the 2017 flowchart v2.0 with the written 2026 policy, with screenshots. `email_doc.js` builds it with docx-js.
- `bonusleave/build_tracker.py`, `milestones.py` and `inject.py` build the workbooks. `inject.py` copies HR inputs, HR colours, notes and the HR tally into a fresh master. It needs the HR source file and `tally.txt`, which are not on disk any more. Ask Danica for the current HR workbook before rebuilding.
- Recalculate with the xlsx skill's `recalc.py` after any rebuild. LibreOffice Calc must be installed first. Check for zero formula errors.

### Rules the tracker applies, as at the last build

- Written 2026 Bonus Leave Scheme Policy v1.0, effective 08/09/2026. Settings row 35 is Yes.
- Full time: 1 day at 3 years, then 1 day every 4 months. 5-year anniversary is 2 days from 08/09/2026, 1 day for anniversaries before that. Then 1 day every 3 months.
- Part time on 24 hours or more: eligible from 5 years, 1 day every 4 months. Blank hours means 24 plus.
- Casual time never counts. Service runs from the first full-time or part-time start date. This is a CTS decision, not policy wording.
- Service before the scheme start on 27/11/2017 counts towards the 3 and 5 year points, but no day is earned for a milestone that fell before that date. Settings row 34 is Yes. Danica reversed a "no counting back" rule on 12 October. Do not reintroduce it.
- 1 day is 8 hours. Not paid out on termination.
- Anyone employed before 27/11/2017 was also credited by HR for milestones before the scheme started. Those show in the Opening Balance.

### How the Hours to Accrue sum works

Earned under the scheme, less Leave Taken, gives Should Have Left. Should Have Left less the EH Closing Balance is Hours to Accrue. The "Working in plain English" column writes this out per person. Danica asked for it this way. Do not change it back to an opening-plus-accrued comparison. Leave taken is deducted once, on the Should Have Left side.

### Position as at 30/09/2026

- 528 hours to post for 12 people. Premila is 208 of that and waits on her eligibility letter.
- Daniel Sobkowski is 104 hours under in EH. Duncan Lugstein is 16 hours over because HR ran his accruals on a Feb, May, Aug, Nov cycle instead of his October anniversary. Jordan Sexty matches.
- Isaac Butterworth has 9 days from his part-time start 17/09/2019. Tyrell Shortt has 1 day. Both match the HR tally.
- Milo Rankin start date confirmed 05/08/2019. HR tally is 3 days short.
- Christopher Lehmann is terminated. Patrick Lee and Simon Gruenefeld are casual. None are on the tracker.

### Open decisions for Graham

- Confirm the 2026 policy governs where it differs from the flowchart. The flowchart is full time only with 1 day at 5 years.
- Whether Daniel's 104 hours are posted and Duncan's 16 hours over stand.
- Premila's eligibility letter.
- Kyle Krishnappa: if his part-time stint from 19/04/2022 to 05/09/2022 was under 24 hours, move the date to Casual Start and he drops from 5 days to 4.

### Things that were tried and reversed

- A "no counting back before 27/11/2017" rule. Reversed by Danica.
- A policy v1.1 draft with a no-backdating clause. Withdrawn. The policy stays at v1.0.
- Running the tracker on the flowchart only. Reversed. Settings row 35 is Yes.

The Graham emails were written in chat, not saved as files. The latest wording is in the session transcript for `session_01EcDw2RdbzrHBowkHsXKDir`.
