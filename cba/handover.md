# Handover: CBA AGM 2026 invoice workbook

Written 8 November 2026 for the next agent. Read the whole file before touching the workbook.

## Read this first

The user is Danica Nelson. She works in finance at Corporate Technology Services Pty Ltd (CTS). Follow `CLAUDE.md` in the repo root for every reply. The rules that matter most here:

- Short sentences. Plain English. Lead with the answer.
- No em dashes anywhere. Not in chat, not in the workbook, not in emails.
- No nested bullets.
- Do not show calculations. Give the result and name the cause in one line.
- Never guess. Say when something is from a search and not a source document.
- Do exactly what she asks. Do not redesign layouts she has approved. She said "your complicating things" when I added sections and helper columns. Keep it simple.
- Australian spelling.

## The job

CBA (Commonwealth Bank of Australia) wants an invoice for its 2026 AGM as soon as possible. Their approving manager went on leave from Friday 30 October 2026.

Duncan Lugstein asked Danica for:

1. A summary of all supplier costs on the job.
2. A table of all staff hours on the job, with dates.

Duncan and Tyler Wood run the job. Bills marked "DL/TW yet to app" are waiting on their approval.

- Xero job (tracking option): **2506208 - CBA AGM 2026 [PRD]**, plus **2506208V - CBA AGM 2026 [VID]** for video.
- Do not mix this up with **26061202 CBA AGM 2027** or **3059 CBA AGM 2025**. Both exist in Xero.
- The event days were around 11 to 14 October 2026.
- Xero organisation: Corporate Technology Services Pty Ltd, organisation id `3993749e-5d5f-4997-bb08-96616f256085`.

## Where things stand

- The workbook is built and Danica has it.
- Two emails to Duncan were drafted for her to send. I do not know if she sent them.
- The last thing she asked for was the Staff Hours by Date tab in its original layout, with formulas that look up each line by Timesheet Id. That is done. It is version v4 below.
- Nothing has been checked in Xero since 30 October 2026. Re-check before you report totals again. The Encore bill or October travel claims may have arrived since.

## The workbook

The latest file is **CBA AGM 2026 Costs and Hours 30102026 v4.xlsx**. Danica has it. A copy sits in `cba/out/`, but that folder is gitignored and does not survive a new container. If you need the file, ask Danica to upload her current copy. She edits it between rounds. Always work from her latest upload and keep every change she made.

Her last upload was named `CBA_AGM_2026_Costs_and_Hours_29102026.xlsx`. The v4 file is that upload run through `cba/update_cba_agm_v8.py`.

### Tabs, in order

1. Summary
2. Supplier Costs
3. Staff Hours (the old monthly tab, kept because her WIP line reads from it)
4. Staff Hours by Date
5. Unclear Items
6. Eight EH timesheet tabs, all named "Timesheet ..." and pasted in by Danica
7. Three "Roster vs Timesheet" tabs. She said to stop using them. They are still in the file and nothing reads them.

### Summary tab (her layout, keep it)

- Row 3: Supplier bills, 49,083.73 ex GST. Note says ExhibitionCo INV8728 is not yet in Xero.
- Row 4: Encore, PO amount, 18,318.43 ex GST. Her note says the PO is attached for review.
- Row 5: Total cost to use, 67,402.16 ex GST. That is bills plus the Encore PO.
- Row 6: WIP, 17,257.50. Danica added this line. Her note says October WIP is not calculated yet.
- Row 7: Staff hours, EH timesheets, 385.58 hours.
- Rows 10 to 15: Notes for the invoice, numbered 1 to 5. Danica renumbered these herself.
- Rows 17 to 22: Costs with GST table. Supplier bills, Encore PO, WIP, then Total supplier costs: 84,659.66 ex GST, 8,466.05 GST, 93,125.71 inc GST. She added WIP here, with GST at 10 percent.
- Rows 25 onwards: Staff hours by employee, with hours, days, first date and last date, read from Staff Hours by Date.

### Supplier Costs tab

- Rows 4 to 17: one row per bill. Columns: Bill date, Supplier, Bill number, Description, Account, Ex GST, GST, Inc GST (formula), Status, Notes. Danica also typed notes in column L.
- Row 18: Total bills. Rows 20 and 21: Paid and Unpaid.
- Rows 23 to 25: Encore PO block. Row 27: Bills plus Encore PO amount.

### Staff Hours by Date tab (the v4 layout she approved)

- Row 3 headers: Date, Employee, Location, Start, Finish, Unpaid break, Hours worked, Status, Work type, Timesheet Id, Source tab.
- Rows 4 to 161: 158 lines, sorted by date.
- Columns J and K are typed values. Every other column is INDEX/MATCH on the Timesheet Id in the Source tab. Example for hours: `=ROUND(INDEX('Timesheet Jun - Aug 2025'!$R:$R,MATCH($J4,'Timesheet Jun - Aug 2025'!$E:$E,0))*24,2)`
- Row 162: Total, 385.58.
- Rows 165 to 181: Hours by employee (SUMIFS, and a distinct-days count).
- Rows 184 to 198: By month, split PRD and VID.
- Rows 201 to 205: Check. SUMIFS over all eight timesheet tabs, less the one duplicate line, compared to the total. The difference is 0.

It does not add new lines by itself. To add a line, add a row with the Timesheet Id and Source tab, then copy the formulas down.

## Supplier costs (from Xero, checked to 30 October 2026)

All matched by bill text containing 2506208 or CBA AGM 2026. Amounts are ex GST.

- Red Bee Media, INV7060003977, 3 Jul 2026, captioning test, 365.00, paid.
- GoGet, INV10498230926, 30 Sep 2026, one line only: Tyler chair review 21/09, 72.00, paid. The rest of that bill is other jobs.
- Elan Projects (PopUp WiFi), IAU35330, 1 Oct 2026, 6,500.00, paid.
- Pop-Up Gardens, INV2070, 1 Oct 2026, 2,905.00, paid.
- Valiant Corporate, INV87154, 1 Oct 2026, chair hire, 2,178.72, paid.
- Tony Taouk (New Office), CTS-20261007, 7 Oct 2026, acoustic screens, 3,072.73, paid.
- Officeworks, INV632570641, 7 Oct 2026, 98.18, unpaid.
- Start Webcast, INV-202222, 14 Oct 2026, 11,624.60, unpaid. Waiting on DL/TW approval. Danica said to include it and say so.
- Access Innovation Media, INV22272, INV22273, INV22274 and INV22275, 19 Oct 2026: 240.00, 840.00, 620.00 and 150.00, unpaid, waiting on DL/TW approval.
- Deaf Services, INV409016, 26 Oct 2026, interpreter on 14/10, 282.50, unpaid. Waiting on DL/TW approval and confirmation of which credit card.
- ExhibitionCo, INV8728, 20,135.00, not in Xero. Danica added this row. Duncan gave her the invoice on 29/10/2026. I checked Xero: no ExhibitionCo bill or PO.

Not yet billed:

- Encore Event Technologies, PO-00016949, 9 Oct 2026, 18,318.43 ex GST. Danica's decision: use the PO amount and label it as the PO amount everywhere.

Danica checked my list against her Xero screen and it matched, apart from Deaf Services, which I then added.

## Decisions Danica has made

- Encore: use the PO amount and say it is the PO amount.
- Start Webcast: include it as billed and say it is still waiting on approval.
- October GoGet, taxi and expense claims: none yet. Say none are included.
- Charge-out rates to mention: 130 dollars an hour for production, 95 dollars an hour for video. These are the rates on the revenue tracker.
- Swan Taxis bill 13062026, 57.28: left out. It belongs to the CBA AGM 2027 job, 26061202, and the revenue tracker already holds it in that job's WIP. Danica first thought "2027" meant FY27. It does not. It is a separate job.
- Unclear Items tab: left out of the totals. These say "CBA AGM" with no year. By date they look like the 2025 AGM.
- Staff hours: use the plain EH timesheet tabs, not the roster vs timesheet report.
- Emails to Duncan: a summary only. The detail lives in the workbook.

## Staff hours (from Danica's EH timesheet tabs)

- 158 lines at location CBA AGM 2026 [PRD] or [VID], from 4 June 2025 to 28 October 2026.
- Total: 385.58 hours. Production 378.33, video 7.25.
- By person: Tyler Wood 166.83, Duncan Lugstein 92.50, Mitchell Healey 30.50, John Chae 26.00, Wade Tonna 16.00, Daniel Sobkowski 12.00, Sebastien Hathaway 9.75, Ruairidh Howard 9.00, Gabriela Panza Romanini Alchaar 7.50, Milo Rankin 6.00, Jaazaniah Salanoa 4.00, Jack Cunningham 2.50, Adam Arnou 2.00, Alana Carling 1.00.
- October 2026 is 247 hours.

How it was counted:

- Rejected timesheets are left out.
- The "Apr - May 2025" tab actually runs to 30 June 2025, so it overlaps "Jun - Aug 2025". One line, Timesheet Id 207177482 (Duncan, 4 June 2025, 2 hours), appears in both and is counted once.
- Hours come from the EH Duration column, which is already net of unpaid breaks.
- Nine lines have work type "Time in Lieu Approved". They were worked on the job, so they are included.
- One line is still Submitted (Tyler, 28 Oct 2026, 2 hours). It is included.

Danica asked for a triple check of the total and of Duncan's time. All three ways agreed:

1. The in-sheet SUMIFS check.
2. An independent search of every column of every tab.
3. Finish less start less break, checked against Duration for every line.

Duncan has 92.50 hours across 48 lines, on 41 days. His September (8.25) and August (2.50) hours match the earlier PRD labour hour reports.

Open questions on hours:

- Jaaz, 27 July 2026: the line is coded to Production - Quotes, but its comment says "CBA AGM 2026 start 1430 end 1530". It may be 1 hour on the wrong code. It is not counted. This was put to Duncan in the draft email.
- Five lines show 0 hours in EH: Duncan 3 Jul 2026, Tyler 1 Oct 2026, and Daniel 27 Feb, 23 Mar and 12 Jun 2026. They do not change the total.
- Tyler, 30 June 2026: 20 minutes, which shows as 0.33 hours. That is why June 2026 is 10.08.

## WIP on the revenue tracker

Source: `FY27_Revenue_Tracker__29092026.xlsx`, WIP Movements tab, rows 17 to 21.

- June 2026 opening balance: 6,390.00 PRD and 142.50 VID.
- July accrual: 2,535.00.
- August accrual: 1,787.50.
- September accrual: 6,402.50. The tracker shows 6,042.50, which is a typo already on the earlier "Fix on tracker" sheet.
- Total: 17,257.50. October is not calculated.

Things to flag if they come up:

- August WIP is 1.5 hours more than the August timesheets.
- September WIP is 0.5 hours less than the September timesheets.
- Danica put WIP inside "Total supplier costs" on the Summary tab. WIP is our own labour at charge-out rates, not a supplier cost. If the invoice also charges the staff hours, the labour is counted twice. I told her once. Do not keep pushing it, but raise it if she builds the invoice from that total.

## Still open

- The Encore final bill. Check Xero for it.
- Approvals from Duncan and Tyler: Start Webcast, the four Access Innovation bills, and Deaf Services, including which credit card.
- Any October travel, GoGet, taxi or expense claims for the event week. Check Xero.
- Whether Jaaz's 1 hour on 27 July should be recoded to CBA AGM 2026.
- October WIP has not been calculated.
- Whether to delete the three Roster vs Timesheet tabs. I offered and she has not answered.

## Scripts in `cba/`

All committed on branch `claude/admiring-johnson-2e6wah`. Each one takes an input file and an output file. Run them with `python3 -I script.py IN.xlsx OUT.xlsx`. Then recalculate with the xlsx skill's `scripts/recalc.py OUT.xlsx 400` and check for formula errors.

- `build_cba_agm_costs.py`: first build from scratch. Superseded.
- `update_cba_agm_v2.py` to `update_cba_agm_v4.py`: earlier passes on Danica's edited files (GST block, ExhibitionCo, Deaf Services row insert).
- `update_cba_agm_v5.py`: first Staff Hours by Date from the EH timesheet tabs, with hard values.
- `update_cba_agm_v6.py`: the same layout, each cell pointing at a fixed source row.
- `update_cba_agm_v7.py`: sections plus SMALL(IF) array lookups. **Danica rejected this as too complicated. Do not use it.**
- `update_cba_agm_v8.py`: **current.** The v6 layout, with INDEX/MATCH on Timesheet Id.

The scripts assert on her exact cell layout, for example Summary A7 text and B21 formula. If she has edited the file since, the asserts will fail. That is on purpose. Read her new file first and adjust. Do not force it.

Note that LibreOffice will not evaluate AGGREGATE with an array argument. SMALL(IF()) written as an ArrayFormula does work. MINIFS and MAXIFS need the `_xlfn.` prefix.

## Other sources used

- Xero, through the Xero connector: bills, purchase orders and contacts. Bill line items do not return tracking. Match on bill number and description text instead.
- Earlier Claude sessions: "Month-on-month analysis workbook" and "Power BI report from Employment Hero timesheets". These gave the early monthly hour figures on the old Staff Hours tab.
- The repo's `wip/` scripts and the month-end recording transcripts in `mejvid/transcripts/`, for how WIP is accrued on this job. Labour hours only, no costs accrued, per the recording.

## Emails drafted (for reference)

1. To Duncan, 29 October: the supplier invoices received, Encore still to come at the PO amount, no event-week travel yet, and some bills waiting on his or Tyler's approval. A summary only.
2. To Duncan, 30 October: the hours now come from the full EH timesheets (385.58 hours, 14 staff), Deaf Services was added, ExhibitionCo is included but not in Xero, WIP is to September, and the Jaaz 27 July question.
