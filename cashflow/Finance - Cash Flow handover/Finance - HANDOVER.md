# Handover - CTS weekly cash flow forecast (Danica Nelson)

Written 8 Nov 2026. Files in this zip carry a "Finance - " prefix; the repo copies in `cashflow/` do not.
`WesleyNelson-O33/Claude` repo, branch `claude/awesome-bell-g1wh0e`.

## 1. Who and what

Danica Nelson is the accountant at Corporate Technology Services Pty Ltd (CTS),
a Sydney AV company. Xero organisation id `3993749e-5d5f-4997-bb08-96616f256085`.
She asked for a week by week cash flow forecast in Excel, forecast against
actual, that reconciles to the real bank balance (cheque + savings only).

She presents the file to management. Everything in it is written in her voice,
first person, plain English. People are initials in notes and chat: GC (Graham
Cherry), DL (Duncan Lugstein), JS (Jordan Sexty), MH (Mitch), TW (Tyler Wood),
Em is Em. Words in full, no abbreviations like "sal sac".

Read `/home/user/Claude/CLAUDE.md` first. It is her writing guide: short
sentences, lead with the answer, no em dashes anywhere, no calculations shown,
headings, no nested bullets. She is dyslexic and listens to replies.

## 2. Files (newest first)

| File | What it is |
|---|---|
| `CTS Cash Flow Forecast - Nov to Dec 2026.xlsx` | Eight weeks, Thu 5 Nov to Thu 31 Dec 2026. The live forecast. Built by `year_end.py`. |
| `CTS Cash Flow Forecast - 04112026.xlsx` | Five weeks, 29 Oct to 2 Dec. Week 1 actual. Superseded by the Nov to Dec file but she has it. |
| `CTS Cash Flow Forecast - 29102026 - bank actuals.xlsx` | The October period (28 Sep to 4 Nov), all five weeks actual, tied to the CommBiz cheque statement. Forecast columns frozen as numbers. |
| `CTS Cash Flow Forecast - 12102026.xlsx` | Older copy of the October file. Ignore. |
| `CTS Cash Flow Forecast - Oct 2026.xlsx` | First build. Ignore. |
| `build_cashflow.py`, `voice.py`, `user_edits.py` | Original builder. Not used any more, the workbooks are now edited in place. |
| `inject_values.py` | Writes cached values into every formula cell so viewers show numbers. Must run after every edit. |
| `update_actuals_12oct.py`, `update_ledgers_12oct.py`, `creditors_run_tab.py`, `actuals_to_28oct.py`, `roll_forward_29oct.py`, `update_4nov.py`, `data_4nov.py`, `bank_actuals.py`, `year_end.py` | The edit scripts in order. Each loads the previous workbook, edits with openpyxl, saves. `data_4nov.py` holds the open invoice and bill lists at 4 Nov. |

File naming: she names files by the date (ddmmyyyy). Her "29102026" file was
the October period, not the 29 Oct forecast. Check the Cash Flow A2 title
before assuming which period a file covers.

## 3. Workbook layout

Tabs: Cash Flow, Assumptions, Payroll, Receivables, Payables, Creditors Run.
(Older files also have Complete AV, a build up of the INV41565 split.)

Cash Flow tab. Rows are fixed and the same in every file:

- 5 to 7 headers, 8 "Actuals entered" flag per week (1 = done)
- 10 opening bank, 13 trading income, 14 interest, 15 total income
- 19 net wages, 20 super, 21 PAYG on the monthly IAS, 22 payroll tax, 23 reimbursements, 24 subtotal
- 27 to 31 cost of sales (equipment, sub contract, hires, client subscriptions, travel and freight), 32 subtotal
- 35 to 47 overheads (rent, IT, advertising, professional fees, dues, phones, recruitment, insurance, office, electricity, bank fees, credit card settlements, other), 48 creditors run, 49 subtotal
- 52 BAS, 53 subtotal, 55 total expenses, 57 net, 58 closing, 59 actual bank per CommBiz (typed), 60 difference, 63 to 66 memo, 69 to 72 the four risks (she wants exactly four)

Columns: each week has a Forecast column and a yellow Actual column.
Five week files: weeks in B/C, D/E, F/G, H/I, J/K, totals L/M, variance N, notes O.
Eight week file: weeks in B/C through P/Q, totals R/S, variance T, notes U.

Opening of week n = previous week's typed bank balance if its flag is 1, else
the previous forecast closing. Forecast cells are SUMIFS off the Receivables
(expected date, col H), Payables (planned date col G, category col F, method
col E not "Card", col K = "Yes") and Payroll tabs, using the week table on
Assumptions rows 44 onward. Monthly estimates on Assumptions are spread
/5 in the five week files and *12/52 in the eight week file.

Assumptions tab. Blue cells are inputs. Key ones: B5 start date, B6 opening
(cheque + savings), B9 gross per fortnight, B10 net percent, B12 super 12
percent, B14 to B16 pay dates (B60 is run 4 in the Dec file), B17/B18 payroll
tax, B19 its date, B21 BAS toggle, B22 to B25 BAS, B27 to B37 monthly
estimates, B39 small jobs estimate, B40 interest, week table from row 44,
IAS inputs rows 54 to 57 (Dec file), Nov payroll tax rows 58 to 59.

Payroll tab. Runs in rows 5 to 7 (5 to 8 in the Dec file). Fortnight is
Saturday to Friday and the money leaves the bank on the Tuesday after.
Super direct debits the day after (HeroClear / Employment Hero).

Creditors Run tab. She pastes the Xero aged payables detail export from row 11
(columns A to K). B4 is the run date, B5 the switch (1 = use the paste), L is an
optional pay on date per line, M an x to exclude (card bills). Cash Flow row 48
sums it by week. When B5 is 1, open "Xero bill" rows on the Payables tab drop
out of the category lines (col K) so nothing is counted twice. Templates stay
in the category lines.

Payables tab rows 4 to 16 are a "payments this week" block with statuses
AGREED - PAY (green), Due this week, AWAITING (amber), PAID (grey).

## 4. How actuals are done

The only reliable source is the CommBiz bank statement. She sent the cheque
account statement (06 2498 10084326) for 2 Aug to 1 Nov as a PDF. The October
file is tied to it to the cent. Rules used:

- Bank line is cheque + savings. Transfers between the two are left out.
  Savings is tracked off the transfers on the cheque statement plus month end
  interest (430.17 on 30 Sep, 411.39 on 30 Oct). Savings: 410,000 at 28 Sep,
  250,000 at 7 Oct, 550,000 at 14 Oct, 400,000 at 21 and 28 Oct, 400,411.39
  at 1 Nov. She keyed 150,000 of transfers out of savings for 2 to 4 Nov, so
  about 250,411 after that. Not confirmed.
- Payment batches on the statement ("CTS ddmmyyyy") go on the creditors run
  line with their make up in the note.
- CTSWAGES lines are net wages, HeroClear is super, Revenue NSW is payroll tax.
- Card spend sits in the card settlements line (transfers to xx4491 PL card).

Xero as a fallback: the cash basis P&L for a date range gives receipts and
payments by account, ex GST, grossed up x1.1. Xero total cash balance
(get_cash_position with request_date) less 456,746.27 (the gap between Xero
total cash and cheque + savings on 5 Oct) estimates cheque + savings. Both are
estimates and were replaced once the statement came.

Pay run net is about 79 percent of gross (6 Oct 79,736 on 101,246; 20 Oct
83,860 on 111,704; 3 Nov 90,994 on 115,121).

## 5. Position at 4 Nov (last data pull)

- Cheque 63,204.22 at 1 Nov per statement. Estimate 29,216.74 at 4 Nov.
- Cheque + savings estimate 279,628.13 at 4 Nov. Opening of the Nov to Dec file.
- Xero aged receivables 726,549.38, 56 invoices. Aged payables 84,568.84, 29 bills.
- Aware Super paid everything in October (237,987 on 13 Oct). Complete AV INV41565 fully paid (139,086 on 29 Sep, 39,033 on 8 Oct).
- Q1 BAS is 115,000, did not go on 28 Oct, assumed 25 Nov through the agent.
- Brookvale insurance 66,239 paid in the 2 Nov creditors run.
- Overdue debtors about 85k: Downer 39,664, Microsoft 13,846, Ventia 10,016, ASX 5,748, ON24 3,006.
- Forecast closing 31 Dec about 243,000.

## 6. Open items she still has to do or confirm

1. Type the CommBiz cheque + savings balance at 4 Nov into Assumptions B6 of the Nov to Dec file and K59 of the October file.
2. Other states payroll tax for September (2,684) is not on the cheque statement. Was it paid?
3. SOR-1423275 for 12,540 on 27 Oct and the CTS ad hoc payments on 23 Oct (1,862.63 and 54.12). What were they?
4. Headroom AV bills INV15 and INV-0033, both 2,992, same job. Duplicate?
5. Vega 1,969 still open in Xero. She thought it was paid.
6. Confirm the BAS date with the agent.
7. Nov to Dec file: December has only the 10k a week small jobs estimate for new work. The December creditors run has nothing in it yet.

## 7. Weekly update procedure

1. Get the CommBiz statement lines for the week (cheque account) or an Account Transactions export from Xero for cheque and savings.
2. Fill the week's Actual column off the bank, by the row map above. Keep transfers between cheque and savings out.
3. Type cheque + savings closing into row 59, set the flag in row 8 to 1. Row 60 should be nil.
4. Refresh Receivables and Payables to Xero today (aged receivables, aged payables, UNPAID invoices and bills). Mark paid items, add new ones with an expected date.
5. Paste the Xero aged payables detail into Creditors Run from row 11, set the run date.
6. Rewrite the notes (column O or U) and the four risks in her voice.
7. Run `inject_values.inject(path)` on the saved file, check no `#` errors, send the file, commit and push.

## 8. Xero tool notes (MCP)

- get_invoices status UNPAID returns only invoices not yet due. Overdue ones are missing. Use get_aged_receivables (has a cursor, 50 per page) for the full open list.
- get_bills status UNPAID does include overdue bills. Aged payables lists a few more (drafts).
- get_bank_account_transactions returns unreconciled items only. It cannot give the bank rec.
- get_profit_and_loss with accounting_basis CASH works for any date range. Payroll journals are dated at the pay run, not the bank date.
- get_cash_position with request_date and comparison_date gives Xero total cash on a day (all accounts including cards and others).
- Xero data refreshes a few times a day. Check last_refreshed.
- Gen-e INV-10532 showed 11,850.16 after a credit, then was paid in full 13,627.68 on 2 Nov.

## 9. Environment notes

- Run Python on her uploaded files with `python3 -I` and copy them into a scratch folder first.
- LibreOffice is not available for recalculation. `inject_values.py` uses the `formulas` package to compute and write cached values. SUMIFS with wildcard criteria and some functions do not evaluate in it; SUMPRODUCT with LEFT does.
- Set `wb.calculation.fullCalcOnLoad = True` before saving so Excel recalculates.
- The container gets rebuilt. Reinstall with `python3 -m pip install openpyxl formulas`, and `git checkout -B claude/awesome-bell-g1wh0e origin/claude/awesome-bell-g1wh0e` if the checkout is on another branch.
- Commit messages end with the Co-Authored-By and Claude-Session trailers used in the git log.

## 10. Things she has pushed back on

- "ewww" at the first layout. She wants her navy and grey styling kept.
- Only four risks.
- Formulas visible in cells (no hard coded results) but viewers must show numbers, hence inject_values.
- Actuals must come from the bank rec, not the Xero P&L. She sent the statement when told the connection cannot read the rec.
- She did not want a Summary tab in the workbook. She wanted the summary in chat to present from.
- She corrected the pay date logic: fortnight Sat to Fri, paid the Tuesday after. The 6 Oct run was in week 1 of the October file.
