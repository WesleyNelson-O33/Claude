# Handover

Handover from the session that read through the CTS accounts procedures and helped with Milo's Brisbane travel. Written 8 November 2026.

## Read this first

The user is Danica. She is dyslexic and listens to most replies with text to speech.

Follow `CLAUDE.md` in this repository exactly. The rules that matter most:

- Short sentences, plain language, answer first.
- No em dashes anywhere. This is a hard rule.
- No nested bullets.
- Do not show calculations. Give the result and name the cause in one line.
- Australian context and spelling.
- Never guess. Say when something is from a search or your own knowledge rather than her documents.
- For payroll and month end, do it exactly as the recordings or last month's file show. If the source does not cover it, say so and ask.

She has said more than once that replies are too long. Match the length to the question.

When she says "don't do anything yet, just read", she means it. Read and summarise only.

She will do bookings, payments and checkouts herself. Never go to a checkout or pay for anything.

## What this session did

1. Read 45 documents she uploaded: 5 older MYOB-era procedures, ACC001 to ACC032, the Accounts Glossary, the Accounts Handbook, the 2024 Accounts Procedure Master File, two expense reimbursement procedures, the Onsite Client Register, the 2026 Overtime Procedure draft, and the salary sacrifice transfer form.
2. Pulled the exact timesheet approval wording for payroll.
3. Searched for a Wednesday and Friday timesheet rule. It is not in any document.
4. Searched for a sales incentive policy. There isn't one.
5. Helped with Milo Rankin's Brisbane travel booking, and wrote messages to Mitch, Milo, Jordan and IT.

No files were changed in this repository apart from this handover.

## Where the documents are

The uploads were at `/root/.claude/uploads/2906c928-926d-5f7d-8064-97fbecb06256/`. A new session will probably not have them. Ask Danica to upload them again if you need the wording.

The BAS and IAS procedures contain the ATO portal password in plain text. Do not copy it into any file or reply.

## The procedures in one line each

All current procedures are Xero and Employment Hero (EH). Anything mentioning MYOB or ADP is out of date.

- **ACC001 Credit Card Reconciliation.** PL and DL cards. Each tax invoice is entered as a bill and paid from the card account. PL card 1-1131, DL card 1-1130. PL tops up 2,500 a week, with a 20,000 limit.
- **ACC002 Petty Cash.** Receipts are entered as bills and paid from 1-1140 straight away. The balance is always 200.
- **ACC003 Creditor Payments.** Weekly run, usually Wednesday, paid from 11110 Cheque Account. Reference is CTS DDMMYY. MA approves in Xero, MD approves in CommBiz. The cheque account keeps a 20,000 buffer.
- **ACC004 Cost Centre Management.** A job number is created in Xero Tracking Categories when a Zoho deal is Closed Won, and the same location is created in EH Payroll. Any state other than NSW must go to the MA because of payroll tax.
- **ACC005 AR Invoicing, v2.0.** Onsite from repeating invoices. Production from Current RMS drafts. Consulting from Qwilr quotes. Ariba for CBA and Ventia. Everything is logged in the FY27 Revenue Tracker. Card surcharge is 1.9%.
- **ACC006 Bills Entry.** Convert from a PO where there is one. Head of department approves cost of sales, GM or MD approves overhead. No attachment needed under 82.50.
- **ACC007 EH Split File.** Three tabs updated from EH after each pay run, then the PRT pivots. Feeds payroll tax.
- **ACC008 Utilisation Report.** Fortnightly data from the pay run audit Earnings Details tab. A monthly labour hours report goes to each department head.
- **ACC009 Debtor Chase.** Weekly, Wednesday or Thursday. Email past 30 days, after MA agrees. Phone past 60 days. Exceptions for Ariba, Ammacus and ON24.
- **ACC010 Bank Feeds.** Daily. Bank charges 61200, interest 82000. Liability payments go to 21410, 21420, 21430, 21440, 21500 or 21700.
- **ACC011 Salary Sacrifice.** Super and tech salary sacrifice set up in EH. Tech is sacrificed at the GST exclusive price. Tracked in the Salary Sacrifice Balance Sheet Reconciliation file.
- **ACC012 Purchase Orders.** Raised in Xero, tax exclusive, quote attached. Copied to a draft bill when the invoice arrives. Never delete lines on a partial shipment.
- **ACC013 Expense Reimbursement.** Claimed in the EH app and approved by the manager. Paid in the pay run. Journal done per ACC023.
- **ACC014 Expense Accrual.** Overhead accruals go to 2-1210. Regular accruals are KPMG, 80/20 Studdert, Energy Australia, instant asset write-off and payroll tax.
- **ACC015 AL, LSL, TIL Provision.** Monthly from the EH Leave Balances report. Accounts are 2-1230, 2-1240 and 2-1410. Reverse last month's TIL journal first.
- **ACC016 Prepayments.** 1-1510. Under 300 is recoded to expense. Monthly CSV journal import.
- **ACC017 Bonus Leave.** Added by hand in EH in the last pay run of the month, for the next month. Never paid out. Not booked in the accounts.
- **ACC018 Flights and Accommodation.** Needs an approved travel request form. Always PL card and the accounts email. Flex fares for production. No Tigerair or Jetstar. Hotels with free cancellation up to 24 hours before. Confirmations go to the traveller with their manager copied in. Any price increase goes to the head of department before booking.
- **ACC019 P&L Analysis.** Check the Xero P&L, export it with the custom Account Transactions report, then work through the tabs until every check is 0 or TRUE.
- **ACC020 Management Reports.** Due by the 10th business day of the month.
- **ACC021 WIP.** Account 11300. Three journal types: auto-reverse, reclass and non-reversing. Rates are Onsite 600 a day, Production 115 or 130 an hour, Video 125 an hour, Lead Consultant 165 an hour, Consultant 145 an hour.
- **ACC022 Payroll Tax.** NSW monthly by the 7th business day. QLD half yearly. SA, ACT, VIC and WA at year end. Journals credit 21440 and debit 65106.
- **ACC023 Payroll Reclass, v2.0.** Automated Payroll Allocation Template. Payroll and super journals are BAS Excluded. Credits: 21410 net pay, Instapay and child support. 21420 PAYG and STSL. 21430 super salary sacrifice. 21270 tech salary sacrifice.
- **ACC024 Revenue Tracker.** One FY27 tracker for all departments. Must reconcile to the Xero P&L.
- **ACC025 Superannuation.** Reconcile 21430 and pay through HeroClear.
- **ACC026 Payroll Processing, v3.0.** The current payroll procedure. Super is paid every pay run through HeroClear.
- **ACC027 New Customer Onboarding.** SME or Large Enterprise form and trade references. 30 day terms.
- **ACC028 Peninsula Referral Program.** Referral income paid to CTS by Peninsula, coded to Income Other with job 9000. Not a staff incentive.
- **ACC029 Approval Process.** Purchase requests must follow the Delegation of Authority Matrix. The matrix is a screenshot only, so the dollar limits are unknown.
- **ACC030 GoGet.** Personal use is repaid by buying the office pantry top-up. Over 100 needs MD approval.
- **ACC031 Contra Payments.** Uses account 11135. Needs MD approval in writing.
- **ACC032 Timesheets Audit.** Audit three employees after each pay run against EH, the roster and the shared calendar.

## Timesheet approval: the exact wording

Danica is processing payroll for the first time in years. She wanted wording she can quote.

ACC026, Timesheets Approval, step 1:

> "The deadline for timesheets approval is at 10:30am on the day of payroll processing. However, managers should be approving timesheets on a weekly basis."

Payroll is processed on the Monday (ACC007 and ACC008 step 1). The pay period runs Saturday to Friday, and pay day is the following Tuesday (ACC026).

ACC026, Timesheets Approval, step 2a:

> "If any timesheets are under the "Submitted" status, it will not be captured in the payrun. The Green approve button should show 0 before commencing the payroll processing."

Accounts Handbook, Payroll Processing:

> "Ensure timesheets are fully approved by HoD/GM/MD before starting the payroll process."

> "Ensure that the Accounts team do not to approve any timesheets/leave that does not fall within your approval authority (check the Org Chart if in doubt)."

Overtime needs three things, from the Handbook: flagged by the head of department at their Ops Meeting, approved in writing, and approved in EH.

**Not found anywhere:** a rule that timesheets are approved every Wednesday and Friday, or a 10am Monday deadline. Danica had heard this. She was told to get it confirmed in writing by the MD or EGM if she wants to rely on it.

## Open issues found in the documents

These are the problems to keep in mind. Raise them only when relevant.

1. **Payday Super.** ACC026 v3.0 pays super every pay run, which matches Payday Super. But ACC025, ACC023 step 9 and the Handbook still say quarterly. From general knowledge, not her documents, Payday Super started 1 July 2026.
2. **Overtime draft, step 8.** It says overtime is processed only "when job/task is confirmed and invoice is paid". From general knowledge, this may breach the Fair Work rule that pay is made at least monthly. Flagged to check with HR or Peninsula.
3. **Who approves the pay run.** ACC026 says the EGM. Its note says to remind the MD. Unresolved.
4. **WIP account.** 11300 is correct. ACC014 says 1-1310 and the Glossary says 1-11310. Both are typos. The 2024 chart has only 1-1300.
5. **The two expense procedures conflict.** The manager version says decline a claim with no attachment. The staff version says the manager can approve after raising it with the GM or MD.
6. **No dollar limits anywhere.** The Delegation of Authority Matrix is a picture only.
7. **Old tracker names.** ACC005 and ACC009 mention the "2026 Monthly Revenue Tracker". ACC024 confirms FY27 is current.
8. **ACC032** is titled "ACC024 Revenue Tracker" by mistake.
9. **Bonus leave, 5 years.** ACC017 says 5 days a year but describes a pattern that adds up to 6.
10. **Card surcharge.** The Handbook says 1.8% in one place and 1.9% in another. ACC005 says 1.9%.
11. **ACC028** gives no GST code for referral income.
12. **ACC010** lists "Precision Training, and Consulting" as the super payable payee. That looks wrong.
13. **Super salary sacrifice cap.** The Handbook says 30,000 of salary sacrifice a year. From general knowledge, 30,000 is the concessional cap and includes the employer's super.
14. **Client register and price book don't match.** The register has Ricoh. The price book has Qantas, Woolworths, JPM and APN.
15. **Sales incentives.** No policy exists in the documents. ACC026 step 4h lists "Bonuses – EOTQ, TOTQ, Management Award, Referrals" with no rules or amounts.
16. **Office phone number.** Not in any document. The office address in the Handbook is 347 Kent Street.

## Milo Rankin's Brisbane travel

**The request.** Travel form for the APA Quarterly All Hands, 4 November 2026. Cost centre 26061607. Approved by Mitch Healey. Budget: flights 500, extra baggage 150, hotel 275, transport 2 x 190, per diem 125.

**Flights: booked by Danica.** Virgin, Flex fare, Sydney to Brisbane on 3 November (VA921) and back on 4 November (VA962). The total was 518.99, which is 18.99 over budget. Mitch approved the extra. Booked through the Virgin account once Danica got access.

**Baggage: not booked.** The Flex fare includes one 23kg bag each way. Milo was told to ask Emelda (Emelda Waiya, accounts) if he needed more.

**Still unknown, check with Danica:**

- Whether the hotel was booked. The request was the ibis Styles Brisbane Elizabeth Street, Standard Queen Room, one night on 3 November, with breakfast.
- Whether the 125 per diem was added to Milo's pay.
- What the transport line of 2 x 190 was for. Mitch was never asked.
- Whether the oversize item was sorted.
- Whether the flight and hotel bills have been entered in Xero against cost centre 26061607 and closed off on the PL card.

The trip dates have now passed, so what's left is reconciliation, not booking.

## Other loose ends

- **Dashlane 2FA.** Danica can't get Dashlane 2FA emails. An email to IT was drafted. Kevin in IT helped last time. Status unknown.
- **Adobe.** Sorted.
- **Browser.** In this session, a headless browser search on Booking.com was blocked by the session's permission check. Live price lookups need her to allow it. Never go to checkout either way.

## The rest of this repository

The repository has earlier payroll, payroll tax, WIP, accruals and reporting work: folders `ptax`, `wip`, `accruals`, `reporting`, `manuals`, `checklists` and others. This session did not touch or review any of it. Check the git log and `CLAUDE.md` before working there.
