"""Narrative rewrite: maps the build script note text to the house voice. Applied at the end of build_cashflow.py."""
VOICE = {
'Blue = input you can change. Black = formula. Bright yellow = key assumption to confirm. All $ AUD, GST inclusive (cash actually moving).':
"Blue = change it if you need to. Black = formula, leave it alone. Bright yellow = the ones I still want to double check. Everything is GST inclusive because that's what actually hits the bank.",
'Mon 28 Sep 2026. Five weeks: Wk1 28 Sep-4 Oct, Wk2 5-11 Oct, Wk3 12-18 Oct, Wk4 19-25 Oct, Wk5 26 Oct-1 Nov.':
"Mon 28 Sep. Week 1 runs through to Wed 7 Oct so it picks up the pay run and payroll tax, then it's weekly from there (Thu to Wed). Week table is at the bottom of this tab.",
'Xero cash position 28 Sep 2026 (cheque, savings and credit cards combined). Payroll is paid from the cheque account with top-ups from savings - check the cheque balance separately before each pay run.':
"Cheque $25,155 + savings $410,000 at 28 Sep. Not counting the other accounts - they're not there for day to day. Payroll comes out of cheque, topped up from savings, so check the cheque balance before every pay run.",
'Cash Flow tab has Forecast and Actual columns for each week.':
'Each week has a Forecast and an Actual column on the Cash Flow tab.',
'Average of the last three known pay runs: 11 Aug $108,325, 25 Aug $115,340 (GL journals #69396, #69669) and the two September runs $217,062 combined (Xero cash P&L 1-28 Sep). See Payroll tab.':
'Average of the last three runs we know: 11 Aug $108,325, 25 Aug $115,340 (GL journals #69396, #69669) and the two Sep runs $217,062 between them (Xero cash P&L to 28 Sep). Detail on the Payroll tab.',
'25 Aug run: gross ~$115k, ABA file ~$88-89k (Payroll Part 6 transcript, 03:16-04:12) = 76.8%. Balance is PAYG withholding and salary-sacrifice deductions.':
'25 Aug run - gross $115.3k, ABA file $88.5k = 76.8% (Part 6 video, 03:16-04:12). The gap is PAYG and sal sac.',
'Formula: 1 - net %. Paid to the ATO via IAS/BAS, not on payday.':
'1 less net %. Goes to the ATO with the IAS/BAS, not on pay day.',
'Superannuation Guarantee 12%. 25 Aug run: SG $13,677 on gross $115,340 = 11.9%. Salary-sacrifice super (~$1,050 a run) is in the same super batch but comes out of gross pay, so it is not added here.':
"SG is 12%. 25 Aug: $13,677 SG on $115,340 gross. Sal sac super (~$1k a run) goes out in the same super batch but comes off gross, so it's not added on here.",
'PLACEHOLDER - not visible in Xero through the connector. Per diems ($75/night) are paid inside payroll and are already in net pay. Enter the typical expense-claim amount paid with each pay run.':
"Placeholder - I can't see expense claims in Xero. Per diems ($75 a night) go through payroll so they're already in net pay. Put the usual expense claim amount in here.",
'Pay date is always the Tuesday of the processing week (Payroll Part 2 transcript; Checklist step 13). Mon 5 Oct is the Labour Day public holiday (NSW), so run 1 must be prepared, finalised and uploaded on Fri 2 Oct. Super is direct-debited the same week (Checklist step 69); Payday Super applies from 1 Jul 2026.':
'Always the Tuesday of processing week (Part 2 video, Checklist step 13). Mon 5 Oct is a PH so run 1 has to be done and uploaded on Fri 2 Oct. Super direct debits the same week (step 69) - payday super since 1 Jul 26.',
'Tue 20 Oct.':
'Tue 20 Oct.',
'Tue 3 Nov. Only hits the forecast if the week table runs that far.':
'Tue 3 Nov (FE 30 Oct). Only hits the forecast if the week table runs that far - it does at the moment.',
'August actual accrual $6,120.40 + $174.67 (GL journals #69671, #70351). NSW monthly return due the 7th of the following month. Replace with the lodged figure.':
"Sep return isn't lodged yet so I've used Aug: $6,120.40 + $174.67 (GL #69671, #70351). Due the 7th of the following month. Update once it's lodged.",
'August actual accrual $2,700 (GL #69670; July $3,000). All due on the 7th of the following month.':
'Aug actual $2,700 (GL #69670), July was $3,000. All due on the 7th too.',
'Wed 7 Oct 2026.':
'Wed 7 Oct.',
'Formula: PAYG % x September gross wages $217,062 (Xero cash P&L). July and August withholding were paid on the monthly IAS (21 Aug, 21 Sep), so no IAS falls in October.':
'PAYG % x Sep gross $217,062 (Xero cash P&L). Jul and Aug went on the IAS (21 Aug, 21 Sep) so nothing falls in Oct.',
'Rough: GST on ~$1.5m quarterly receipts less GST on ~$1.0m purchases. Replace with the Xero Activity Statement figure.':
'Rough guess - GST on ~$1.5m in, less GST on ~$1.0m out. Replace with the activity statement figure from Xero.',
'Unknown - not visible through the Xero connector. Enter the ATO instalment if one applies.':
"Can't see this in Xero. Put the ATO instalment in here if there is one.",
'COSTS NOT YET BILLED - MONTHLY ESTIMATES (spread evenly over the 5 weeks)':
'COSTS NOT BILLED YET - MONTHLY ESTIMATES (spread over the 5 weeks)',
'September cash P&L: Sub-Contract Labour INTEGRATION $31,990 + PRD $7,154. Bills for October work paid inside the month.':
'Sep cash P&L: Integration $31,990 + PRD $7,154. Oct work that gets billed and paid inside the month.',
'September cash P&L: Service hires $17,344 + Equipment hires $11,427.':
'Sep cash P&L: service hires $17,344 + equipment hires $11,427.',
'Beyond the Complete AV / Crestron / Fredon bills already listed. Budget flags October as a heavy Production month ($587k income), so this could run higher.':
'On top of the Complete AV / Crestron / Fredon bills already listed. Oct is a big Production month in the budget ($587k income) so this could easily be higher.',
'September cash P&L: Direct Travel ONS $4,120 + PRD $1,572 + freight ~$1,300.':
'Sep: ONS travel $4,120 + PRD $1,572 + ~$1,300 freight.',
'September cash P&L: Subscriptions & Licences expense $3,573 beyond the itemised Appspace/Spacera/Eptura bills.':
'Sep: $3,573 outside the Appspace / Spacera / Eptura bills.',
'September cash P&L: Recruitment $11,742 (elevated by a casual concierge campaign). Half that assumed.':
"Sep was $11,742 with the casual concierge campaign. I've assumed half that.",
'Jun-Aug cash P&L: Accounting fees $7,733 and Legal $2,900 over three months.':
'Jun-Aug: accounting $7,733 + legal $2,900 over 3 months.',
'Jun-Aug cash P&L: Other insurance $5,073 + workers comp $2,336 over three months (~$2.5k/month).':
'Jun-Aug: other insurance $5,073 + workers comp $2,336 over 3 months, call it $2.5k a month.',
'Jun-Aug cash P&L: staff amenities, office supplies, staff entertainment ~$4k/month; September ran lower.':
'Jun-Aug ran ~$4k a month across amenities, supplies and staff entertainment. Sep was lower.',
'September cash P&L: Bank charges $2,311 + Stripe fees $271 (Sep included a one-off).':
'Sep: bank charges $2,311 + Stripe $271, but Sep had a one-off in it.',
"September transfers from cheque to cards: $15,000 + $4,224.75 (PL) and $5,718.88 (DL) = $24,944 (Xero bank transactions). Covers every bill marked 'paid via CC', which are excluded from the Payables totals to avoid double counting.":
'Sep transfers to the cards: $15,000 + $4,224.75 (PL) + $5,718.88 (DL) = $24,944. Covers every bill marked paid via CC - those are left out of the Payables totals so nothing gets counted twice.',
"September evidence: invoices dated and paid in Sep (Ubank $10,867, RSNSW $3,587, CSIRO $1,971, Garvan, Elkiem, Bankwest event, St Vincent's) = ~$20k. Spread over weeks 3-5.":
"Sep: invoices raised and paid in the same month (Ubank $10,867, RSNSW $3,587, CSIRO $1,971, Garvan, Elkiem, Bankwest event, St Vincent's) came to ~$20k. Spread over weeks 3-5.",
'September $648.92 (Xero cash P&L). Lands at month end (week 5).':
'Sep was $648.92. Lands at month end.',
'Fortnightly pay cycle. Pay date = Tuesday of the processing week. Super is direct-debited the same week. PAYG withholding is paid later via IAS/BAS.':
'Fortnightly. Pay date is the Tuesday of processing week. Super direct debits the same week. PAYG goes later with the IAS/BAS.',
'Fortnight ends Fri 2 Oct. Mon 5 Oct is the Labour Day public holiday (NSW), so do the full process on Fri 2 Oct: chase timesheet approvals first thing, reconcile, finalise with Graham, upload the ABA to CommBiz and send the payroll email for Duncan to approve. Transfer the payroll funds from savings on the Friday. Paid Tue 6 Oct.':
'FE Fri 2 Oct. Mon 5 Oct is a PH so the whole thing gets done on Fri 2 Oct - chase approvals first thing, reconcile, finalise with Graham, ABA into CommBiz, payroll email to Duncan. Move the funds from savings on the Friday. Paid Tue 6 Oct.',
'Last pay run of the month: accrue bonus leave (Checklist step 48) - no cash effect.':
'Last run of the month - accrue bonus leave (Checklist step 48). No cash effect.',
'Paid Tue 3 Nov. Only counts if the forecast window covers 3 Nov.':
'FE 30 Oct, paid Tue 3 Nov. Only counts if the week table runs to 3 Nov.',
'Payroll tax for October wages is due 7 Nov - outside this forecast.':
'Payroll tax on Oct wages is due 7 Nov - after this forecast.',
'Xero cash-basis P&L 1-28 Sep 2026: all Salaries & Wages accounts; Direct + Indirect Superannuation':
'Xero cash P&L 1-28 Sep: all salaries & wages accounts; direct + indirect super',
'GL journals #69669 / #69663 (the run in the training videos)':
'GL journals #69669 / #69663 (the run in the training videos)',
'Aug runs plus the two Sep runs, over four fortnights':
'Aug runs plus the two Sep runs, over four fortnights',
'SG is 12%. The batch ratio above is higher because the GL super accounts include salary-sacrifice super (25 Aug: $13,677 SG + $1,052 salary sacrifice). The forecast uses 12%; the salary-sacrifice portion is deducted from gross pay so it is not double counted.':
"SG is 12%. The batch % above is higher because the GL super accounts include sal sac super (25 Aug: $13,677 SG + $1,052 sal sac). Forecast uses 12% - the sal sac part comes off gross so it isn't counted twice.",
'FY27 budget has payroll tax at ~$7.4k/month; actuals are running ~$9k/month, so actuals are used.':
"Budget has payroll tax at ~$7.4k a month but actuals are running ~$9k, so I've used actuals.",
"Source: Xero aged receivables 28 Sep 2026 ($580,292.39, 38 invoices) plus September contract invoices still to be raised (Xero repeating templates). Expected date (blue) is a judgement from each customer's recent paying pattern - change it and the forecast moves.":
'Xero aged receivables at 28 Sep ($580,292.39, 38 invoices) plus the Sep contract invoices we still have to raise (from the repeating templates). Expected date (blue) is my call based on how each customer has been paying - change it and the forecast moves.',
'APA paid its last invoice the day after issue':
'APA paid the last one the day after we sent it',
'On due date':
'On the due date',
'ASX paid its Aug invoices ~4 days late':
'ASX paid the Aug ones ~4 days late',
'Due date + a day':
'Due date plus a day',
'Small, overdue - chase this week':
'Small and overdue - chase this week',
'Aware pays on or near due date':
'Aware pay on the due date',
'KEY RECEIPT. Complete AV materials payment ($136,163) is released the week after this lands. If Aware slips, hold Complete AV.':
'THE BIG ONE. Complete AV materials ($136,163) is going out this week ahead of it - chase Aware so it lands on time.',
'Due Sun 18 Oct - expect Fri 16 Oct. Funds the Appspace bill ($20,985) due 24 Oct.':
'Due Sun 18 Oct so Fri 16 Oct. Covers the Appspace bill ($20,985) due 24 Oct.',
'Due Sat 24 Oct - expect Fri 23 Oct':
'Due Sat 24 Oct - call it Fri 23 Oct',
'Due Sun 25 Oct - expect Fri 23 Oct':
'Due Sun 25 Oct - call it Fri 23 Oct',
'4 days overdue - assume paid this week':
'4 days overdue - should land this week',
'Due 30 Sep; allow a week':
'Due 30 Sep, give it a week',
'A month overdue - chase. Assume week 2':
'A month overdue - chase. Week 2.',
'Slow payer - assume 3 weeks after due':
'Slow payers - 3 weeks after due',
'Staff recovery - likely via payroll':
'Staff recovery - probably through payroll',
'DOUBTFUL - 3 months overdue. Excluded from October. Escalate.':
'DOUBTFUL - 3 months overdue. Left out of Oct. Needs escalating.',
'A month overdue - usually a PO/portal issue. Assume week 2':
'A month overdue - usually a PO / portal issue. Week 2.',
'A month overdue - chase':
'A month overdue - chase',
'Assume paid with the July one':
'Assume it comes with the July one',
'DOUBTFUL - almost 4 months overdue. Excluded from October. Escalate.':
'DOUBTFUL - nearly 4 months overdue. Left out of Oct. Needs escalating.',
'A month overdue. PwC paid INV-10577 on its due date, so these look like PO problems - chase':
'A month overdue. PwC paid INV-10577 on the day, so these two look like PO issues - chase',
'As above':
'Same',
'Allow two weeks after due':
'Give it two weeks after due',
'Allow a week':
'Give it a week',
'8 days overdue - assume this week':
'8 days overdue - this week',
'LARGEST OVERDUE - a month late. Chase now; assume week 2':
'BIGGEST OVERDUE - a month late. Chase now. Week 2.',
'Ventia is paying ~3 weeks late':
'Ventia are running ~3 weeks late',
'Two weeks after due':
'Two weeks after due',
'2 days overdue - assume this week':
'2 days overdue - this week',
'Allow a few days after due':
'A few days after due',
'Aug invoice (31 Aug) paid 14 Sep - Deloitte pays ~2 weeks after issue. Repeating template $113,987.99/month':
'Aug invoice (31 Aug) paid 14 Sep - Deloitte pay ~2 weeks after we invoice. Template $113,987.99 a month',
'Aug invoices (31 Aug) paid 8 Sep - CBA pays ~8 days after issue':
'Aug invoices (31 Aug) paid 8 Sep - CBA pay ~8 days after invoice',
'Aug invoice (31 Aug) paid 3 Sep - Bankwest pays within days':
'Aug invoice (31 Aug) paid 3 Sep - Bankwest pay within days',
'As above (template amount; actual varies)':
'Same (template amount, actual varies)',
"PwC's Jul contract invoices are still unpaid - assume November":
"PwC's July contract invoices are still unpaid - assume Nov",
'Aware pays on due date':
'Aware pay on the due date',
'of which open in Xero at 28 Sep (should equal $580,292.39)':
'of which open in Xero at 28 Sep (should be $580,292.39)',
'of which expected inside the forecast window':
'of which expected inside the forecast',
'of which expected after 1 Nov (PwC Sep invoices, Meet Magic, Property NSW)':
'of which pushed past the forecast (PwC Sep invoices, Meet Magic, Property NSW)',
"Source: Xero aged payables 28 Sep 2026 ($283,042.53, 24 bills; Complete AV split per the 28 Sep meeting) and Xero repeating-bill templates. 'Card' rows are paid on the PL/DL credit cards and are covered by the credit-card settlement line on the Cash Flow tab, so they are excluded from the category totals. The P&L category drives which Cash Flow line each bill lands on.":
"Xero aged payables at 28 Sep ($283,042.53, 24 bills - Complete AV split per the 28 Sep meeting) plus the repeating bill templates for Oct. Card rows get paid on the PL/DL cards and sit in the credit card line on the Cash Flow tab, so they're left out of the category totals. The P&L category decides which line each bill lands on.",
"Pay once Aware's matching invoice INV-10635 ($23,317) is received (expected 16 Oct)":
"Pay once Aware's Appspace invoice INV-10635 ($23,317) is in (expected 16 Oct)",
'Overdue - pay this week':
'Overdue - pay this week',
'Direct debit via PL credit card':
'DD via the PL card',
'Earlier Expedia bookings went on the PL card':
'Went on the PL card last time',
"Overdue; bill note 'DH to confirm'. Assume released week 2":
"Overdue - DH to confirm. Assume it's released week 2",
'On due date. October rent will be invoiced ~22 Oct and due ~22 Nov':
'On the due date. Oct rent gets invoiced ~22 Oct, due ~22 Nov',
'BPAY on due date':
'BPAY on the due date',
'Bill note: GC advised hold until he confirms. Assumed released week 4':
'GC said hold until he confirms. Assumed released week 4',
'Bill note: check with DL before payment':
'Check with DL before paying',
'Earlier Virgin bookings went on the PL card':
'Went on the PL card last time',
'Refund to PL card':
'Refund back to the PL card',
'Repeating bill template, direct debit':
'Template, DD',
'Repeating bill template':
'Template',
'Template $3,850/month (Sep cash P&L shows $5,000 business advisory). Date assumed':
'Template $3,850 a month (Sep cash P&L had $5,000). Date assumed',
'Repeating bill, next due 31 Oct':
'Template, due 31 Oct',
'Repeating bill, due 14 Oct':
'Template, due 14 Oct',
'Repeating bill, due 5 Oct':
'Template, due 5 Oct',
'Repeating bill every 3 months, due 8 Oct':
'Quarterly, due 8 Oct',
'PL card - in credit card settlement line':
'PL card - in the credit card line',
'PL/DL cards - approx':
'PL / DL cards - approx',
'Invoiced ~22 Oct, due ~22 Nov - AFTER this forecast':
'Invoiced ~22 Oct, due ~22 Nov - after this forecast',
'Due 8 Nov - AFTER this forecast':
'Due 8 Nov - after this forecast',
'TOTAL open bills in Xero (should equal $283,042.53)':
'TOTAL open bills in Xero (should be $283,042.53)',
'of which paid by card (excluded - see credit card line)':
'of which paid by card (left out - see the credit card line)',
'Recurring templates (non-card) falling inside the forecast':
'Templates (non-card) inside the forecast',
'Non-card bills dated after 1 Nov (Complete AV labour hold, Oct rent, Oct mobiles)':
'Non-card bills dated past the forecast (Complete AV labour hold, Oct rent, Oct mobiles)',
"How to use: forecast columns are formulas driven by the Receivables, Payables, Payroll and Assumptions tabs. Each week, type the actual cash movements into the yellow Actual column, enter the closing bank balance from CommBiz, and set 'Actuals entered' to 1. The next week's opening balance then rebases to the real bank balance, and the difference row shows anything still unexplained.":
"How this works: the Forecast columns are formulas off the Receivables, Payables, Payroll and Assumptions tabs. Each week, put the actual movements in the yellow Actual column, type in the closing bank balance from CommBiz, and set 'Actuals entered' to 1. The next week then starts from the real bank balance and the difference row shows anything that's been missed.",
"Set to 1 once the week's actuals and bank balance are in. Drives the opening balance of the following week and the variance column.":
"Set to 1 once the week's actuals and bank balance are in. This drives next week's opening balance and the variance column.",
"Week 1 = Xero cash position 28 Sep (all accounts). Later weeks = prior week's actual bank balance once actuals are entered, otherwise the prior forecast closing.":
"Week 1 = cheque + savings at 28 Sep. After that it's last week's actual bank balance once the flag is set, otherwise last week's forecast closing.",
'Receivables tab by expected receipt date (open invoices, overdue invoices and the September contract invoices to be raised), plus the October quick-pay allowance from the Assumptions tab in weeks 3-5.':
'Receivables tab by expected date - everything open, the overdue ones, and the Sep contract invoices we still have to raise - plus the $20k quick-pay allowance in weeks 3-5.',
'Assumptions tab. Savings interest at month end.':
'Savings interest, month end.',
'Payroll tab: gross x net %. Pay runs Tue 6 Oct and Tue 20 Oct (ABA file via CommBiz).':
'Payroll tab - gross x net %. Pay runs Tue 6 Oct, Tue 20 Oct and Tue 3 Nov (ABA through CommBiz).',
'Payroll tab: gross x 12%. Direct-debited by Employment Hero in the pay-run week.':
'Gross x 12%. EH direct debits it in the pay run week.',
"No IAS falls in October: July and August withholding were paid 21 Aug and 21 Sep; September's goes on the Q1 BAS (see BAS line).":
'Nothing in Oct - Jul and Aug were paid 21 Aug and 21 Sep, Sep goes on the Q1 BAS (see the BAS line).',
'Assumptions tab. September wages, due 7 Oct. August actuals used as proxy ($6,295 NSW + $2,700 other states).':
'Sep wages, due 7 Oct. Aug numbers used until the Sep return is lodged ($6,295 NSW + $2,700 other states).',
'Assumptions tab placeholder per pay run - replace with the real expense-claim run.':
'Placeholder per pay run - swap for the real expense claim run.',
'Payables tab (Complete AV materials $136,163 wk 3, Fredon $24,708 wk 2, Crestron $7,103 wk 4) plus the unbilled estimate from the Assumptions tab. Complete AV labour $41,956 is held past 1 Nov.':
'Payables tab - Complete AV materials $136,163 (this week, see the Complete AV tab), Fredon $24,708, Crestron $7,103 - plus the unbilled estimate. Complete AV labour $41,956 is held past the end of the forecast.',
'Assumptions tab estimate (Sep cash P&L ~$39k), spread evenly.':
'Estimate (Sep was ~$39k), spread evenly.',
'Assumptions tab estimate (Sep cash P&L ~$29k), spread evenly.':
'Estimate (Sep was ~$29k), spread evenly.',
'Payables tab (Appspace, Spacera, Eptura, LiveU) plus unbilled estimate.':
'Payables tab (Appspace, Spacera, Eptura, LiveU) plus the unbilled estimate.',
'Payables tab (non-card) plus Assumptions estimate. Card-paid flights/accommodation sit in the credit card line.':
'Payables tab (non-card) plus the estimate. Flights and accommodation on the card sit in the credit card line.',
'Payables tab: Investa September rent $29,281 due 22 Oct. October rent falls in November.':
'Investa Sep rent $29,281 due 22 Oct. Oct rent falls in Nov.',
'Payables tab: First Focus CORE + Azure direct debit ~16 Oct.':
'First Focus CORE + Azure, DD around 16 Oct.',
'Payables tab: RJW website balance $6,600 (wk 1) and RJW digital marketing $4,675 (wk 4).':
'RJW website balance $6,600 (check with DL first) and RJW digital marketing $4,675.',
'Payables tab (8020 Advisors retainer, Employsure) plus Assumptions estimate for accounting/legal.':
'8020 Advisors retainer and Employsure, plus the estimate for accounting / legal.',
'Payables tab non-card items (CEO Institute quarterly). Card subscriptions sit in the credit card line.':
'Non-card only (CEO Institute quarterly). Card subscriptions are in the credit card line.',
'Payables tab: Optus mobiles BPAY, Telstra Nighthawk. Optus internet is on the card.':
'Optus mobiles BPAY, Telstra Nighthawk. Optus internet is on the card.',
'Assumptions estimate; Seek ads are on the card.':
'Estimate. Seek ads go on the card.',
'Assumptions estimate (~$2.5k/month across business insurance and workers comp).':
'Estimate, ~$2.5k a month.',
'Payables tab (Ricoh printing DD, Posh plants) plus Assumptions estimate.':
'Ricoh printing DD, Posh plants, plus the estimate.',
'Energy Australia is direct-debited to the PL card, so it sits in the credit card line (shows nil here).':
'Energy Australia DDs the PL card so it shows in the credit card line, not here.',
'Assumptions estimate (bank charges, Stripe fees).':
'Estimate - bank charges and Stripe.',
"Assumptions tab: ~$25k/month transferred to the PL and DL cards. Covers every 'Card' row on the Payables tab (subscriptions, travel, electricity, Seek, Goget, Employment Hero).":
'~$25k a month transferred to the PL and DL cards. Covers every Card row on the Payables tab (subscriptions, travel, electricity, Seek, Goget, EH).',
'Payables tab: Persona Health (on hold until GC confirms).':
'Persona Health - on hold until GC confirms.',
"Opening + net cash flow. Actual column fills in once the week's flag is set to 1; the 5-week Actual total shows the latest completed week.":
"Opening + net. The Actual column fills in once the week's flag is 1, and the total column shows the latest completed week.",
'Type the real closing balance (all accounts) at the end of each week.':
'Type the real closing balance (cheque + savings) in at the end of each week.',
'Should be nil once every movement is entered. A balance here means a receipt or payment is missing from the Actual column.':
"Should be nil once everything's in. If there's a number here, a receipt or payment is missing from the Actual column.",
'Checklist steps 63-65: transfer from savings before the ABA upload, leaving ~$20k buffer.':
'Checklist 63-65 - transfer from savings before the ABA upload, leave ~$20k buffer.',
'Weeks of net payroll + super the closing balance would cover with no further receipts.':
'Weeks of net pay + super the closing balance would cover if nothing else came in.',
'PwC September contract invoices, Meet Magic, Property NSW.':
'PwC Sep contract invoices, Meet Magic, Property NSW.',
'Complete AV labour & programming $41,956 (held for Aware sign-off), October rent, October mobiles.':
'Complete AV labour $41,956 (held for Aware sign off), Oct rent, Oct mobiles.',
'Basis / source':
'Notes',
'Basis for expected date':
'Why that date',
}
