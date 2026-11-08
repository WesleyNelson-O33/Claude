# CTS Finance Revenue Risk Profile: handover

## What it is

One Excel workbook that shows management where FY27 revenue is at risk.

It answers two questions. What happens if pipeline does not come through. What happens if a big customer walks away.

It works by department, and against budget.

## The files

The workbook is CTS Finance Revenue Risk Profile.xlsx.

The script that builds it is finance_build_risk.py.

The source files are in the source folder. One is the Zoho export from 12/10/2026. The other is the FY27 budget, CTS Budget FY27 Final DN.

Everything is in the revenue-risk folder of the repo, on branch claude/keen-lamport-r5vcdn.

## Where things stand

Actuals are loaded for July and August 2026 only.

September is not loaded yet. It needs the Xero Account Transactions export.

The last actuals month on Settings is set to 31/08/2026. Change it to 30/09/2026 once September is in.

The budget is loaded. The FY27 revenue budget is 6,304,600.65.

The Zoho export has 663 deals and 133 customers.

## Each month

Step 1. Settings. Change the last month of P&L actuals to the new month end.

Step 2. P&L Dump. Clear columns A to O. Run the Xero Account Transactions report from 1 July to the month end. Paste it into cell A1, with headings in row 1. It is the same export that goes into GL_Paste in the Financial Controller Pack.

Step 3. Zoho Dump. Clear columns A to J. Run the Zoho report All Deals by Stage. Paste the whole export into cell A1, exactly as it comes out.

Step 4. Adjustments. Add any signed or likely revenue that is not in Zoho.

Step 5. Dashboard. Check the Data checks box at the bottom. Every line should say OK.

Step 6. Risk Summary. This is the page for management. It updates by itself.

## The sheets

Risk Summary is the one page for management. It has the headline tiles, the contract and non-contract chart, the budget table, the rest of year table, and the key risks in plain words.

Dashboard has the full department table, the scenario levers, the top 10 customers and the data checks.

Monthly Outlook shows each month by department. Actuals for past months, forecast for future months.

Customer Risk lists every Zoho customer with their backlog, pipeline and share.

Zoho Dump, P&L Dump, Adjustments and Budget are the four sheets you paste or type into. Their tabs are orange.

Account Map and Settings hold the mappings and the levers.

## Where each number comes from

Actual revenue. Income accounts from the Xero Account Transactions export. Credit less debit. Department comes from the Cost Centres column. If that is blank, it uses the bracket tag on the job number. This is the same method as the Financial Controller Pack.

Contract and non-contract. Contract is any account with subcategory Contract Income in the pack's account master. Everything else in Income is non-contract.

Backlog. Zoho Closed Won deals with a closing date after the last actuals month.

Pipeline. Open Zoho deals. Shown at full value, and weighted by stage probability.

Budget. FY27 Budget detailed sheet, rows 3 to 5, July to June.

## Checks done

July total income in the workbook is 459,048.71. That matches Xero exactly.

The budget totals match the budget file for all three lines.

Backlog, pipeline and top customer figures were checked against a separate count of the Zoho file. They match.

## Decisions taken from the source files

The budget has three revenue lines, not seven departments. Onsite. Production, which includes Video. Consulting, which includes Integration. This comes from the PRD/VID, ONS and CONS/INT split in the budget file.

Zoho department Support maps to Onsite. The budget summary calls the Onsite line Support.

Cost centre tags are copied from the Financial Controller Pack. CTS maps to Admin.

## Assumptions still to confirm

These are mine, not from a source. They need sign off.

Stage probabilities are placeholders. Verbal Approval 90 percent. Negotiation 70 percent. Proposal Sent 50 percent. Building Value Proposition 20 percent. Management should set them on Settings.

Risk rating thresholds on Settings are placeholders. Management should set them.

Zoho department Multi-Service maps to Unallocated.

It is not known whether Zoho amounts include GST. The P&L is ex GST.

## Open items

Item 1. Load the September Xero Account Transactions export.

Item 2. Decide how to forecast the recurring Onsite contracts. They are not in Zoho as future deals. Until they are, the Onsite forecast looks far short of budget. Either list them on Adjustments, or agree a method such as the monthly average of actuals.

Item 3. Get the stage probabilities and risk thresholds from management.

Item 4. Until items 2 and 3 are done, every department rates High. Do not show the risk ratings to management yet. The tiles, chart and budget table are fine to show.

Item 5. Fourteen open Zoho deals have no amount. Three open deals, worth about 140 thousand, are past their closing date. Ask sales to update Zoho.

## Known quirk

Two cells in the P&L Dump, A2393 and A2394, show #N/A. They come from the GL data copied out of the Financial Controller Pack. They are not formulas and do not affect any figure.

## Rebuilding the workbook

Only needed if the layout changes. Run finance_build_risk.py from the revenue-risk folder. It reads the Zoho export and the budget from the source folder, and the GL data and account master from the Financial Controller Pack. Then recalculate the workbook before opening it.
