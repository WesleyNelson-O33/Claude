# Handover: Xero review and July to September 2026 BAS

Written 8 November 2026 for the next agent.

## Read this first

The user is Danica. Read CLAUDE.md before you write anything to her. The key rules:

- She is dyslexic and listens to replies with text to speech.
- Short sentences. Lead with the answer. Headings. One level of bullets.
- No em dashes, ever.
- Don't show the working unless she asks.
- Do payroll and month end work exactly the way the recordings and last month's file did it. Where neither covers something, ask her.
- Say plainly when something comes from general rules or a search rather than her files.

## The business

- Corporate Technology Services Pty Ltd, often called CTS.
- Xero organisation ID: 3993749e-5d5f-4997-bb08-96616f256085
- The financial year runs from 1 July to 30 June. FY27 started on 1 July 2026.
- GST is lodged quarterly.

## What the Xero connector can and can't see

It can see:

- Bills and sales invoices, with line items
- The P&L and balance sheet totals
- Aged payables and aged receivables
- Unreconciled bank and card items only
- The chart of accounts, contacts and tracking categories
- Repeating bills, repeating invoices and purchase orders

It can't see:

- Manual journals
- Reconciled spend money and receive money
- Credit notes
- The GST report and the activity statement
- Payroll

For any of those, ask Danica to export the report from Xero and upload it. She has already done this once with the Journal Report.

## Where things stand

The July to September 2026 BAS has been reviewed and is ready to process. It has not been confirmed as lodged.

### Fixed by Danica in Xero

- Asteron Life Insurance bill PN81461659-4, 19 August. The GST of 110.12 was removed. Verified in Xero.
- ActivateUTS bill 108146, 5 August. GST of 263.72 was added and the total stays 2,900.90. Verified in Xero.
- The Old Clare Hotel soft drinks line now carries GST. It shows in journal 71329.
- ATO payment of 169,885.00 on 24 August. Danica says she recoded it from 21420 PAYG Withholding Payable to 21500 BAS Transactions. **Not verified.** The Journal Report she uploaded was exported before the fix.

### Confirmed by Danica, leave as is

- The ON24, Inc. invoices INV-10552 and INV-10589 stay with GST. Every ON24 invoice since February 2023 has charged GST, and the work is done in Australia for Australian events.
- Brookvale travel insurance INV1914759. The 222.50 premium line with no GST is correct.
- All alcohol lines on venue bills have no GST. This is the house rule. It covers the five CTS Day After Party bills on 1 August, Hotel CBD INV8054807, Woodcut Restaurant INV806106 and The Flames Galore INV27092026.

### Small items, left on purpose

Each is a few cents or a few dollars of GST:

- Resolution X I26-1570, card surcharge 31.57, no GST
- K&F Concept INV15-403950, 29.99, no GST
- Three Expedia discount lines with no GST: INV73530202683410, ITN73535242035475 and ITN73550041732636
- Four The AVIT Group bills a few cents out
- Diversified Communications INTSEC26-213, a 5.19 fee with no GST

### Optional tidy ups, not BAS issues

- Asteron PN81461659-4 is still coded to 62300 Other Insurance. The other two Asteron bills are on 11510 Prepayments.
- First Focus IT INV167160 has interest of 198.58 coded to 64700 IT managed services.

## What's left to do on the BAS

1. **Check the ATO payment fix.** Ask Danica for a new Journal Report covering at least 1 to 31 August. Check that journal 69661, or whatever replaced it, now debits 21500 and not 21420. Run `bas/journal_detail.py` against it.
2. **She should run the GST Reconciliation report in Xero** for 1 July to 30 September 2026 before lodging.
3. **Post the BAS journal.** Copy last quarter's manual journal "BAS Apr-Jun 2026 - Manual", journal ID 68727, dated 30 June 2026. Use the same layout and date the new one 30 September 2026. Last quarter's version credits 21500 BAS Transactions with the total, debits 21310 GST with GST collected, credits 21310 GST with GST paid, and debits 21420 PAYG Withholding Payable with the PAYG amount. Take the figures from the activity statement, not from this note.
4. **Check the lodgement date.** The standard ATO due date for this quarter was 28 October 2026. Agent lodgement is often later. That is a general ATO rule, not from her files. Ask whether it has been lodged.

PAYG withholding was not checked. It comes from payroll.

### GST totals from the scan

These are rough figures, from bills and invoices only, July to September. They leave out spend money, receive money, journals and card transactions. **Do not use them for lodging.** Use the Xero reports instead.

- GST on bills: about 104,663.65. Part of this was re-keyed by hand, so it needs a cross-check.
- GST on sales invoices: about 160,560.95.

## Other items from the first Xero review, 19 October

None of these change the BAS. Their status is unknown unless shown otherwise.

1. **Card and direct debit bills still showing unpaid.** As of 19 October: Employment Hero 2,248.84, Optus 639.15, Goget 629.52, Energy Australia 696.25, Flatworld 92.21 and Ricoh 210.20. Each needs matching to its card or bank payment.
2. **October unreconciled items.** Woolworths 92.30 on the PL Credit Card, 3 October. A 10,000 transfer to the PL Credit Card on 7 October. A batch payment of 112,761.41 from the Cheque Account on 16 October.
3. **Meet Magic INV-10464, 5,075.07.** It was due on 26 June and has not been paid.
4. **Overdue debtors.** Microsoft INV-10554 for 13,845.70, Generation-e INV-10532 for 11,850.16 and ON24 INV-10589 for 3,005.75. ON24 INV-10552 has since been paid.
5. **New Brookvale insurance policies from 30 September.** The premiums are in 11510 Prepayments. The monthly release needs to start in October. The broker fees and stamp duty were expensed in September, so check that against last year's renewals.
6. **Appspace for Aware Super.** This looks handled. The Journal Report shows manual reclass journals for it, IDs 71179, 71180, 71190 and 71192.
7. **Contact Group INV457215.** The title says DTTL Hobart, but the line text says Deloitte Bathurst St. Check the job.
8. **Annual leave expense was negative for the quarter**, at minus 6,720.59. Worth a check against the leave journals.
9. **The 500,000 non-current asset increase is resolved.** It was a transfer to CTS Holdings, account 12600, on 14 July.
10. **Invoice reference typos.** Garvan INV-10680 says 2025 instead of 2026. Noisebox INV-10681 has a date of 13.10.16.

## Files

- `bas/parse_journal_report.py` turns a Xero Journal Report xlsx into a pickle. Usage: `python3 -I bas/parse_journal_report.py <report.xlsx> <out.pkl>`.
- `bas/journal_gst_check.py` lists the July to September manual journals that touch GST, spend money and receive money with their GST rate, and credit notes. Usage: `python3 -I bas/journal_gst_check.py <out.pkl>`.
- `bas/journal_detail.py` prints chosen journal IDs, whether each is live or reversed, and the monthly movement on 21310 and 21420. Edit the ID list at the top before you run it.

The uploaded Journal Report was not committed. It covers 1 January to 31 October 2026 and is now out of date anyway. Ask Danica for a fresh export.

The Journal Report includes every edit. An edited document shows up three times: the original, a "Reversed:" copy and the new version. The scripts handle this by matching "Reversal of ID" titles.

## Earlier work in this repo

The September 2026 month end work is in the other folders. That covers WIP, accruals, leave, payroll tax, staff expenses and the tracker. The git log describes each step. CLAUDE.md sets out the approved layout for adjustment journal schedules.
