# Handover: Creditors Run Review, 23 October 2026

## Read this first

This file hands over a creditors run review to a new agent.

The review was done on 23 October 2026. Today is later than that. Bills may have been paid, edited or added since. Re-pull the bills from Xero before acting on anything below.

Nothing in this file has been confirmed as fixed or done by Danica. Treat every open item as still open until she says otherwise.

Follow CLAUDE.md in this repo for every reply to Danica. The main points are short sentences, lead with the answer, no em dashes, no nested bullets, no shown calculations, and say plainly when something is from a web search or cannot be verified.

## What Danica asked for

She asked for a check of the creditors run for 23 October 2026. She wanted to know four things.

1. Are all invoices processed accurately?
2. Do the banking details on each invoice match the supplier contact in Xero?
3. Is the GST in Xero accurate per the invoice PDF attached to each bill?
4. Is the ABN shown on each invoice?

## The short answer given to her

Items 2, 3 and 4 could not be checked. The Xero connector cannot open the PDFs attached to bills. It also does not return bank details or ABNs on contacts. It only returns contact ID, name and email.

Item 1 was checked as far as Xero data allows. GST maths, coding, comparison to past bills from the same supplier, and matching to purchase orders.

Danica was told to do the bank detail, GST to PDF, and ABN checks herself in Xero.

## Where things are

Xero organisation is Corporate Technology Services Pty Ltd. Organisation ID is 3993749e-5d5f-4997-bb08-96616f256085. It is the only connected organisation.

There is no recording, manual or past workings file for a creditors run in this repo. A search of the transcripts found nothing on how the run is done. So there is no source process to check against. If Danica supplies one, check the run against it.

No files were created or changed in the repo for this review, apart from this handover.

## How the bills were pulled

Unpaid bills, issued from 1 July 2026, due on or before 23 October 2026, with line items. That returned 11 bills.

History for comparison was pulled for each supplier from 1 July 2025 to 22 October 2026. Purchase orders from 1 September 2026 were also pulled. The chart of accounts was pulled to check account names and default tax types.

## The 11 bills in the run

Each line gives supplier, total, GST, due date, Xero bill ID, and how it is paid if the narration says so.

Avetta. 892.00. GST nil. Due 23 Oct. Bill ID a29bb0c1-1fe7-4c0c-948a-0e4cc9b56dcd. Tagged paid via PL CC.

Energy Australia. 696.25. GST 63.30. Due 14 Oct. Bill ID 9056b342-b238-4e81-bb4b-85e4269cfc0d. Tagged dd via PL CC.

Employsure Pty Ltd. 880.00. GST 80.00. Due 5 Oct. Bill ID 7e7124d8-a387-4a58-bf10-ba2fff1fd174. Repeating bill, 42 of 60. No payment tag.

Optus Billing Services. 639.15. GST 58.11. Due 23 Oct. Bill ID 97149de5-6c51-4cc1-b928-98d1c7ab37d3. Tagged paid via PL CC.

Seek Limited. 555.50. GST 50.50. Due 21 Oct. Bill ID fd3067fb-d69b-432f-8740-b9abb883b2d3. Tagged paid via PL CC.

Mailchimp. 38.54. GST nil. Due 18 Oct. Bill ID 48365c95-abbe-4e6d-b8be-231586b11350. Tagged paid via PL CC.

Virgin Airlines. 716.00. GST 65.09. Due 20 Oct. Bill ID 52b31283-fec0-4600-b72a-8f87a63bad9a. MH App. No payment tag.

Expedia Inc. 233.83. GST 21.26. Due 20 Oct. Bill ID 4f79b891-0c82-4d9c-b547-68c2a036db44. MH App. No payment tag.

LinkedIn Singapore. 6,996.00. GST 636.00. Due 22 Oct. Bill ID e748e4fa-925e-4d0a-9456-84abe1110a66. DL App. No payment tag.

The CEO Institute. 2,310.00. GST 210.00. Due 8 Oct. Bill ID 04896162-e8d5-40a6-a60e-6c9f54bbed80. DL App. No payment tag.

Audio Visual Events. 3,958.63. GST 359.88. Due 22 Oct. Bill ID 07d39afe-0c06-485b-ae09-410fd566f0fc. Linked to PO-00016948. No payment tag.

## Findings: flagged for Danica to check

### Audio Visual Events, 3,958.63

The amount matches PO-00016948 exactly, including GST.

The description does not match the PO. The PO reference says "QX26-1295, CTS, Lexi Captioning, 23rd Oct". The bill narration says "Equipment hires for 2603105 Cochlear AGM 26.10.2026". The job number on the bill may be wrong. She needs to check the invoice to confirm which job it belongs to.

The coding is unusual for this supplier. 640 is on 51100 Equipment Purchase, against a line called "Sales". 285 is on 51315 Sub-Contract Labour PRD, against "Pre/Post Production". Past bills from this supplier are almost always only 51200 Equipment Hires and 51400 Service Hires.

### Energy Australia, 696.25

This was due on 14 October and was still unpaid in Xero on 23 October. It is tagged as a direct debit on the PL credit card. Past Energy Australia bills are all marked paid. The likely cause is the card payment has gone through but has not been reconciled in Xero. This is not confirmed. She was told not to include it in a bank payment and to check the card feed first.

### Employsure, 880.00

There is no invoice attached. Every Employsure bill back to July 2025 has the same amount, coding and GST, and none have an attachment. So it looks consistent, but there is nothing to check GST or ABN against.

## Findings: looked right on Xero data

LinkedIn Singapore matches last year's bill exactly. That was "LinkedIn Job Slots (Contract Year 2/3)", dated 1 Sep 2025, 6,996.00 with 636.00 GST, coded to 11510 Prepayments, quantity 3 at 2,120.

The CEO Institute matches past quarterly bills. 2,100 plus GST, coded to 11510 Prepayments. It is overdue.

Avetta matches the bill of 11 November 2025 exactly. Same three lines, 892.00 total, no GST both times, coded to 11510 Prepayments. GST nil still needs checking against the PDF.

Optus has the same line layout as September 2026. Line GST adds up to the bill GST.

Seek, Virgin and Expedia have GST at one eleventh of the total, or 10 percent on top for Seek. Coding matches their past bills.

Mailchimp has no GST, which matches every past month back to July 2025.

## Which bills belong in the bank run

Five bills are tagged as paid on the PL credit card. They should stay out of the bank batch. These are Avetta, Energy Australia, Optus, Seek and Mailchimp.

Virgin and Expedia have no payment tag. Their past bills were all paid on the day of issue, which suggests a card. This was not confirmed. Danica was asked to check.

That leaves Employsure, LinkedIn, CEO Institute and Audio Visual Events as the likely bank payments. This is an inference from narration tags, not a source document.

## ABN follow up

Danica asked for ABNs to be checked on each invoice. This could not be done, for the same reason as above. She was given a list of what to look for.

Australian suppliers must show an ABN. These are Audio Visual Events, CEO Institute, Employsure, Energy Australia, Optus, Virgin and Seek. Employsure cannot be checked until a copy of the invoice is obtained.

Overseas suppliers. LinkedIn Singapore is the one that matters most. If its invoice shows an ARN rather than an ABN, it is on the ATO simplified GST system. It then cannot issue a tax invoice, and the 636.00 GST cannot be claimed. Last year's bill should be checked for the same thing. Expedia has the same question on its 21.26 GST. Avetta and Mailchimp charge no GST, so there is no GST credit at risk on those.

Rules given to her. These came from an ATO web search on 23 October 2026, not from a recording or last month's file. She was told this plainly and told to check with her manager before holding back any payment.

If an Australian supplier quotes no ABN, the payer generally must withhold 47 percent, unless an exception applies. One exception is a total payment of 75 dollars or less.

A valid tax invoice is needed to claim GST credits on purchases over 82.50 dollars including GST.

Businesses registered under simplified GST get an ATO reference number, not an ABN, and cannot issue tax invoices.

Sources used.

https://www.ato.gov.au/businesses-and-organisations/hiring-and-paying-your-workers/payg-withholding/payments-you-need-to-withhold-from/withholding-from-suppliers/payments-you-don-t-withhold-from

https://www.ato.gov.au/businesses-and-organisations/not-for-profit-organisations/your-organisation/dealing-with-suppliers/withholding-in-business-transactions

https://www.ato.gov.au/businesses-and-organisations/international-tax-for-business/gst-for-non-resident-businesses/gst-registration-for-non-resident-businesses/simplified-gst-registration

https://www.ato.gov.au/businesses-and-organisations/gst-excise-and-indirect-taxes/gst/in-detail/rules-for-specific-transactions/international-transactions/australian-business-importing-goods-and-services

## Open items for the next agent

1. Re-pull the 11 bills by bill ID and see which have since been paid or changed.
2. Ask Danica whether she did the bank detail, GST to PDF and ABN checks, and what she found. Do not assume.
3. Ask whether the Audio Visual Events job number was confirmed or corrected.
4. Ask whether the Energy Australia card payment was found and reconciled.
5. Ask what the LinkedIn invoice shows, ABN or ARN. If ARN, the 636.00 GST is a real problem for this bill and possibly last year's bill too. Do not decide the fix. Ask her and her manager.
6. Ask whether she has a recording or past workings for the creditors run. If she does, check the run against it, per the CLAUDE.md rule to do the work exactly as shown.
7. Confirm with her which bills are paid by card and which by bank, for Virgin and Expedia in particular.

## Connector notes for the next agent

The Xero tools used were get_connected_organisations, get_bills, get_purchase_orders and get_chart_of_accounts.

get_bills with include_line_items set to true returns large results. Over about 30 bills it overflows and saves to a file. Use jq on the saved file.

get_contacts returns only ID, name and email. No bank details, no ABN, no address.

No Xero tool reads bill attachments. has_attachments only says whether one exists.

In the chart of accounts, 11510 Prepayments and 65128 Recruitment both default to BAS Excluded. Bills on those accounts that show GST must have had the tax rate changed on the line. That is consistent with past bills, so it was not flagged.
