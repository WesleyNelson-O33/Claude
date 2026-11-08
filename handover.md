# Handover: Scene Change comms hire PO

## Where things stand

The PO for the Scene Change comms hire has not been raised by me. I only drafted it.

Three things are still missing. The job number, the quote number, and the price for 8 units.

Danica was about to email Tyler Wood for the job number. I do not know if she sent it or got an answer.

This handover was written on 8 November 2026. The hire itself was on 29 and 30 October 2026. So the first job for the next agent is to check Xero. The PO may already exist, or Scene Change may already have sent a bill.

## Who you are writing for

Danica is dyslexic and listens to replies with text to speech. Read CLAUDE.md in this repo before you reply to her. The key rules are below, but CLAUDE.md wins if anything differs.

Short sentences. Plain words. Answer first.

No em dashes, ever. This is a hard rule.

No nested bullets. Keep replies short. She has said more than once that replies are too long.

Do not show the working. Give the answer and the one line cause.

Never guess and present it as fact. Say clearly when something is from Xero, from the email, or is your own judgement.

When she asks for "Xero format to copy and paste", she wants each Xero field as its own code block, in the order the Xero screen asks for it.

## The source document

She uploaded a PDF of an Outlook email thread. Subject "Re: Comms hire 29th & 30th Oct". It sits in the CTS Accounts Team mailbox.

The upload path was:
/root/.claude/uploads/829863d0-dba1-5b3e-8578-8244e1eb4257/29ecdcf3-Re__Comms_hire_29th__30th_Oct_-_CTS_Accounts_Team_-_Outlook_-_PO_Request.pdf

That path may not exist in a new session. If it is gone, ask Danica to upload it again.

The Read tool did not show the PDF text. Use pdftotext -layout to read it.

## What the email thread says

27 Oct, 10:07am. Tyler Wood asks Jared Lattouf at Scene Change for a comms hire for Thursday morning and a Friday show. He says Riedel can't facilitate it. He asks for 10 belt packs and pro headsets, and 2 antennas with arms. Delivery or pick up from 23/33 Mary St, Surry Hills.

27 Oct, 10:29am. Jared offers it for $2,500 ex GST delivered.

27 Oct, 10:40am. Tyler accepts and asks for a quote. He wants delivery before midday Thursday and pick up Friday after 1:30pm. The contact on site is Jaazaniah Salanoa, called Jaaz, on 0422 474 355.

27 Oct, 10:48am. Tyler reduces the order to 8 packs and headsets.

27 Oct, 2:26pm. Jared sends the formal quote for 8 units, delivered and picked up. The quote was an attachment. It is NOT in the PDF. So the quote number and the price for 8 units are unknown.

28 Oct, 10:59am. Tyler asks the accounts team to send Jared a PO as soon as possible.

The thread never names the client or the job.

Contacts:
Tyler Wood, Senior AV Production Project Manager, tyler.wood@ctsav.com.au, 0424 755 997.
Jared Lattouf, Technical Director, Scene Change Sydney, jlattouf@scenechange.com.au, 0416 928 068.
Jaazaniah Salanoa, Jaazaniah.Salanoa@ctsav.com.au, 0422 474 355.

## What I found in Xero

Organisation: Corporate Technology Services Pty Ltd.
Organisation ID: 3993749e-5d5f-4997-bb08-96616f256085

Supplier contact: Scene Change Pty Ltd.
Contact ID: 37313532-69f9-46da-ac83-722b84f8c930
The email saved on the contact is sydney@scenechange.com.au. Tyler wants the PO sent to Jared, so the send to address must be changed.

Scene Change has no POs and no bills in Xero from 1 July 2026 to 28 October 2026. So there is no past Scene Change coding to copy.

The PO number pattern is PO-000169xx. The last one I saw was PO-00016958 on 23 Oct.

The reference pattern on past POs is the quote number, a dash, then a short description. For example "Q-26-3367.01 - SYD - CTS AV - Suncorp Investor Day 29.10.2026".

## The account code and where it came from

I used account 51200. This is my own match, not a rule from a recording or manual.

It comes from the Riedel Communications bills. Riedel is the usual comms hire supplier. Their equipment hire lines are coded 51200. Their delivery and pick up line is coded 58215. Their preparation line is coded 51400.

The bills I checked were INV00010213, dated 13 Aug 2026, and INV00010410, dated 15 Oct 2026.

Danica has not confirmed 51200. Tell her it is a match to Riedel, not a confirmed rule.

## Tracking categories

There are two tracking categories. Cost Centres and Job Numbers.

Cost Centres options are CONSULTING, CTS, INTEGRATION, ONSITE, PRODUCTION and VIDEO.

I used PRODUCTION. The reason is that Tyler is the production project manager, Production is copied on the email, and every job that fits the dates is a PRD job. Danica has not confirmed it.

Job Numbers has 469 options. The Xero tool output is too large to show in full. Save it to a file and search it with jq.

## The job number problem

These jobs fit the dates of 29 and 30 October:

26091604 - The George Institute - SYD CV Symposium - 29-30.10.26 [PRD]

26032502 - 3YS Owls - Governance Top 100 2026 - 30.10.26 [PRD]

26061606 - MUFG - Endeavour AGM - 30.10.26 [PRD]

2609402 - MUFG - SUN Investor Day - 29.10.26 [PRD]

2610605 - Bankwest - Stage Hire - 29.10.26 [PRD] also exists. But it is a stage hire job, so it is unlikely.

The George Institute is the only job that runs on both days. That is a guess, not a source. Do not code it there without Tyler confirming.

## The Riedel PO, and why we dropped it

There is an open Riedel PO, PO-00016951, for the Suncorp Investor Day on 29 Oct. It totals 1,829.52 and was raised on 15 Oct.

I first suggested asking Tyler if Scene Change replaces it. Danica questioned this. I agreed it was weak, because nothing in the email links this hire to Suncorp. "Riedel can't facilitate it" most likely means Riedel can't take this new job.

We removed it from the email to Tyler. Do not raise it again unless Tyler's answer points to Suncorp.

## The PO draft

The top of the PO:

Contact: Scene Change Pty Ltd

Date: 28 Oct 2026

Delivery date: 29 Oct 2026

Reference: [QUOTE NUMBER] - Comms Hire - 29th & 30th Oct

Amounts are: Tax Exclusive

The line item:

Description: Comms hire - 8 x belt packs and pro headsets - delivered and picked up

Qty: 1

Unit price: [EX GST TOTAL FROM QUOTE]

Account: 51200

Tax rate: GST on Expenses

Cost Centres: PRODUCTION

Job Numbers: [JOB NUMBER FROM TYLER]

If the quote shows delivery as its own line, add a second line. Use account 58215, the same cost centre and the same job number.

The delivery details:

Delivery address: 23/33 Mary St, Surry Hills NSW 2010

Attention: Jaazaniah Salanoa

Telephone: 0422 474 355

Delivery instructions: Deliver before midday Thursday 29 Oct. Pick up Friday 30 Oct after 1:30pm. Contact Jaaz on 0422 474 355 on arrival.

Send the PO to: jlattouf@scenechange.com.au

Note. The original request also asked for 2 antennas with arms. Tyler's reduction to 8 only mentions packs and headsets. Check the quote to see if the antennas are still included.

## The email Danica approved for Tyler

To: tyler.wood@ctsav.com.au

Subject: Job number for Scene Change comms hire PO

Hi Tyler,

I'm raising the PO for the Scene Change comms hire on 29 and 30 October, but I need the job number before I can send it to Jared.

Could you please confirm which job this one should go to?

Thanks,
Danica

I suggested she reply on the existing thread so the quote stays attached.

## What the next agent should do

First, check Xero for any PO or bill for Scene Change Pty Ltd dated from 28 October. Use the contact ID above.

Second, ask Danica if Tyler replied with the job number, and if she has Jared's quote.

Third, once you have the quote, fill in the quote number and price. Check if delivery is a separate line and if the antennas are included.

Fourth, give her the finished PO in copy and paste format, one code block per field.

## Tool notes

The Xero tools can list POs, but they do not show PO line items or account codes. Bills with include_line_items set to true do show account codes, but not tracking.

The tracking categories result is very large. It saves to a file. Search it with jq rather than reading it all.

The repo has payroll, WIP and month end material. It has no PO or accounts payable guidance. So there is no recording or manual for PO coding. Everything above on coding is a match to past Xero data.
