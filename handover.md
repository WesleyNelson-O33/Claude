# Handover

Written 8 November 2026. This picks up a session that started on 16 October 2026.

## Read this first

Read `CLAUDE.md` in this folder before doing anything. It sets how to write for Danica, and it overrides any default style.

The rules that matter most are below.

- Danica is dyslexic and listens to replies with text to speech. Short sentences. Plain language. Headings.
- No em dashes anywhere. This is a hard rule.
- Lead with the answer. Keep it short. Don't show the calculations unless she asks.
- No nested bullets.
- Australian context and spelling.
- Payroll and month end work is done exactly as the recordings and last month's file show. Never invent a method, a date, an account or a split. If the source doesn't cover it, ask.
- Say plainly when something comes from a web search rather than a source document, or when it can't be verified.

## Where things are

- Repo: `wesleynelson-o33/claude`, cloned at `/home/user/Claude`.
- Working branch: `claude/hopeful-albattani-efl4fg`. Commit and push here only. Don't open a pull request unless Danica asks.
- Xero organisation: Corporate Technology Services Pty Ltd. Organisation ID `3993749e-5d5f-4997-bb08-96616f256085`.
- Monthly P&L reporting file: `reporting/CTS P&L Reporting.xlsx`, built by `reporting/build_pnl.py` and `reporting/build_accounts.py`.
- Financial Controller pack: `pack/CTS Financial Controller Pack.xlsx`, built by `pack/build_pack.py`. The budget file is `pack/CTS Budget FY27 Final - 3 Way fixed.xlsx`.
- Other folders: `accruals`, `wip`, `wages`, `leave`, `ptax` (payroll tax), `payexp`, `checklists`, `manuals` (procedure manuals and recording transcripts), `tools`.

## Task 1. Creditors run review. Done, waiting on Danica.

### What she asked

Check the creditors run before she approves payment. Is the GST right? Is the invoice number right? Does the attached invoice match?

### What was done

There was no creditors run file in the repo or in Google Drive. So the review covered all 32 unpaid bills in Xero as at 16 October 2026. They totalled 235,698.37.

Every bill was checked in Xero for GST at 10%, line maths, invoice number format, attachment present, and past payments to the same supplier.

### The big limit

The Xero connector can't open the attached PDFs. It only says whether an attachment exists. So no bill was checked against the actual supplier invoice. Danica was told this.

### Bills to hold until the PDF is checked

1. The AVIT Group, INV017153016, 94,225.16, due 16 October. Most lines are half quantities, which looks like a 50% progress claim, but the bill doesn't say so. No earlier BW Perth bill from AVIT exists in Xero. The MacBook Pro laptop line is entered twice with the same amount. If that is a keying error, the bill is 5,051.07 too high. The GST is also 9 cents above 10% of the net because of line rounding.
2. Appspace, Inc., INV00133660, 20,985.43. A US company charging 1,907.77 of Australian GST. That GST can only be claimed if the PDF is a tax invoice with an ABN. If there's no ABN, ask Appspace to reissue it without GST. This came from the ATO website, not a source document.
3. Luminous Reputation Management, INV7200, 4,823.50, the balance invoice. The 50% deposit was INV7201, the same amount, paid on 28 August. The balance has a lower invoice number than the deposit, so it could be a duplicate or a typo.

### Bills to take out of the bank file

These are paid another way. Putting them in the ABA file would pay them twice.

- Paid on the PL credit card: Employment Hero 2,248.84, Goget 629.52, Flatworld 92.21, Optus office 639.15, Energy Australia 696.25.
- Direct debit: Ricoh 210.20.
- BPAY: ASIC 342.00, Optus Lugstein home internet 101.20.
- Employsure "42 of 60", 880.00. They never send an invoice, which is normal for them. But instalment 41 was paid a month late, on 6 October. That debit may really have been October's, which would mean 42 is already paid. Check the bank.

### Smaller checks

- Persona Health, INV-18315, 996.60. It was due 10 September and is over a month overdue. Find out why it was held.
- Brookvale travel insurance, INV1914759. There's no GST on a 222.50 premium line. That's fine if it's the international part. Its broker fee is coded 62400, while the other Brookvale broker fees use 62300.
- Brookvale INV1920113 and INV1918462. Each total is 1 cent below the net plus GST. Check them against the PDFs.
- Contact Group INV457215. It's coded to the Hobart fitout, but the line says "Bathurst St Office refresh". Hobart has a Bathurst Street, so it may be fine.

### Bills that passed

Spacera, Contact Group INV457214, Crestron, Easton Belle, the other five Brookvale bills, Posh Services, Type Transcripts, WordBirds, Riedel, both Resolution X bills, Megadeck and RJW. ASIC and Flatworld correctly have no GST.

### Still open

- Danica hasn't said whether she approved the run or checked the PDFs.
- It is now 8 November. Some of these bills will since have been paid. Pull a fresh aged payables report from Xero before acting on anything above.

## Task 2. P&L column insert. Not started.

### What she asked

"In P&L FY27 column insert column U to X, moving PY to Z etc. And insert in column U (Month), column V (Month budget), column W Var, and then Var %."

### Where it got to

Nothing was changed. The first attempt, which only opened the workbooks to list their sheets, was rejected by Danica. She then asked for this handover. Nothing has been done since.

### What the next agent needs to find out

- Which file. The likely one is `reporting/CTS P&L Reporting.xlsx`, but it might be the pack. Ask her, or confirm by finding a sheet or column headed FY27 with PY to its right.
- Which month. Probably October 2026, but confirm.
- Where the Month budget comes from. Probably `pack/CTS Budget FY27 Final - 3 Way fixed.xlsx`. Copy the method from last month's file. Don't invent one.
- Which way the variance is signed. Copy any existing variance columns in the same file.

### How to do it when cleared

- Insert four columns at U, so everything from U onwards moves four columns right. PY should land in Z, and the columns after it should move too. She said Z, not Y. Check why before inserting. One more column may already sit between them.
- Headers: U is Month, V is Month budget, W is Var, X is Var %.
- Use formulas, not typed values. Var % should return blank when the budget is zero.
- Openpyxl doesn't update formula references when columns are inserted. If the build scripts generate the file, change the script and rebuild rather than editing the xlsx. Otherwise, check every formula, total and named range after the insert.
- Tell her exactly what changed, in plain words.

## Other context from this repo

Recent commits cover the September WIP adjustment schedule, the FY27 revenue tracker fixes, the September P&L trading income by cost centre, and the wage split check. The September WIP corrections are recorded as not yet done in Xero. Never mark them as done unless Danica says so. See the adjustment journal rules in `CLAUDE.md`.

## Tools and access

- Xero: read only. It can list bills, contacts, aged payables and the P&L. It can't open attachments or post anything.
- Google Drive: searching for a creditors run file found nothing.
- Gmail wasn't available, so supplier invoices couldn't be pulled from email.
- BigQuery and Slack need to be authorised in the claude.ai connector settings before they can be used.
