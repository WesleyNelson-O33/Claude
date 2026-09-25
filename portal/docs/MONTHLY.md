# The monthly routine

Around twenty minutes once the exports are run. Eight templates carry the
month; the ninth, Forecast, only changes when an assumption does.

## 1. Update the templates

In the synced `templates` folder. Each tab is shaped to its export, so it is
a straight paste; the Read me tab in each file says exactly what to run.

| Template | Run in | Paste |
|---|---|---|
| 01 Xero P&L | Xero, Profit and Loss, year to date, one column per month | The account rows, whole year to date each time |
| 02 Xero GL Transactions | Xero, the saved custom report Account transactions for P&L analysis, from 1 July, sorted by account code | All thirteen columns, whole year to date |
| 03 Employment Hero Earnings | Employment Hero, Earnings Details, each pay run | Whole export including its title row, one pay run under the last |
| 04 Zoho Deals | Zoho CRM, Deals, export | Placeholder shape until a real export is matched |
| 05 OnRent Orders | OnRent, orders list | Placeholder shape until a real export is matched |
| 06 Qwilr Quotes | Qwilr, quotes list | Placeholder shape until a real export is matched |
| 07 Budget | Once a year and when the forecast moves | Same layout as the P&L |
| 08 Config | Change the reporting month on the Reporting tab | The one setting that moves every month |
| 09 Forecast | Usually nothing | A growth rate, a method, or a typed override when you know something the history does not |
| 10 Cash and Commitments | Bank and card statements at month end | One row per bank account and per credit card with the month end balance; keep every month. Commitments only when one changes |
| 11 Xero Aged Receivables | Xero, Aged Receivables Detail as at month end | One row per unpaid invoice. Read by the cash flow build; safe to fill now |
| 12 Xero Aged Payables | Xero, Aged Payables Detail as at month end | One row per unpaid bill. Same |

Before you paste the ledger: set Stripe fee rows to `9000 - OFFICE / ADMIN [CTS]`
and fill in the contact on revenue rows from manual journals, one spelling per
client. Convert the values to number.

## 2. Build

Open the portal from the synced folder, **Admin, Build**.

1. **Choose portal folder.** Edge asks you to confirm the folder. Pick the synced
   `CTS Business Portal` folder. One click.
2. **Read templates.** The checks appear: ledger against P&L, rows read, accounts
   not in the chart, locations with no department, back pays reallocated,
   placeholder split bases. Fix anything red in the template and read again.
3. **Write data files.** Last month's files go to `data/_previous`, the new
   ones are written, and a copy of the new data files and the eight templates
   as pasted goes into `archive/YYYY-MM`. The portal reloads on the new month.

OneDrive syncs the data files out. Everyone's portal is current the next time
they open it.

The archive is the record. Every month's folder holds exactly what the portal
showed and exactly what was pasted to get there, so a question in March about
an August number is answered from `archive/2026-08`, not from memory. Nothing
in it is read by the portal and nothing in it needs tidying.

## 3. Commentary

Every page has a commentary block under its heading. Anyone who is signed in
with a role above Viewer can add a note for the reporting month; a department
head's note carries their department. A note is a draft in that browser until
finance opens **Commentary** and clicks Publish, which writes
`data/CTS_commentary_data.js` into the folder the same way Build writes the
rest. OneDrive carries it to everyone, and the monthly emails carry the
published notes for the pages each tier reads.

Write the way the manual asks: explain the variance rather than restate it,
dollars against dollars, and say whether a miss was lost or moved.

## 4. Email

Publish the commentary first, then on Distribution **Write this month's emails** puts one message per recipient
into `outbox/YYYY-MM` as JSON and HTML. From there:

- **Now:** open the HTML beside it and copy it into Outlook, or click Open in
  Outlook for a plain text version.
- **Automated:** a Power Automate flow watches the outbox and sends each JSON
  as an email. See `AUTOMATION.md`. Once it is on, step 3 is one click.

## If a check is red

- **Ledger and P&L disagree.** The paste did not cover every month, or a row
  slipped. Re-paste the whole year to date in both.
- **Account not in the chart.** A new Xero account, or a renamed one. Add it to
  the chart (Config, coming) or fix the spelling in the paste.
- **Location not on the Locations tab.** Add the row on the Earnings template,
  say which department it is and whether it is client work.
- **Split bases placeholder.** Type the live Staff and Office Dept percentages
  from the budget workbook on the Config Split bases tab and set Confirmed.

## At the end of a financial year

There is no rollover step. The templates run three years across and the
portal reads the month from each column heading, so the July paste goes into
the July column and the new year appears on its own. The Build page confirms
it: the Calendar check names the years it built, and on a July build it says
which year has started.

What does need doing, in June, before the first build of the new year:

1. **Budget.** Put the new year's budget into its twelve columns in
   `07 Budget.xlsx`. Without it July is measured against nothing, and the
   revenue schedule has no target.
2. **Public holidays.** Add the coming year's holidays to the Holidays tab of
   `08 Config.xlsx`. The Build page warns if a year in the calendar has none,
   because plain weekdays overstate capacity and understate utilisation.
3. **Split bases.** If the budget workbook moved the Staff or Office Dept
   percentages, type them on the Split bases tab and set Confirmed.
4. **Departments and users.** Check they still match the org chart.

Then in July:

- **P&L.** The year to date paste is July only. Leave the old year's columns
  as they are; the revenue shares and the same month last year come off them.
- **GL.** The Xero transactions report restarts from 1 July, so the paste is
  short for a while. The ledger to P&L control only compares months both cover.
- **Earnings.** Nothing changes; pay runs keep going under the last.

Every third year, when the last year across is the one you are in, add twelve
more month columns to the right of the P&L and Budget tabs, headed the same
way. Columns older than two years back can be deleted once the archive has
them; the prior year is the oldest the portal uses.

