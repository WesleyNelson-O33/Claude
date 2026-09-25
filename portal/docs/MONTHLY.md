# The monthly routine

Around twenty minutes once the exports are run.

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
   ones are written, the portal reloads on the new month.

OneDrive syncs the data files out. Everyone's portal is current the next time
they open it.

## 3. Email

Still on Build, **Write this month's emails** puts one message per recipient
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
