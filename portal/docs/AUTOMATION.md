# Automating the monthly email

Designed so the switch from manual to automatic is a flow you turn on, not a
change to the portal.

## How it is built

The portal never sends email itself. On Build, **Write this month's emails**
renders one message per recipient on the Config Distribution tab and writes it
into `outbox/YYYY-MM/`:

```
outbox/2026-08/
  _manifest.json          month, count, file list
  duncan.json             { to, name, tier, subject, html, text, month, status: "pending" }
  duncan.html             the same message, openable in a browser
  jordan.json
  jordan.html
```

Three tiers, decided per recipient on the Distribution tab:

| Tier | Who | Gets |
|---|---|---|
| 1 | Executive | Company result, departments after the split, top clients under the four headings, utilisation, what is worth a comment |
| 2 | Department head | Their own department in depth, one line on the rest |
| 3 | Finance | The controls: does the ledger tie, what is uncoded, what rests on a placeholder, what the build did |

Whatever sends the mail reads that folder. The portal does not need to know
what it is. That is the whole design: rendering is deterministic and reviewed,
dispatch is pluggable.

## Why it is triggered, not scheduled

Nothing goes out until somebody has clicked Build and read the checks. O33
made the same call, and for the same reason: automated scheduling risks
distributing unreviewed data. What is automated is everything after the click.

## Stage 1, now: copy into Outlook

Open `outbox/YYYY-MM/<name>.html` in a browser, select all, copy, paste into a
new Outlook message. The formatting survives. Or use Open in Outlook on the
Distribution page for a plain text version.

## Stage 2: Power Automate sends the outbox

A cloud flow in the CTS Microsoft 365 tenant, about ten minutes to set up, no
code. Once on, step 3 of the monthly routine is one click.

1. In Power Automate, **Create, Automated cloud flow**.
2. Trigger: **When a file is created in a folder** (SharePoint). Site: the CTS
   site. Library: `CTS Business Portal`. Folder: `/outbox`. Tick *Include
   subfolders*. Add a condition that the file name ends in `.json` and does not
   start with `_`.
3. Action: **Get file content** (SharePoint), on the trigger's file identifier.
4. Action: **Parse JSON** on the file content, with this schema:

   ```json
   { "type": "object", "properties": {
     "to": { "type": "string" }, "name": { "type": "string" },
     "subject": { "type": "string" }, "html": { "type": "string" },
     "text": { "type": "string" }, "month": { "type": "string" },
     "status": { "type": "string" } } }
   ```

5. Condition: `status` equals `pending` and `to` is not empty.
6. Action: **Send an email (V2)** (Office 365 Outlook). To: `to`. Subject:
   `subject`. Body: `html`. Switch the body editor to code view so the HTML is
   sent as HTML, not as text. Send from the finance mailbox.
7. Action: **Update file** to write the JSON back with `status` set to `sent`
   and a `sentAt` timestamp, so the Distribution page can show it went.
8. Save, and turn the flow on.

Test it by building with one recipient on the Distribution tab set to Send Yes
and your own address.

## Stage 3, later: Graph from the portal

The O33 way: an Azure app registration with `Mail.Send` delegated, MSAL in the
portal, and a Send button that calls `/me/sendMail` from the signed-in finance
user's own mailbox. It removes the flow but adds an app registration, a
redirect URI and a launcher on port 3000 for local sign in. Not worth it until
the flow is the thing slowing you down.

## What is deliberately not automated

- **The build.** A browser cannot run on a schedule. Finance clicks it.
- **The distribution list.** It is a tab in Config, rebuilt each month, so an
  address change is a template edit and not a flow edit.
