# Hosting the portal on SharePoint and OneDrive

One folder, synced by everyone, opened from the synced copy. No server, no
launcher, no sign on, nothing installed.

## Set up once

1. **Create a document library** in the CTS SharePoint site called
   `CTS Business Portal`. A library rather than a folder inside another one,
   so it can carry its own permissions.
2. **Copy the whole `portal` folder into it**: the HTML, `CTS_bi_bundle.js`,
   `CTS_build.js`, `CTS_email.js`, `vendor`, `data`, `templates`, `docs`.
3. **Set permissions** on the library: Finance gets Edit, everyone else who
   should see the portal gets Read. This is the real access control. If someone
   cannot open the library they cannot sync it and there is no portal to open.
4. **Sync it** from SharePoint with the Sync button. Everyone who should see
   the portal does the same.
5. **Mark it Always keep on this device.** Right-click the synced folder in
   File Explorer and choose it. A file that is only in the cloud reads slowly
   or fails, and the ledger file is a few megabytes. The O33 handover flagged
   exactly this trap.

## Opening it

Open the synced folder in File Explorer and double-click
`CTS Business Intelligence Portal.html`. Pin it to the taskbar or make a
desktop shortcut.

**Clicking the file in OneDrive or SharePoint on the web does not run it.**
The web view shows it as a file to download. It has to be opened from the
synced folder on a PC. This is the single most common confusion and it is
worth telling people up front.

## Who does what

| Person | Does | Needs |
|---|---|---|
| Finance | Updates the templates monthly, clicks Build | Edit on the library, Edge or Chrome |
| Everyone else | Opens the portal | Read on the library, any browser |

## Rollback

Every Build copies the previous data files into `data/_previous` before
writing. To go back a month, copy the contents of `_previous` over `data`.

## A department head who must not see the ledger

Give them a second library with a restricted build in it. The Build page can
write a data set that omits the ledger, or carries only one department. That is
the only version of "cannot see" that holds; the tab visibility inside the
portal only decides what people are shown.
