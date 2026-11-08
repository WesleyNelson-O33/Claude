# CLAUDE.md — CTS payroll support

This repo supports Danica Nelson (Finance Operations Manager, CTS, Sydney) running fortnightly payroll in
Employment Hero (EH). The agent's job: maintain the payroll check workbook, check every employee's EH pay
lines against timesheets, rosters, the shared calendar and manager approvals, tell Danica exactly what to
fix in EH, and write her notes, emails and Teams messages.

## Start of every session
1. Read every file in `memory/`. Reuse the wording there exactly; do not invent new formats.
2. Read the `Corrections & Checks` tab of `payroll/Payroll_Tracker_Template.xlsx` (the register of
   recurring checks and known EH problems).
3. Check `list_triggers` for the carried-over reminder for the current pay run.
4. If a handover document was supplied by Danica, read it. It is not in this repo on purpose.

## Hard rules
- **No employee data in this repo.** No names with rates, hours, tax or leave details. The template must
  stay clean. The `memory/` folder is written without employee names. Handover notes with names are given
  to Danica as a file, never committed.
- **Delete every copy of her uploads** from the scratchpad after use. Corrected files go back only to her
  via SendUserFile, same filename. Never put a TFN in an email.
- **Only what Danica explicitly asks to remember goes in `memory/`** ("commit that to memory",
  "remember that", "embed in memory").
- **Facts only.** Answer only from the material she gave this run. If a source does not cover something,
  say "nothing you gave me covers X". Never infer an approval.
- Australian context and spelling.

## How Danica wants answers
- Short. One person at a time when she says so. Under time pressure, only what changes money.
- Corrections as a short copyable list. EH Notes in the exact format of existing notes:
  `DD/MM/YYYY - HH:MM to HH:MM. <short reason>`.
- Payroll notes written in her voice (first person), with `___` for pending answers.
- "Is this on any list?" → answer against every source she gave this run in one reply.
- Overtime: always say "N hours at x1.5", never "1.5 OT". The Hours box is hours, the multiplier is in the
  category/rate. Explain x1.5 over-80 OT and 10-hour-break cases to Graham in one plain line each.
- Don't argue twice. State the fact once, then do what she decides.

## Workbook mechanics
- Patch with openpyxl (`python3 -m pip install openpyxl`; plain `pip` targets the wrong interpreter).
- Read values with `load_workbook(data_only=True)`. LibreOffice recalc times out on this workbook; set
  `wb.calculation.fullCalcOnLoad = True` and let Excel recalc.
- Never insert columns in `Fortnightly` with openpyxl (cross-sheet references do not shift).
- Structural fixes go into both her live tracker and `payroll/Payroll_Tracker_Template.xlsx`; commit the
  template with a one-line message and push to the working branch.
- Key checks: Staff column C blank → gross 0 ("I'm out $18k"); duplicate Staff rows; unmapped pay
  categories (Check Summary line 14); hours reconciliation excludes Leave Without Pay (row 139).
- EH re-import of timesheets wipes manual pay-run edits. Tell her to fix lines last, export, never import again.

## Standing policy (details in memory/STANDARD_NOTES.md)
- Award: Live Performance Award 2020 (MA000081) Part 8. Casual OT >8 h/day 175%/225%; >38 h/week
  (excluding daily OT) 175%/225%; meal break penalty 225% casual / 200% permanent for the break period;
  <10 h break 200%/225%; Sunday and PH 200%/225%; casual minimum engagement 3 h; casual OT on base rate.
- CTS Overtime Matrix for full-time: 80 h per fortnight; >80 h paid OT x1.5; TIL in 15-min blocks.
- Full-time staff are not paid the 200% less-than-10-hour-break rate (salary buffer). Flag it each time.
- OT/TIL only recognised when on the manager's approval list; EH auto casual OT stands but is listed.

## Git
- Work on the designated `claude/...` branch, commit small, push after every commit, never create a PR
  unless asked. End commit messages with the attribution footer the harness provides.
