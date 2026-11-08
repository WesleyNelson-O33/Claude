# Handover

This file is for the next agent picking up Danica's work.

Read it fully before you touch anything.

---

## Read this first

Danica is dyslexic. She reads most replies with text to speech.

The file `CLAUDE.md` in this repo holds her writing rules. They are not
suggestions. Follow them exactly.

The three that get broken most often:

- No em dashes anywhere. Use a comma or a full stop. This is a hard rule.
- Give the answer, not the working. Do the maths, keep it to yourself, report
  the result. Name the cause in one line.
- Match the length to the question. She has said more than once that replies
  are too long.

She is using this for payroll and statutory reporting. If something cannot be
verified, say so. Never fill a gap with a guess that reads like fact.

---

## The most important rule in this repo

Do the work exactly as the recordings show it. Never embellish.

Everything for payroll and month end is done the way the recordings show, and
the way last month's file did it. Nothing added. Nothing improved. No method
of your own.

The source is the recording, or last month's file. Nothing else.

Where neither covers something, say so and ask. Do not fill the gap, even with
something that looks sensible.

If you are unsure whether something came from the source, it did not. Ask.

This rule exists because it was broken before. Take it seriously.

---

## What this repo is

It is the documentation and tooling for finance and payroll at Corporate
Technology Services, called CTS.

It holds three kinds of thing.

Training manuals, built from recordings of the people who did the job before
Danica. Python scripts that build and fix the month end workbooks. Transcripts
of the source recordings.

It does not hold the live workbooks. Those live on the Y drive and on Danica's
machine.

---

## Current state

The branch is `claude/ecstatic-einstein-ikdcr5`.

The working tree is clean. Nothing is unpushed. There are 50 commits.

The remote is `https://github.com/WesleyNelson-O33/Claude`.

---

## How the repo is laid out

- `CLAUDE.md` holds Danica's writing rules and the adjustment schedule layout.
- `manuals/` holds the finished Word manuals, plus `specs/` which builds them.
- `tools/` holds the manual builders, the redaction script and the transcriber.
- `wip/` holds the September WIP and revenue tracker scripts.
- `ptax/` holds the payroll tax workbook scripts.
- `leave/` holds the annual leave and long service leave provision scripts.
- `accruals/` holds the September expense accruals script.
- `payexp/` holds the staff expense reimbursement journal script.
- `wages/` holds the wage split check against Employment Hero.
- `pack/` holds the Financial Controller Pack and the FY27 budget.
- `reporting/` holds the P&L reporting workbook builder.
- `mejvid/` and `ptaxvid/` hold recording transcripts and their READMEs.
- `checklists/` holds the payroll processing checklist workbook.

---

## How the build pattern works

Scripts are committed. Workbook outputs are not.

Look at `.gitignore`. Every `out/` folder is ignored, and so are the payroll
tax workbooks. The outputs are rebuilt by running the script.

The manuals work differently. Each manual is a JSON spec in
`manuals/specs/`, rendered to Word by `tools/build_manual.js`. The master
manual binds several specs with `tools/build_master.js`.

If you change a manual, change the spec and rebuild. Never edit the Word file
directly. The spec is the source.

The renderer needs the `docx` npm package. It is not committed. Run
`npm install docx` first.

Screenshots in the manuals have employee names and pay rates blurred out.
There is redaction tooling in `tools/redact.py` and the `manuals/redaction-*.json`
files. Keep that convention.

---

## What is finished

- The Payroll Processing Manual, Parts 1 to 7, plus a bound complete version.
- A listening version of the payroll walkthrough.
- The Payroll Processing Checklist workbook.
- The Payroll Handover questions document and the meeting prep document.
- The Monthly Reporting Manual, Parts 1 to 3, plus a bound version.
- The Payroll Tax monthly process manual.
- The Month-End Journals WIP manual.
- The CTS Financial Controller Pack and the FY27 budget.
- The Payroll Workflow document, covering the fortnightly cycle end to end.
- Seven approver email templates, in `checklists/Payroll Approval Emails.txt`.

There is no Part 4 of the payroll manual. Nothing is missing. Video 4 covered a
correction to a salaried employee, so it sits at the end of Part 3b.

---

## What is outstanding

### The September WIP corrections are not done

This is the live item. Treat it as the first thing to check.

The September WIP Adjustment Schedule has been rebuilt with Danica's own
narration. The corrections it describes have not been made in Xero, on the WIP
schedule, or on the revenue tracker.

The schedule says so at the top. The Done column is unticked. The ending
position is worded as still to do.

Do not mark anything as done unless Danica tells you it is done. That is a
standing rule in `CLAUDE.md`.

The script that sets this status is `wip/wip_edit_schedule_status.py`.

### The repository is public

Anyone can read this repo. It is not private.

That matters. The manuals name the people who approve the pay run and the
payment. They describe the banking process, the payment file, the transfer
wording and the account buffer. They include the wording and the recipients of
the payroll approval email.

The risk is not that someone reads the manual. It is that the approval chain,
the platform and the exact email wording together are enough to write a
convincing fake payment approval request.

Danica was told about this. No decision has been made yet. Raise it again if it
is still public when you pick this up.

Employee names and pay rates are blurred in the screenshots, so that part is
handled. No account numbers or passwords are in the repo.

---

## The branch split in two, and has been merged back

This is worth understanding, because it nearly cost 41 commits.

The branch split at `9404d23`, the budget workbook fix. Two different lines of
work then grew from that same point.

One line was payroll: the end to end workflow document and the approver
responsibility model. Those two commits were pushed to the remote.

The other line was month end: payroll tax, WIP, leave, accruals, staff
expenses and the revenue tracker. That was 41 commits. It only ever existed
inside a session container, which is temporary.

At some point the remote was forced back to the payroll line. So the remote
held the payroll work, and the container held the month end work, and neither
knew about the other.

They have now been merged. Nothing was forced and nothing was deleted. No file
was touched by both lines, so the merge was clean.

If you ever see this branch reject a push again, do not force it. Fetch, look
at what is on each side, and merge.

### What came back in the merge

- A Payroll Workflow document, covering the full fortnightly cycle end to end.
- A responsibility model moving the timesheet, roster, overtime and time in lieu
  checks onto approving managers rather than payroll.
- Seven email templates for chasing timesheet approvals each fortnight.
- Two new block types in the manual renderer, for grids and email templates.

One thing did not come back. A payroll notes register, holding one fortnight of
notes, was never pushed. Danica was sent a copy at the time. Ask her for it
before rebuilding it.

### Two decisions in that work are still only drafts

Neither was signed off by Duncan or Graham:

- The responsibility model itself, and the liability wording.
- The rule that time not approved by the Monday cut-off is paid in the next pay
  run.

Do not treat either as agreed policy. They were drafts.

---

## Things that will trip you up

### The fortnight calendar is derived, not sourced

The pay period runs Saturday to Friday. Prep starts on the Friday the period
ends. The timesheet cut-off is 10am on the Monday. The pay date is the Tuesday
of the processing week.

That much is from the recordings.

Any specific fortnight dates were worked out by counting forward in 14 day
steps from the example in Part 2. Always confirm against the Lookup tab in the
timesheet spreadsheet before relying on a date.

### Public holidays were never covered

The recordings do not say what happens when the processing Monday or the
Tuesday pay date falls on a public holiday. Ask. Do not work it out.

### Two approvers have no documented backup

Graham approves the finalised pay run. Duncan approves the payment in CommBiz.
Neither has a written backup. This is listed as an open question in the handover
document and has not been answered.

### Employment Hero gets things wrong on purpose

Several known faults are written into the manual. They are real, not mistakes in
the writing.

It silently removes a thirty minute meal break. It can pay the same hours as
both daily and weekly overtime. It rounds a salaried employee's day to eight
hours in both directions. A rate updated in HR sometimes fails to reach payroll.

Payroll has to catch all four. The manual says how.

### The spreadsheet has a known fault

Total hours do not handle leave without pay correctly. The leave without pay
column is left out of total hours but needed for the dollar value. Verify that
person's total by hand.

This was found during training and never fixed.

---

## Where the source material lives

The recordings are on GitHub releases on this repo, not in the tree. They are
too large for the tree.

The payroll tax recordings are on the `payrolltax` release. The month end
journals recording is on the `Month-endJournals` release.

Only transcripts are committed. To rebuild audio or frames, download the release
assets and run `tools/transcribe.py`.

---

## How to start

Read `CLAUDE.md` first. Then read this file. Then ask Danica what she is working
on right now.

Do not assume the last commit is the current task. She moves between payroll,
month end, payroll tax and reporting depending on where she is in the month.
