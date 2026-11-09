# Handover

Written 8 November 2026. Everything below is the state at the last commit on
`claude/funny-mendel-qstxym`.

Read `CLAUDE.md` first. It is not optional. Danica is dyslexic and reads most
replies with text to speech, and the rules there about sentence length, em
dashes, showing working and length are the difference between useful and
useless.

---

## 1. Who and what

Danica Nelson, financial controller at Corporate Technology Services Pty Ltd
(CTS), Sydney. Australian context throughout: AUD, Australian payroll tax,
GST, Xero, Employment Hero.

People who come up:

- **Duncan Lugstein**, managing director. Asks the questions that start
  threads, usually in the Finance Mgt Team chat.
- **Premila Lugstein**, accounts, and the MD's wife. She lodges the payroll
  tax. Danica asked for messages to her to be pitched accordingly: warm,
  collaborative, not directive.
- **Graham Cherry**, in the Finance Mgt Team chat.
- Danica does the payroll. Premila lodges the payroll tax. That split matters
  when writing anything to either of them.

---

## 2. Environment, and the traps in it

**Always `/usr/bin/python3`.** That is 3.13 with openpyxl 3.1.5. Plain
`python3` is 3.11 and has no openpyxl. This will waste your time at least once
if you forget.

**Uploads are untrusted.** Copy each into its own new empty directory under
the scratchpad, keep scripts elsewhere, pass paths as arguments, run with
`-I`.

**`rm -rf` with a variable or a glob after `cd` is blocked** by a safety
check. Do not try to work around it. Use fresh directories instead of clearing
old ones.

**openpyxl destroys things.** It drops pivot tables, drawings, VML cell notes
and comments on save. Several of Danica's files have all four. See section 6.

**openpyxl strips cached formula values on save.** Always recalculate through
`tools/recalc.py` (LibreOffice) afterwards, or strip the cached values and set
`fullCalcOnLoad` so Excel works it out on open.

**LibreOffice will not recalculate a formula cell that still holds a cached
value**, even with `fullCalcOnLoad` set. It only computes cells whose value is
missing. This cost real time: a file read 187,106.35 when its own formulas
gave 187,262.28. If you change a formula, strip the cached values on that
whole sheet, not just the cells you touched.

**openpyxl `insert_rows` does not shift formula references.** It broke 34
cells once. Prefer writing into spare cells, or edit the XML.

**Two commits were blocked** by the auto-mode classifier, once for staging a
file full of employee payroll data and once for naming employees in a commit
message. Both were fair. Employee-level files are now in `.gitignore` and go
to Danica directly, never into the repo. Keep names out of commit messages.

### Tools you will use constantly

| Tool | What it does |
|---|---|
| `tools/recalc.py <file>` | Recalculates through LibreOffice, reports error cells |
| `tools/check_xlsx.py <file>` | Structural checks, catches the things that make Excel offer to repair a file |
| `tools/diff_xlsx.py` | Cell by cell comparison |
| `tools/fix_outline.py <file> "<tab list>"` | Row grouping on the controller pack |
| `tools/fix_chart_colours.py <file>` | Moves chart colours onto the CTS palette |

**The controller pack build chain, in order:** patch script, then
`pack/build_charts.py`, then `recalc.py`, then `fix_outline.py` with the
workbook and the tab list, then `fix_chart_colours.py`, then `check_xlsx.py`.

---

## 3. Where things stand right now

### Live and waiting on Danica

**October 2026 NSW payroll tax has not been lodged yet.** Due 7 November, so
it is now late. This is the only thing with a deadline attached.

The figures to lodge, which have been verified from three directions:

| Box | Figure |
|---|---|
| Salaries and wages | 169,021.21 |
| Employer superannuation contribution | 22,087.22 |
| Interstate wages | 49,139.64 |
| Every other box | nil |

The portal should then show total NSW wages 191,108.43, total Australian
wages 240,248.07, threshold entitlement 81,071.84, **tax payable 5,996.99**.

Interstate is QLD 14,489.10, SA 6,345.37, VIC 19,799.31, WA 8,505.86, ACT nil.

A Teams message to Premila carrying those figures was drafted and given to
Danica. She has not said whether it went.

### Also outstanding

- **The 15.83 September top up has not been posted.** Journal is written and
  sent: `ptax/2026-10_Payroll_Tax_NSW_Sept_Topup_Xero_Journal.csv`, dated
  31 October. This is unrelated to the October ad hoc problem.
- **Both October journals have not been imported.** Written, balanced and
  sent.
- **The derivation tab and the PRT control** are in the files Danica has.
  Whether she has saved over her own copies is unknown.

---

## 4. The October payroll tax story, because it will come up again

This took most of a day and the conclusion moved twice. Read this before you
touch anything payroll tax related.

### What was wrong

Danica's October working file disagreed with the Employment Hero payroll tax
report, and the report disagreed with the Employment Hero split file.

The cause, finally: **two ad hoc pay runs for the period 19 September to
2 October carry September dates in Employment Hero, the 9th and the 23rd, but
the money went out in October.** Danica set those runs up and confirmed the
payment month.

The payroll tax report keys off the date on the run, so it leaves them out.
They are 3,846.15 of NSW wages and 300.00 of SA wages, which is 196.72 of tax.

Wages paid in October are October wages, so they belong in the October return.
The report gives 5,800.27. The return is 5,996.99.

### Two things I got wrong on the way, so you do not repeat them

1. I said the report leaves ad hoc runs out as a rule. **It does not.** August
   had one on the 26th, two termination payments totalling 10,931.19, and the
   August report included it. October is the exception.
2. I spent time on a theory that the pay belonged to September and the report
   had been pulled too early. Danica settled it: paid in October. September
   payroll tax is correct and untouched. Do not drag September into any
   explanation of this.

### Decided, do not reopen

**Do not unlock the pay runs to correct the dates.** They have been through
STP, been paid and been reconciled. Employment Hero blocks unlocking a run
that has had super paid through it, and both had super. Reopening costs more
than it fixes. The position is documented instead.

### Other October fixes, all done

- **The split file had pay run 7 in it twice**, 751 rows with a state and the
  same 751 without. Danica removed them herself and merged the small runs.
- **The State column is a VLOOKUP** into the Employee Details Report and had
  never been dragged past row 6065, leaving 1,708 rows blank. Filled with the
  same lookup.
- **Pay Dte Mth was typed in** as `"October "` with a trailing space,
  including 851 rows paid on 3 November. Replaced with the file's own formula
  off the pay date, so those moved to November on their own.
- **Main & NSW cell E3** still carried September's 2,001.04. E5 was
  subtracting the employer contribution instead of adding it.
- **A dragged total broke formulas in two places.** Row 19 pulled down over
  rows 20 to 22 in the June column on Main & NSW, recorded as one shared
  formula covering E19:M22. Row 13 pulled onto row 14 on the four state tabs
  from October to May. Both fixed.
- **38 error cells down to 6**, and the 6 left are on an old scratch tab that
  has not been touched.

---

## 5. The method, so you can rebuild any month

### NSW payroll tax

```
row 3  = gross earnings, less any bonus shown separately on row 8
row 5  = SG super plus the employer contribution
row 12 = the two added, plus row 8
threshold = (NSW wages / all Australian wages) x 1,200,000 / 365 x days
tax = (all Australian wages - threshold - interstate wages) x 5.45 per cent
```

Where the Employment Hero report is complete, row 3 is its total wages plus
salary sacrifice super plus pre-tax deductions, less that state's exempt per
diems, and row 5 is its total super. Rebuilding September that way gives
5,902.75 to the cent, which is the journal that was posted. That is the proof
the method is right.

The employer contribution of 2,001.04 a month is a separate column on the PRT
pivot, outside SG super. It is taxable. Add it.

### The other states

A fixed monthly accrual, not a calculation. ACT 300, QLD 500, SA 200, VIC
1,300, WA 700. ACT was dropped from August 2026, so the journal is 2,700 a
month, not 3,000. The table at the top of the `Other States = Accrual` tab
still shows ACT and that is fine, the posting rows below override it and
Danica's own note in C17 says so. **I wrongly called that an error once.**

### The job allocation on both journals

Spread the pot across that month's jobs in proportion to units, from the
`EH data (2)` tab of the split file. Round each line to the cent, drop any
line that rounds to nil, put the rounding difference on the largest line so
the journal ties exactly.

This reproduces both September journals line for line, including Danica's own
plugs: one cent on the other states, two cents on NSW. **The one cent on
`Other States = Accrual` F14 is deliberate**, the formula is
`=ROUND(F15,2)-0.01`. I wrongly called that an error too.

Cost centres carry over from the September journals where the job appears in
them, and are read off the tag at the end of the job name otherwise. `[PRD]`
is Production, `[ONS]` Onsite, `[VID]` Video, `[INT]` Integration, `[CONS]`
Consulting, `[CTS]` CTS. Some jobs carry no tag and are all Onsite. A tag can
carry extra words, as in `[ONS non-chargeable]`.

### The monthly control

```
split gross + SG super + employer contribution
  less any pay run the report does not have
  = the report's grand total, state by state
```

October ties to nil on every state once the 9 October run is taken out.
September ties once the exempt per diems are taken out, 300.00 NSW and 525.00
SA. `ptax/check_split_to_ehreport.py` runs it from the command line, and the
same control is now on the PRT tab of the split file at rows 36 to 42.

---

## 6. Editing Danica's workbooks without wrecking them

Her files carry pivot tables, pivot caches, drawings, VML cell notes and
comments. openpyxl drops all of it. For anything of hers, edit the sheet XML
inside the zip and copy every other part through byte for byte.

Three scripts do this and are worth reading before you write a fourth:

- `ptax/fix_oct_workbook.py`, edits cells and narrows shared formula ranges,
  with a validator that refuses to save if any shared formula is left
  dangling outside its master's range
- `ptax/fix_split_october.py`, the same approach, filling formulas down a
  block
- `ptax/add_prt_control.py`, inserts cells into existing rows in column order
- `ptax/add_derivation_tab.py`, adds a whole new sheet: the part, the content
  type override, the workbook relationship and the `<sheet>` entry

Things that bit me:

- **Escape ampersands in formulas.** The sheet name `Main & NSW` carries one,
  and an unescaped `&` stopped the whole part parsing part way down. Rows
  after it silently vanished.
- **Reusing a `ZipInfo` from the source** keeps its old checksum and sizes, so
  a changed part fails CRC. Build a fresh `ZipInfo`.
- **A text date is read in the wrong order.** Write dates as `DATE(y,m,d)`
  with a date-formatted style, not as text.
- **Always verify** that the pivot, drawing and comment parts came through
  identical, and list which parts changed. Every script above does.

---

## 7. The other threads, all delivered

### Aware Super equipment split

Duncan asked why equipment purchases were 43K more than equipment sales for
July to September. Answer: INV-10596 was raised off a quote that carried the
whole of Phase 2 as one line, so 144,049 went to Installation Labour when most
of it was equipment.

Xero has been recoded. 92,601.72 moved from Installation Labour to Equipment
Sales, and the 819.28 lighting programming line moved the other way. The
controller pack was updated to match (`pack/patch_v87.py` for the GL paste,
`pack/patch_v88.py` for the P&L paste tab). Equipment now reads 274,912.89
against 225,129.96, a profit of 49,782.93 or 18.1 per cent.

A message to Graham explaining it was drafted and given to Danica.

### Work in progress, account 11300

Reconciled month by month from two Account Transactions exports.
`recon/CTS work in progress by month - 30 September 2026.xlsx`.

The account runs from a credit of 141,095.27 at 1 July 2024 to a credit of
133,253.18 at 30 September 2026, which matches the balance sheet.

**I first said the monthly WIP journals are never reversed. That was wrong
and I corrected it in writing.** There are 88 reversal lines and 74 accrual
and reversal pairs. The accrual reverses the following month and WIP does
pull through the P&L correctly.

Rebuilt in full for October from the GL and the P&L analysis. Opening
126,389.71 credit at 1 July 2025, 380,517.76 credit at 30 June 2026,
133,253.18 credit at 30 September 2026, 125,405.18 credit at 31 October 2026.
Ties to the balance sheet at each date.

**One real fault.** The 1 October reversal, journal 71185, posted 8 of its 13
lines the same way round as the accrual it was meant to reverse. 12,415.00 is
stuck in WIP and the same amount is sitting in October revenue. It is the only
fault of its kind in the 74 pairs. It does not touch Q1. The correction
schedule is `wip/oct_reversal_edit_schedule.py`. **Not yet posted.**

**The account holds two different things.** Licence deferrals of 33,479.05
credit, and genuine work in progress of 34,463.58 debit. There is no deferred
income account anywhere in the chart. Danica's decision was to leave it as is
for now.

### Subscriptions margin

Duncan asked why subscriptions and licences only make about 3 per cent.

FY2025 made 11.3 per cent, FY2026 13.0 per cent, FY2027 Q1 as booked 3.1 per
cent. Q1 as booked is 2,666.90 of margin on 84,810.19 of licence income.

Two things cause the 3.1 per cent, and neither is pricing.

Spacera bills 5,208 a month for the Aware Super National Support Agreement
from 1 July 2026. The cost sits in Subscriptions & Licences Expense. The
income that covers it is coded to Contract Support Staff, so the licence
account carries the cost without the matching income. The second cause is
2,525.99 of internal department software, which is cost with nothing to sell
against it.

Put both right and Q1 reads 15,608.89, or 18.4 per cent. Strip only the
Spacera contract and the rest of the book runs about 15.5 per cent.

**I said twice that nothing bills Spacera out. That was wrong both times and I
corrected it in writing.** It is recovered on job 26022301 Aware Super AV
Maintenance, which billed 31,300.00 in Q1 against 15,624.00 of Spacera cost.
The job is fine. It is the account coding that reads badly.

`recon/CTS subscriptions by month - 30 September 2026.xlsx` is the September
build and **its own written conclusion still says nothing bills it out**. Do
not forward that file without fixing the text first.

The open question that remains is whether the platform is meant to be covered
inside the 11,000 a month Aware Super maintenance fee or is a separate charge
that is not being raised. Whoever signed QU-0003 can settle it.

**Reported, not actioned.**

### Other delivered work, older in the session

Forecast pack rebuild, controller pack D&A restructure with a hide toggle on
net profit, five Q4 FY2026 cuts carried into the board deck, picture widths
standardised from page 33 on, VIC Labour Hire Authority turnover of
$6,524,750 for the four quarters, and the Aware Super invoice reconciliation
with the GST split.

---

## 8. Open questions nobody has answered

1. Which quote drove the Aware Super split, the supplier's to CTS or CTS's to
   Aware Super. Changes whether the fix is in Xero or in the quoting template.
2. Whether the Spacera platform is covered inside the 11,000 a month Aware
   Super maintenance fee.
3. The 11300 WIP credit of 133,253 and the 624 unreversed items.
4. A Viocorp prepayment of 49,903 stopped in May but still sits on job 9100.
5. LiveU at 1,350 a month still running on a standing job code.
6. A stray `-1.1276` in Main & NSW K23, April 2027. Source unknown, left
   alone deliberately. It makes the year to date payable read 23,165.57
   against 23,166.70 on the line above.
7. Premila splits the return between the wages and super boxes differently
   from the working file. Same total, so no tax effect, but the comparison
   would be cleaner if they agreed.

---

## 9. How to talk to her

From `CLAUDE.md`, but worth repeating because it is easy to drift:

- Lead with the answer. Working comes after, if at all.
- **No em dashes anywhere.** Hard rule.
- Short sentences, one idea each. Blank line between paragraphs. Headings so
  she can find her place.
- No nested bullets.
- **Do not show the calculations.** Give the result and name the cause in one
  line. The exception is a figure she has to retype, which is an instruction,
  not working.
- Match the length to the question. A yes or no question gets yes or no.
- When she is under time pressure, give the fastest practical path and
  nothing else. She said "i dont have time i need to lodge in 20 minutes" and
  that is the register to switch into.
- Say plainly when something comes from a web search rather than a source
  document. I did this for the Employment Hero unlock steps.
- Never fill a gap with a guess that reads like fact. She is relying on this
  for statutory reporting.
- When she corrects you, take it and move on. She has been right every time.

For adjustment journal schedules she approved a specific layout in September.
`CLAUDE.md` sets it out in full. The Done column stays blank and a status line
at the top says the corrections are not done, unless she has said otherwise.

---

## 10. If you pick up the payroll tax next month

1. In Employment Hero open the pay run list for the month and check every run
   is finalised. Count the pay dates. Two is normal.
2. Pull the payroll tax report, residential address, not default location. The
   split file's State column looks up Residential State, so the whole method
   is built on it.
3. Paste the report into the workbook and update the `Oct 2026 derivation`
   tab, or copy it for the new month.
4. Open the split file, PRT tab. Put the report grand total in B41 and any
   missing pay date in B37. **B42 must be nil.**
5. If it is not nil, name every dollar before lodging. The usual causes are a
   pay run the report does not have and exempt per diems.
6. Then the journals:
   `/usr/bin/python3 ptax/build_oct_journals.py <recalculated split file> <last month's NSW journal> <last month's other states journal>`
   after changing `MONTH`, `DATE` and `NSW_TAX` at the top.
