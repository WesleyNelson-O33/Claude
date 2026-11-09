# How to build these packs yourself

Read this once, then open `Example Pack.xlsx` next to it. Everything I built for
you — the GL analysis pack, the CTS pack, the audit report — is the same four
pieces repeated at larger scale.

---

## 1. The one rule

**One sheet you paste into. Every other sheet is formulas pointing at it.**

That is the whole idea. If any number is typed into two places, the pack is
already broken — it will disagree with itself the first time one of them
changes and nobody will know which one is right.

Your current process breaks this rule in a few places. The clearest example:
the `Sales Revenue Comparison` sheet in your existing file reads typed
constants on the Revenue tab, including `=229761.49+290` and `=41411.43-290`.
Someone typed a reclassification directly into a formula. It works, but nothing
downstream knows it happened and no control would ever catch it.

---

## 2. The four pieces

Every pack has these, in this order:

**Input** — one sheet, a plain grid, nothing above it. You paste the Xero
export here and do nothing else.

**Engine** — a hidden working sheet that adds the raw rows up **once**, by
account and by month. Nobody looks at it.

**Report** — the sheets you actually read. These contain no arithmetic on raw
data, only references to the Engine.

**Check** — controls that prove nothing fell through the gaps. A pack nobody
checks is a pack nobody should trust.

The reason the Engine exists at all is speed. If ten report sheets each scan
5,000 raw rows, Excel does that work ten times over on every keystroke. Scan
once, read the answer ten times.

---

## 3. What you can do in Excel today, with no code

You do not need Python to get 80% of this. Four techniques:

### Make the paste zone start at A1

Excel refuses to paste a whole-sheet copy (`Ctrl+A` in Xero, or a full-column
selection) anywhere except cell A1. That is why the Data sheet in the example
has no title, no logo, no instructions above the header row. Put anything
above your paste zone and you will get "We can't paste that here" forever.
Instructions go on a separate sheet.

### SUMIFS, not lookups

```excel
=SUMIFS(Data!$D:$D, Data!$B:$B, $A2, Data!$A:$A, ">="&$C$1, Data!$A:$A, "<="&$D$1)
```

Plain English: add up column D, where column B equals the account in A2, and
the date in column A falls between the two dates in C1 and D1.

This is the single most useful formula in management reporting. It replaces
almost every pivot table you currently rebuild by hand. Unlike a pivot, it
updates the instant new rows are pasted — no refresh, no "change data source",
no re-dragging fields.

### Row grouping for collapsible reports

Select the detail rows, then `Alt` + `Shift` + `→` (or Data ▸ Group). You get
the +/- buttons down the left edge.

One setting you must change, and it is the one everyone misses: Data ▸ Outline
dialog launcher ▸ **untick "Summary rows below detail"**. Financial reports put
the total at the top of the section, not the bottom. Leave that ticked and the
+/- buttons attach to the wrong rows and the collapse behaviour looks random.

### Controls that tie

At minimum, two:

- Total pasted, less total on the Engine, equals nil. Proves no account got
  dropped on the way through.
- Total on the Engine, less total on the Report, equals nil. Proves no line got
  left off the face of the report.

Write them as `=IF(ROUND(x-y,2)=0,"yes","NO")` so a person reads a word, not a
number they have to interpret at 6pm.

---

## 4. When Python earns its place

Not for a ten-row report. Use it when:

- The same block of formulas repeats hundreds of times (12 months × 190
  accounts × 5 departments is 11,400 cells — nobody drags that correctly).
- The structure changes when the data does (add a department, every sheet needs
  a new block).
- You want the build to be repeatable and reviewable. A script is a written
  record of exactly how the pack was constructed. A hand-built workbook is not.

The library is `openpyxl`. You write formulas as **strings** — openpyxl does not
calculate anything, it just writes the text of the formula into the cell. Excel
works out the answer when the file opens.

### Start here

`template.py` in this folder is a complete, working, ~130-line version of the
whole pattern. Run it:

```bash
python3 template.py
```

Then open the file it writes, change something near the top of the script
(add an account to `ACCOUNTS`, add a month), and run it again. That loop —
change, run, open, look — is how you learn this. Nothing else works as well.

### Then read the real ones

Same folder structure, just bigger:

- `reporting/gl_analysis/build_gl_analysis.py` — 354,143 formulas
- `reporting/cts_pack/build_cts_pack.py` — 533,460 formulas
- `reporting/audit/build_tech_audit.py` — the auditor report

They look intimidating. They are not: they are the four pieces above, wrapped
in `for` loops.

### The build order

Always three steps, always in this order:

```bash
python3 build_whatever.py                          # writes formulas, no answers
python3 /mnt/skills/public/xlsx/scripts/recalc.py "Whatever.xlsx"   # works out the answers
python3 reporting/tools/fix_outline.py "Whatever.xlsx" "Sheet1,Sheet2"  # repairs step 2
```

Step 3 exists because LibreOffice strips the outline settings when it
recalculates. Skip it and your collapse buttons break.

---

## 5. Traps I hit building yours

These cost me real time. Learn them free.

**Paste zone must start at A1.** Covered above. Excel will not budge on this.

**`OR()` does not short-circuit.** In most programming languages,
`OR(A, B)` stops if A is true. Excel evaluates both sides always. So
`=OR(A1="", ABS(A1)>100)` throws `#VALUE!` on blank cells, because `ABS("")`
is evaluated even when the first test already passed. Use nested `IF`s instead.
This produced 682 errors in one build.

**`summaryBelow` must be off when totals sit above detail.** Covered above.

**LibreOffice recalc strips the outline settings.** All three of your packs had
lost them before I noticed — I had checked the file after building and the
settings were there, but the recalc afterwards removed them. Check the *final*
file, not the intermediate one. That is a general lesson: verify what ships.

**Use numbers, not text, as your match keys.** `SUMIF` on a text key like
`"41100|2026|07"` is dramatically slower than on a number like `41100207`.
Build a composite numeric key in a working column:

```excel
=account_index*1000 + fy_offset*100 + period
```

On the CTS pack this was the difference between a pack that opened in seconds
and one that took minutes.

**Guard your division.** `#DIV/0!` spreads across a sheet fast. Always
`=IF(denominator=0, "", numerator/denominator)` — and make sure the guard tests
the *same* cell you are dividing by. I shipped 30 `#DIV/0!` errors to you on
Tech_Spend because my guard checked column B and my division used column F.

**Watch column collisions.** I once wrote a month picker into columns D and E
of a Lists sheet that already held sign and net-amount lookups. Date serials
overwrote the multipliers and the pack quietly reported figures in the
billions. Map out what lives in which column before you write.

---

## 6. The honest summary

The Excel-only route gets you a pack that is genuinely better than what you
have now: one paste zone, SUMIFS everywhere, collapsible sections, two controls
that tie. Build that first. It is a weekend, not a project.

Go to Python when the repetition gets absurd or when you need the build to be
auditable. Not before.
