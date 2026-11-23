# CTS Monthly Business Review

`CTS Monthly Business Review - <Month Year>.pptx`, built from snips of the
Financial Controller Pack. The file is named after whatever month the pack's
Setup tab is on, so last month's deck is never overwritten.

Every table and every chart on a content slide is a picture of the pack's own
cells, with the pack's own formatting. Nothing is redrawn in PowerPoint. Only
the title, agenda, introduction, section dividers, actions table and closing
slide are native, because none of them is data.

Every figure in the commentary and the speaker notes is read from the pack too.
None is typed into the slide code, so a re-paste of the pack followed by a
rebuild moves the words as well as the pictures.

## Rebuilding it next month

    cd deck/build
    python3.13 snip.py              # renders every snip from the pack
    python3.13 extract.py           # figures for the commentary and notes
    NODE_PATH=$PWD/node_modules node make.js

Point `PACK` in `snip.py` and `extract.py` at the new pack, and that is the only
edit. Both read the reporting month off the pack's own Setup tab, work out how
many months of the year are closed, and set the month names, the comparatives
and the action dates from it.

`snip.py` takes about twenty minutes. `extract.py` takes about five.

### Why the pack has to be recalculated first

The pack carries the previous month's cached results until Excel opens it, and
openpyxl only ever reads that cache. Both scripts therefore put the pack through
LibreOffice with the cached answers stripped out, which forces every formula to
be worked out again. Read the pack with openpyxl directly and you get last
month's figures with this month's headings.

## The two monthly pastes the deck cannot do without

The people and utilisation section needs two things that are pasted in by hand
each month:

1. the month's hours on the pack's **Utilisation** tab, columns by month, and
2. that month's block in **`deck/data/util_her_basis.json`**, which is her own
   monthly Utilisation Report.

Until the reporting month is in both, `make.js` prints a warning and leaves the
whole section out, along with its agenda line, rather than printing it at nil.
The actions table then carries a line asking for the paste.

## Monthly and quarterly

The deck is one file serving both reviews, so the slides have to come apart
cleanly. A monthly slide talks about the month and the year to date only. It
does not refer across to the quarter, because the quarterly slides may be
hidden. Where a quarterly figure is needed it comes from the pack's own Q1 tabs,
which stay fixed on July to September after the pack moves on a month.

## How the snips are made

`snip.py` takes a copy of the pack, hides every sheet except the one it wants so
LibreOffice exports a single page, sets the print area to the range, converts to
PDF and trims the white margin. Sheets are hidden rather than deleted, because
deleting them would turn every cross-sheet formula into #REF.

The Summary snips collapse the outline to the on-screen view: departments and
sub-departments, no account detail. Some snips hide working columns so a table
can sit beside its chart.

`snip.py` renders 51 ranges. The deck places 21 of them. The rest are there to
be dropped on to a slide when a month needs them.

## Where each figure comes from

`extract.py` writes `data.json`. The mapping, by block:

| Block | Tab | Cells |
|---|---|---|
| `pl` | Summary, Actual FY26 | month D, budget L, YTD E, YTD budget G; last year off Actual FY26 |
| `dept` | Overhead Split section 7, Summary | rows 143 to 147; revenue and cost of sales off the Summary |
| `staff` | Overheads MoM | rows 162 to 166 |
| `fte` | Overhead Split | rows 113 to 125, column of the reporting month |
| `util` | Utilisation | blocks of five from row 9, column of the reporting month |
| `top10` | Top 10 Clients | rows 7 to 18 |
| `dso`, `ageing`, `debtors` | DSO | rows 6 to 14, 19 to 24, 31 to 40 |
| `other_opex`, `cost_dist`, `labour_mix` | Cost Analysis | rows 6 to 14, 49 to 58, 63 to 78 |
| `income_by_month` | P&L FY27, Budget FY27 | row 6, columns B to M |
| `q1`, `fy` | Q1 Company Summary, Budget FY27, Actual FY26 | fixed Q1 and twelve month totals |

One figure is not read straight off a cell. `Top 10 Clients` C18 and C19, the
"everyone else" and "total income" figures for the month, are wired to a fixed
column on `P&L FY27` rather than to the reporting month, so they still show
September after the pack moves on. `extract.py` works the month's "everyone
else" out from the month's total income instead, and the cell needs fixing in
the pack.

## Charts added to the pack for this deck

Three graphs the deck needed did not exist in the pack, so they were built there
rather than in PowerPoint:

- Cost Analysis, revenue by department against budget and last year
- Utilisation, hours by department, chargeable, non-chargeable and leave
- Overheads MoM, employment expense, departmental against office staff

Department colours follow `pack/data/dept_colours.json`.
