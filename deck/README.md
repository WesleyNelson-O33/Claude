# CTS Monthly Business Review

`CTS Monthly Business Review - September 2026.pptx`, 22 slides, built from snips
of the Financial Controller Pack.

Every table and every chart on a content slide is a picture of the pack's own
cells, with the pack's own formatting. Nothing is redrawn in PowerPoint. Only
the title, agenda, introduction, section dividers, actions table and closing
slide are native, because none of them is data.

## Rebuilding it next month

    cd deck/build
    python3.13 snip.py              # renders every snip from the pack
    python3.13 extract.py           # figures for the commentary and notes
    NODE_PATH=$PWD/node_modules node make.js

Both scripts read whatever month the pack's Setup tab is on, so the only thing
to change each month is the pack.

## How the snips are made

`snip.py` takes a copy of the pack, hides every sheet except the one it wants so
LibreOffice exports a single page, sets the print area to the range, converts to
PDF and trims the white margin. Sheets are hidden rather than deleted, because
deleting them would turn every cross-sheet formula into #REF.

The Summary snips collapse the outline to the on-screen view: departments and
sub-departments, no account detail. Some snips hide working columns so a table
can sit beside its chart.

## Slide to range

| Slide | Sheet | Range |
|---|---|---|
| P&L summary, income | Summary | A5:I58, outline collapsed |
| P&L summary, cost and profit | Summary | A59:I181, header row repeated |
| Department result | Overhead Split | A140:H148, section 7 |
| Revenue against budget and last year | Cost Analysis | F190:W212 |
| Operating cost distribution | Cost Analysis | I33:X59 |
| Other operating expenses | Cost Analysis | I4:X30 |
| Labour mix | Cost Analysis | I61:X87 |
| Employment expense split | Overheads MoM | A160:P178 |
| Top ten customers | Top 10 Clients | A5:N20 |
| Receivables and DSO | DSO | A5:S45 |
| Utilisation by department | Utilisation | A66:P84 |
| Revenue by month against budget | Cost Analysis | I117:X140 |

## Charts added to the pack for this deck

Three graphs the deck needed did not exist in the pack, so they were built there
rather than in PowerPoint:

- Cost Analysis, revenue by department against budget and last year
- Utilisation, hours by department, chargeable, non-chargeable and leave
- Overheads MoM, employment expense, departmental against office staff

The pack now carries 20 charts. Department colours follow
`pack/data/dept_colours.json`.
