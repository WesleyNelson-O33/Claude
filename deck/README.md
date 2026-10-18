# CTS Monthly Business Review

`CTS Monthly Business Review - September 2026.pptx`, 25 slides, built from the
Financial Controller Pack.

## Rebuilding it next month

    cd deck/build
    python3.13 extract.py          # reads the pack, writes data.json
    NODE_PATH=$PWD/node_modules node make.js

`extract.py` reads whatever month the pack's Setup tab is set to, so the only
thing to change each month is the pack itself.

## Where each slide comes from

| Slide | Source in the pack |
|---|---|
| Introduction, P&L summary | Summary, Budget FY27, Actual FY26 |
| P&L per department, profit made and lost | Summary plus Overhead Split sections 4 and 5 |
| Three months to date | Summary columns B to I |
| Cost distribution, other opex, labour mix | Cost Analysis |
| Employment expense split | Overheads MoM rows 133 to 151, Overhead Split section 6 |
| Top ten customers | Top 10 Clients |
| Receivables and DSO | DSO |
| Utilisation, profit on utilisation | Utilisation |
| Revenue by month | Cost Analysis, total income by month |

Department colours follow `pack/data/dept_colours.json`: Support red, Production
teals, Consulting blues.
