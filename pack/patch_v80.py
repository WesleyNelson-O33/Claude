"""Fix the forecast overhead references, which were one month early.

'Budget FY27' holds its departmental overhead split at rows 197 to 226, and
that block has its own header at row 196 running column C for July through
column N for June.  The rest of the tab runs column B for July through column
M for June, with N the FY total.

'P&L by Dept' read the block by the tab's layout, so every forecast month
picked up the month before, July read the block's label column and got nothing,
and June was never read at all.  The three departments' forecast overhead came
to the company's figure for the previous month, which left the four blocks
23,471 worse than the company across October to June and put the FY total out
by the same.

Each forecast reference now points one column right, at the month the block's
own header gives it.
"""
import re
import openpyxl
from openpyxl.utils import get_column_letter as gl, column_index_from_string as ci

WB = 'pack/CTS Financial Controller Pack - v20.xlsx'
# the departmental overhead rows on Budget FY27, which the block lays out C..N
BLOCK = range(197, 227)
# the overhead lines on each P&L by Dept block that carry a budget reference
ROWS = [61, 62, 63, 64, 66, 68, 138, 139, 140, 141, 143, 145,
        188, 189, 190, 191, 193, 195]
COLS = [gl(c) for c in range(2, 14)]
REF = re.compile(r"'Budget FY27'!([A-Z]{1,2})(\d+)")

wb = openpyxl.load_workbook(WB)
ws = wb['P&L by Dept']
bud = wb['Budget FY27']
assert bud['C196'].value == "='Budget FY27'!B$5", bud['C196'].value

def bump(m):
    col, row = m.group(1), int(m.group(2))
    if row not in BLOCK:
        return m.group(0)
    return f"'Budget FY27'!{gl(ci(col) + 1)}{row}"

changed = 0
for row in ROWS:
    for col in COLS:
        cell = ws[f'{col}{row}']
        new = REF.sub(bump, cell.value)
        if new != cell.value:
            cell.value = new
            changed += 1

print(f'{changed} references moved one column right')
print('  example, Support employment, October:')
print('   ', ws['E193'].value)
wb.save(WB)
