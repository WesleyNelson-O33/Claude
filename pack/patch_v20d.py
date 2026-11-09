"""Collapsible P&L by Dept, and the department colour coding.

Two things:
  1. The same +/- grouping the P&L and Summary tabs have, so the department
     P&L opens and closes the same way.
  2. Her department colours on the column headings, so the colour on a
     chart and the colour on the heading are the same thing.
"""
import os, json, copy
import openpyxl
from openpyxl.styles import PatternFill, Font
from openpyxl.worksheet.properties import Outline

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = f'{HERE}/CTS Financial Controller Pack - v20.xlsx'
P = json.load(open(f'{HERE}/data/pldept.json'))
D = json.load(open(f'{HERE}/data/dept_colours.json'))
B = json.load(open(f'{HERE}/data/brand.json'))

wb = openpyxl.load_workbook(OUT)
ws = wb['P&L by Dept']

# ----------------------------------------------- 1. the +/- grouping -------
# Same convention as P&L FY27: the summary line sits ABOVE its detail, and
# the detail starts closed.
ws.sheet_properties.outlinePr = Outline(applyStyles=False, summaryBelow=False,
                                        summaryRight=False)

TOP = {P['month_rows']['Total Income'], P['ytd_rows']['Total Income']}
grouped = 0
for r in range(1, ws.max_row + 1):
    v = ws.cell(r, 1).value
    if not isinstance(v, str) or not v.strip():
        continue
    dim = ws.row_dimensions[r]
    if v.startswith(' '):                 # indented = account category detail
        dim.outlineLevel = 2
        dim.hidden = True
        grouped += 1
    elif r in TOP:                        # Total Income, the outermost line
        dim.outlineLevel = 0
        dim.hidden = False
    else:                                 # every other P&L line
        dim.outlineLevel = 1
        dim.hidden = False
print('detail rows grouped and closed:', grouped)

ws['A3'] = ('=TEXT(Setup!$B$5,"mmm yyyy")&"   |   click the + and - buttons on '
            'the left to open a category down to subcategory detail"')
ws['A3']._style = copy.copy(wb['P&L FY27']['A2']._style)

# ------------------------------------------- 2. department colour coding ---
# Production is the teals, Consulting is the blue and the purple, Support is
# the red. Admin and Total are charcoal and deep blue, which no department uses.
HDR = {'Total': B['deep_blue']}
HDR.update(D['departments'])
LIGHT = {B['bright_teal'].upper()}        # needs dark text on it

painted = []
for hdr in (P['month_rows']['Total Income'] - 1, P['ytd_rows']['Total Income'] - 1):
    for c in range(2, 9):
        name = ws.cell(hdr, c).value
        col = HDR.get(str(name).strip()) if name else None
        if not col:
            continue
        cell = ws.cell(hdr, c)
        cell.fill = PatternFill('solid', fgColor=col)
        f = copy.copy(cell.font)
        f.color = openpyxl.styles.Color(rgb='FF' + (B['charcoal'] if col.upper()
                                                    in LIGHT else B['white']))
        cell.font = f
        painted.append(f'{name}={col}')
print('headings coloured:', len(painted))
for x in painted[:8]:
    print('  ', x)

# ------------------------------------------- 3. one stray balance sheet row -
# Office/Comp. Equip-at Cost was left at the top level, so it stayed on screen
# when Non-Current Assets was closed. Same rule as every other detail row.
bs = wb['Balance_Sheet']
fixed = 0
for r in range(6, bs.max_row + 1):
    v = bs.cell(r, 1).value
    if isinstance(v, str) and v.startswith(' ') and not (bs.row_dimensions[r].outlineLevel or 0):
        bs.row_dimensions[r].outlineLevel = 1
        bs.row_dimensions[r].hidden = True
        fixed += 1
        print('balance sheet row', r, 'grouped:', v.strip())
print('balance sheet rows corrected:', fixed)

wb.save(OUT)
print('saved', OUT)
