"""Add the reconciling line beside the equipment line on the Aware Super job.

Danica has marked the equipment line yellow. It reads as a loss of 63,332.58
because 124,635.75 of equipment was bought and only 61,303.17 was invoiced as
equipment, with the rest invoiced as installation labour.

This puts the reclassification next to it, and the same entry the other way on
the labour line, so the two reconcile back to the job. The job's invoiced
total, cost and margin do not move. Only the split between the two lines does.

The basis is equipment at nil margin, which is the least that has to move for
the equipment line to stop reading as a loss. If equipment should instead earn
the job's own 20.2%, more has to move again, and the note on the sheet says by
how much.

Everything else in the workbook is left exactly as she saved it.
"""
from copy import copy

import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side

OUT = 'recon/CTS revenue against cost - Jul to Sep 2026.xlsx'
TAB = 'Aware Super job'

NAVY, GREY, YELLOW = 'FF242C39', 'FFF4F4F4', 'FFFFFF00'
BAND = PatternFill('solid', start_color=NAVY, end_color=NAVY)
TOTF = PatternFill('solid', start_color=GREY, end_color=GREY)
MARK = PatternFill('solid', start_color=YELLOW, end_color=YELLOW)
RULE = Side(style='thin', color='FFC9CED6')
HEAD = Font(name='Calibri', sz=11, b=True, color='FFFFFFFF')
BOLD = Font(name='Calibri', sz=11, b=True, color=NAVY)
BODY = Font(name='Calibri', sz=11, color=NAVY)
NOTE = Font(name='Calibri', sz=10, i=True, color='FF6B7585')
MONEY = '#,##0.00;[RED](#,##0.00);"-"'
PCT = '0.0%;[RED](0.0%);"-"'

wb = openpyxl.load_workbook(OUT)
ws = wb[TAB]

# clear the old block and rebuild it from row 89
for r in range(89, ws.max_row + 2):
    for c in range(1, 8):
        cell = ws.cell(r, c)
        cell.value = None
        cell.fill = PatternFill()
        cell.border = Border()

r = 89
ws.cell(r, 1, 'HOW THE SPLIT GOES WRONG, AND WHAT RECONCILES IT').font = HEAD
for c in range(1, 8):
    ws.cell(r, c).fill = BAND
r += 1
for c, h in ((1, 'Line'), (5, 'Invoiced'), (6, 'Cost'), (7, 'Margin')):
    cell = ws.cell(r, c, h)
    cell.font = BOLD
    cell.border = Border(bottom=RULE)
r += 1

EQ_INV, EQ_COST = 61303.17, 124635.75
LB_INV, LB_COST = 144049.00, 39317.06

blocks = [
    # label, invoiced, cost, style
    ('Equipment, as invoiced', EQ_INV, EQ_COST, 'mark'),
    ('Reclassify to equipment, the shortfall above', '=F{0}-E{0}'.format(r), None, 'move'),
    ('Equipment, reconciled', None, None, 'sub'),
    (None, None, None, 'gap'),
    ('Labour and other, as invoiced', LB_INV, LB_COST, 'plain'),
    ('Reclassify out of labour, the same entry the other way', None, None, 'move2'),
    ('Labour and other, reconciled', None, None, 'sub2'),
    (None, None, None, 'gap'),
    ('The job', None, None, 'total'),
]
rows = {}
for label, inv, cost, kind in blocks:
    if kind == 'gap':
        r += 1
        continue
    ws.cell(r, 1, label).font = BOLD if kind in ('sub', 'sub2', 'total') else BODY
    if kind == 'mark':
        ws.cell(r, 5, inv).number_format = MONEY
        ws.cell(r, 6, cost).number_format = MONEY
        for c in range(1, 8):
            ws.cell(r, c).fill = MARK
        rows['eq'] = r
    elif kind == 'move':
        ws.cell(r, 5, f'=F{rows["eq"]}-E{rows["eq"]}').number_format = MONEY
        ws.cell(r, 6, 0).number_format = MONEY
        rows['eqmove'] = r
    elif kind == 'sub':
        ws.cell(r, 5, f'=E{rows["eq"]}+E{rows["eqmove"]}').number_format = MONEY
        ws.cell(r, 6, f'=F{rows["eq"]}+F{rows["eqmove"]}').number_format = MONEY
        rows['eqnew'] = r
    elif kind == 'plain':
        ws.cell(r, 5, inv).number_format = MONEY
        ws.cell(r, 6, cost).number_format = MONEY
        rows['lb'] = r
    elif kind == 'move2':
        ws.cell(r, 5, f'=-E{rows["eqmove"]}').number_format = MONEY
        ws.cell(r, 6, 0).number_format = MONEY
        rows['lbmove'] = r
    elif kind == 'sub2':
        ws.cell(r, 5, f'=E{rows["lb"]}+E{rows["lbmove"]}').number_format = MONEY
        ws.cell(r, 6, f'=F{rows["lb"]}+F{rows["lbmove"]}').number_format = MONEY
        rows['lbnew'] = r
    elif kind == 'total':
        ws.cell(r, 5, f'=E{rows["eqnew"]}+E{rows["lbnew"]}').number_format = MONEY
        ws.cell(r, 6, f'=F{rows["eqnew"]}+F{rows["lbnew"]}').number_format = MONEY
        rows['tot'] = r
    cell = ws.cell(r, 7, f'=E{r}-F{r}')
    cell.number_format = MONEY
    cell.font = BOLD if kind in ('sub', 'sub2', 'total') else BODY
    if kind in ('sub', 'sub2'):
        for c in range(1, 8):
            ws.cell(r, c).border = Border(top=RULE)
            ws.cell(r, c).font = BOLD
    if kind == 'total':
        for c in range(1, 8):
            ws.cell(r, c).fill = TOTF
            ws.cell(r, c).font = BOLD
    for c in (5, 6):
        ws.cell(r, c).font = copy(ws.cell(r, 1).font)
    r += 1

# the margin percentages beside the reconciled lines
for key, col in (('eq', 8), ('eqnew', 8), ('lb', 8), ('lbnew', 8), ('tot', 8)):
    cell = ws.cell(rows[key], col, f'=IF(E{rows[key]}=0,"",G{rows[key]}/E{rows[key]})')
    cell.number_format = PCT
    cell.font = BOLD if key in ('eqnew', 'lbnew', 'tot') else BODY
ws.cell(90, 8, 'Margin %').font = BOLD
ws.cell(90, 8).border = Border(bottom=RULE)
ws.cell(89, 8).fill = BAND
ws.cell(rows['tot'], 8).fill = TOTF
ws.cell(rows['eq'], 8).fill = MARK

r += 1
EVEN_EQ = EQ_COST / (1 - (205352.17 - 163952.81) / 205352.17)
for t in [f'The yellow line is the one that is wrong. {EQ_COST:,.2f} of '
          f'equipment was bought and {EQ_INV:,.2f} was invoiced as equipment. '
          f'The rest of the job was invoiced as installation labour.',
          'The reclassification moves the shortfall from labour to equipment. '
          'The job\'s invoiced total, its cost and its margin do not change. '
          'Only the split between the two lines does.',
          'This puts equipment at nil margin, which is the least that has to '
          'move for the line to stop reading as a loss. If equipment should '
          f'earn the job\'s own 20.2%, the equipment line would be '
          f'{EVEN_EQ:,.0f} and labour {205352.17 - EVEN_EQ:,.0f}, so more again '
          'would have to move.',
          'The right split is whatever the quote actually priced. That is the '
          'thing to check, and it is the same question on every job where the '
          'supplier bills one lump sum and the client is invoiced another.']:
    ws.cell(r, 1, t).font = NOTE
    r += 1

ws.column_dimensions['G'].width = 15
ws.column_dimensions['H'].width = 12
wb.save(OUT)
print('saved', OUT)
print('  equipment line marked at row', rows['eq'])
print('  reconciling line at row', rows['eqmove'])
