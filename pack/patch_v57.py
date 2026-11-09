"""
Revenue by month and department, as a chart block.

The deck shows revenue three ways already: the income table on the P&L summary,
each department against budget for the month, and the company total by month
against budget. What it has not had is the mix: how much of each month's revenue
came from which department.

This block carries it. The months come off the P&L FY27 header row so it follows
the pack, and each month is filled in only once it is closed on Setup, which is
the same convention the income by month block uses.
"""
import os, sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

HERE = os.path.dirname(os.path.abspath(__file__))
WB = sys.argv[1] if len(sys.argv) > 1 else f'{HERE}/CTS Financial Controller Pack - v20.xlsx'

NAVY, WHITE, GREYL = 'FF242C39', 'FFFFFFFF', 'FFF4F4F4'
BAND = PatternFill('solid', start_color=NAVY, end_color=NAVY)
TOTAL = PatternFill('solid', start_color=GREYL, end_color=GREYL)
RULE = Side(style='medium', color=NAVY)
MONEY = '#,##0;[RED]\\(#,##0\\);\\-'

# department -> its income row on the P&L FY27 tab
DEPTS = [('Support', 41), ('Events', 8), ('Video', 16),
         ('Integration', 25), ('Consulting', 33)]
# kept clear of the revenue by department chart, which sits at F190
FIRST, LAST = 232, 243            # twelve months
HEAD, BANDROW, TOTROW = 231, 230, 244

wb = openpyxl.load_workbook(WB)
ws = wb['Cost Analysis']
assert ws['A190'].value.startswith('REVENUE BY DEPARTMENT'), ws['A190'].value

# clear the first attempt, which sat under the chart above it
from openpyxl.styles import PatternFill as _P
for r in range(205, 220):
    for c in range(1, 8):
        cell = ws.cell(r, c)
        cell.value = None
        cell.fill = _P(fill_type=None)
        cell.border = Border()

ws.cell(BANDROW, 1, 'REVENUE BY MONTH AND DEPARTMENT')
for c in range(1, 7):
    cell = ws.cell(BANDROW, c)
    cell.fill = BAND
    cell.font = Font(name='Calibri', sz=12, b=True, color=WHITE)

for c, text in enumerate(['Month'] + [d for d, _ in DEPTS], start=1):
    cell = ws.cell(HEAD, c, text)
    cell.fill = BAND
    cell.font = Font(name='Calibri', sz=10, b=True, color=WHITE)
    cell.border = Border(bottom=RULE)
    cell.alignment = Alignment(horizontal='left' if c == 1 else 'right')

for i, r in enumerate(range(FIRST, LAST + 1)):
    col = chr(ord('B') + i)                      # B..M on the P&L FY27 header
    ws.cell(r, 1, f"='P&L FY27'!{col}$5").number_format = 'mmm-yy'
    ws.cell(r, 1).font = Font(name='Calibri', sz=11, color=NAVY)
    for j, (_, src) in enumerate(DEPTS):
        cell = ws.cell(r, 2 + j)
        cell.value = (f'=IF($A{r}<=Setup!$B$5,'
                      f"SUMIF('P&L FY27'!$B$5:$M$5,$A{r},'P&L FY27'!$B{src}:$M{src}),\"\")")
        cell.number_format = MONEY
        cell.font = Font(name='Calibri', sz=11, color=NAVY)

ws.cell(TOTROW, 1, 'Total')
for j in range(len(DEPTS)):
    col = chr(ord('B') + j)
    ws.cell(TOTROW, 2 + j, f'=SUM({col}{FIRST}:{col}{LAST})').number_format = MONEY
for c in range(1, 7):
    cell = ws.cell(TOTROW, c)
    cell.fill = TOTAL
    cell.font = Font(name='Calibri', sz=11, b=True, color=NAVY)

wb.save(WB)
print(f'Cost Analysis rows {BANDROW} to {TOTROW} added: revenue by month and department')
