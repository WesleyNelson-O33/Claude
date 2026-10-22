"""
Budget on the revenue by month chart.

The stack shows which department billed each month. The budget line across it
answers the next question, which is whether the month was ahead or behind, and
by how much, without needing a second chart.

Budget comes off the Budget FY27 total income row for the same month, on the
same closed month test as the actuals, so the line stops where the columns do.
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
FIRST, LAST, HEAD, BANDROW, TOTROW = 232, 243, 231, 230, 244

wb = openpyxl.load_workbook(WB)
ws = wb['Cost Analysis']
assert ws.cell(BANDROW, 1).value == 'REVENUE BY MONTH AND DEPARTMENT'

cell = ws.cell(HEAD, 7, 'Budget')
cell.fill = BAND
cell.font = Font(name='Calibri', sz=10, b=True, color=WHITE)
cell.border = Border(bottom=RULE)
cell.alignment = Alignment(horizontal='right')
ws.cell(BANDROW, 7).fill = BAND

for r in range(FIRST, LAST + 1):
    cell = ws.cell(r, 7)
    cell.value = (f'=IF($A{r}<=Setup!$B$5,'
                  f"SUMIF('Budget FY27'!$B$5:$M$5,$A{r},'Budget FY27'!$B6:$M6),\"\")")
    cell.number_format = MONEY
    cell.font = Font(name='Calibri', sz=11, color=NAVY)

ws.cell(TOTROW, 7, f'=SUM(G{FIRST}:G{LAST})').number_format = MONEY
ws.cell(TOTROW, 7).fill = TOTAL
ws.cell(TOTROW, 7).font = Font(name='Calibri', sz=11, b=True, color=NAVY)

wb.save(WB)
print('Cost Analysis column G added: budget for the revenue by month chart')
