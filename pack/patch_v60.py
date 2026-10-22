"""
Revenue against budget, month by month.

The first attempt at this ranked the departments for the year to date. What was
asked for was each closed month with its actual and its budget side by side:
July actual and July budget, August actual and August budget, and so on.

The block runs the full twelve months so nothing has to be added later; the
chart takes the closed ones. The year to date pair sits on the end of the table
for reference and is quoted in the commentary rather than charted, because it
cannot sit next to the months in one contiguous series.
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
PCT = '0.0%;[RED]-0.0%;\\-'
BANDROW, HEAD, FIRST, LAST = 260, 261, 262, 273
YTD = 275

wb = openpyxl.load_workbook(WB)
ws = wb['Cost Analysis']
assert ws.cell(250, 1).value == 'REVENUE AGAINST BUDGET, YEAR TO DATE'

ws.cell(BANDROW, 1, 'REVENUE AGAINST BUDGET, MONTH BY MONTH')
for c in range(1, 6):
    cell = ws.cell(BANDROW, c)
    cell.fill = BAND
    cell.font = Font(name='Calibri', sz=12, b=True, color=WHITE)

for c, text in enumerate(['Month', 'Actual', 'Budget', 'Var $', 'Var %'], start=1):
    cell = ws.cell(HEAD, c, text)
    cell.fill = BAND
    cell.font = Font(name='Calibri', sz=10, b=True, color=WHITE)
    cell.border = Border(bottom=RULE)
    cell.alignment = Alignment(horizontal='left' if c == 1 else 'right')

for i, r in enumerate(range(FIRST, LAST + 1)):
    col = chr(ord('B') + i)
    ws.cell(r, 1, f"='P&L FY27'!{col}$5").number_format = 'mmm-yy'
    ws.cell(r, 1).font = Font(name='Calibri', sz=11, color=NAVY)
    ws.cell(r, 2, f'=IF($A{r}<=Setup!$B$5,'
                  f"SUMIF('P&L FY27'!$B$5:$M$5,$A{r},'P&L FY27'!$B6:$M6),\"\")")
    ws.cell(r, 3, f'=IF($A{r}<=Setup!$B$5,'
                  f"SUMIF('Budget FY27'!$B$5:$M$5,$A{r},'Budget FY27'!$B6:$M6),\"\")")
    ws.cell(r, 4, f'=IF($B{r}="","",$B{r}-$C{r})')
    ws.cell(r, 5, f'=IF($B{r}="","",IFERROR($D{r}/$C{r},0))')
    for c in range(2, 5):
        ws.cell(r, c).number_format = MONEY
    ws.cell(r, 5).number_format = PCT
    for c in range(2, 6):
        ws.cell(r, c).font = Font(name='Calibri', sz=11, color=NAVY)

ws.cell(YTD, 1, 'Year to date')
ws.cell(YTD, 2, f'=SUM(B{FIRST}:B{LAST})').number_format = MONEY
ws.cell(YTD, 3, f'=SUM(C{FIRST}:C{LAST})').number_format = MONEY
ws.cell(YTD, 4, f'=B{YTD}-C{YTD}').number_format = MONEY
ws.cell(YTD, 5, f'=IFERROR(D{YTD}/C{YTD},0)').number_format = PCT
for c in range(1, 6):
    cell = ws.cell(YTD, c)
    cell.fill = TOTAL
    cell.font = Font(name='Calibri', sz=11, b=True, color=NAVY)

wb.save(WB)
print(f'Cost Analysis rows {BANDROW} to {YTD} added: revenue against budget, month by month')
