"""
Same month across years, and the quarter labelled by calendar year.

What was wanted is September 2022 beside September 2023 beside September 2024,
not twelve months with a line per year. So the reporting month is pulled out of
the year on year table and set out one column per year, and it follows Setup:
in October it becomes October against October.

The year labels are the calendar month and year rather than the financial year,
which is how the comparison was described. September 2026 sits in FY27, so FY27
reads Sep 2026. The quarter labels follow the same rule: the first quarter of
FY27 is July to September 2026, so it reads Q1 2026.
"""
import os, sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

HERE = os.path.dirname(os.path.abspath(__file__))
WB = sys.argv[1] if len(sys.argv) > 1 else f'{HERE}/CTS Financial Controller Pack - v20.xlsx'

NAVY, WHITE = 'FF242C39', 'FFFFFFFF'
BAND = PatternFill('solid', start_color=NAVY, end_color=NAVY)
RULE = Side(style='medium', color=NAVY)
MONEY = '#,##0;[RED]\\(#,##0\\);\\-'

FY = [23, 24, 25, 26, 27]
MFIRST, MLAST = 282, 293          # the year on year table this reads from
BANDROW, HEAD, FIRST = 316, 317, 318
Q1FIRST = 298

# the calendar year the month falls in: July to June means the first half of a
# financial year sits in the previous calendar year
CAL = "(2000+{fy}-IF(MONTH(Setup!$B$5)>=Setup!$B$6,1,0))"

wb = openpyxl.load_workbook(WB)
ws = wb['Cost Analysis']
assert ws.cell(281, 1).value == 'Month'

ws.cell(BANDROW, 1, 'REVENUE FOR THE MONTH ON SETUP, YEAR ON YEAR')
for c in range(1, 4):
    cell = ws.cell(BANDROW, c)
    cell.fill = BAND
    cell.font = Font(name='Calibri', sz=12, b=True, color=WHITE)

for c, text in enumerate(['Month', 'Revenue'], start=1):
    cell = ws.cell(HEAD, c, text)
    cell.fill = BAND
    cell.font = Font(name='Calibri', sz=10, b=True, color=WHITE)
    cell.border = Border(bottom=RULE)
    cell.alignment = Alignment(horizontal='left' if c == 1 else 'right')

for i, fy in enumerate(FY):
    r = FIRST + i
    col = chr(ord('B') + i)
    ws.cell(r, 1, f'=TEXT(Setup!$B$5,"mmm")&" "&{CAL.format(fy=fy)}')
    ws.cell(r, 1).font = Font(name='Calibri', sz=11, color=NAVY)
    ws.cell(r, 2, f'=IFERROR(INDEX({col}${MFIRST}:{col}${MLAST},'
                  f'MATCH(TEXT(Setup!$B$5,"mmm"),$A${MFIRST}:$A${MLAST},0)),"")')
    ws.cell(r, 2).number_format = MONEY
    ws.cell(r, 2).font = Font(name='Calibri', sz=11, b=(fy == 27), color=NAVY)

# the quarter, relabelled by the calendar year it falls in
for i, fy in enumerate(FY):
    ws.cell(Q1FIRST + i, 1, f'="Q1 "&(2000+{fy}-1)')
    ws.cell(Q1FIRST + i, 1).font = Font(name='Calibri', sz=11, color=NAVY)

wb.save(WB)
print(f'Cost Analysis rows {BANDROW} to {FIRST + 4} added: the reporting month '
      f'year on year; the quarter relabelled by calendar year')
