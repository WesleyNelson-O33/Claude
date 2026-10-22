"""
Revenue by month across years, and a first quarter comparison.

The pack holds two years of actual revenue: FY27 on the P&L FY27 tab and FY26
on the Actual FY26 tab, which already aligns last year's months onto this
year's column headers. FY23, FY24 and FY25 are not in the pack and are not in
any file I hold, so those three columns are left empty for a paste.

Both charts read whichever year columns have figures in them, so pasting a year
in is all that is needed. Nothing has to be rebuilt.

FY26 and FY27 are references rather than pastes, so they follow the pack.
"""
import os, sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

HERE = os.path.dirname(os.path.abspath(__file__))
WB = sys.argv[1] if len(sys.argv) > 1 else f'{HERE}/CTS Financial Controller Pack - v20.xlsx'

NAVY, WHITE, GREYL, AMBER = 'FF242C39', 'FFFFFFFF', 'FFF4F4F4', 'FFFFF4CE'
BAND = PatternFill('solid', start_color=NAVY, end_color=NAVY)
TOTAL = PatternFill('solid', start_color=GREYL, end_color=GREYL)
PASTE = PatternFill('solid', start_color=AMBER, end_color=AMBER)
RULE = Side(style='medium', color=NAVY)
MONEY = '#,##0;[RED]\\(#,##0\\);\\-'

YEARS = ['FY23', 'FY24', 'FY25', 'FY26', 'FY27']
BANDROW, HEAD, FIRST = 280, 281, 282          # twelve months of revenue by year
LAST = FIRST + 11
TOTROW = LAST + 1
Q1BAND, Q1HEAD, Q1FIRST = 296, 297, 298       # the first quarter, year by year

wb = openpyxl.load_workbook(WB)
ws = wb['Cost Analysis']
assert ws.cell(260, 1).value == 'REVENUE AGAINST BUDGET, MONTH BY MONTH'

def band(row, cols, text):
    ws.cell(row, 1, text)
    for c in range(1, cols + 1):
        cell = ws.cell(row, c)
        cell.fill = BAND
        cell.font = Font(name='Calibri', sz=12, b=True, color=WHITE)

def header(row, labels):
    for c, text in enumerate(labels, start=1):
        cell = ws.cell(row, c, text)
        cell.fill = BAND
        cell.font = Font(name='Calibri', sz=10, b=True, color=WHITE)
        cell.border = Border(bottom=RULE)
        cell.alignment = Alignment(horizontal='left' if c == 1 else 'right')

band(BANDROW, 6, 'REVENUE BY MONTH, YEAR ON YEAR')
header(HEAD, ['Month'] + YEARS)
for i, r in enumerate(range(FIRST, LAST + 1)):
    col = chr(ord('B') + i)                    # B..M on the yearly tabs
    ws.cell(r, 1, f"=TEXT('P&L FY27'!{col}$5,\"mmm\")")
    ws.cell(r, 1).font = Font(name='Calibri', sz=11, color=NAVY)
    for j, year in enumerate(YEARS):
        cell = ws.cell(r, 2 + j)
        if year == 'FY27':
            cell.value = f"='P&L FY27'!{col}6"
        elif year == 'FY26':
            cell.value = f"='Actual FY26'!{col}6"
        else:
            cell.fill = PASTE                  # waiting on the Xero export
        cell.number_format = MONEY
        cell.font = Font(name='Calibri', sz=11, color=NAVY)

ws.cell(TOTROW, 1, 'Full year')
for j in range(len(YEARS)):
    col = chr(ord('B') + j)
    ws.cell(TOTROW, 2 + j, f'=SUM({col}{FIRST}:{col}{LAST})').number_format = MONEY
for c in range(1, 7):
    cell = ws.cell(TOTROW, c)
    cell.fill = TOTAL
    cell.font = Font(name='Calibri', sz=11, b=True, color=NAVY)

band(Q1BAND, 3, 'FIRST QUARTER, YEAR ON YEAR')
header(Q1HEAD, ['Year', 'Jul to Sep'])
for i, year in enumerate(YEARS):
    r = Q1FIRST + i
    ws.cell(r, 1, year).font = Font(name='Calibri', sz=11, color=NAVY)
    col = chr(ord('B') + i)
    ws.cell(r, 2, f'=SUM({col}{FIRST}:{col}{FIRST + 2})').number_format = MONEY
    ws.cell(r, 2).font = Font(name='Calibri', sz=11, b=(year == 'FY27'), color=NAVY)

wb.save(WB)
print(f'Cost Analysis rows {BANDROW} to {Q1FIRST + 4} added: revenue year on year '
      f'and the first quarter, with FY23 to FY25 awaiting a paste')
