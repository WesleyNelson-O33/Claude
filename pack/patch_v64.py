"""
A Xero History tab, and the FY25 column of the year on year tables filled from it.

The year on year tables were built with FY23, FY24 and FY25 left empty for a
paste, because the pack holds only FY26 and FY27. FY25 is now pulled from Xero
and lives on its own tab so the source is visible and so the figure beside each
month cannot drift from the month it belongs to.

Xero holds the last three financial years only. FY25, FY26 and FY27 can be
pulled. FY23 and FY24 cannot, so those two columns stay as a paste.

The FY26 and FY27 months on this tab are references to the pack's own actuals
tabs, not a second copy of them. Three of those months were checked against
Xero and agree: July 2025 at 320,808.06, August 2025 at 551,450.89, September
2025 at 356,323.23 against 356,323.22 in the pack, and September 2026 at
561,944.53 on both sides.
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
BODY = Font(name='Calibri', sz=11, color=NAVY)
NOTE = Font(name='Calibri', sz=10, color='FF6B7585')

# Total Income, accrual basis, pulled from Xero on 22 October 2026.
FY25 = [
    ('Jul 2024', 351628.82), ('Aug 2024', 588230.29), ('Sep 2024', 440243.09),
    ('Oct 2024', 948991.15), ('Nov 2024', 579564.81), ('Dec 2024', 748430.22),
    ('Jan 2025', 918288.99), ('Feb 2025', 566037.74), ('Mar 2025', 389012.14),
    ('Apr 2025', 391802.42), ('May 2025', 593566.87), ('Jun 2025', 484560.44),
]

SHEET = 'Xero History'
HEADROW, FIRST = 9, 10            # thirty six months, Jul 2024 to Jun 2027
LAST = FIRST + 35
TBAND, THEAD, TFIRST = LAST + 2, LAST + 3, LAST + 4   # the three year totals

wb = openpyxl.load_workbook(WB)
if SHEET in wb.sheetnames:
    del wb[SHEET]
ws = wb.create_sheet(SHEET, wb.sheetnames.index('Budget FY26') + 1)
ws.sheet_properties.tabColor = '005358'

def band(row, cols, text):
    ws.cell(row, 1, text)
    for c in range(1, cols + 1):
        cell = ws.cell(row, c)
        cell.fill = BAND
        cell.font = Font(name='Calibri', sz=12, b=True, color=WHITE)

def header(row, labels, money_from=2):
    for c, text in enumerate(labels, start=1):
        cell = ws.cell(row, c, text)
        cell.fill = BAND
        cell.font = Font(name='Calibri', sz=10, b=True, color=WHITE)
        cell.border = Border(bottom=RULE)
        cell.alignment = Alignment(horizontal='right' if c >= money_from else 'left')

band(1, 5, 'REVENUE HISTORY FROM XERO')
for i, line in enumerate([
    'Revenue here is Total Income on the accrual basis, the same line the pack calls '
    'Total Income. It excludes other income such as interest.',
    'FY25 was pulled from Xero on 22 October 2026 and is held as figures, because the '
    'pack has no FY25 tab. Those figures do not move when the pack rolls forward.',
    'FY26 and FY27 are references to the Actual FY26 and P&L FY27 tabs, so they follow '
    'the pack. FY27 shows closed months only.',
    'Xero holds the last three financial years. FY23 and FY24 cannot be pulled, so those '
    'two columns on Cost Analysis are still a paste from an older pack.',
], start=3):
    ws.cell(i, 1, line).font = NOTE

band(HEADROW - 1, 5, 'MONTHLY REVENUE, THREE FINANCIAL YEARS')
header(HEADROW, ['Month', 'Financial year', 'Revenue', 'Source'], money_from=3)
ws.cell(HEADROW, 4).alignment = Alignment(horizontal='left')

for i, (label, value) in enumerate(FY25):
    r = FIRST + i
    ws.cell(r, 1, label).font = BODY
    ws.cell(r, 2, 'FY25').font = BODY
    ws.cell(r, 3, round(value, 2)).font = BODY
    ws.cell(r, 3).number_format = MONEY
    ws.cell(r, 4, 'Xero, pulled 22 Oct 2026').font = NOTE

for i in range(12):                                   # FY26, from the pack
    r = FIRST + 12 + i
    col = chr(ord('B') + i)
    ws.cell(r, 1, f"=TEXT('Actual FY26'!{col}$5,\"mmm\")&\" \"&(YEAR('Actual FY26'!{col}$5)-1)")
    ws.cell(r, 1).font = BODY
    ws.cell(r, 2, 'FY26').font = BODY
    ws.cell(r, 3, f"='Actual FY26'!{col}6").font = BODY
    ws.cell(r, 3).number_format = MONEY
    ws.cell(r, 4, 'Pack, Actual FY26').font = NOTE

for i in range(12):                                   # FY27, closed months only
    r = FIRST + 24 + i
    col = chr(ord('B') + i)
    ws.cell(r, 1, f"=TEXT('P&L FY27'!{col}$5,\"mmm\")&\" \"&YEAR('P&L FY27'!{col}$5)")
    ws.cell(r, 1).font = BODY
    ws.cell(r, 2, 'FY27').font = BODY
    ws.cell(r, 3, f"=IF('P&L FY27'!{col}$5<=Setup!$B$5,'P&L FY27'!{col}6,\"\")").font = BODY
    ws.cell(r, 3).number_format = MONEY
    ws.cell(r, 4, 'Pack, P&L FY27').font = NOTE

band(TBAND, 5, 'FINANCIAL YEAR TOTALS')
header(THEAD, ['Financial year', 'Full year', 'First quarter'])
for i, year in enumerate(['FY25', 'FY26', 'FY27']):
    r, src = TFIRST + i, FIRST + 12 * i
    ws.cell(r, 1, year).font = Font(name='Calibri', sz=11, b=True, color=NAVY)
    ws.cell(r, 2, f'=SUM(C{src}:C{src + 11})')
    ws.cell(r, 3, f'=SUM(C{src}:C{src + 2})')
    for c in (2, 3):
        ws.cell(r, c).number_format = MONEY
        ws.cell(r, c).font = Font(name='Calibri', sz=11, b=True, color=NAVY)
    for c in range(1, 6):
        ws.cell(r, c).fill = TOTAL

for col, w in (('A', 16), ('B', 15), ('C', 14), ('D', 26), ('E', 4)):
    ws.column_dimensions[col].width = w
ws.sheet_view.showGridLines = False
ws.freeze_panes = f'A{FIRST}'

# the FY25 column of the year on year table now reads this tab
ca = wb['Cost Analysis']
assert ca.cell(281, 4).value == 'FY25'
for i in range(12):
    cell = ca.cell(282 + i, 4)
    cell.value = f"='{SHEET}'!$C${FIRST + i}"
    cell.fill = PatternFill(fill_type=None)
    cell.number_format = MONEY
    cell.font = BODY
ca.cell(294, 1, 'FY23 and FY24 are outside the three financial years Xero holds, so those '
                'two columns are still a paste. FY25 comes from the Xero History tab.')
ca.cell(294, 1).font = NOTE

wb.save(WB)
print(f'{SHEET} tab added: thirty six months of revenue, FY25 from Xero and '
      f'FY26 and FY27 from the pack. Cost Analysis FY25 column wired to it.')
