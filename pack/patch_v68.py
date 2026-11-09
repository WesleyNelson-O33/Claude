"""Spend and profit blocks for the cost and profit pages.

Three charts come off these: spend for the month against budget and the two
years before it, the same for the first quarter, and net profit above and below
the line for both.

Spend is cost of sales plus overheads, which is row 59 plus row 164 on the
yearly tabs. Net profit is row 180, so it is the pack's own figure and includes
other income. Everything follows Setup: the month comes from the period number
in B8, so in October this becomes October without being touched.

FY26 and FY27 are the pack's own tabs. FY25 is not in the pack, so the first
quarter of it was pulled from Xero on 22 October 2026 and sits in its own block
with the three months showing. Only that quarter was pulled, so the FY25 line
goes blank from period four onward rather than showing something that was never
checked.
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
BOLD = Font(name='Calibri', sz=11, b=True, color=NAVY)
NOTE = Font(name='Calibri', sz=10, color='FF6B7585')

# Xero, accrual basis, pulled 22 October 2026. Cost of sales, overheads,
# revenue, net profit. Net profit is revenue less both, plus other income.
FY25_Q1 = [
    ('Jul 2024', 222400.77, 227574.26, 351628.82, -97579.89),
    ('Aug 2024', 269327.95, 207880.28, 588230.29, 111437.22),
    ('Sep 2024', 216749.49, 203160.95, 440243.09,  20737.43),
]

XQ_BAND, XQ_HEAD, XQ_FIRST = 110, 111, 112          # the Xero quarter
XQ_TOT = XQ_FIRST + 3
M_BAND, M_HEAD, M_FIRST = 118, 119, 120             # the month on Setup
Q_BAND, Q_HEAD, Q_FIRST = 126, 127, 128             # the first quarter
PM_BAND, PM_HEAD, PM_FIRST = 134, 135, 136          # net profit, the month
PQ_BAND, PQ_HEAD, PQ_FIRST = 140, 141, 142          # net profit, the quarter

MONTHS = ('mmm')
COLS = ['Period', 'Cost of sales', 'Overheads', 'Revenue', 'Net profit', 'Total spend']
ROWS = {'cos': 59, 'ovh': 164, 'rev': 6, 'np': 180}

wb = openpyxl.load_workbook(WB)
ws = wb['Xero History']
assert ws['A9'].value == 'Month'

def band(row, cols, text):
    ws.cell(row, 1, text)
    for c in range(1, cols + 1):
        ws.cell(row, c).fill = BAND
        ws.cell(row, c).font = Font(name='Calibri', sz=12, b=True, color=WHITE)

def header(row, labels):
    for c, t in enumerate(labels, start=1):
        cell = ws.cell(row, c, t)
        cell.fill = BAND
        cell.font = Font(name='Calibri', sz=10, b=True, color=WHITE)
        cell.border = Border(bottom=RULE)
        cell.alignment = Alignment(horizontal='left' if c == 1 else 'right')

def money(row, col, value, font=BODY):
    cell = ws.cell(row, col, value)
    cell.number_format = MONEY
    cell.font = font

# ---- the Xero quarter, held as figures because the pack has no FY25 tab -----
band(XQ_BAND, 5, 'FY25 FIRST QUARTER, PULLED FROM XERO')
header(XQ_HEAD, COLS[:5])
for i, (lbl, cos, ovh, rev, np_) in enumerate(FY25_Q1):
    r = XQ_FIRST + i
    ws.cell(r, 1, lbl).font = BODY
    for c, v in enumerate((cos, ovh, rev, np_), start=2):
        money(r, c, round(v, 2))
ws.cell(XQ_TOT, 1, 'Q1 total').font = BOLD
for c in range(2, 6):
    money(XQ_TOT, c, f'=SUM({chr(64+c)}{XQ_FIRST}:{chr(64+c)}{XQ_FIRST+2})', BOLD)
for c in range(1, 6):
    ws.cell(XQ_TOT, c).fill = TOTAL
ws.cell(XQ_FIRST - 1, 7, 'Only this quarter was pulled. The FY25 line below goes '
                         'blank from period four until the rest is pulled.').font = NOTE

# ---- the month on Setup -----------------------------------------------------
YR = 'YEAR(Setup!$B$5)'
band(M_BAND, 6, 'SPEND AND PROFIT, THE MONTH ON SETUP')
header(M_HEAD, COLS)
SRC = [
    (f'=TEXT(Setup!$B$5,"mmm")&" "&({YR}-2)', None),
    (f'=TEXT(Setup!$B$5,"mmm")&" "&({YR}-1)', 'Actual FY26'),
    (f'=TEXT(Setup!$B$5,"mmm")&" "&{YR}', 'P&L FY27'),
    (f'="Budget "&TEXT(Setup!$B$5,"mmm")&" "&{YR}', 'Budget FY27'),
]
for i, (label, tab) in enumerate(SRC):
    r = M_FIRST + i
    ws.cell(r, 1, label).font = BOLD if tab == 'P&L FY27' else BODY
    for c, key in enumerate(('cos', 'ovh', 'rev', 'np'), start=2):
        sign = '-' if key in ('cos', 'ovh') else ''
        if tab is None:
            f = (f'=IF(Setup!$B$8>3,"",INDEX(${chr(64+c)}${XQ_FIRST}:'
                 f'${chr(64+c)}${XQ_FIRST+2},Setup!$B$8))')
        else:
            f = (f"={sign}INDEX('{tab}'!$B${ROWS[key]}:$M${ROWS[key]},Setup!$B$8)")
        money(r, c, f, BOLD if tab == 'P&L FY27' else BODY)
    money(r, 6, f'=IF(B{r}="","",B{r}+C{r})', BOLD if tab == 'P&L FY27' else BODY)

# ---- the first quarter ------------------------------------------------------
band(Q_BAND, 6, 'SPEND AND PROFIT, THE FIRST QUARTER')
header(Q_HEAD, COLS)
QSRC = [
    (f'="Q1 "&(Setup!$B$7-3)', None),
    (f'="Q1 "&(Setup!$B$7-2)', 'Actual FY26'),
    (f'="Q1 "&(Setup!$B$7-1)', 'P&L FY27'),
    (f'="Budget Q1 "&(Setup!$B$7-1)', 'Budget FY27'),
]
for i, (label, tab) in enumerate(QSRC):
    r = Q_FIRST + i
    ws.cell(r, 1, label).font = BOLD if tab == 'P&L FY27' else BODY
    for c, key in enumerate(('cos', 'ovh', 'rev', 'np'), start=2):
        sign = '-' if key in ('cos', 'ovh') else ''
        if tab is None:
            f = f'=${chr(64+c)}${XQ_TOT}'
        else:
            f = f"={sign}SUM('{tab}'!$B${ROWS[key]}:$D${ROWS[key]})"
        money(r, c, f, BOLD if tab == 'P&L FY27' else BODY)
    money(r, 6, f'=B{r}+C{r}', BOLD if tab == 'P&L FY27' else BODY)

# ---- net profit, the three months then the three quarters -------------------
# The month and the quarter are two separate blocks, because the deck carries a
# monthly slide and a quarterly slide and either one can be hidden. They must
# never share a chart.
#
# Profit and loss are split into their own columns so the chart can colour them
# without the build script having to read the figures. A profit lands in the
# profit column and a loss in the loss column, the other goes blank, and the two
# series sit on top of each other so each period shows one bar.
def profit_block(band_row, head_row, first_row, title, src_first):
    band(band_row, 4, title)
    header(head_row, ['Period', 'Net profit', 'Profit', 'Loss'])
    for i in range(3):
        r = first_row + i
        ws.cell(r, 1, f'=$A${src_first + i}').font = BODY
        money(r, 2, f'=$E${src_first + i}')
        money(r, 3, f'=IF(N($B{r})>0,$B{r},"")')
        money(r, 4, f'=IF(N($B{r})<=0,$B{r},"")')

profit_block(PM_BAND, PM_HEAD, PM_FIRST,
             'NET PROFIT AND NET LOSS, THE MONTH', M_FIRST)
profit_block(PQ_BAND, PQ_HEAD, PQ_FIRST,
             'NET PROFIT AND NET LOSS, THE QUARTER', Q_FIRST)

wb.save(WB)
print(f'Xero History rows {XQ_BAND} to {PQ_FIRST + 2} added: spend and profit for '
      f'the month and the quarter, and net profit and net loss as two blocks')
