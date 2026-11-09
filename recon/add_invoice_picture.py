"""Aware Super invoice INV-10596: what it was, what it is now, all on formulas.

Every figure on this page is calculated. The only inputs are the transaction
lines already on the Aware Super job tab, which this reads with SUMIF. Change a
line there and this page follows.

Goes in as a tab in front of that detail. Nothing else in the workbook moves.
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as gl

OUT = 'recon/CTS revenue against cost - Jul to Sep 2026.xlsx'
SRC = 'Aware Super job'
TAB = 'Aware invoice, was and now'

# the line detail on the source tab: codes in B, amounts in F
RC, RA = f"'{SRC}'!$B$6:$B$20", f"'{SRC}'!$F$6:$F$20"      # revenue
CC, CA = f"'{SRC}'!$B$25:$B$80", f"'{SRC}'!$F$25:$F$80"    # cost

NAVY, WHITE, GREY = 'FF242C39', 'FFFFFFFF', 'FFF4F4F4'
MINT, AMBER = 'FFE8F4F3', 'FFFFF4D6'
BAND = PatternFill('solid', start_color=NAVY, end_color=NAVY)
TOTF = PatternFill('solid', start_color=GREY, end_color=GREY)
WAS = PatternFill('solid', start_color=AMBER, end_color=AMBER)
NOW = PatternFill('solid', start_color=MINT, end_color=MINT)
RULE = Side(style='thin', color='FFC9CED6')
THICK = Side(style='medium', color=NAVY)
HEAD = Font(name='Calibri', sz=11, b=True, color=WHITE)
BOLD = Font(name='Calibri', sz=11, b=True, color=NAVY)
BODY = Font(name='Calibri', sz=11, color=NAVY)
NOTE = Font(name='Calibri', sz=10, i=True, color='FF6B7585')
MONEY = '#,##0.00;[RED](#,##0.00);"-"'
PCT = '0.0%;[RED](0.0%);"-"'


def band(ws, row, c1, text):
    ws.cell(row, 1, text).font = HEAD
    for c in range(1, c1 + 1):
        ws.cell(row, c).fill = BAND


def put(ws, r, c, formula, bold=False, fill=None, fmt=MONEY):
    cell = ws.cell(r, c, formula)
    cell.font = BOLD if bold else BODY
    cell.number_format = fmt
    if fill:
        cell.fill = fill
    return cell


wb = openpyxl.load_workbook(OUT)
if TAB in wb.sheetnames:
    del wb[TAB]
ws = wb.create_sheet(TAB, wb.sheetnames.index(SRC))

ws['A1'] = 'Aware Super, invoice INV-10596'
ws['A1'].font = Font(name='Calibri', sz=16, b=True, color=NAVY)
ws['A2'] = ('Job 2604704, Seminar Room Audio. Raised 9 September 2026. The '
            'supplier bill is Claim 2 of 3, so a third claim is still to come.')
ws['A2'].font = NOTE
ws['A3'] = (f"Every figure is a formula. The only inputs are the lines on the "
            f"{SRC} tab, read by account code.")
ws['A3'].font = NOTE

# --------------------------------------------------- what was invoiced
r = 5
band(ws, r, 9, 'WHAT WAS INVOICED TO AWARE SUPER')
r += 1
INV = [('42300  Installation labour', f'=SUMIF({RC},"42300",{RA})'),
       ('42600  Equipment sales, the charges',
        f'=SUMIFS({RA},{RC},"42600",{RA},">0")'),
       ('42600  Equipment sales, the credits',
        f'=SUMIFS({RA},{RC},"42600",{RA},"<0")'),
       ('42800  Consulting, accrued then reversed, nets to nil',
        f'=SUMIF({RC},"42800",{RA})')]
inv_first = r
for label, f in INV:
    ws.cell(r, 1, label).font = BODY
    put(ws, r, 4, f)
    r += 1
ws.cell(r, 1, 'Invoice total').font = BOLD
put(ws, r, 4, f'=SUM(D{inv_first}:D{r - 1})', bold=True)
for c in range(1, 5):
    ws.cell(r, c).fill = TOTF
R_INV = r

# ------------------------------------------------------- what it cost
r += 2
band(ws, r, 9, 'WHAT IT COST')
r += 1
COST = [('51100  Equipment purchase, Complete AV Solutions INV41565',
         f'=SUMIF({CC},"51100",{CA})'),
        ('51325  Sub-contract labour, same bill', f'=SUMIF({CC},"51325",{CA})'),
        ('58225  Freight, same bill', f'=SUMIF({CC},"58225",{CA})'),
        ('55025 and 56025  Own staff, payroll journals',
         f'=SUMIF({CC},"55025",{CA})+SUMIF({CC},"56025",{CA})')]
cost_first = r
for label, f in COST:
    ws.cell(r, 1, label).font = BODY
    put(ws, r, 4, f)
    r += 1
R_EQC = cost_first                       # the equipment purchase row
ws.cell(r, 1, 'Cost total').font = BOLD
put(ws, r, 4, f'=SUM(D{cost_first}:D{r - 1})', bold=True)
for c in range(1, 5):
    ws.cell(r, c).fill = TOTF
R_COST = r

# ----------------------------------------------- was and now, side by side
r += 2
band(ws, r, 9, 'THE TWO LINES, BEFORE AND AFTER')
r += 1
ws.cell(r, 2, 'AS INVOICED').font = BOLD
ws.cell(r, 6, 'RECONCILED, EQUIPMENT AT COST').font = BOLD
for c in range(2, 6):
    ws.cell(r, c).fill = WAS
for c in range(6, 10):
    ws.cell(r, c).fill = NOW
r += 1
for c, h in ((1, 'Line'), (2, 'Invoiced'), (3, 'Cost'), (4, 'Margin'), (5, '%'),
             (6, 'Invoiced'), (7, 'Cost'), (8, 'Margin'), (9, '%')):
    cell = ws.cell(r, c, h)
    cell.font = BOLD
    cell.border = Border(bottom=THICK)
    if c > 1:
        cell.alignment = Alignment(horizontal='right')
r += 1

R_EQ, R_LB = r, r + 1
ws.cell(R_EQ, 1, 'Equipment').font = BODY
put(ws, R_EQ, 2, f'=D{inv_first + 1}+D{inv_first + 2}', fill=WAS)
put(ws, R_EQ, 3, f'=D{R_EQC}', fill=WAS)
put(ws, R_EQ, 6, f'=C{R_EQ}', fill=NOW)       # equipment invoiced at its cost
put(ws, R_EQ, 7, f'=C{R_EQ}', fill=NOW)

ws.cell(R_LB, 1, 'Labour and other').font = BODY
put(ws, R_LB, 2, f'=D{R_INV}-B{R_EQ}', fill=WAS)
put(ws, R_LB, 3, f'=D{R_COST}-C{R_EQ}', fill=WAS)
put(ws, R_LB, 6, f'=D{R_INV}-F{R_EQ}', fill=NOW)
put(ws, R_LB, 7, f'=D{R_COST}-G{R_EQ}', fill=NOW)

for rr in (R_EQ, R_LB):
    put(ws, rr, 4, f'=B{rr}-C{rr}', fill=WAS)
    put(ws, rr, 5, f'=IF(B{rr}=0,"",D{rr}/B{rr})', fill=WAS, fmt=PCT)
    put(ws, rr, 8, f'=F{rr}-G{rr}', fill=NOW)
    put(ws, rr, 9, f'=IF(F{rr}=0,"",H{rr}/F{rr})', fill=NOW, fmt=PCT)

r = R_LB + 1
R_JOB = r
ws.cell(r, 1, 'The job').font = BOLD
for c in (2, 3, 4, 6, 7, 8):
    put(ws, r, c, f'=SUM({gl(c)}{R_EQ}:{gl(c)}{R_LB})', bold=True)
for c in (5, 9):
    base = 'B' if c == 5 else 'F'
    num = 'D' if c == 5 else 'H'
    put(ws, r, c, f'={num}{r}/{base}{r}', bold=True, fmt=PCT)
for c in range(1, 10):
    ws.cell(r, c).fill = TOTF
    ws.cell(r, c).border = Border(top=RULE)
r += 1
ws.cell(r, 1, 'Moved from labour to equipment').font = BOLD
put(ws, r, 6, f'=F{R_EQ}-B{R_EQ}', bold=True, fill=NOW)

# ------------------------------------------------------ the other basis
r += 2
band(ws, r, 9, "IF EQUIPMENT SHOULD EARN THE JOB'S OWN MARGIN INSTEAD")
r += 1
ws.cell(r, 1, 'The version above puts equipment at nil margin, which is the '
              'least that has to move. This is the other end of it, with both '
              "lines on the job's own margin.").font = NOTE
r += 1
for c, h in ((6, 'Invoiced'), (7, 'Cost'), (8, 'Margin'), (9, '%')):
    cell = ws.cell(r, c, h)
    cell.font = BOLD
    cell.border = Border(bottom=RULE)
    cell.alignment = Alignment(horizontal='right')
r += 1
R_EQ2, R_LB2 = r, r + 1
ws.cell(R_EQ2, 1, 'Equipment').font = BODY
put(ws, R_EQ2, 6, f'=C{R_EQ}/(1-E{R_JOB})')
put(ws, R_EQ2, 7, f'=C{R_EQ}')
ws.cell(R_LB2, 1, 'Labour and other').font = BODY
put(ws, R_LB2, 6, f'=B{R_JOB}-F{R_EQ2}')
put(ws, R_LB2, 7, f'=C{R_LB}')
for rr in (R_EQ2, R_LB2):
    put(ws, rr, 8, f'=F{rr}-G{rr}')
    put(ws, rr, 9, f'=IF(F{rr}=0,"",H{rr}/F{rr})', fmt=PCT)
r = R_LB2 + 1
ws.cell(r, 1, 'The job').font = BOLD
for c in (6, 7, 8):
    put(ws, r, c, f'=SUM({gl(c)}{R_EQ2}:{gl(c)}{R_LB2})', bold=True)
put(ws, r, 9, f'=H{r}/F{r}', bold=True, fmt=PCT)
for c in range(1, 10):
    ws.cell(r, c).fill = TOTF
r += 1
ws.cell(r, 1, 'Moved from labour to equipment').font = BOLD
put(ws, r, 6, f'=F{R_EQ2}-B{R_EQ}', bold=True)

# --------------------------------------------------------------- notes
r += 2
for t in ['What happened. The supplier billed one lump sum and it was split '
          'mostly to equipment. Aware Super was invoiced one lump sum and it '
          'was split mostly to labour. Same job, two different splits.',
          'What does not change. The invoice total, the cost and the job '
          'margin are the same in every version above. Only the split between '
          'the two lines moves.',
          'What does change. On the company profit and loss, equipment reads '
          'as a 23.5% loss for the quarter. This job is the whole of that.',
          'What to check. What the quote actually priced for equipment. That '
          'is the real split, and the same question applies on every job where '
          'the supplier bills one lump sum and the client is invoiced another.']:
    ws.cell(r, 1, t).font = NOTE
    r += 1

for c, w in ((1, 52), (2, 15), (3, 15), (4, 15), (5, 9),
             (6, 15), (7, 15), (8, 15), (9, 9)):
    ws.column_dimensions[gl(c)].width = w
ws.sheet_view.showGridLines = False
wb.save(OUT)
print('added tab:', TAB)
print(f'  invoice total row {R_INV}, cost total row {R_COST}, job row {R_JOB}')
