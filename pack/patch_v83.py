"""A recode check, so a change in Xero after the paste cannot go unseen.

A bill recoded from one account to another after month end leaves the company
result alone, so nothing in the pack moves. But the two accounts can sit on
different allocation keys, and then the departmental result moves with nothing
to show for it. That is what happened to July: 8,544.44 of First Focus IT went
from Dues & Subscriptions to IT Network Service & Support on 20 August, and
Support picked up 2,995.30 of overhead it had not carried when the pack was
built.

Two checks, because they catch different things.

Section 2 is the pack against Xero itself, month by month, on the five headline
lines. The Xero side is a hard pull with a date on it, and it has to be
refreshed to mean anything. It catches whatever has moved in Xero since the GL
was pasted, including a reclassification that nets to nothing at company level,
because it reads the largest line out rather than the net.

Section 3 is the GL paste against the pasted Xero P&L, account by account, for
all three months. Those two normally come out of the same export, so it reads
clean; it earns its keep when one is refreshed and the other is not, and it
names the account rather than just the month.
"""
import os
import sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter as gl

HERE = os.path.dirname(os.path.abspath(__file__))
WB = sys.argv[1] if len(sys.argv) > 1 else f'{HERE}/CTS Financial Controller Pack - v20.xlsx'

NAVY, WHITE, GREYL, RED = 'FF242C39', 'FFFFFFFF', 'FFF4F4F4', 'FFF33844'
BAND = PatternFill('solid', start_color=NAVY, end_color=NAVY)
TOTAL = PatternFill('solid', start_color=GREYL, end_color=GREYL)
RULE = Side(style='medium', color=NAVY)
MONEY = '#,##0.00;[RED]\\(#,##0.00\\);\\-'
BODY = Font(name='Calibri', sz=11, color=NAVY)
BOLD = Font(name='Calibri', sz=11, b=True, color=NAVY)
HEADF = Font(name='Calibri', sz=11, b=True, color=WHITE)
NOTE = Font(name='Calibri', sz=10, color='FF6B7585')
FLAG = Font(name='Calibri', sz=11, b=True, color=RED)

# Xero, accrual basis, pulled 26 October 2026, positive as Xero shows them.
PULLED = '26 October 2026'
XERO = {'Jul': (459048.71, 282514.16, 202397.94, 1583.82, -24279.57),
        'Aug': (603466.54, 335516.00, 208721.68,  861.35,  60090.21),
        'Sep': (561944.53, 374130.12, 230233.37,  955.03, -41463.93)}
for m, (inc, cos, exp, oi, net) in XERO.items():
    assert abs((inc - cos - exp + oi) - net) < 0.005, f'{m} does not add up'

# What the pull on that date showed as having moved, so she can act on it now.
# Positive as Xero shows them; the difference is the paste less Xero.
MOVED = [('Other Insurance', 'Aug', 5073.18, 5183.30),
         ('Staff Entertainment', 'Aug', 7620.26, 7355.63)]

LINES = [('Total income', 6, 1), ('Cost of sales', 59, -1),
         ('Operating expenses', 164, -1), ('Other income', 57, 1),
         ('Net profit', 180, 1)]
# PL_Paste: first row, last row, category, sign against the GL
SECT = [(8, 25, 'Income', 1), (29, 64, 'Cost of Sales', -1),
        (70, 120, 'Expenses', -1), (124, 125, 'Other Income', 1)]
MONTHS = ('Jul', 'Aug', 'Sep')
PASTE_COL = {'Jul': 'D', 'Aug': 'C', 'Sep': 'B'}      # PL_Paste runs newest first

S1, S2_BAND, S2_HEAD, S2_FIRST = 8, 17, 18, 19
S2_ROLL = S2_FIRST + len(LINES) * 3 + 1               # 35
S2_FLAG = S2_ROLL + 1
MV_BAND = S2_FLAG + 2
MV_HEAD = MV_BAND + 1
MV_FIRST = MV_HEAD + 1
MV_NET = MV_FIRST + len(MOVED)
S3_BAND = MV_NET + 3
S3_DATE, S3_HEAD, S3_FIRST = S3_BAND + 1, S3_BAND + 2, S3_BAND + 3
N_ACC = sum(b - a + 1 for a, b, _, _ in SECT)
S3_LAST = S3_FIRST + N_ACC - 1

wb = openpyxl.load_workbook(WB)
if 'Recode Check' in wb.sheetnames:
    del wb['Recode Check']
ws = wb.create_sheet('Recode Check', wb.sheetnames.index('PL_Check') + 1)
pv = openpyxl.load_workbook(WB, data_only=True)['PL_Paste']


def band(row, cols, text):
    ws.cell(row, 1, text).font = HEADF
    for c in range(1, cols + 1):
        ws.cell(row, c).fill = BAND


def money(row, col, value, fill=None):
    c = ws.cell(row, col, value)
    c.font = BODY
    c.number_format = MONEY
    if fill:
        c.fill = fill
    return c


ws['A1'] = 'Recode check'
ws['A1'].font = Font(name='Calibri', sz=14, b=True, color=NAVY)
for i, t in enumerate([
        'A bill recoded from one account to another after month end leaves the company '
        'result alone, so nothing in the pack moves.',
        'The two accounts can sit on different allocation keys, though, and then the '
        'departmental result moves with nothing to show for it.',
        'Section 2 is the pack against Xero, so it catches what has moved in Xero since '
        'the GL was pasted. Refresh the Xero rows or it means nothing.',
        'Section 3 is what the last Xero pull found. Section 4 is the GL paste against '
        'the pasted Xero P&L, account by account, which names the account itself.']):
    ws.cell(3 + i, 1, t).font = NOTE

# ---------------------------------------------------------------- section 1
band(S1, 5, '1. HAS ANYTHING MOVED')
rows1 = [
    ('Months out against Xero', f'=COUNTIF($C${S2_FLAG}:$E${S2_FLAG},"CHECK")', '0'),
    ('Largest line out against Xero, any month',
     f'=MAX($C${S2_ROLL}:$E${S2_ROLL})', MONEY),
    ('Accounts out against the P&L paste, section 4',
     f'=COUNTIF($L${S3_FIRST}:$L${S3_LAST},"CHECK")', '0'),
    ('Largest account out against the P&L paste, section 4',
     f'=MAX($M${S3_FIRST}:$M${S3_LAST})', MONEY),
]
for i, (lbl, f, fmt) in enumerate(rows1):
    ws.cell(S1 + 1 + i, 1, lbl).font = BODY
    c = ws.cell(S1 + 1 + i, 4, f)
    c.font = BODY
    c.number_format = fmt
r = S1 + 1 + len(rows1)
ws.cell(r, 1, 'Status').font = BOLD
c = ws.cell(r, 4, f'=IF(AND($D${S1 + 1}=0,$D${S1 + 3}=0),'
                  f'"OK, nothing has moved","CHECK, see the sections below")')
c.font = FLAG
c.alignment = Alignment(horizontal='left')
ws.cell(r + 2, 1, f'The Xero rows in section 2 were pulled on {PULLED}, accrual basis. '
                  f'Re-run the Xero P&L and overwrite them before relying on it.').font = NOTE

# ---------------------------------------------------------------- section 2
band(S2_BAND, 5, '2. THE PACK AGAINST XERO, BY MONTH')
ws.cell(S2_HEAD, 1, 'Line').font = BOLD
ws.cell(S2_HEAD, 2, 'Basis').font = BOLD
for i in range(3):
    c = ws.cell(S2_HEAD, 3 + i, f'=TEXT(EDATE(Setup!$B$5,{i - 2}),"mmm yyyy")')
    c.font = BOLD
    c.alignment = Alignment(horizontal='right')
for c in range(1, 6):
    ws.cell(S2_HEAD, c).border = Border(bottom=RULE)

r = S2_FIRST
for idx, (lbl, plrow, sign) in enumerate(LINES):
    ws.cell(r, 1, lbl).font = BODY
    ws.cell(r, 2, 'per the pack').font = NOTE
    ws.cell(r + 1, 2, 'per Xero').font = NOTE
    ws.cell(r + 2, 2, 'difference').font = NOTE
    for i, mon in enumerate(MONTHS):
        money(r, 3 + i, f"='P&L FY27'!{gl(2 + i)}{plrow}")
        money(r + 1, 3 + i, round(XERO[mon][idx] * sign, 2))
        money(r + 2, 3 + i, f'=ROUND({gl(3 + i)}{r}-{gl(3 + i)}{r + 1},2)')
        ws.cell(r + 2, 3 + i).font = BOLD
    r += 3

ws.cell(S2_ROLL, 1, 'Largest line out').font = BOLD
ws.cell(S2_ROLL, 2, 'the month').font = NOTE
diffs = [S2_FIRST + 2 + 3 * i for i in range(len(LINES))]
for i in range(3):
    col = gl(3 + i)
    money(S2_ROLL, 3 + i,
          '=MAX(' + ','.join(f'ABS({col}{d})' for d in diffs) + ')', TOTAL)
    ws.cell(S2_ROLL, 3 + i).font = BOLD
ws.cell(S2_FLAG, 1, 'Flag').font = BOLD
for i in range(3):
    c = ws.cell(S2_FLAG, 3 + i,
                f'=IF({gl(3 + i)}{S2_ROLL}>0.005,"CHECK","OK")')
    c.font = FLAG
    c.alignment = Alignment(horizontal='right')

# ------------------------------------------------- what had moved on the pull
band(MV_BAND, 5, f'3. WHAT HAD MOVED WHEN XERO WAS PULLED ON {PULLED.upper()}')
for i, t in enumerate(['Account', 'Month', 'per the paste', 'per Xero', 'difference']):
    c = ws.cell(MV_HEAD, 1 + i, t)
    c.font = BOLD
    c.border = Border(bottom=RULE)
    c.alignment = Alignment(horizontal='right' if i else 'left')
for i, (acct, mon, paste, xero) in enumerate(MOVED):
    r = MV_FIRST + i
    ws.cell(r, 1, acct).font = BODY
    ws.cell(r, 2, mon + ' 2026').font = BODY
    money(r, 3, paste)
    money(r, 4, xero)
    money(r, 5, round(paste - xero, 2))
ws.cell(MV_NET, 1, 'Net, which is the August figure in section 2').font = BOLD
money(MV_NET, 5, f'=ROUND(SUM($E${MV_FIRST}:$E${MV_NET - 1}),2)', TOTAL)
ws.cell(MV_NET, 5).font = BOLD
ws.cell(MV_NET + 1, 1,
        'This block is the pull on that date, so it does not move on its own. When a '
        'month flags in section 2, re-run the Xero P&L and paste it over PL_Paste. '
        'Section 4 then names the account itself.').font = NOTE

# ---------------------------------------------------------------- section 4
band(S3_BAND, 13, '4. EVERY ACCOUNT, THE GL AGAINST THE PASTED XERO P&L')
for i in range(3):
    c = ws.cell(S3_DATE, 3 + 3 * i, f'=EDATE(Setup!$B$5,{i - 2})')
    c.font = BOLD
    c.number_format = 'mmm yyyy'
head = ['Account', 'Category']
for _ in range(3):
    head += ['per the GL', 'per the P&L', 'difference']
head += ['Moved', 'Largest']
for i, t in enumerate(head):
    c = ws.cell(S3_HEAD, 1 + i, t)
    c.font = BOLD
    c.border = Border(bottom=RULE)
    c.alignment = Alignment(horizontal='right' if i else 'left', wrap_text=True)

CL_A = 'Cleanup!$A$5:$A$20004'
CL_I = 'Cleanup!$I$5:$I$20004'
CL_B = 'Cleanup!$B$5:$B$20004'
r = S3_FIRST
for a, b, cat, sign in SECT:
    for src in range(a, b + 1):
        assert pv.cell(src, 1).value, f'PL_Paste row {src} has no account'
        ws.cell(r, 1, f"='PL_Paste'!$A${src}").font = BODY
        ws.cell(r, 2, cat).font = NOTE
        for i, mon in enumerate(MONTHS):
            c0 = 3 + 3 * i
            money(r, c0, f'=SUMIFS({CL_A},{CL_I},$A{r},{CL_B},{gl(c0)}${S3_DATE})')
            money(r, c0 + 1, f"={sign}*'PL_Paste'!${PASTE_COL[mon]}${src}")
            money(r, c0 + 2, f'=ROUND({gl(c0)}{r}-{gl(c0 + 1)}{r},2)')
        c = ws.cell(r, 12, f'=IF($M{r}>0.005,"CHECK","")')
        c.font = FLAG
        c.alignment = Alignment(horizontal='right')
        money(r, 13, f'=MAX(ABS($E{r}),ABS($H{r}),ABS($K{r}))')
        r += 1
assert r - 1 == S3_LAST, f'{r - 1} vs {S3_LAST}'

ws.column_dimensions['A'].width = 40
ws.column_dimensions['B'].width = 15
for c in range(3, 14):
    ws.column_dimensions[gl(c)].width = 15
ws.freeze_panes = ws[f'C{S3_HEAD + 1}']
ws.sheet_view.showGridLines = False

wb.save(WB)
print(f'Recode Check: section 2 rows {S2_FIRST} to {S2_FLAG}, '
      f'{N_ACC} accounts in rows {S3_FIRST} to {S3_LAST}')
