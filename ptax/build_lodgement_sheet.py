"""The October 2026 payroll tax return, laid out the way Revenue NSW shows
it, so the figures can be typed straight off the page and checked after.

Same fields and the same order as the return history summary that came back
for September, with September alongside for comparison.

The figures come from the split file, not from the Employment Hero payroll
tax report, because the report leaves out the two ad hoc pay runs for the
period 19 September to 2 October. That money went out in October, so it is
October wages. The report gives 5,800.27 and the return is 5,996.99.
"""
import sys

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as gl

OUT = 'ptax/CTS payroll tax return - October 2026.xlsx'

NAVY, WHITE, GREY = 'FF242C39', 'FFFFFFFF', 'FFF4F4F4'
BAND = PatternFill('solid', start_color=NAVY, end_color=NAVY)
TOTF = PatternFill('solid', start_color=GREY, end_color=GREY)
AMBER = PatternFill('solid', start_color='FFFFF3D6', end_color='FFFFF3D6')
GREEN = PatternFill('solid', start_color='FFE4F2E4', end_color='FFE4F2E4')
RULE = Side(style='thin', color='FFC9CED6')
BOX = Border(top=RULE, bottom=RULE, left=RULE, right=RULE)
HEAD = Font(name='Calibri', sz=11, b=True, color=WHITE)
BOLD = Font(name='Calibri', sz=11, b=True, color=NAVY)
BODY = Font(name='Calibri', sz=11, color=NAVY)
NOTE = Font(name='Calibri', sz=10, i=True, color='FF6B7585')
TITLE = Font(name='Calibri', sz=14, b=True, color=NAVY)
MONEY = '#,##0.00;[RED](#,##0.00);"-"'

RATE, THRESHOLD, DAYS_YEAR, DAYS = 0.0545, 1200000, 365, 31
SAL, SUPER = 169021.21, 22087.22
QLD, SA, VIC, WA, ACT = 14489.10, 6345.37, 19799.31, 8505.86, 0.0

DETAILS = [
    ('Calculation source', 'Web Form'),
    ('Calculation period', '1 Oct 2026 to 31 Oct 2026'),
    ('Lodgement version number', '1'),
    ('Return due date', '7 Nov 2026'),
    ('Calculation based on', 'Actual'),
    ('Government department?', 'No'),
    ('Claiming threshold?', 'Yes'),
    ('Reason for claiming threshold', 'Not grouped'),
    ('Entitled to rebate for GP contractor payments?', 'No'),
]

# label, what to do, the October figure or None for a nil box,
# and what September came back as
LINES = [
    ('Salaries and wages', 'type', SAL, 158232.73),
    ('Fringe benefits', 'nil', 0.0, 0.0),
    ('Employer superannuation contribution', 'type', SUPER, 24019.72),
    ('Termination payments', 'nil', 0.0, 0.0),
    ('Contractor payments', 'nil', 0.0, 0.0),
    ('Employment agency payments', 'nil', 0.0, 0.0),
    ('Allowances', 'nil', 0.0, 0.0),
    ('Bonuses/Commissions', 'nil', 0.0, 0.0),
    ('Directors fees', 'nil', 0.0, 0.0),
    ('Shares and options', 'nil', 0.0, 0.0),
    ('NSW apprentice and trainee wages', 'nil', 0.0, 0.0),
]


def band(ws, row, labels, widths=None, height=30):
    for i, t in enumerate(labels, start=1):
        c = ws.cell(row, i, t)
        c.fill, c.font, c.border = BAND, HEAD, BOX
        c.alignment = Alignment(horizontal='center', wrap_text=True,
                                vertical='center')
    ws.row_dimensions[row].height = height
    if widths:
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[gl(i)].width = w


def para(ws, row, text, cols=5, font=None, height=28):
    c = ws.cell(row, 1, text)
    c.font = font or BODY
    c.alignment = Alignment(wrap_text=True, vertical='top')
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=cols)
    ws.row_dimensions[row].height = height


def main():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'October 2026 return'
    ws['A1'] = 'Payroll tax return, October 2026'
    ws['A1'].font = TITLE
    ws['A2'] = 'Corporate Technology Services Pty Ltd, client 119723076'
    ws['A2'].font = BODY
    ws['A3'] = ('Laid out the way Revenue NSW shows it, so the figures can be '
                'typed straight off this page and checked against the screen '
                'after.')
    ws['A3'].font = NOTE
    ws.merge_cells('A3:E3')
    ws.row_dimensions[3].height = 28

    r = 5
    para(ws, r, 'RETURN DETAILS', font=BOLD, height=18)
    for i, (k, v) in enumerate(DETAILS):
        n = r + 1 + i
        ws.cell(n, 1, k).font = BODY
        c = ws.cell(n, 2, v)
        c.font = BOLD
        c.alignment = Alignment(horizontal='left')
    r = r + 1 + len(DETAILS) + 1

    para(ws, r, 'NSW WAGE COMPONENTS', font=BOLD, height=18)
    hdr = r + 1
    band(ws, hdr, ['Line', 'October 2026', 'What to do',
                   'September, as it came back', 'Movement'],
         [46, 16, 22, 24, 14])
    first = hdr + 1
    for i, (lbl, what, oct_v, sep_v) in enumerate(LINES):
        n = first + i
        ws.cell(n, 1, lbl).font = BODY
        c = ws.cell(n, 2, oct_v)
        c.number_format, c.font = MONEY, BOLD if what == 'type' else BODY
        ws.cell(n, 3, 'Type this in' if what == 'type' else 'Leave at nil'
                ).font = BODY
        d = ws.cell(n, 4, sep_v)
        d.number_format, d.font = MONEY, BODY
        e = ws.cell(n, 5, f'=B{n}-D{n}')
        e.number_format, e.font = MONEY, BODY
        if what == 'type':
            for k in range(1, 6):
                ws.cell(n, k).fill = AMBER
    last = first + len(LINES) - 1

    rows = [
        ('Total NSW wages', f'=SUM(B{first}:B{last})', 182252.45,
         'The portal works this out'),
        ('Interstate wages', f'=B{last+6}', 61800.00, 'Type this in'),
        ('Total Australian wages', None, 244052.45,
         'The portal works this out'),
        ('Threshold entitlement', None, 73654.59,
         'The portal works this out'),
        ('NSW apprentice and trainee offset', '0', 0.0, 'Leave at nil'),
        ('GP contractor rebate', '0', 0.0, 'Leave at nil'),
        ('Tax payable', None, 5918.58, 'The portal works this out'),
    ]
    n = last + 1
    tn, iw = n, n + 1
    ws.cell(n, 1, rows[0][0]).font = BOLD
    ws.cell(n, 2, rows[0][1]).number_format = MONEY
    ws.cell(n, 2).font = BOLD
    ws.cell(n, 3, rows[0][3]).font = BODY
    ws.cell(n, 4, rows[0][2]).number_format = MONEY
    ws.cell(n, 5, f'=B{n}-D{n}').number_format = MONEY

    n += 1
    ws.cell(n, 1, 'Interstate wages').font = BOLD
    ws.cell(n, 2).number_format = MONEY   # filled in below, once the
    #                                       interstate block exists
    ws.cell(n, 2).font = BOLD
    ws.cell(n, 3, 'Type this in').font = BODY
    ws.cell(n, 4, 61800.00).number_format = MONEY
    ws.cell(n, 5, f'=B{n}-D{n}').number_format = MONEY
    for k in range(1, 6):
        ws.cell(n, k).fill = AMBER

    n += 1
    taw = n
    ws.cell(n, 1, 'Total Australian wages').font = BOLD
    ws.cell(n, 2, f'=B{tn}+B{iw}').number_format = MONEY
    ws.cell(n, 3, 'The portal works this out').font = BODY
    ws.cell(n, 4, 244052.45).number_format = MONEY
    ws.cell(n, 5, f'=B{n}-D{n}').number_format = MONEY

    n += 1
    thr = n
    ws.cell(n, 1, 'Threshold entitlement').font = BOLD
    ws.cell(n, 2, f'=ROUND((B{tn}/B{taw})*({THRESHOLD}/{DAYS_YEAR}*{DAYS}),2)'
            ).number_format = MONEY
    ws.cell(n, 3, 'The portal works this out').font = BODY
    ws.cell(n, 4, 73654.59).number_format = MONEY
    ws.cell(n, 5, f'=B{n}-D{n}').number_format = MONEY

    for lbl, sep_v in (('NSW apprentice and trainee offset', 0.0),
                       ('GP contractor rebate', 0.0)):
        n += 1
        ws.cell(n, 1, lbl).font = BODY
        ws.cell(n, 2, 0).number_format = MONEY
        ws.cell(n, 3, 'Leave at nil').font = BODY
        ws.cell(n, 4, sep_v).number_format = MONEY
        ws.cell(n, 5, f'=B{n}-D{n}').number_format = MONEY

    n += 1
    ws.cell(n, 1, 'TAX PAYABLE').font = BOLD
    ws.cell(n, 2, f'=ROUND((B{taw}-B{thr}-B{iw})*{RATE},2)'
            ).number_format = MONEY
    ws.cell(n, 3, 'The portal works this out').font = BODY
    ws.cell(n, 4, 5918.58).number_format = MONEY
    ws.cell(n, 5, f'=B{n}-D{n}').number_format = MONEY
    for k in range(1, 6):
        ws.cell(n, k).fill, ws.cell(n, k).border = TOTF, BOX
        ws.cell(n, k).font = BOLD
    tax = n

    n += 2
    para(ws, n, 'THE INTERSTATE FIGURE', font=BOLD, height=18)
    n += 1
    band(ws, n, ['State', 'October 2026', '', '', ''])
    istart = n + 1
    for i, (st, v) in enumerate((('QLD', QLD), ('SA', SA), ('VIC', VIC),
                                 ('WA', WA), ('ACT', ACT))):
        m = istart + i
        ws.cell(m, 1, st).font = BODY
        c = ws.cell(m, 2, v)
        c.number_format, c.font = MONEY, BODY
        if st == 'ACT':
            ws.cell(m, 3, 'nil this month').font = NOTE
    m = istart + 5
    ws.cell(m, 1, 'Interstate wages').font = BOLD
    c = ws.cell(m, 2, f'=SUM(B{istart}:B{m-1})')
    c.number_format, c.font, c.fill, c.border = MONEY, BOLD, TOTF, BOX
    ws.cell(m, 1).fill, ws.cell(m, 1).border = TOTF, BOX
    # now the interstate line up in the return can point at that total
    ws.cell(iw, 2).value = f'=B{m}'

    n = m + 2
    para(ws, n, 'CHECK IT AGAINST THE SCREEN BEFORE YOU SUBMIT', font=BOLD,
         height=18)
    for i, (lbl, ref) in enumerate((('Total NSW wages', tn),
                                    ('Total Australian wages', taw),
                                    ('Threshold entitlement', thr),
                                    ('Tax payable', tax))):
        q = n + 1 + i
        ws.cell(q, 1, lbl).font = BODY
        c = ws.cell(q, 2, f'=B{ref}')
        c.number_format, c.font = MONEY, BOLD
        ws.cell(q, 3, 'should match the portal').font = NOTE
    n = n + 6

    para(ws, n, 'WHY THIS DIFFERS FROM THE EMPLOYMENT HERO REPORT',
         font=BOLD, height=18)
    for i, t in enumerate([
        'The Employment Hero payroll tax report gives 5,800.27. This return '
        'is 5,996.99.',
        'Two ad hoc pay runs for the period 19 September to 2 October carry '
        'September dates in Employment Hero but were paid in October. The '
        'report leaves them out, so they appear in no payroll tax report at '
        'all.',
        'They are 3,846.15 of New South Wales wages and 300.00 of South '
        'Australian wages, which is 196.72 of tax.',
        'Wages paid in October are October wages, so they belong here. '
        'Lodging on the report alone would understate the return.',
        'The full working is on the Oct 2026 derivation tab of the payroll '
        'tax workbook.',
    ]):
        para(ws, n + 1 + i, t, height=30)

    ws.freeze_panes = f'A{hdr+1}'
    wb.save(OUT)
    print(f'  written to {OUT}')
    return 0


if __name__ == '__main__':
    sys.exit(main())
