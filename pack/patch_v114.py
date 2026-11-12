"""Put the Q1 department P&L into the controller pack.

Danica's own layout and headings: the five departments across the top under
their reporting group banners, two years side by side, Sales through to Gross
Profit, then the overhead block and the result underneath.

Nothing is typed. Every figure is a formula into the pack's own sheets:

    departments, both years   'P&L FY27' and 'Actual FY26', columns B to D
    group overheads           'P&L by Dept'
    headcount                 'Overhead Split'

The three cost of sales headings are built from the pack's sixteen cost
lines, in this grouping:

    Labour            Sub-Contract Labour, Direct Wages & Salaries, Direct
                      Salaries, Direct Wages, Direct Superannuation,
                      Direct Workers Comp
    Equipment & Misc  Equipment Purchase, Equipment Hires, Service hires,
                      Subscriptions & Licences, Equipment Maintenance
                      Support, Stamp Duty, Damage Waiver, Supplier Per
                      Diems, Credit Card fees paid COGS
    Transport         Other Direct Expenses

That grouping is shown on the sheet so it can be checked. It does not tie
exactly to the old report, so it may need moving.

The lower half is this year only. FY26 overheads were never split by
department, which the pack says itself at 'Overhead Split' row 179.

Run it with no arguments.
"""
# Every Font here names Calibri, or LibreOffice renders the default as serif.
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

PACK = 'CTS Financial Controller Pack - v20.xlsx'
SHEET = 'Q1 Dept P&L'
INK, TEAL, BAND, LINE = '242C39', '3FD0C9', 'EDEFF2', 'C9CED6'
GREEN, AMBER, GREY = 'DCEFDC', 'FFF2CC', 'F2F2F2'
F27, F26 = "'P&L FY27'", "'Actual FY26'"
BYD = "'P&L by Dept'"
OS = "'Overhead Split'"

# income total row, first of the sixteen cost of sales rows
DEPT = {'Support': (41, 131), 'Events': (8, 62), 'Video': (16, 79),
        'Integration': (25, 97), 'Consulting': (33, 114), 'General Admin': (49, 148)}
LAB, EQ, TR = (2, 9, 10, 11, 12, 13), (0, 1, 3, 4, 5, 6, 7, 8, 15), (14,)
COLS = [('Support', 3), ('Events', 5), ('Video', 7),
        ('Integration', 9), ('Consulting', 11), ('CTS', 13)]
BAN = [('Support', 3, 4), ('Production', 5, 8), ('Consulting', 9, 12), ('CTS', 13, 14)]
SUB = [('Managed Services', 3, 4), ('Events', 5, 6), ('Video', 7, 8),
       ('Integration', 9, 10), ('Consulting', 11, 12), ('CTS', 13, 14)]
# overhead block base row, gross profit row, on 'P&L by Dept'
GRP = {'Support': (187, 197), 'Production': (60, 70),
       'Consulting': (137, 147), 'General Admin': (237, 247)}
GCOLS = [('Support', 3, 4), ('Production', 5, 8), ('Consulting', 9, 12), ('CTS', 13, 14)]
HEADROW = {'Support': 6, 'Events': 7, 'Video': 8, 'Integration': 9,
           'Consulting': 10, 'Admin': 11}
NUM = '#,##0;[Red](#,##0)'
PCT = '0.0%;[Red]-0.0%'


def q(sheet, row):
    return 'SUM(%s!B%d:D%d)' % (sheet, row, row)


def cos(sheet, dept, offsets):
    first = DEPT[dept][1]
    return '+'.join(q(sheet, first + o) for o in offsets)


def main():
    pk = openpyxl.load_workbook(PACK)
    if SHEET in pk.sheetnames:
        del pk[SHEET]
    ws = pk.create_sheet(SHEET, pk.sheetnames.index('P&L by Dept') + 1)
    thin = Side(style='thin', color=LINE)

    def put(r, c, v, bold=False, italic=False, fmt=None, fill=None, right=False, size=9):
        cell = ws.cell(r, c, v)
        cell.font = Font(name='Calibri', bold=bold, italic=italic, size=size, color=INK)
        if fmt:
            cell.number_format = fmt
        if fill:
            cell.fill = PatternFill('solid', fgColor=fill)
        cell.alignment = Alignment(horizontal='right' if right else 'left')
        cell.border = Border(bottom=thin)
        return cell

    ws.cell(1, 2, 'Q1 DEPARTMENT P&L').font = Font(name='Calibri', bold=True, size=11, color=INK)
    ws.cell(2, 2, 'Q1 FY27 (Jul to Sep 2026) against Q1 FY26 (Jul to Sep 2025). '
                  'Quarterly, not monthly.').font = Font(
        name='Calibri', italic=True, size=9, color='7F7F7F')

    for name, c1, c2 in BAN:
        cell = ws.cell(4, c1, name)
        cell.font = Font(name='Calibri', bold=True, size=10, color='FFFFFF')
        cell.fill = PatternFill('solid', fgColor=INK)
        cell.alignment = Alignment(horizontal='center')
        cell.border = Border(bottom=Side(style='medium', color=TEAL))
        ws.merge_cells(start_row=4, start_column=c1, end_row=4, end_column=c2)
    for name, c1, c2 in SUB:
        cell = ws.cell(5, c1, name)
        cell.font = Font(name='Calibri', bold=True, size=9, color='FFFFFF')
        cell.fill = PatternFill('solid', fgColor='3C4858')
        cell.alignment = Alignment(horizontal='center')
        ws.merge_cells(start_row=5, start_column=c1, end_row=5, end_column=c2)
    put(6, 2, 'Q1 (Jul - Sep)', bold=True)
    for _, c1, _ in SUB:
        for j, lab in enumerate(('Q1 FY27', 'Q1 FY26')):
            cell = ws.cell(6, c1 + j, lab)
            cell.font = Font(name='Calibri', bold=True, size=9, color='FFFFFF')
            cell.fill = PatternFill('solid', fgColor=INK)
            cell.alignment = Alignment(horizontal='center')

    def row(r, label, fn, fmt=NUM, bold=False, italic=False, fill=None, ind=0):
        put(r, 2, ('    ' * ind) + label, bold=bold, italic=italic, fill=fill)
        for name, c1 in COLS:
            for j, sh in enumerate((F27, F26)):
                put(r, c1 + j, fn(name, sh), bold=bold, italic=italic,
                    fmt=fmt, fill=fill, right=True)

    r = 8
    put(r, 2, 'Sales', bold=True, fill=GREY)
    for c in range(3, 15):
        put(r, c, None, fill=GREY)
    r += 1
    sales = lambda n, sh: '=' + q(sh, DEPT[n][0] if n != 'CTS' else 6)
    row(r, 'Sales Revenue', sales, bold=True)
    sr = r
    r += 1
    row(r, 'Income share',
        lambda n, sh: '=IFERROR(%s/%s,"")' % (q(sh, DEPT[n][0] if n != 'CTS' else 6), q(sh, 6)),
        fmt=PCT, italic=True)
    r += 2
    put(r, 2, 'Cost of Sales', bold=True, fill=GREY)
    for c in range(3, 15):
        put(r, c, None, fill=GREY)
    r += 1
    for lab, offs in (('Labour', LAB), ('Equipment & Misc', EQ), ('Transport', TR)):
        row(r, lab, lambda n, sh, o=offs: '=-(%s)' % (
            cos(sh, n, o) if n != 'CTS'
            else '+'.join(cos(sh, d, o) for d in DEPT)), ind=1)
        r += 1
    row(r, 'Total Cost of Sales',
        lambda n, sh: '=-' + (q(sh, DEPT[n][1] - 1) if n != 'CTS' else q(sh, 59)), bold=True)
    r += 2
    gp = r
    row(r, 'Gross Profit', lambda n, sh: '=%s+%s' % (
        q(sh, DEPT[n][0] if n != 'CTS' else 6),
        q(sh, (DEPT[n][1] - 1) if n != 'CTS' else 59)), bold=True, fill=GREEN)
    r += 1
    row(r, 'Gross Profit %', lambda n, sh, g=gp, s=sr: '=IFERROR(%s%d/%s%d,"")' % (
        '', 0, '', 0), fmt=PCT, italic=True, fill=GREEN)
    for name, c1 in COLS:
        for j in (0, 1):
            cl = L(c1 + j)
            ws.cell(r, c1 + j, '=IFERROR(%s%d/%s%d,"")' % (cl, gp, cl, sr))
    r += 2

    cell = ws.cell(r, 2, 'OVERHEADS AND RESULT, BY REPORTING GROUP, Q1 FY27 ONLY')
    cell.font = Font(name='Calibri', bold=True, size=10, color='FFFFFF')
    cell.fill = PatternFill('solid', fgColor=INK)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=14)
    r += 1
    cell = ws.cell(r, 2, 'FY26 overheads were never split by department, so this half is this '
                         'year only. General Admin is inside the CTS column. The pack says the '
                         "same at 'Overhead Split' row 179.")
    cell.font = Font(name='Calibri', italic=True, size=8, color=INK)
    cell.fill = PatternFill('solid', fgColor=AMBER)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=14)
    r += 2

    def grow(r, label, fn, fmt=NUM, bold=False, italic=False, fill=None, ind=0):
        put(r, 2, ('    ' * ind) + label, bold=bold, italic=italic, fill=fill)
        for name, c1, c2 in GCOLS:
            cell = ws.cell(r, c1, fn(name))
            cell.font = Font(name='Calibri', bold=bold, italic=italic, size=9, color=INK)
            cell.number_format = fmt
            cell.alignment = Alignment(horizontal='center')
            if fill:
                for k in range(c1, c2 + 1):
                    ws.cell(r, k).fill = PatternFill('solid', fgColor=fill)
            for k in range(c1, c2 + 1):
                ws.cell(r, k).border = Border(bottom=thin)
            if c2 > c1:
                ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)

    def g(key, off):
        if key != 'CTS':
            return '=' + q(BYD, GRP[key][0] + off)
        return '=' + '+'.join(q(BYD, GRP[k][0] + off) for k in GRP)

    def gpf(key):
        if key != 'CTS':
            return '=' + q(BYD, GRP[key][1])
        return '=' + '+'.join(q(BYD, GRP[k][1]) for k in GRP)

    gprow = r
    grow(r, 'Gross Profit', gpf, bold=True, fill=GREEN)
    r += 1
    put(r, 2, 'Other Expenses', bold=True, fill=GREY)
    for c in range(3, 15):
        put(r, c, None, fill=GREY)
    r += 1
    for lab, off in (('Advertising & Promotion Exp', 3), ('Business Insurances', 2),
                     ('Employment Expenses', 5)):
        grow(r, lab, lambda k, o=off: g(k, o), ind=1)
        r += 1
    for lab, off in (('Office Staff', 7), ('Departmental Staff', 6)):
        grow(r, lab, lambda k, o=off: g(k, o), italic=True, ind=2, fill=GREY)
        r += 1
    for lab, off in (('General & Administrative Exp', 1), ('Operating Expenses', 4)):
        grow(r, lab, lambda k, o=off: g(k, o), ind=1)
        r += 1
    opex = r
    grow(r, 'Total OpEx', lambda k: g(k, 0), bold=True)
    r += 2
    pl = r
    grow(r, 'Profit/Loss', lambda k: None, bold=True, fill=GREEN)
    for name, c1, c2 in GCOLS:
        ws.cell(r, c1, '=%s%d+%s%d' % (L(c1), gprow, L(c1), opex))
    r += 1
    SALES = {'Support': q(F27, 41),
             'Production': '%s+%s' % (q(F27, 8), q(F27, 16)),
             'Consulting': '%s+%s' % (q(F27, 25), q(F27, 33)),
             'CTS': q(F27, 6)}
    grow(r, 'Profit/Loss %', lambda k: None, fmt=PCT, italic=True)
    for name, c1, c2 in GCOLS:
        ws.cell(r, c1, '=IFERROR(%s%d/(%s),"")' % (L(c1), pl, SALES[name]))
    r += 2
    grow(r, 'GPM if Dep Employment Exp move to COS', lambda k: None, fmt=PCT, bold=True, fill=GREY)
    for name, c1, c2 in GCOLS:
        ds = (q(BYD, GRP[name][0] + 6) if name != 'CTS'
              else '+'.join(q(BYD, GRP[k][0] + 6) for k in GRP))
        ws.cell(r, c1, '=IFERROR((%s%d+%s)/(%s),"")' % (L(c1), gprow, ds, SALES[name]))
    r += 2
    HC = {'Support': 'B6', 'Production': 'B7+%s!B8' % OS,
          'Consulting': 'B9+%s!B10' % OS, 'CTS': 'B6+%s!B7+%s!B8+%s!B9+%s!B10+%s!B11' % (OS, OS, OS, OS, OS)}
    pe = r
    grow(r, 'Profitability Per Employee QTR', lambda k: None, bold=True)
    for name, c1, c2 in GCOLS:
        ws.cell(r, c1, '=IFERROR(%s%d/(%s!%s),"")' % (L(c1), pl, OS, HC[name]))
    r += 1
    grow(r, 'Headcount used', lambda k: None, fmt='#,##0', italic=True)
    for name, c1, c2 in GCOLS:
        ws.cell(r, c1, '=%s!%s' % (OS, HC[name]))
    r += 2
    note = ws.cell(r, 2, 'COST OF SALES GROUPING.  Labour: Sub-Contract Labour, Direct Wages & '
                         'Salaries, Direct Salaries, Direct Wages, Direct Superannuation, Direct '
                         'Workers Comp.   Equipment & Misc: Equipment Purchase, Equipment Hires, '
                         'Service hires, Subscriptions & Licences, Equipment Maintenance Support, '
                         'Stamp Duty, Damage Waiver, Supplier Per Diems, Credit Card fees paid '
                         'COGS.   Transport: Other Direct Expenses.   Please check this, it does '
                         'not tie exactly to the old report.')
    note.font = Font(name='Calibri', italic=True, size=8, color=INK)
    note.fill = PatternFill('solid', fgColor=AMBER)
    note.alignment = Alignment(wrap_text=True, vertical='top')
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=14)
    ws.row_dimensions[r].height = 34

    ws.column_dimensions['A'].width = 2
    ws.column_dimensions['B'].width = 36
    for c in range(3, 15):
        ws.column_dimensions[L(c)].width = 12
    ws.freeze_panes = 'C7'
    pk.save(PACK)
    print('  written to a new sheet %r, last row %d' % (SHEET, r))
    print("  snip range '%s'!B4:N%d" % (SHEET, pe + 1))


if __name__ == '__main__':
    main()
