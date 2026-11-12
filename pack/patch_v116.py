"""Put the Q1 worked hours by department, both years, into the pack.

Danica asked for the departmental Q1 figures for FY26 and FY27 together.
Worked hours is the only measure the pack holds for both years by
department. The 'WORKED HOURS BY COST CENTRE' block on the Utilisation
sheet has the prior year at rows 50 to 54 and this year at rows 58 to 62,
both pasted from the same report, so the two are the same measure and can
sit beside each other.

The billable, non-billable and leave split only exists for FY27, in the
block at rows 9 to 43. There is no FY26 version of it, so this stops at
worked hours and says so on the sheet. The FY27 split is carried
underneath for reference, not as a comparison.

Everything is a formula into the Utilisation sheet. Nothing typed.

Run it with no arguments.
"""
# Every Font here names Calibri, or LibreOffice renders the default as serif.
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

PACK = 'CTS Financial Controller Pack - v20.xlsx'
SHEET = 'Utilisation Q1'
UTIL = 'Utilisation'
TOP = 32                                  # under the per employee table
INK, TEAL, BAND, LINE = '242C39', '3FD0C9', 'EDEFF2', 'C9CED6'
AMBER = 'FFF2CC'
# department, prior year row, this year row, and the FY27 detail first row
DEPTS = [('Support', 50, 58, 9), ('Events', 51, 59, 15), ('Video', 52, 60, 21),
         ('Integration', 53, 61, 27), ('Consulting', 54, 62, 33)]
Q1 = 'D%d:F%d'                            # Jul, Aug, Sep in both blocks
NUM = '#,##0;[Red](#,##0)'
PCT = '0.0%;[Red]-0.0%'


def main():
    pk = openpyxl.load_workbook(PACK)
    if SHEET not in pk.sheetnames or UTIL not in pk.sheetnames:
        print('STOPPED: the pack needs both %r and %r' % (SHEET, UTIL))
        return
    u = pk[UTIL]
    if str(u.cell(50, 2).value).strip() != 'Support' or str(u.cell(58, 2).value).strip() != 'Support':
        print('STOPPED: rows 50 and 58 on %r are %r and %r, expected Support twice. '
              'The worked hours block has moved.'
              % (UTIL, u.cell(50, 2).value, u.cell(58, 2).value))
        return
    ws = pk[SHEET]
    for r in range(TOP, TOP + 26):        # rebuilt each run
        for c in range(1, 10):
            ws.cell(r, c).value = None

    thin = Side(style='thin', color=LINE)

    def head(r, cols):
        for c, h in enumerate(cols, 1):
            cell = ws.cell(r, c, h)
            cell.font = Font(name='Calibri', bold=True, color='FFFFFF', size=10)
            cell.fill = PatternFill('solid', fgColor=INK)
            cell.alignment = Alignment(horizontal='left' if c == 1 else 'right')
            cell.border = Border(bottom=Side(style='medium', color=TEAL))

    def body(r, n, bold=False, band=False, pcts=()):
        for c in range(1, n + 1):
            cell = ws.cell(r, c)
            cell.font = Font(name='Calibri', bold=bold, size=9, color=INK)
            cell.alignment = Alignment(horizontal='left' if c == 1 else 'right')
            if c > 1:
                cell.number_format = PCT if c in pcts else NUM
            cell.border = Border(bottom=thin)
            if band:
                cell.fill = PatternFill('solid', fgColor=BAND)

    ws.cell(TOP, 1, 'WORKED HOURS BY DEPARTMENT, Q1 FY27 AGAINST Q1 FY26').font = Font(
        name='Calibri', bold=True, size=10, color=INK)
    r = TOP + 1
    head(r, ['Department', 'Q1 FY27', 'Q1 FY26', 'Var hours', 'Var %'])
    r += 1
    first = r
    for name, prior, this, _ in DEPTS:
        ws.cell(r, 1, name)
        ws.cell(r, 2, "=SUM(%s!%s)" % (UTIL, Q1 % (this, this)))
        ws.cell(r, 3, "=SUM(%s!%s)" % (UTIL, Q1 % (prior, prior)))
        ws.cell(r, 4, '=B%d-C%d' % (r, r))
        ws.cell(r, 5, '=IFERROR(D%d/C%d,"")' % (r, r))
        body(r, 5, pcts=(5,))
        r += 1
    last = r - 1
    ws.cell(r, 1, 'Total')
    for c in (2, 3, 4):
        ws.cell(r, c, '=SUM(%s%d:%s%d)' % (L(c), first, L(c), last))
    ws.cell(r, 5, '=IFERROR(D%d/C%d,"")' % (r, r))
    body(r, 5, bold=True, band=True, pcts=(5,))
    r += 3

    ws.cell(r, 1, 'THE FY27 SPLIT, FOR REFERENCE').font = Font(
        name='Calibri', bold=True, size=10, color=INK)
    r += 1
    cell = ws.cell(r, 1, 'Billable, non-billable and leave are only held for FY27. There is no '
                         'FY26 version, so the table above stops at worked hours.')
    cell.font = Font(name='Calibri', italic=True, size=8, color=INK)
    cell.fill = PatternFill('solid', fgColor=AMBER)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    r += 1
    head(r, ['Department', 'Billable hours', 'Non-billable hours', 'Leave hours', 'Chargeable %'])
    r += 1
    first = r
    for name, _, _, det in DEPTS:
        ws.cell(r, 1, name)
        for c, off in ((2, 0), (3, 1), (4, 2)):
            ws.cell(r, c, "=SUM(%s!%s)" % (UTIL, Q1 % (det + off, det + off)))
        ws.cell(r, 5, '=IFERROR(B%d/(B%d+C%d),"")' % (r, r, r))
        body(r, 5, pcts=(5,))
        r += 1
    last = r - 1
    ws.cell(r, 1, 'Total')
    for c in (2, 3, 4):
        ws.cell(r, c, '=SUM(%s%d:%s%d)' % (L(c), first, L(c), last))
    ws.cell(r, 5, '=IFERROR(B%d/(B%d+C%d),"")' % (r, r, r))
    body(r, 5, bold=True, band=True, pcts=(5,))

    pk.save(PACK)
    print('  written to %r, rows %d to %d' % (SHEET, TOP, r))
    print("  snip range '%s'!A%d:E%d for the year on year table" % (SHEET, TOP + 1, TOP + 7))


if __name__ == '__main__':
    main()
