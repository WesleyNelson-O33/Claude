"""Add the October P&L by Cost Centre tab and point Month-End at it.

Month-End section 1 was comparing October's invoicing against a tab called
'Sep 26 P&L by Cost Centre'. That tab is empty, and its own heading still
reads August, so it was copied from August and never filled. Every line of
section 1 therefore read CHECK against a column of zeroes.

This builds 'Oct 26 P&L by Cost Centre' on the same layout as the August
tab, which is the last one that was actually filled, and repoints Month-End
column H at it.

The figures are deliberately left blank. They are a paste from the Xero
Profit and Loss by tracking category, which this connection cannot produce,
and the controller pack's own October ledger is 423,473.70 short of Xero, so
it cannot fill them either. The tab says so on its face.

One change to how September was wired. Month-End H13, the CTS line, pointed
at column I, which is Reallocation, not column H, which is CTS. It is
pointed at H here and flagged on the sheet.

    scripts/oct_cost_centre_tab.py <tracker in> <tracker out>
"""
import copy
import sys

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

SRC = 'Aug-26 P&L by Cost Centre'
NEW = 'Oct 26 P&L by Cost Centre'
OLD = 'Sep 26 P&L by Cost Centre'
TITLE = 'October 2026 - Xero P&L Trading Income split by cost centre'
NOTE = ('NOT YET FILLED. Run the Xero Profit and Loss by tracking category for '
        '1 to 31 October 2026 and paste each account into columns C to H. Column I '
        'is for a reallocation between cost centres, column J adds across. '
        'October trading income per Xero is 1,056,257.95, so column J row 21 must '
        'come to that once WIP and deferrals are allowed for. Do not build this '
        'from the controller pack: its October ledger is 423,473.70 short of Xero '
        'and needs re-pasting first.')


def main():
    if len(sys.argv) < 3:
        print(__doc__.strip())
        return 1
    wb = openpyxl.load_workbook(sys.argv[1])
    for n in (SRC, OLD, 'Month-End'):
        if n not in wb.sheetnames:
            print('STOPPED: the tracker has no %r tab' % n)
            return 1
    if NEW in wb.sheetnames:
        print('STOPPED: %r already exists' % NEW)
        return 1

    src = wb[SRC]
    ws = wb.copy_worksheet(src)
    ws.title = NEW
    wb.move_sheet(ws, offset=wb.sheetnames.index(OLD) - wb.sheetnames.index(NEW))

    ws['A1'] = TITLE
    ws['A2'] = NOTE
    ws['A2'].alignment = Alignment(wrap_text=True, vertical='top')
    ws['A2'].fill = PatternFill('solid', fgColor='FFF2CC')
    ws.row_dimensions[2].height = 56
    # clear August's figures, leave the account list and the totals formulas
    cleared = 0
    for r in range(5, 21):
        for c in range(3, 10):
            if ws.cell(r, c).value is not None:
                ws.cell(r, c).value = None
                cleared += 1
    ws['A23'] = 'Type these into Month-End column H (month = Oct-26):'
    for r in range(24, 30):
        if ws.cell(r, 3).value is not None:
            ws.cell(r, 3).value = None

    m = wb['Month-End']
    before = [m.cell(r, 8).value for r in range(8, 14)]
    for r, col in zip(range(8, 14), ('C', 'D', 'E', 'F', 'G', 'H')):
        m.cell(r, 8).value = "='%s'!%s21" % (NEW, col)
    flag = m.cell(16, 1)
    flag.value = (str(flag.value or '') + '  NOTE 13/11/26: column H now reads the '
                  'Oct 26 P&L by Cost Centre tab. The CTS line used to read the '
                  'Reallocation column, it now reads the CTS column.')
    wb.save(sys.argv[2])
    print('  %r added, %d August figures cleared' % (NEW, cleared))
    print('  Month-End H8:H13 repointed')
    for r, b, a in zip(range(8, 14), before, [m.cell(r, 8).value for r in range(8, 14)]):
        print('     H%-3d was %-44s now %s' % (r, str(b)[:44], a))
    print('  saved %s' % sys.argv[2])
    return 0


if __name__ == '__main__':
    sys.exit(main())
