"""Put the two WIP comparison figures on Month-End onto October.

Rows 23 and 25 are typed inputs and both still held -133,253.18, which is the
September close, while the month being closed is October. So section 2 was
comparing October's tracker balance against September's actuals.

October closes at a credit of 166,317.47 on both, which is the ending position
established from the Work in Progress Transactions export for 1 September to
31 October after journal #72092 was posted. The schedule and Xero agreed on
that figure, which was the point of the October adjustment work.

Where each figure came from, and when, is written beside it on the sheet.

If the September CSIRO accrual is posted, both become 167,687.30.

    scripts/wip_to_october.py <tracker in> <tracker out>
"""
import sys

import openpyxl
from openpyxl.styles import Alignment, Font

OCT = -166317.47
SEP = -133253.18
SRC23 = ('October close per the WIP schedule, from the corrected schedule in the '
         'WIP Adjustment Schedule 31 10 2026 pack. Typed 13-Nov-26.')
SRC25 = ('October close per Xero 11300, from the Work in Progress Transactions export '
         'for 1 Sep to 31 Oct after journal #72092. Typed 13-Nov-26. Becomes '
         '167,687.30 if the September CSIRO accrual is posted.')


def main():
    if len(sys.argv) < 3:
        print(__doc__.strip())
        return 1
    wb = openpyxl.load_workbook(sys.argv[1])
    m = wb['Month-End']
    if str(m.cell(23, 1).value or '')[:28] != 'Closing balance per the WIP ':
        print('STOPPED: Month-End r23 is %r' % m.cell(23, 1).value)
        return 1
    if str(m.cell(25, 1).value or '')[:28] != 'Closing balance per Xero GL ':
        print('STOPPED: Month-End r25 is %r' % m.cell(25, 1).value)
        return 1
    for r in (23, 25):
        v = m.cell(r, 2).value
        if not isinstance(v, (int, float)) or round(v, 2) != SEP:
            print('STOPPED: Month-End B%d is %r, expected the September figure %s'
                  % (r, v, SEP))
            return 1
    month = m.cell(4, 3).value
    print('  month being closed: %s' % (month.strftime('%d %b %Y') if hasattr(month, 'strftime') else month))
    for r, note in ((23, SRC23), (25, SRC25)):
        m.cell(r, 2).value = OCT
        c = m.cell(r, 3)
        c.value = note
        c.font = Font(name='Arial', italic=True, size=8, color='242C39')
        c.alignment = Alignment(wrap_text=True, vertical='top')
        m.merge_cells(start_row=r, start_column=3, end_row=r, end_column=8)
        m.row_dimensions[r].height = 26
    wb.save(sys.argv[2])
    print('  B23 and B25 moved from %s to %s' % (
        '{:,.2f}'.format(SEP), '{:,.2f}'.format(OCT)))
    print('  saved %s' % sys.argv[2])
    return 0


if __name__ == '__main__':
    sys.exit(main())
