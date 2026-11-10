"""Percentages on the benchmark table read as percentages.

The Change column was formatted "+3.2 pts" and two of the benchmark cells
spelled it out, "9 pts from breakeven" and "High twenties per cent". Danica
wants all of it as %.

The figures are percentage points held as plain numbers, 3.17 not 0.0317, so
the format carries a literal % rather than Excel's percent format, which would
multiply by a hundred.
"""
import openpyxl

PACK = 'CTS Financial Controller Pack - v20.xlsx'
SHEET = 'Xero History'
FIRST, LAST = 229, 236
OLD_FMT = '\\+0.0" pts";\\-0.0" pts";\\-'
NEW_FMT = '\\+0.0"%";\\-0.0"%";\\-'
TEXT = {(230, 6): ('9 pts from breakeven', '9% from breakeven'),
        (233, 5): ('High twenties per cent', 'High 20s%')}


def main():
    wb = openpyxl.load_workbook(PACK)
    ws = wb[SHEET]
    n = 0
    for r in range(FIRST, LAST + 1):
        c = ws.cell(r, 4)
        if c.number_format == OLD_FMT:
            c.number_format = NEW_FMT
            n += 1
        elif c.number_format != NEW_FMT:
            print('  row %d has an unexpected format, left alone: %r'
                  % (r, c.number_format))
    for (r, col), (was, now) in TEXT.items():
        cell = ws.cell(r, col)
        if cell.value == was:
            cell.value = now
            print('  %s%d  %r -> %r' % (openpyxl.utils.get_column_letter(col), r, was, now))
        elif cell.value != now:
            print('  %s%d is not what this script expected: %r'
                  % (openpyxl.utils.get_column_letter(col), r, cell.value))
    wb.save(PACK)
    print('  %d change cells reformatted to %%' % n)


if __name__ == '__main__':
    main()
