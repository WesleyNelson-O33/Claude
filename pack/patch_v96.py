"""Fix the year to date departmental table on 'P&L by Dept'.

The block at row 258, "ALL FOUR SIDE BY SIDE, YEAR TO DATE", had two columns
pointing at the wrong rows in every one of its four department blocks.

EBITDA pointed at the EBITDA % row, one below the EBITDA figure.
Net profit pointed at the EBIT % row, not at Net profit at all.

So both columns printed ratios under dollar headings. Income, cost of sales
and gross profit were always right.

After the fix the four net profits sum to the quarter's loss of 5,807.80,
which is what the Summary tab carries.
"""
import sys
import openpyxl

# department row -> (EBITDA row, Net profit row) in column O
FIX = {260: (72, 78),     # Production
       261: (149, 155),   # Consulting
       262: (199, 205),   # Support
       263: (249, 255)}   # General Admin


def main(path):
    wb = openpyxl.load_workbook(path)
    ws = wb['P&L by Dept']
    for row, (ebitda, net) in FIX.items():
        dept = ws.cell(row, 1).value
        for col, target, label in (('E', ebitda, 'EBITDA'), ('F', net, 'Net profit')):
            was = ws[f'{col}{row}'].value
            now = f'=$O${target}'
            ws[f'{col}{row}'] = now
            print('  %-14s %s%-4d %-10s %-9s -> %s' % (dept, col, row, label, was, now))
    wb.save(path)
    print('\nsaved', path)


if __name__ == '__main__':
    main(sys.argv[1])
