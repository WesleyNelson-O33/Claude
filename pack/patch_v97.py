"""Cumulative year to date revenue by department, for the YTD revenue chart.

Danica asked for department lines on the year to date revenue graph. The data
behind that chart only runs by year, so this adds the department cut.

It goes well below everything else on the tab, because the chart itself floats
over columns F to AA from row 34 and the snip prints that area.

Three months of actuals only. The chart's categories already run July to June,
so the three department points land on July, August and September.
"""
import sys
import openpyxl
from openpyxl.utils import get_column_letter as gl

TOP = 240                       # first row of the new block

# department -> the row it sits on in Summary, whose columns B, C and D are
# July, August and September
ROWS = [('Support', 41), ('Events', 8), ('Video', 16),
        ('Integration', 25), ('Consulting', 33)]

MONTHS = ['Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec',
          'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun']


def main(path):
    wb = openpyxl.load_workbook(path)
    ws = wb['Xero History']

    ws.cell(TOP, 1).value = 'YEAR TO DATE REVENUE BY DEPARTMENT, CUMULATIVE'
    ws.cell(TOP + 1, 1).value = 'Month'
    for i, (name, _) in enumerate(ROWS):
        ws.cell(TOP + 1, 2 + i).value = name

    for m, month in enumerate(MONTHS):
        r = TOP + 2 + m
        ws.cell(r, 1).value = month
        if m > 2:
            continue                       # only three months of actuals
        cols = ''.join(f'+Summary!{gl(2 + k)}' for k in range(m + 1))
        for i, (_, srow) in enumerate(ROWS):
            parts = ''.join(f'+Summary!{gl(2 + k)}{srow}' for k in range(m + 1))
            ws.cell(r, 2 + i).value = '=' + parts.lstrip('+')

    wb.save(path)
    print('wrote the department block at rows %d to %d' % (TOP, TOP + 13))
    for i, (name, srow) in enumerate(ROWS):
        print('  %-12s column %s, from Summary row %d' % (name, gl(2 + i), srow))


if __name__ == '__main__':
    main(sys.argv[1])
