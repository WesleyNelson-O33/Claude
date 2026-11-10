"""A quarterly profit and loss with the reporting groups as columns.

Danica asked for a quarterly P&L per department. Snipping each group's own
block gives a tall narrow table, which lands on a landscape slide about two
inches wide and cannot be read. One table with the groups across the top fits
the slide and compares them at a glance, which is the point.

Every figure is a reference to column O of that group's own block, which is
the year to date and at September is the quarter. Nothing is recalculated
here, so the table cannot drift from the blocks above it.

The row for each line is found by its label inside each block, because the
blocks are different lengths. Production has accounts Support does not, so a
fixed offset between them would silently read the wrong line.
"""
import sys
import openpyxl
from openpyxl.utils import get_column_letter as gl

TOP = 270
BLOCKS = [('Production', 4, 79), ('Consulting', 81, 156),
          ('Support', 158, 207), ('General Admin', 208, 257)]

LINES = ['Total Income', 'Other income', 'Cost of Sales', 'Expenses Overheds',
         'Other Expenses', 'Gross profit', 'Gross profit %', 'EBITDA', 'EBITDA %',
         'Depreciation & amortisation', 'Instant asset write off',
         'EBIT', 'EBIT %', 'Net profit', 'Net profit %']

RATIO = {'Gross profit %', 'EBITDA %', 'EBIT %', 'Net profit %'}


def find(ws, label, first, last):
    for r in range(first, last + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, str) and v.strip() == label:
            return r
    return None


def main(path):
    wb = openpyxl.load_workbook(path)
    ws = wb['P&L by Dept']

    ws.cell(TOP, 1).value = 'QUARTERLY P&L BY REPORTING GROUP'
    ws.cell(TOP + 1, 1).value = 'Line'
    for j, (name, _, _) in enumerate(BLOCKS):
        ws.cell(TOP + 1, 2 + j).value = name
    ws.cell(TOP + 1, 2 + len(BLOCKS)).value = 'CTS'
    # the header bar, taken from the table that already sits on this tab
    for c in range(1, 3 + len(BLOCKS)):
        ws.cell(TOP + 1, c)._style = ws.cell(259, min(c, 7))._style

    # wide enough for "General Admin", which otherwise runs into the column
    # beside it when the header is right aligned
    for j in range(len(BLOCKS) + 1):
        ws.column_dimensions[gl(2 + j)].width = 15.5
    ws.column_dimensions['A'].width = 30

    missing = []
    for i, line in enumerate(LINES):
        r = TOP + 2 + i
        ws.cell(r, 1).value = line
        cells = []
        label_done = False
        for j, (name, first, last) in enumerate(BLOCKS):
            src = find(ws, line, first, last)
            col = gl(2 + j)
            if src is None:
                missing.append((line, name))
                ws.cell(r, 2 + j).value = None
                continue
            ws.cell(r, 2 + j).value = '=$O$%d' % src
            # take the number format and the rest of the style from the cell
            # this one reads, so the new table prints the way the block does
            ws.cell(r, 2 + j)._style = ws.cell(src, 15)._style
            if not label_done:
                ws.cell(r, 1)._style = ws.cell(src, 1)._style
                label_done = True
            cells.append('%s%d' % (col, r))
        tot = gl(2 + len(BLOCKS))
        if line in RATIO:
            # a ratio never sums, so take it from the line it belongs to
            base = {'Gross profit %': 'Gross profit', 'EBITDA %': 'EBITDA',
                    'EBIT %': 'EBIT', 'Net profit %': 'Net profit'}[line]
            br = TOP + 2 + LINES.index(base)
            ir = TOP + 2 + LINES.index('Total Income')
            ws.cell(r, 2 + len(BLOCKS)).value = (
                '=IF({t}{i}=0,"",{t}{b}/{t}{i})'.format(t=tot, i=ir, b=br))
            for j in range(len(BLOCKS)):
                c = gl(2 + j)
                ws.cell(r, 2 + j).value = (
                    '=IF({c}{i}=0,"",{c}{b}/{c}{i})'.format(c=c, i=ir, b=br))
        elif cells:
            ws.cell(r, 2 + len(BLOCKS)).value = '=SUM(%s:%s)' % (cells[0], cells[-1])
        if cells:
            ws.cell(r, 2 + len(BLOCKS))._style = ws.cell(r, 2)._style
        print('  %-30s %s' % (line, 'ok' if cells else 'NOT FOUND'))

    if missing:
        print('\nlines not found in a block:')
        for line, name in missing:
            print('   %-30s %s' % (line, name))
    wb.save(path)
    print('\nwrote the table at rows %d to %d' % (TOP, TOP + 1 + len(LINES)))


if __name__ == '__main__':
    main(sys.argv[1])
