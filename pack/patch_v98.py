"""First quarter revenue across three years.

The quarterly revenue chart showed all four quarters side by side across the
years. Danica asked for it to be the first quarter only, compared across the
three years, which is also how the deck pairs a monthly slide with its
quarterly twin. Three Septembers on one slide, three first quarters on another.

The labels match the quarterly net result chart, which already reads
Q1 2024, Q1 2025, Q1 2026.

The block goes at row 420, below every chart on the tab and below every snip
print range, so nothing prints it by accident.
"""
import sys
import openpyxl

TOP = 420
SOURCE = [('Q1 2024', 'B74'), ('Q1 2025', 'C74'), ('Q1 2026', 'D74')]


def main(path):
    wb = openpyxl.load_workbook(path)
    ws = wb['Xero History']
    ws.cell(TOP, 1).value = 'REVENUE, FIRST QUARTER, THREE YEARS'
    ws.cell(TOP + 1, 1).value = 'Period'
    ws.cell(TOP + 1, 2).value = 'Revenue'
    for i, (label, ref) in enumerate(SOURCE):
        ws.cell(TOP + 2 + i, 1).value = label
        ws.cell(TOP + 2 + i, 2).value = '=' + ref
        print('  %-9s = %s' % (label, ref))
    wb.save(path)
    print('wrote the block at rows %d to %d' % (TOP, TOP + 4))


if __name__ == '__main__':
    main(sys.argv[1])
