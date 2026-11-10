"""Put the Q1 utilisation schedule into the controller pack.

It was being built in a workbook of its own and snipped from there, so the
slide was the only thing in the deck not reading from the pack. This writes
it into the pack's Utilisation sheet, under the blocks already there, and the
snip then comes from the pack like every other table.

The conditional formatting is Danica's own, copied off the Jul-Sept sheet of
her utilisation report: a green data bar across the Over / under column,
colour 63C384, scaled from the smallest value to the largest. Her rule covers
M6 to M32, which is every row of the table including the Grand Total, so this
one does the same. Nothing else about the formatting is invented.

Run it with the utilisation report as the only argument.
"""
import os, sys
import openpyxl
from openpyxl.formatting.rule import DataBarRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

PACK = 'CTS Financial Controller Pack - v20.xlsx'
# It gets a sheet of its own. Putting it under the blocks on Utilisation
# meant sharing that sheet's column widths, and those are already set for the
# four other tables and two charts snipped off it. Widening columns A to I to
# fit this table would have squeezed all of them.
SHEET = 'Utilisation Q1'
TOP = 1
SRC_SHEET = 'Jul-Sept'
SRC_FIRST, SRC_LAST = 6, 32                 # her table, Blake Crisford to Grand Total
SRC_COLS = [1, 2, 3, 4, 5, 7, 9, 11, 12]    # B group, C name, D E F H J L M
HEAD = ['Group', 'Name', 'Chargeable', 'Non-chargeable', 'Leave',
        'Workable hours', 'Chargeable %', 'Target %', 'Over / under']
WIDTHS = [13, 22, 12, 16, 9, 15, 13, 10, 13]
INK, TEAL, BAND, LINE = '242C39', '3FD0C9', 'EDEFF2', 'C9CED6'
BAR = '63C384'                              # her data bar colour, off M6:M32
OVER = 9                                    # the Over / under column in this block


def main(src):
    wb = openpyxl.load_workbook(src, data_only=True, read_only=True)
    if SRC_SHEET not in wb.sheetnames:
        print('STOPPED: %s has no %s sheet' % (os.path.basename(src), SRC_SHEET))
        wb.close()
        return
    rows = [r for r in wb[SRC_SHEET].iter_rows(
        min_row=SRC_FIRST, max_row=SRC_LAST, values_only=True)]
    wb.close()
    if not rows or str(rows[-1][1]).strip() != 'Grand Total':
        print('STOPPED: row %d of %s is %r, expected Grand Total. The table has '
              'moved, check the report before running this again.'
              % (SRC_LAST, SRC_SHEET, rows[-1][1] if rows else None))
        return

    pk = openpyxl.load_workbook(PACK)
    # rebuilt from scratch each run, so a second run cannot leave two data
    # bars stacked on the same column or stale rows below the table
    if SHEET in pk.sheetnames:
        del pk[SHEET]
    ws = pk.create_sheet(SHEET, pk.sheetnames.index('Utilisation') + 1)
    thin = Side(style='thin', color=LINE)

    ws.cell(TOP, 1).value = 'UTILISATION PER EMPLOYEE, FY27 Q1'
    ws.cell(TOP, 1).font = Font(bold=True, size=10, color=INK)
    head_row = TOP + 1
    for c, h in enumerate(HEAD, 1):
        cell = ws.cell(head_row, c, h)
        cell.font = Font(bold=True, color='FFFFFF', size=10)
        cell.fill = PatternFill('solid', fgColor=INK)
        cell.alignment = Alignment(horizontal='left' if c <= 2 else 'right',
                                   vertical='center')
        cell.border = Border(bottom=Side(style='medium', color=TEAL))
    ws.row_dimensions[head_row].height = 20

    out = head_row + 1
    group = ''
    for r in rows:
        if r[1]:
            group = str(r[1]).strip()
        name = str(r[2]).strip() if r[2] else ''
        is_total = name == ''
        vals = [group if (name == '' or r[1]) else ''] + [r[i] for i in SRC_COLS[1:]]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(out, c, v)
            cell.font = Font(bold=is_total, size=9, color=INK)
            if c <= 2:
                cell.alignment = Alignment(horizontal='left')
            else:
                cell.alignment = Alignment(horizontal='right')
                cell.number_format = '0.0%' if c >= 7 else '#,##0'
            cell.border = Border(bottom=thin)
            if is_total:
                cell.fill = PatternFill('solid', fgColor=BAND)
        out += 1
    last = out - 1

    for c, w in enumerate(WIDTHS, 1):
        ws.column_dimensions[get_column_letter(c)].width = w

    col = get_column_letter(OVER)
    rng = '%s%d:%s%d' % (col, head_row + 1, col, last)
    ws.conditional_formatting.add(rng, DataBarRule(
        start_type='min', end_type='max', color=BAR, showValue=True))

    pk.save(PACK)
    print('  schedule written to a new sheet %r, rows %d-%d' % (SHEET, head_row, last))
    print('  data bar %s on %s, her rule off %s!M%d:M%d'
          % (BAR, rng, SRC_SHEET, SRC_FIRST, SRC_LAST))
    print("  snip range '%s'!A%d:I%d" % (SHEET, head_row, last))


if __name__ == '__main__':
    if len(sys.argv) != 2:
        print('usage: patch_v108.py <utilisation report .xlsx>')
    else:
        main(sys.argv[1])
