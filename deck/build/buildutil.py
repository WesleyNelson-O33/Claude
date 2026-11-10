"""Build the quarter's utilisation table as a clean styled sheet.

The utilisation report is a 46 sheet workbook and loading it whole runs the
machine out of memory, so only the values of the Jul-Sept sheet are read, in
read only mode, and written into a small workbook styled the way the rest of
the pack's tables are styled. That also makes it match the deck instead of
carrying the pivot's own colours in.
"""
import sys
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

INK, TEAL, BAND, LINE = '242C39', '3FD0C9', 'EDEFF2', 'C9CED6'
# Nine columns, not twelve. The table sits beside a commentary panel now, so
# it has about five and a half inches. Public holiday, total hours and the
# non-chargeable percentage all repeat what is already on the row.
SRC = [3, 4, 5, 7, 9, 11, 12]                     # source columns after group and name
HEAD = ['Group', 'Name', 'Chargeable', 'Non-chargeable', 'Leave',
        'Workable hours', 'Chargeable %', 'Target %', 'Over / under']


def main(src, out):
    wb = openpyxl.load_workbook(src, data_only=True, read_only=True)
    rows = list(wb['Jul-Sept'].iter_rows(min_row=6, max_row=43, values_only=True))
    wb.close()

    nb = openpyxl.Workbook()
    ws = nb.active
    ws.title = 'Utilisation'
    thin = Side(style='thin', color=LINE)

    for c, h in enumerate(HEAD, 1):
        cell = ws.cell(1, c, h)
        cell.font = Font(bold=True, color='FFFFFF', size=10)
        cell.fill = PatternFill('solid', fgColor=INK)
        cell.alignment = Alignment(horizontal='left' if c <= 2 else 'right',
                                   vertical='center')
        cell.border = Border(bottom=Side(style='medium', color=TEAL))
    ws.row_dimensions[1].height = 20

    out_r = 2
    group = ''
    stop = False
    for r in rows:
        if stop or r is None or all(v is None for v in r[1:13]):
            continue
        if r[1]:
            group = str(r[1]).strip()
        # her sheet carries working notes under the Grand Total. The table
        # ends there.
        if group == 'Grand Total':
            stop = True
        name = str(r[2]).strip() if r[2] else ''
        total_row = name == '' and group.endswith('Total')
        is_total = group.endswith('Total') or (name == '' and r[1])
        vals = [group if (name == '' or r[1]) else '', name] + [r[i] for i in SRC]
        if stop:
            vals = ['Grand Total', ''] + [r[i] for i in SRC]
            is_total = True
        for c, v in enumerate(vals, 1):
            cell = ws.cell(out_r, c, v)
            cell.font = Font(bold=is_total, size=9,
                             color=INK)
            if c <= 2:
                cell.alignment = Alignment(horizontal='left')
            else:
                cell.alignment = Alignment(horizontal='right')
                cell.number_format = '0.0%' if c >= 7 else '#,##0'
            cell.border = Border(bottom=thin)
            if is_total:
                cell.fill = PatternFill('solid', fgColor=BAND)
        out_r += 1

    widths = [13, 22, 12, 16, 9, 15, 13, 10, 13]
    for c, w in enumerate(widths, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = w
    nb.save(out)
    print('wrote %d rows to %s' % (out_r - 2, out))


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
