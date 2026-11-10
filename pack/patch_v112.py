"""Shift the per-head years to FY23 to FY27, and give the Q1 row her shape.

Danica relabelled the carried over block herself, FY23 through FY27, so the
pack follows her. See the note written onto the sheet about the last row.

She also asked for the first quarter in the same shape as that table, one
row across the reporting groups, rather than the workings block. This adds
it, pointing at the workings so the two cannot disagree.
"""
# Every Font here names Calibri, or LibreOffice renders the default as serif.
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

PACK = 'CTS Financial Controller Pack - v20.xlsx'
SHEET = 'Xero History'
YEARS = {153: ('FY22', 'FY23'), 154: ('FY23', 'FY24'), 155: ('FY24', 'FY25'),
         156: ('FY25', 'FY26'), 157: ('FY26', 'FY27')}
WORK_FIRST = 273                       # ONS row of the workings block
ROW = 278                              # header of the one line version
INK, TEAL, BAND, LINE = '242C39', '3FD0C9', 'EDEFF2', 'C9CED6'


def main():
    wb = openpyxl.load_workbook(PACK)
    ws = wb[SHEET]

    for r, (was, now) in YEARS.items():
        v = ws.cell(r, 1).value
        if v == now:
            continue
        if v != was:
            print('  STOPPED at A%d: expected %r, found %r' % (r, was, v))
            return
        ws.cell(r, 1).value = now
    print('  years relabelled FY23 to FY27')

    ws.cell(151, 1).value = (
        'Carried over from the Q4 FY2026 Business Advisory pack, page 3. Profit per '
        'head by reporting group. The last row is labelled FY27, which is the year '
        'now running, so it cannot be a completed year. The source pack went to '
        'FY2026. Worth settling before this goes out.')
    ws.cell(151, 1).font = Font(name='Calibri', size=8, italic=True, color='6B7585')

    if ws.cell(ROW, 1).value not in (None, 'Year'):
        print('  STOPPED: row %d is not free, it holds %r' % (ROW, ws.cell(ROW, 1).value))
        return
    thin = Side(style='thin', color=LINE)
    for c, h in enumerate(['Year', 'ONS', 'PRD', 'CONS', 'CTS'], 1):
        cell = ws.cell(ROW, c, h)
        cell.font = Font(name='Calibri', bold=True, color='FFFFFF', size=10)
        cell.fill = PatternFill('solid', fgColor=INK)
        cell.alignment = Alignment(horizontal='left' if c == 1 else 'right',
                                   vertical='center')
        cell.border = Border(bottom=Side(style='medium', color=TEAL))
    ws.cell(ROW + 1, 1, 'Q1 FY27').font = Font(name='Calibri', size=9, bold=True,
                                               color=INK)
    for i in range(4):                               # ONS, PRD, CONS, CTS
        cell = ws.cell(ROW + 1, 2 + i, '=D%d' % (WORK_FIRST + i))
        cell.number_format = '#,##0;(#,##0);-'
        cell.font = Font(name='Calibri', size=9, color=INK)
        cell.alignment = Alignment(horizontal='right')
    for c in range(1, 6):
        ws.cell(ROW + 1, c).border = Border(bottom=thin)
        ws.cell(ROW + 1, c).fill = PatternFill('solid', fgColor=BAND)
    ws.cell(ROW + 2, 1).value = (
        'One quarter, not a year. It cannot be read along the same line as the five '
        'full years above it.')
    ws.cell(ROW + 2, 1).font = Font(name='Calibri', size=8, italic=True, color='6B7585')

    wb.save(PACK)
    print('  one line Q1 row written at rows %d-%d' % (ROW, ROW + 1))


if __name__ == '__main__':
    main()
