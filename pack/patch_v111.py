"""Profit per employee: label the years as financial years, and add Q1 FY27.

The block at Xero History row 150 is annual profit per head for FY22 to FY26,
carried over from the Q4 FY2026 Business Advisory pack. It is full year data
and it stops at FY26, so it carries nothing from this year at all.

Danica asked for the years to read FY22, FY23 and so on, and for a first
quarter figure.

The Q1 block is built from the pack, not typed. Net profit by reporting group
comes from the departmental table on P&L by Dept, and headcount from the
staff key on Overhead Split. Support 49, Production 21 (Events 18 plus Video
3) and Consulting 2 (Integration 1 plus Consulting 1) add to the pack's own
total excluding Admin, 72, so the four rows are consistent with each other.

One quarter is not a year. The Q1 figures sit in their own block, under their
own heading, so nobody reads them along the same line as the five years above.
"""
# Every Font here names Calibri. Without a name openpyxl leaves the workbook
# default, which comes out of LibreOffice as a serif.
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

PACK = 'CTS Financial Controller Pack - v20.xlsx'
SHEET = 'Xero History'
YEARS = {153: 'FY22', 154: 'FY23', 155: 'FY24', 156: 'FY25', 157: 'FY26'}
TOP = 270
INK, TEAL, BAND, LINE = '242C39', '3FD0C9', 'EDEFF2', 'C9CED6'
GOOD, BAD = '005358', 'C41A28'
# label, headcount formula, net profit formula
GROUPS = [
    ('ONS  Support',     "='Overhead Split'!B6",
     "='P&L by Dept'!F262"),
    ('PRD  Production',  "='Overhead Split'!B7+'Overhead Split'!B8",
     "='P&L by Dept'!F260"),
    ('CONS Consulting',  "='Overhead Split'!B9+'Overhead Split'!B10",
     "='P&L by Dept'!F261"),
    ('CTS',              "='Overhead Split'!B12",
     "='P&L by Dept'!F264"),
]


def main():
    wb = openpyxl.load_workbook(PACK)
    ws = wb[SHEET]

    # 1. the years read as financial years
    for r, now in YEARS.items():
        was = ws.cell(r, 1).value
        if str(was) == now:
            continue
        if str(was) not in ('20' + now[2:],):
            print('  STOPPED at A%d: expected %r, found %r'
                  % (r, '20' + now[2:], was))
            return
        ws.cell(r, 1).value = now
    print('  years relabelled FY22 to FY26')

    # 2. the first quarter block
    if ws.cell(TOP, 1).value not in (None, 'PROFIT PER EMPLOYEE, Q1 FY27'):
        print('  STOPPED: row %d is not free, it holds %r'
              % (TOP, ws.cell(TOP, 1).value))
        return
    thin = Side(style='thin', color=LINE)
    ws.cell(TOP, 1).value = 'PROFIT PER EMPLOYEE, Q1 FY27'
    ws.cell(TOP, 1).font = Font(name='Calibri', bold=True, size=10, color=INK)
    ws.cell(TOP + 1, 1).value = (
        'Net profit by reporting group from P&L by Dept, headcount from the staff '
        'key on Overhead Split. One quarter, not a year, so it does not sit on the '
        'same line as the five years above. Admin’s seven staff are outside the '
        'three groups and outside the CTS headcount, as the staff key has them.')
    ws.cell(TOP + 1, 1).font = Font(name='Calibri', size=8, italic=True, color='6B7585')

    head = TOP + 2
    for c, h in enumerate(['Reporting group', 'Headcount', 'Q1 net profit',
                           'Per employee'], 1):
        cell = ws.cell(head, c, h)
        cell.font = Font(name='Calibri', bold=True, color='FFFFFF', size=10)
        cell.fill = PatternFill('solid', fgColor=INK)
        cell.alignment = Alignment(horizontal='left' if c == 1 else 'right',
                                   vertical='center')
        cell.border = Border(bottom=Side(style='medium', color=TEAL))
    for i, (name, hc, pf) in enumerate(GROUPS):
        r = head + 1 + i
        last = (i == len(GROUPS) - 1)
        ws.cell(r, 1, name).font = Font(name='Calibri', size=9, bold=last, color=INK)
        ws.cell(r, 2, hc).number_format = '#,##0'
        ws.cell(r, 3, pf).number_format = '#,##0;(#,##0);-'
        ws.cell(r, 4, '=IF(N(B%d)=0,"",C%d/B%d)' % (r, r, r))
        ws.cell(r, 4).number_format = '#,##0;(#,##0);-'
        for c in range(2, 5):
            ws.cell(r, c).font = Font(name='Calibri', size=9, bold=last, color=INK)
            ws.cell(r, c).alignment = Alignment(horizontal='right')
        for c in range(1, 5):
            ws.cell(r, c).border = Border(bottom=thin)
            if last:
                ws.cell(r, c).fill = PatternFill('solid', fgColor=BAND)

    wb.save(PACK)
    print('  Q1 block written at rows %d-%d' % (head, head + len(GROUPS)))
    print('  snip range A%d:D%d' % (head, head + len(GROUPS)))


if __name__ == '__main__':
    main()
