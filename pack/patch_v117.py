"""Put the Q1 FY27 against Q1 FY26 P&L into the pack, in the Summary layout.

The whole company on one sheet. patch_v118 does the same thing a block per
department, so this one is named for the company to keep the two apart.

Danica asked for the departmental figures for both years, laid out exactly
as the Summary tab. The Summary carries the whole P&L line list down the
side, every department and every account under it, with comparison blocks
across. This does the same thing with one block: Q1 FY27, Q1 FY26, the
variance and the variance per cent.

Line labels, fonts, fills and number formats are read off the Summary sheet
row by row, so the two look the same and stay that way if she restyles one.

Figures are formulas into 'P&L FY27' and 'Actual FY26', columns B to D,
which are July, August and September on both sheets. The per cent rows are
built the way the Summary builds them, as a ratio rather than a sum, and
the profit margin per department block at the bottom follows its own
pairings.

Overheads were never split by department before FY27, so the FY26 column is
empty from 'Expenses Overheds' down for anything departmental. The sheet
says so at the top.

Run it with no arguments.
"""
# Every Font here names Arial, because the Summary sheet does.
import copy
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter as L

PACK = 'CTS Financial Controller Pack - v20.xlsx'
SHEET = 'Q1 Company Summary'
OLD = 'Q1 Summary by Dept'      # what this sheet was called before v118 arrived
SRC = 'Summary'
F27, F26 = "'P&L FY27'", "'Actual FY26'"
LAST = 191                                 # the last row the Summary uses
HEAD = 5
INK, AMBER = '242C39', 'FFF2CC'
# the per cent rows, and the pair they are built from, as the Summary has them
RATIO = {174: (173, 6), 176: (175, 6), 180: (179, 6), 182: (181, 6)}
MARGIN = {184: (7, 60), 185: (8, 61), 186: (16, 78), 187: (24, 95),
          188: (25, 96), 189: (33, 113), 190: (41, 130), 191: (49, 147)}


def main():
    pk = openpyxl.load_workbook(PACK)
    if SRC not in pk.sheetnames:
        print('STOPPED: the pack has no %r sheet' % SRC)
        return
    s = pk[SRC]
    if str(s.cell(6, 1).value).strip() != 'Total Income':
        print('STOPPED at %s!A6: expected %r, found %r. The Summary has moved.'
              % (SRC, 'Total Income', s.cell(6, 1).value))
        return
    for n in (SHEET, OLD):
        if n in pk.sheetnames:
            del pk[n]
    ws = pk.create_sheet(SHEET, pk.sheetnames.index(SRC) + 1)

    ws.cell(1, 1, 'Profit & Loss').font = copy.copy(s.cell(1, 1).font)
    ws.cell(3, 1, 'Q1 FY27 (Jul to Sep 2026) against Q1 FY26 (Jul to Sep 2025)').font = Font(
        name='Arial', bold=True, size=10, color=INK)
    note = ws.cell(4, 1, 'Overheads were never split by department before FY27, so the FY26 '
                         'column is empty on the departmental overhead lines.')
    note.font = Font(name='Arial', italic=True, size=8, color=INK)
    note.fill = PatternFill('solid', fgColor=AMBER)
    ws.merge_cells('A4:E4')

    for c, h in enumerate(['Line', 'Q1 FY27', 'Q1 FY26', 'Var $', 'Var %'], 1):
        cell = ws.cell(HEAD, c, h)
        cell.font = copy.copy(s.cell(HEAD, 5).font)
        cell.fill = copy.copy(s.cell(HEAD, 5).fill)
        cell.alignment = Alignment(horizontal='left' if c == 1 else 'right')

    def q(sheet, row):
        return 'SUM(%s!B%d:D%d)' % (sheet, row, row)

    for r in range(6, LAST + 1):
        label = s.cell(r, 1).value
        a = ws.cell(r, 1, label)
        a.font = copy.copy(s.cell(r, 1).font)
        a.fill = copy.copy(s.cell(r, 1).fill)
        if label in (None, ''):
            continue
        if r in RATIO:
            top, base = RATIO[r]
            ws.cell(r, 2, '=IF(B%d=0,"",B%d/B%d)' % (base, top, base))
            ws.cell(r, 3, '=IF(C%d=0,"",C%d/C%d)' % (base, top, base))
            ws.cell(r, 4, '=IFERROR(B%d-C%d,"")' % (r, r))
            ws.cell(r, 5, None)
        elif r in MARGIN:
            inc, cos = MARGIN[r]
            ws.cell(r, 2, '=IF(B%d=0,"",(B%d+B%d)/B%d)' % (inc, inc, cos, inc))
            ws.cell(r, 3, '=IF(C%d=0,"",(C%d+C%d)/C%d)' % (inc, inc, cos, inc))
            ws.cell(r, 4, '=IFERROR(B%d-C%d,"")' % (r, r))
            ws.cell(r, 5, None)
        elif r == 183:
            continue                        # the block heading, no figures
        else:
            ws.cell(r, 2, '=' + q(F27, r))
            ws.cell(r, 3, '=' + q(F26, r))
            ws.cell(r, 4, '=B%d-C%d' % (r, r))
            ws.cell(r, 5, '=IFERROR(D%d/ABS(C%d),0)' % (r, r))
        for c in range(2, 6):
            cell = ws.cell(r, c)
            if cell.value is None:
                continue
            cell.font = copy.copy(s.cell(r, 5).font)
            cell.fill = copy.copy(s.cell(r, 5).fill)
            # the Summary's own money and per cent formats, row by row
            src = s.cell(r, 9 if c == 5 else 5).number_format
            cell.number_format = src
            cell.alignment = Alignment(horizontal='right')

    ws.column_dimensions['A'].width = s.column_dimensions['A'].width or 44
    for c in range(2, 6):
        ws.column_dimensions[L(c)].width = 14
    ws.freeze_panes = 'B6'
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = 'portrait'
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    pk.save(PACK)
    print('  written to a new sheet %r, rows %d to %d' % (SHEET, HEAD, LAST))
    print('  styles copied off %r row by row' % SRC)


if __name__ == '__main__':
    main()
