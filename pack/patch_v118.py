"""Q1 FY27 against Q1 FY26, a block per department, in the Summary layout.

Danica asked for the Summary tab's look and its collapsing, but per
department. So this writes one block per reporting group, each the full
P&L for that group down to net profit, with four columns across: Q1 FY27,
Q1 FY26, the variance and the variance per cent.

The blocks are taken from 'P&L by Dept', which already holds the P&L split
that way, and the row grouping is copied off it so the detail folds up the
same. The fonts, fills and number formats are read off the Summary sheet
by indent depth, so it reads like the Summary.

This year comes from 'P&L by Dept' columns B to D. Last year comes from
'Actual FY26' on a row map built here and checked label against label
before anything is written. If a label does not match the script stops
rather than write a wrong figure.

Overheads were never split by department before FY27. So in each block the
FY26 column runs to gross profit and stops. Gross profit and its margin are
worked out from that department's own income and cost of sales, which do
exist for both years. Everything below is this year only and the sheet
says so.

Run it with no arguments.
"""
# Every Font here is copied off the Summary sheet, which uses Arial.
import copy
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter as L

PACK = 'CTS Financial Controller Pack - v20.xlsx'
SHEET = 'Q1 Dept Summary'
BYD, FY26, SUM = 'P&L by Dept', 'Actual FY26', 'Summary'
INK, AMBER = '242C39', 'FFF2CC'

# department, first and last row of its block on 'P&L by Dept',
# then the rows that carry its totals on both sheets
BLOCKS = [
    ('Production',    4,  79, dict(inc=7,   cos=25,  gp=70,  f_inc=7,  f_cos=60)),
    ('Consulting',    81, 156, dict(inc=84,  cos=102, gp=147, f_inc=24, f_cos=95)),
    ('Support',       158, 207, dict(inc=161, cos=170, gp=197, f_inc=41, f_cos=130)),
    ('General Admin', 208, 257, dict(inc=211, cos=220, gp=247, f_inc=49, f_cos=147)),
]


def depth(label):
    return len(label) - len(label.lstrip()) if isinstance(label, str) else 0


def main():
    pk = openpyxl.load_workbook(PACK)
    for n in (BYD, FY26, SUM):
        if n not in pk.sheetnames:
            print('STOPPED: the pack has no %r sheet' % n)
            return
    d, a, s = pk[BYD], pk[FY26], pk[SUM]

    # style by indent depth, read off the Summary
    STYLE = {}
    for dep, row in ((0, 6), (4, 7), (8, 8), (12, 9)):
        STYLE[dep] = (copy.copy(s.cell(row, 1).font), copy.copy(s.cell(row, 1).fill),
                      copy.copy(s.cell(row, 5).font), copy.copy(s.cell(row, 5).fill),
                      s.cell(row, 5).number_format, s.cell(row, 9).number_format)

    # row map, this year's block row to last year's row, checked by label
    pairs = {}
    for name, lo, hi, k in BLOCKS:
        runs = [(k['inc'], k['f_inc'], 1)]
        # the seven income detail lines sit under each income line
        if name in ('Production', 'Consulting'):
            runs += [(k['inc'] + 1, k['f_inc'] + 1, 8), (k['inc'] + 9, k['f_inc'] + 9, 8),
                     (k['cos'], k['f_cos'], 1), (k['cos'] + 1, k['f_cos'] + 1, 17),
                     (k['cos'] + 18, k['f_cos'] + 18, 17)]
        else:
            runs += [(k['inc'] + 1, k['f_inc'] + 1, 7),
                     (k['cos'], k['f_cos'], 1), (k['cos'] + 1, k['f_cos'] + 1, 16)]
        for b0, f0, n in runs:
            for i in range(n):
                b, f = b0 + i, f0 + i
                lb = str(d.cell(b, 1).value or '').strip()
                lf = str(a.cell(f, 1).value or '').strip()
                # the two header lines are named differently on the two sheets:
                # 'P&L by Dept' says Total Income and Cost of Sales, 'Actual FY26'
                # names the department. Everything else must match word for word.
                header = (b == k['inc'] and lb == 'Total Income') or \
                         (b == k['cos'] and lb == 'Cost of Sales')
                if lb != lf and not header:
                    print('STOPPED: %s r%d is %r but %s r%d is %r' % (BYD, b, lb, FY26, f, lf))
                    return
                pairs[b] = f
    print('  row map built and checked, %d lines' % len(pairs))

    if SHEET in pk.sheetnames:
        del pk[SHEET]
    ws = pk.create_sheet(SHEET, pk.sheetnames.index(SUM) + 1)
    ws.sheet_properties.outlinePr.summaryBelow = s.sheet_properties.outlinePr.summaryBelow

    ws.cell(1, 1, 'Profit & Loss').font = copy.copy(s.cell(1, 1).font)
    ws.cell(2, 1, 'Q1 FY27 (Jul to Sep 2026) against Q1 FY26 (Jul to Sep 2025), '
                  'by department').font = Font(name='Arial', bold=True, size=10, color=INK)
    note = ws.cell(3, 1, 'Overheads were never split by department before FY27. So in each block '
                         'the FY26 column runs to gross profit and stops. Everything below gross '
                         'profit is this year only.')
    note.font = Font(name='Arial', italic=True, size=8, color=INK)
    note.fill = PatternFill('solid', fgColor=AMBER)
    ws.merge_cells('A3:E3')

    out = 5
    for name, lo, hi, k in BLOCKS:
        ttl = ws.cell(out, 1, name.upper())
        ttl.font = Font(name='Arial', bold=True, size=11, color='FFFFFF')
        ttl.fill = PatternFill('solid', fgColor=INK)
        ws.merge_cells(start_row=out, start_column=1, end_row=out, end_column=5)
        out += 1
        for c, h in enumerate(['Line', 'Q1 FY27', 'Q1 FY26', 'Var $', 'Var %'], 1):
            cell = ws.cell(out, c, h)
            cell.font = copy.copy(s.cell(5, 5).font)
            cell.fill = copy.copy(s.cell(5, 5).fill)
            cell.alignment = Alignment(horizontal='left' if c == 1 else 'right')
        out += 1
        gp_out = None
        for r in range(k['inc'], hi + 1):
            label = d.cell(r, 1).value
            if label in (None, ''):
                continue
            dep = min(depth(label), 12)
            fnt, fil, nfnt, nfil, money, pc = STYLE[dep - (dep % 4)]
            cell = ws.cell(out, 1, label)
            cell.font, cell.fill = copy.copy(fnt), copy.copy(fil)
            lvl = getattr(d.row_dimensions.get(r), 'outlineLevel', 0) or 0
            if lvl:
                ws.row_dimensions[out].outlineLevel = lvl
            ws.cell(out, 2, "=SUM('%s'!B%d:D%d)" % (BYD, r, r))
            if r == k['gp']:
                gp_out = out
                ws.cell(out, 3, "=SUM('%s'!B%d:D%d)+SUM('%s'!B%d:D%d)"
                        % (FY26, k['f_inc'], k['f_inc'], FY26, k['f_cos'], k['f_cos']))
            elif r in pairs:
                ws.cell(out, 3, "=SUM('%s'!B%d:D%d)" % (FY26, pairs[r], pairs[r]))
            elif r == k['gp'] + 1:          # gross profit %
                ws.cell(out, 2, '=IFERROR(B%d/B%d,"")' % (gp_out, gp_out - (k['gp'] - k['inc'])))
                ws.cell(out, 3, '=IFERROR(C%d/C%d,"")' % (gp_out, gp_out - (k['gp'] - k['inc'])))
            if ws.cell(out, 3).value is not None:
                ws.cell(out, 4, '=IFERROR(B%d-C%d,"")' % (out, out))
                ws.cell(out, 5, '=IFERROR(D%d/ABS(C%d),"")' % (out, out))
            # a line is a percentage line only if its own name says so
            is_pct = str(label).replace(' ', '').endswith('%')   # some labels carry a trailing space
            if is_pct:
                # ws.cell(r, c, None) does not clear, it is read as no value given
                ws.cell(out, 5).value = None   # per cent of a per cent means nothing
            for c in range(2, 6):
                cl = ws.cell(out, c)
                if cl.value is None:
                    continue
                cl.font, cl.fill = copy.copy(nfnt), copy.copy(nfil)
                cl.number_format = pc if (is_pct or c == 5) else money
                cl.alignment = Alignment(horizontal='right')
            out += 1
        out += 1

    ws.column_dimensions['A'].width = s.column_dimensions['A'].width or 44
    for c in range(2, 6):
        ws.column_dimensions[L(c)].width = 14
    ws.freeze_panes = 'B7'
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = 'portrait'
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    pk.save(PACK)
    print('  written to a new sheet %r, last row %d' % (SHEET, out))


if __name__ == '__main__':
    main()
