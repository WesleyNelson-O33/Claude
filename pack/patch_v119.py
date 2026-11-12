"""Q1 FY27 against Q1 FY26, the departments side by side, in the Summary layout.

Danica asked for the departments next to each other, not stacked one under
the other. So this rewrites 'Q1 Dept Summary' with the P&L lines running
down the page and the departments running across: Support, Production,
Consulting, General Admin, then CTS. Each department has Q1 FY27, Q1 FY26,
the variance and the variance per cent, so the two years sit next to each
other inside the department.

The line list is one shared spine, read off the Support block on
'P&L by Dept', because every department uses the same account names.
Production and Consulting hold their income and cost of sales split into
two sub-departments, so those lines add the two sub-rows together.

This year comes from 'P&L by Dept' columns B to D, which are July, August
and September. Last year comes from 'Actual FY26' columns B to D on a row
map built here and checked label against label before anything is written.
If a label does not match the script stops rather than write a wrong figure.

Every percentage line is worked out on this sheet, as the value line over
total income. A quarterly percentage is not the sum of three monthly
percentages, so none of them is summed off the source.

Overheads were never split by department before FY27. So the FY26 column
runs to gross profit and stops, and the sheet says so.

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

SPINE = 161            # the Support block on 'P&L by Dept', 46 lines
NLINES = 46
PCT = {37: 36, 39: 38, 43: 42, 45: 44}      # percentage line -> its value line
GP, INCOME, COS = 36, 0, 9                  # spine positions


def flat(inc, cos):
    """This year's rows for a department whose block is a single flat list."""
    return {i: [inc + i] for i in range(NLINES)}


def split(inc, cos, a_inc, a_cos, b_inc, b_cos, tail):
    """This year's rows for a department split into two sub-departments."""
    m = {0: [inc], 8: [inc + 17], 9: [cos]}
    for i in range(1, 8):
        m[i] = [a_inc + i - 1, b_inc + i - 1]
    for i in range(10, 26):
        m[i] = [a_cos + i - 10, b_cos + i - 10]
    for i in range(26, 36):
        m[i] = [tail + i - 26]
    for i, off in ((GP, 10), (38, 12), (40, 14), (41, 15), (42, 16), (44, 18)):
        m[i] = [tail + off]
    return m


def fy26_flat(inc, cos):
    m = {0: [inc], 9: [cos]}
    for i in range(1, 8):
        m[i] = [inc + i]
    for i in range(10, 26):
        m[i] = [cos + i - 9]
    return m


def fy26_split(inc, cos, a_inc, a_cos, b_inc, b_cos):
    m = {0: [inc], 9: [cos]}
    for i in range(1, 8):
        m[i] = [a_inc + i - 1, b_inc + i - 1]
    for i in range(10, 26):
        m[i] = [a_cos + i - 10, b_cos + i - 10]
    return m


# name, this year's row map, last year's row map
DEPTS = [
    ('Support', flat(161, 170), fy26_flat(41, 130)),
    ('Production', split(7, 25, 9, 27, 17, 44, 60), fy26_split(7, 60, 9, 62, 17, 79)),
    ('Consulting', split(84, 102, 86, 104, 94, 121, 137), fy26_split(24, 95, 26, 97, 34, 114)),
    ('General Admin', flat(211, 220), fy26_flat(49, 147)),
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

    labels = [d.cell(SPINE + i, 1).value for i in range(NLINES)]
    if str(labels[0]).strip() != 'Total Income' or str(labels[GP]).strip() != 'Gross profit':
        print('STOPPED: the Support block on %r is not where it was' % BYD)
        return

    # check every source row carries the label the spine says it should
    for name, this, last in DEPTS:
        for i in range(NLINES):
            want = str(labels[i]).strip()
            for r in this.get(i, []):
                got = str(d.cell(r, 1).value or '').strip()
                # a sub-department's own rows are headed by the sub-department
                # name on the total lines, which is the one allowed difference
                if got != want and i not in (0, 9):
                    print('STOPPED: %s wants %r at %s r%d but found %r'
                          % (name, want, BYD, r, got))
                    return
            for r in last.get(i, []):
                got = str(a.cell(r, 1).value or '').strip()
                if got != want and i not in (0, 9):
                    print('STOPPED: %s wants %r at %s r%d but found %r'
                          % (name, want, FY26, r, got))
                    return
    print('  row maps checked against both sources, %d lines by %d departments'
          % (NLINES, len(DEPTS)))

    # style by indent depth, read off the Summary
    STYLE = {}
    for dep, row in ((0, 6), (4, 7), (8, 8), (12, 9)):
        STYLE[dep] = (copy.copy(s.cell(row, 1).font), copy.copy(s.cell(row, 1).fill),
                      copy.copy(s.cell(row, 5).font), copy.copy(s.cell(row, 5).fill),
                      s.cell(row, 5).number_format, s.cell(row, 9).number_format)

    if SHEET in pk.sheetnames:
        del pk[SHEET]
    ws = pk.create_sheet(SHEET, pk.sheetnames.index(SUM) + 1)
    ws.sheet_properties.outlinePr.summaryBelow = s.sheet_properties.outlinePr.summaryBelow

    groups = [n for n, _, _ in DEPTS] + ['CTS']
    last_col = 1 + 4 * len(groups)

    ws.cell(1, 1, 'Profit & Loss').font = copy.copy(s.cell(1, 1).font)
    ws.cell(2, 1, 'Q1 FY27 (Jul to Sep 2026) against Q1 FY26 (Jul to Sep 2025), '
                  'departments across').font = Font(name='Arial', bold=True, size=10, color=INK)
    note = ws.cell(3, 1, 'Overheads were never split by department before FY27. So the FY26 '
                         'column runs to gross profit and stops. Everything below gross profit '
                         'is this year only.')
    note.font = Font(name='Arial', italic=True, size=8, color=INK)
    note.fill = PatternFill('solid', fgColor=AMBER)
    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=last_col)

    # two header rows: the department across its four columns, then the columns
    for g, name in enumerate(groups):
        c0 = 2 + 4 * g
        ttl = ws.cell(5, c0, name.upper())
        ttl.font = Font(name='Arial', bold=True, size=11, color='FFFFFF')
        ttl.fill = PatternFill('solid', fgColor=INK)
        ttl.alignment = Alignment(horizontal='center')
        ws.merge_cells(start_row=5, start_column=c0, end_row=5, end_column=c0 + 3)
        for k, h in enumerate(['Q1 FY27', 'Q1 FY26', 'Var $', 'Var %']):
            cell = ws.cell(6, c0 + k, h)
            cell.font = copy.copy(s.cell(5, 5).font)
            cell.fill = copy.copy(s.cell(5, 5).fill)
            cell.alignment = Alignment(horizontal='right')
    hl = ws.cell(6, 1, 'Line')
    hl.font = copy.copy(s.cell(5, 5).font)
    hl.fill = copy.copy(s.cell(5, 5).fill)

    top = 7                                    # first line row
    def rowof(i):
        return top + i

    for i in range(NLINES):
        label = labels[i]
        r = rowof(i)
        dep = min(depth(label), 12)
        fnt, fil, nfnt, nfil, money, pc = STYLE[dep - (dep % 4)]
        cell = ws.cell(r, 1, label)
        cell.font, cell.fill = copy.copy(fnt), copy.copy(fil)
        lvl = getattr(d.row_dimensions.get(SPINE + i), 'outlineLevel', 0) or 0
        if lvl:
            ws.row_dimensions[r].outlineLevel = lvl

        for g, name in enumerate(groups):
            c0 = 2 + 4 * g
            f27 = L(c0)
            f26 = L(c0 + 1)
            if name == 'CTS':
                others = [L(2 + 4 * k) for k in range(len(DEPTS))]
                o26 = [L(3 + 4 * k) for k in range(len(DEPTS))]
                if i in PCT:
                    ws.cell(r, c0, '=IFERROR(%s%d/%s%d,"")'
                            % (f27, rowof(PCT[i]), f27, rowof(INCOME)))
                    ws.cell(r, c0 + 1, '=IFERROR(%s%d/%s%d,"")'
                            % (f26, rowof(PCT[i]), f26, rowof(INCOME)))
                else:
                    ws.cell(r, c0, '=' + '+'.join('%s%d' % (x, r) for x in others))
                    vals = ['%s%d' % (x, r) for x in o26]
                    ws.cell(r, c0 + 1, '=IF(COUNT(%s)=0,"",SUM(%s))'
                            % (','.join(vals), ','.join(vals)))
            elif i in PCT:
                this, last = DEPTS[g][1], DEPTS[g][2]
                ws.cell(r, c0, '=IFERROR(%s%d/%s%d,"")'
                        % (f27, rowof(PCT[i]), f27, rowof(INCOME)))
                if PCT[i] == GP:
                    ws.cell(r, c0 + 1, '=IFERROR(%s%d/%s%d,"")'
                            % (f26, rowof(GP), f26, rowof(INCOME)))
            else:
                this, last = DEPTS[g][1], DEPTS[g][2]
                src = this.get(i)
                if src:
                    ws.cell(r, c0, '=' + '+'.join(
                        "SUM('%s'!B%d:D%d)" % (BYD, x, x) for x in src))
                if i == GP:
                    # gross profit last year is that department's own income
                    # plus its own cost of sales, the two lines FY26 does hold
                    ws.cell(r, c0 + 1, '=%s%d+%s%d' % (f26, rowof(INCOME), f26, rowof(COS)))
                else:
                    src26 = last.get(i)
                    if src26:
                        ws.cell(r, c0 + 1, '=' + '+'.join(
                            "SUM('%s'!B%d:D%d)" % (FY26, x, x) for x in src26))

            # the variance pair, only where last year exists
            if ws.cell(r, c0 + 1).value is not None and i not in PCT:
                ws.cell(r, c0 + 2, '=IFERROR(%s%d-%s%d,"")' % (f27, r, f26, r))
                ws.cell(r, c0 + 3, '=IFERROR(%s%d/ABS(%s%d),"")' % (L(c0 + 2), r, f26, r))
            elif i in PCT and ws.cell(r, c0 + 1).value is not None:
                ws.cell(r, c0 + 2, '=IFERROR(%s%d-%s%d,"")' % (f27, r, f26, r))

            for k in range(4):
                cl = ws.cell(r, c0 + k)
                if cl.value is None:
                    continue
                cl.font, cl.fill = copy.copy(nfnt), copy.copy(nfil)
                cl.number_format = pc if (i in PCT or k == 3) else money
                cl.alignment = Alignment(horizontal='right')

    foot = rowof(NLINES) + 1
    f = ws.cell(foot, 1, 'This year is read from the P&L by Dept sheet, July to September. '
                         'Last year is read from the Actual FY26 sheet, July to September. '
                         'Every percentage is worked out here, not summed.')
    f.font = Font(name='Arial', italic=True, size=8, color=INK)
    ws.merge_cells(start_row=foot, start_column=1, end_row=foot, end_column=last_col)

    ws.column_dimensions['A'].width = s.column_dimensions['A'].width or 44
    for c in range(2, last_col + 1):
        ws.column_dimensions[L(c)].width = 13
    ws.freeze_panes = 'B7'
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    pk.save(PACK)
    print('  written: %r, %d lines down, %d departments across, last column %s'
          % (SHEET, NLINES, len(groups), L(last_col)))


if __name__ == '__main__':
    main()
