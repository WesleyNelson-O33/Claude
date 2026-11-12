"""Q1 Dept Summary with the same departments as the Q1 Dept P&L tab.

Danica said the departments were not the same ones. The Q1 Dept P&L tab
reads Support, then Events and Video under Production, then Integration
and Consulting under Consulting, then CTS. This sheet now carries that
same set, with the reporting group total sitting next to its two children
so nothing is lost.

Columns across, each with Q1 FY27, Q1 FY26, Var $ and Var %:

    Support | Events | Video | Production total |
    Integration | Consulting | Consulting total | General Admin | CTS

The P&L lines run down the page, one shared spine read off the Support
block on 'P&L by Dept', because every department uses the same account
names.

Overheads are only held at reporting group level. The pack has no Events
or Video overhead split and never had one. So on the four sub-department
columns the sheet runs to gross profit and stops, and says so. The group
columns next to them carry the whole way down.

FY26 has no departmental overheads at all, so the FY26 column stops at
gross profit everywhere.

A sub-department has no gross profit row of its own in the pack, so its
gross profit is its own income plus its own cost of sales, the same way
the Q1 Dept P&L tab does it.

This year is 'P&L by Dept' columns B to D, July to September. Last year is
'Actual FY26' columns B to D. Every row map is checked label against label
before anything is written, and the script stops rather than write a wrong
figure. Every percentage is worked out on the sheet, never summed off
three monthly percentages.

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
INK, AMBER, TOTAL = '242C39', 'FFF2CC', 'E8EDF2'

SPINE = 161                                  # the Support block on 'P&L by Dept'
NLINES = 46
INCOME, OTHER_INC, COS, OVH, GP = 0, 8, 9, 26, 36
PCT = {37: 36, 39: 38, 43: 42, 45: 44}       # percentage line -> its value line


def flat(inc):
    """A department whose block on the source is one flat list."""
    return {i: [inc + i] for i in range(NLINES)}


def split(inc, a_inc, b_inc, cos, a_cos, b_cos, tail):
    """A reporting group whose income and cost of sales are split in two."""
    m = {INCOME: [inc], OTHER_INC: [inc + 17], COS: [cos]}
    for i in range(1, 8):
        m[i] = [a_inc + i, b_inc + i]
    for i in range(10, 26):
        m[i] = [a_cos + i - 9, b_cos + i - 9]
    for i in range(OVH, 36):
        m[i] = [tail + i - OVH]
    for i, off in ((GP, 10), (38, 12), (40, 14), (41, 15), (42, 16), (44, 18)):
        m[i] = [tail + off]
    return m


def sub(inc, cos):
    """A sub-department: income and cost of sales only, no overheads."""
    m = {INCOME: [inc], COS: [cos]}
    for i in range(1, 8):
        m[i] = [inc + i]
    for i in range(10, 26):
        m[i] = [cos + i - 9]
    return m


def pair(inc, a_inc, b_inc, cos, a_cos, b_cos):
    """Last year for a reporting group whose block is split in two."""
    m = {INCOME: [inc], COS: [cos]}
    for i in range(1, 8):
        m[i] = [a_inc + i, b_inc + i]
    for i in range(10, 26):
        m[i] = [a_cos + i - 9, b_cos + i - 9]
    return m


# band, department, this year's map, last year's map, is it a total column
COLS = [
    ('Support',     'Managed Services', flat(161),                         sub(41, 130),  False),
    ('Production',  'Events',           sub(8, 26),                        sub(8, 61),  False),
    ('Production',  'Video',            sub(16, 43),                       sub(16, 78), False),
    ('Production',  'Production total', split(7, 8, 16, 25, 26, 43, 60),   pair(7, 8, 16, 60, 61, 78),  True),
    ('Consulting',  'Integration',      sub(85, 103),                      sub(25, 96), False),
    ('Consulting',  'Consulting',       sub(93, 120),                      sub(33, 113), False),
    ('Consulting',  'Consulting total', split(84, 85, 93, 102, 103, 120, 137), pair(24, 25, 33, 95, 96, 113), True),
    ('General Admin', 'General Admin',  flat(211),                         sub(49, 147),  True),
    ('CTS',         'CTS',              None,                              None,      True),
]
# which columns add up to CTS
CTS_PARTS = ['Managed Services', 'Production total', 'Consulting total', 'General Admin']
# the columns that carry the lines below gross profit
DEEP = set(CTS_PARTS + ['CTS'])


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
    if str(labels[INCOME]).strip() != 'Total Income' or str(labels[GP]).strip() != 'Gross profit':
        print('STOPPED: the Support block on %r is not where it was' % BYD)
        return

    checked = 0
    for band, name, this, last in [(c[0], c[1], c[2], c[3]) for c in COLS]:
        if this is None:
            continue
        for sheet, m in ((d, this), (a, last)):
            for i, rows in m.items():
                want = str(labels[i]).strip()
                for r in rows:
                    got = str(sheet.cell(r, 1).value or '').strip()
                    # a total line is headed by the department's own name on
                    # the source, which is the one allowed difference
                    if got != want and i not in (INCOME, COS):
                        print('STOPPED: %s wants %r at %s r%d but found %r'
                              % (name, want, sheet.title, r, got))
                        return
                    checked += 1
    print('  row maps checked against both sources, %d source rows' % checked)

    STYLE = {}
    for dep, row in ((0, 6), (4, 7), (8, 8), (12, 9)):
        STYLE[dep] = (copy.copy(s.cell(row, 1).font), copy.copy(s.cell(row, 1).fill),
                      copy.copy(s.cell(row, 5).font), copy.copy(s.cell(row, 5).fill),
                      s.cell(row, 5).number_format, s.cell(row, 9).number_format)

    if SHEET in pk.sheetnames:
        del pk[SHEET]
    ws = pk.create_sheet(SHEET, pk.sheetnames.index(SUM) + 1)
    ws.sheet_properties.outlinePr.summaryBelow = s.sheet_properties.outlinePr.summaryBelow

    last_col = 1 + 4 * len(COLS)
    col0 = {name: 2 + 4 * g for g, (_, name, _, _, _) in enumerate(COLS)}

    ws.cell(1, 1, 'Profit & Loss').font = copy.copy(s.cell(1, 1).font)
    ws.cell(2, 1, 'Q1 FY27 (Jul to Sep 2026) against Q1 FY26 (Jul to Sep 2025), '
                  'departments across').font = Font(name='Arial', bold=True, size=10, color=INK)
    note = ws.cell(3, 1,
                   'Overheads are only held by reporting group. There is no Events, Video, '
                   'Integration or Consulting overhead split in the pack, so those four columns '
                   'run to gross profit and stop. The total column beside them carries the rest. '
                   'FY26 was never split by department below gross profit either, so every FY26 '
                   'column stops there. A sub-department gross profit is its own income plus its '
                   'own cost of sales, the same as the Q1 Dept P&L tab.')
    note.font = Font(name='Arial', italic=True, size=8, color=INK)
    note.fill = PatternFill('solid', fgColor=AMBER)
    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=last_col)

    # band row: the reporting group over its columns
    g = 0
    while g < len(COLS):
        band = COLS[g][0]
        n = 1
        while g + n < len(COLS) and COLS[g + n][0] == band:
            n += 1
        c0 = 2 + 4 * g
        cell = ws.cell(4, c0, band.upper())
        cell.font = Font(name='Arial', bold=True, size=11, color='FFFFFF')
        cell.fill = PatternFill('solid', fgColor=INK)
        cell.alignment = Alignment(horizontal='center')
        ws.merge_cells(start_row=4, start_column=c0, end_row=4, end_column=c0 + 4 * n - 1)
        g += n

    for gi, (band, name, _, _, is_total) in enumerate(COLS):
        c0 = 2 + 4 * gi
        cell = ws.cell(5, c0, name)
        cell.font = Font(name='Arial', bold=True, size=10, color=INK)
        cell.fill = PatternFill('solid', fgColor=TOTAL)
        cell.alignment = Alignment(horizontal='center')
        ws.merge_cells(start_row=5, start_column=c0, end_row=5, end_column=c0 + 3)
        for k, h in enumerate(['Q1 FY27', 'Q1 FY26', 'Var $', 'Var %']):
            hc = ws.cell(6, c0 + k, h)
            hc.font = copy.copy(s.cell(5, 5).font)
            hc.fill = copy.copy(s.cell(5, 5).fill)
            hc.alignment = Alignment(horizontal='right')
    hl = ws.cell(6, 1, 'Line')
    hl.font = copy.copy(s.cell(5, 5).font)
    hl.fill = copy.copy(s.cell(5, 5).fill)

    top = 7
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

        for band, name, this, last, is_total in COLS:
            c0 = col0[name]
            f27, f26 = L(c0), L(c0 + 1)
            deep = name in DEEP
            if i > GP + 1 and not deep:
                continue                       # no overhead split for a sub-department
            if name == 'CTS':
                parts27 = ['%s%d' % (L(col0[p]), r) for p in CTS_PARTS]
                parts26 = ['%s%d' % (L(col0[p] + 1), r) for p in CTS_PARTS]
                if i in PCT:
                    ws.cell(r, c0, '=IFERROR(%s%d/%s%d,"")'
                            % (f27, rowof(PCT[i]), f27, rowof(INCOME)))
                    if PCT[i] == GP:
                        ws.cell(r, c0 + 1, '=IFERROR(%s%d/%s%d,"")'
                                % (f26, rowof(GP), f26, rowof(INCOME)))
                else:
                    ws.cell(r, c0, '=' + '+'.join(parts27))
                    ws.cell(r, c0 + 1, '=IF(COUNT(%s)=0,"",SUM(%s))'
                            % (','.join(parts26), ','.join(parts26)))
            elif i in PCT:
                ws.cell(r, c0, '=IFERROR(%s%d/%s%d,"")'
                        % (f27, rowof(PCT[i]), f27, rowof(INCOME)))
                if PCT[i] == GP:
                    ws.cell(r, c0 + 1, '=IFERROR(%s%d/%s%d,"")'
                            % (f26, rowof(GP), f26, rowof(INCOME)))
            elif i == GP:
                src = this.get(i)
                if src:                        # a group has its own gross profit row
                    ws.cell(r, c0, '=' + '+'.join("SUM('%s'!B%d:D%d)" % (BYD, x, x) for x in src))
                else:                          # a sub-department has none, so build it
                    ws.cell(r, c0, '=%s%d+%s%d' % (f27, rowof(INCOME), f27, rowof(COS)))
                ws.cell(r, c0 + 1, '=%s%d+%s%d' % (f26, rowof(INCOME), f26, rowof(COS)))
            else:
                src = this.get(i)
                if src:
                    ws.cell(r, c0, '=' + '+'.join("SUM('%s'!B%d:D%d)" % (BYD, x, x) for x in src))
                src26 = last.get(i)
                if src26 and i <= GP:
                    ws.cell(r, c0 + 1, '=' + '+'.join(
                        "SUM('%s'!B%d:D%d)" % (FY26, x, x) for x in src26))

            if ws.cell(r, c0 + 1).value is not None:
                ws.cell(r, c0 + 2, '=IFERROR(%s%d-%s%d,"")' % (f27, r, f26, r))
                if i not in PCT:
                    ws.cell(r, c0 + 3, '=IFERROR(%s%d/ABS(%s%d),"")' % (L(c0 + 2), r, f26, r))

            for k in range(4):
                cl = ws.cell(r, c0 + k)
                if cl.value is None:
                    continue
                cl.font, cl.fill = copy.copy(nfnt), copy.copy(nfil)
                cl.number_format = pc if (i in PCT or k == 3) else money
                cl.alignment = Alignment(horizontal='right')

    foot = rowof(NLINES) + 1
    f = ws.cell(foot, 1, 'This year is read from P&L by Dept, July to September. Last year is '
                         'read from Actual FY26, July to September. Every percentage is worked '
                         'out here, not summed.')
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
    print('  written: %r, %d lines down, %d columns across, last column %s'
          % (SHEET, NLINES, len(COLS), L(last_col)))


if __name__ == '__main__':
    main()
