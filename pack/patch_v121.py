"""Q1 Dept Report: the P&L in Danica's own department report layout.

She sent the report she actually uses (September 2025 pack, sheet PL, the
Q1 block) and asked for the same thing on FY27 against FY26. So this builds
a new tab laid out exactly that way:

    Support | Production (Events, Video) | Consulting (Integration,
    Consulting) | CTS

    Sales Revenue, Income share, then cost of sales as Labour,
    Equipment & Misc and Transport, Gross Profit and its margin, then the
    overheads, Total OpEx, Profit/Loss and its margin, the combined
    result per reporting group, the GPM if departmental employment cost
    moved into cost of sales, and profitability per employee.

FY27 is live formula, read out of the pack. Income and cost of sales come
off the Engine sheet, Q1 being its columns E to G. The overheads are built
the same way 'P&L by Dept' builds them: the department's own cost off the
Engine, plus its share of the admin overhead off the Overhead Split sheet,
where column U is the category and columns P to T are each department's
share.

FY26 is a pasted block. It is her own published Q1 FY26 report, taken from
the file she sent, because the pack has no departmental overhead split for
last year and inventing one would be worse than pasting hers. The sheet
says on its face where that block came from and that it does not
recalculate.

Costs are shown positive and subtracted, the way her report shows them.
The pack holds them negative, so they are negated here.

Two deviations from her sheet, both flagged on it:
  - Credit Card fees paid gets its own line. The pack carries it as a
    separate overhead and dropping it would stop Total OpEx tying.
  - The GPM line carries both years rather than one.

The row grouping is written with summaryBelow off, because her layout puts
the heading above its detail. That is what stops the plus and minus
buttons working when it is left on.

Run it with no arguments.
"""
import copy
import json
import os
import openpyxl
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

PACK = 'CTS Financial Controller Pack - v20.xlsx'
DATA = os.path.join('data', 'fy26_q1_dept.json')
SHEET = 'Q1 Dept Report'

INK, WHITE = '242C39', 'FFFFFF'
RED, TEAL, BLUE = 'FF4D5A', '3FD0C9', '5B5BE8'
GREEN_F, RED_F = 'C6EFCE', 'FFC7CE'
GREEN_T, RED_T = '006100', '9C0006'
MONEY = '$#,##0;-$#,##0'
PCT = '0.0%'

# department -> first Engine row of its block, and its Overhead Split share column
ENG = {'Managed Services': 5, 'Events': 36, 'Video': 67,
       'Integration': 98, 'Consulting': 129}
SHARE = {'Managed Services': 'P', 'Events': 'Q', 'Video': 'R',
         'Integration': 'S', 'Consulting': 'T'}
DEPTS = ['Managed Services', 'Events', 'Video', 'Integration', 'Consulting', 'CTS']
BANDS = [('Support', ['Managed Services'], RED),
         ('Production', ['Events', 'Video'], TEAL),
         ('Consulting', ['Integration', 'Consulting'], BLUE),
         ('CTS', ['CTS'], None)]

# offsets inside an Engine block
INCOME = tuple(range(0, 7))
LABOUR = (9, 16, 17, 18, 19, 20)
EQUIP = (7, 8, 10, 11, 12, 13, 14, 15, 22)
TRANS = (21,)
OWN = {'General & Administrative Exp': 23, 'Business Insurances': 24,
       'Advertising & Promotion Exp': 25, 'Operating Expenses': 26,
       'Departmental Staff': 27, 'Credit Card fees paid': 28}
# the admin categories each line also picks up
ADMIN = {'General & Administrative Exp': ('General & Administrative Exp', 'Other'),
         'Business Insurances': ('Business Insurances',),
         'Advertising & Promotion Exp': ('Advertising & Promotion Exp',),
         'Operating Expenses': ('Operating Expenses',),
         'Credit Card fees paid': ('Credit Card fees paid',),
         'Office Staff': ('Employment Expenses',)}

# row -> label, indent, outline level
LINES = [
    (4, 'Sales Revenue', 0, 0), (5, 'Income share', 0, 0),
    (6, 'Cost of Sales', 0, 0),
    (7, 'Labour', 1, 1), (8, 'Equipment & Misc', 1, 1), (9, 'Transport', 1, 1),
    (11, 'Gross Profit', 0, 0), (12, 'Gross Profit %', 0, 0),
    (13, 'Other Expenses', 0, 0),
    (14, 'Advertising & Promotion Exp', 1, 1), (15, 'Business Insurances', 1, 1),
    (16, 'Employment Expenses', 1, 1),
    (17, 'Office Staff', 2, 2), (18, 'Departmental Staff', 2, 2),
    (19, 'General & Administrative Exp', 1, 1), (20, 'Operating Expenses', 1, 1),
    (21, 'Credit Card fees paid', 1, 1),
    (22, 'Total OpEx', 0, 0),
    (24, 'Profit/Loss', 0, 0), (25, 'Profit/Loss %', 0, 0),
]
PCT_ROWS = (5, 12, 25)


def q1(start, offs):
    return '+'.join('SUM(Engine!$E%d:$G%d)' % (start + o, start + o) for o in offs)


def spread(col, cats):
    out = []
    for c in cats:
        out.append("SUMPRODUCT(('Overhead Split'!$U$23:$U$86=\"%s\")"
                   "*'Overhead Split'!$%s$23:$%s$86"
                   "*('Overhead Split'!$C$23:$C$86+'Overhead Split'!$D$23:$D$86"
                   "+'Overhead Split'!$E$23:$E$86))" % (c, col, col))
    return '+'.join(out)


def main():
    if not os.path.exists(DATA):
        print('STOPPED: %r is missing, so last year has no source' % DATA)
        return
    fy26 = json.load(open(DATA))
    pk = openpyxl.load_workbook(PACK)
    for n in ('Engine', 'Overhead Split', 'Setup', 'Summary'):
        if n not in pk.sheetnames:
            print('STOPPED: the pack has no %r sheet' % n)
            return
    e = pk['Engine']
    # check the Engine blocks are where this script thinks they are
    for d, s in ENG.items():
        tag = str(e.cell(s, 2).value or '').strip()
        want = {'Managed Services': 'ONSITE', 'Events': 'PRODUCTION'}.get(d, d.upper())
        if tag != want:
            print('STOPPED: Engine r%d is %r, expected %r for %s' % (s, tag, want, d))
            return
        for off, lab in ((0, 'Contract Income'), (21, 'Other Direct Expenses'),
                         (27, 'Employment Expenses')):
            got = str(e.cell(s + off, 4).value or '').strip()
            if got != lab:
                print('STOPPED: Engine r%d is %r, expected %r' % (s + off, got, lab))
                return
    print('  Engine blocks checked for all five departments')

    if SHEET in pk.sheetnames:
        del pk[SHEET]
    ws = pk.create_sheet(SHEET, pk.sheetnames.index('Summary') + 1)
    ws.sheet_properties.outlinePr.summaryBelow = False   # her headings sit above

    col = {}
    c = 2
    for d in DEPTS:
        col[d] = c
        c += 2
    last = c - 1
    thin = Side(style='thin', color='B0B6BF')
    box = Border(left=thin, right=thin, top=thin, bottom=thin)

    # band row and department row
    for name, members, colour in BANDS:
        c0, c1 = col[members[0]], col[members[-1]] + 1
        b = ws.cell(1, c0, name)
        b.font = Font(name='Arial', bold=True, size=10, color=INK)
        b.alignment = Alignment(horizontal='center')
        b.border = box
        ws.merge_cells(start_row=1, start_column=c0, end_row=1, end_column=c1)
        for d in members:
            cell = ws.cell(2, col[d], d)
            cell.font = Font(name='Arial', bold=True, size=10,
                             color=INK if colour is None else WHITE)
            if colour:
                cell.fill = PatternFill('solid', fgColor=colour)
            cell.alignment = Alignment(horizontal='center')
            cell.border = box
            ws.merge_cells(start_row=2, start_column=col[d], end_row=2,
                           end_column=col[d] + 1)
    t = ws.cell(2, 1, 'Q1 (Jul - Sep 2026)')
    t.font = Font(name='Arial', bold=True, size=10, color=INK)

    ws.cell(3, 1, 'Sales').font = Font(name='Arial', bold=True, size=10, color=INK)
    for d in DEPTS:
        for k, h in enumerate(('Q1 FY27', 'Q1 FY26')):
            hc = ws.cell(3, col[d] + k, h)
            hc.font = Font(name='Arial', bold=True, size=9, color=INK)
            hc.alignment = Alignment(horizontal='right')

    # the lines
    for r, label, indent, lvl in LINES:
        cell = ws.cell(r, 1, ('    ' * indent) + label)
        head = label in ('Cost of Sales', 'Other Expenses')
        cell.font = Font(name='Arial', bold=(indent == 0), size=9, color=INK,
                         italic=label in ('Office Staff', 'Departmental Staff'))
        if lvl:
            ws.row_dimensions[r].outlineLevel = lvl
        if label in ('Office Staff', 'Departmental Staff'):
            cell.fill = PatternFill('solid', fgColor='EDEDED')
        if head:
            continue

        for d in DEPTS:
            c27, c26 = col[d], col[d] + 1
            f27 = None
            if d != 'CTS':
                s, sh = ENG[d], SHARE[d]
                if label == 'Sales Revenue':
                    f27 = '=%s' % q1(s, INCOME)
                elif label == 'Labour':
                    f27 = '=-(%s)' % q1(s, LABOUR)
                elif label == 'Equipment & Misc':
                    f27 = '=-(%s)' % q1(s, EQUIP)
                elif label == 'Transport':
                    f27 = '=-(%s)' % q1(s, TRANS)
                elif label == 'Office Staff':
                    f27 = '=-(%s)' % spread(sh, ADMIN['Office Staff'])
                elif label == 'Departmental Staff':
                    f27 = '=-(%s)' % q1(s, (OWN['Departmental Staff'],))
                elif label in OWN:
                    f27 = '=-(%s+%s)' % (q1(s, (OWN[label],)), spread(sh, ADMIN[label]))
            else:
                parts = [L(col[x]) for x in DEPTS if x != 'CTS']
                if label not in ('Income share', 'Gross Profit', 'Gross Profit %',
                                 'Total OpEx', 'Profit/Loss', 'Profit/Loss %'):
                    f27 = '=' + '+'.join('%s%d' % (p, r) for p in parts)

            if label == 'Income share':
                f27 = '=IFERROR(%s4/%s$4,"")' % (L(c27), L(col['CTS']))
            elif label == 'Gross Profit':
                f27 = '=%s4-%s7-%s8-%s9' % tuple([L(c27)] * 4)
            elif label == 'Gross Profit %':
                f27 = '=IFERROR(%s11/%s4,"")' % (L(c27), L(c27))
            elif label == 'Employment Expenses':
                f27 = '=%s17+%s18' % (L(c27), L(c27))
            elif label == 'Total OpEx':
                f27 = '=%s14+%s15+%s16+%s19+%s20+%s21' % tuple([L(c27)] * 6)
            elif label == 'Profit/Loss':
                f27 = '=%s11-%s22' % (L(c27), L(c27))
            elif label == 'Profit/Loss %':
                f27 = '=IFERROR(%s24/%s4,"")' % (L(c27), L(c27))

            if f27:
                ws.cell(r, c27, f27)

            # last year, her own published figures
            v = None
            if label in fy26['lines']:
                v = fy26['lines'][label].get(d)
                if label == 'Office Staff' and d != 'CTS':
                    # her report left this blank per department, but it is just
                    # her own employment expense less her own departmental staff
                    v = None
                    ws.cell(r, c26, '=%s16-%s18' % (L(c26), L(c26)))
            if label == 'Income share':
                ws.cell(r, c26, '=IFERROR(%s4/%s$4,"")' % (L(c26), L(col['CTS'] + 1)))
            elif label == 'Gross Profit %':
                ws.cell(r, c26, '=IFERROR(%s11/%s4,"")' % (L(c26), L(c26)))
            elif label == 'Profit/Loss %':
                ws.cell(r, c26, '=IFERROR(%s24/%s4,"")' % (L(c26), L(c26)))
            elif label == 'Credit Card fees paid':
                pass                      # her report folds this elsewhere
            elif v is not None:
                ws.cell(r, c26, v)

            for cc in (c27, c26):
                cl = ws.cell(r, cc)
                if cl.value is None:
                    continue
                cl.font = Font(name='Arial', bold=(indent == 0), size=9, color=INK,
                               italic=label in ('Office Staff', 'Departmental Staff'))
                cl.number_format = PCT if r in PCT_ROWS else MONEY
                cl.alignment = Alignment(horizontal='right')
                if label in ('Office Staff', 'Departmental Staff'):
                    cl.fill = PatternFill('solid', fgColor='EDEDED')

    # combined result per reporting group
    labels = [('Profit/Loss combined Q1 FY27', 27, 0),
              ('Profit/Loss combined % Q1 FY27', 28, 0),
              ('Profit/Loss combined Q1 FY26', 29, 1),
              ('Profit/Loss combined % Q1 FY26', 30, 1)]
    for text, r, off in labels:
        lc = ws.cell(r, 1, text)
        lc.font = Font(name='Arial', bold=True, size=9, color=INK)
        for name, members, colour in BANDS:
            c0 = col[members[0]] + off
            c1 = col[members[-1]] + 1
            terms = ['%s%d' % (L(col[m] + off), 24) for m in members]
            rev = ['%s%d' % (L(col[m] + off), 4) for m in members]
            if '%' in text:
                cell = ws.cell(r, c0, '=IFERROR((%s)/(%s),"")'
                               % ('+'.join(terms), '+'.join(rev)))
                cell.number_format = PCT
            else:
                cell = ws.cell(r, c0, '=' + '+'.join(terms))
                cell.number_format = MONEY
            cell.font = Font(name='Arial', bold=True, size=9, color=INK)
            cell.alignment = Alignment(horizontal='center')
            cell.border = box
            if c1 > c0:
                ws.merge_cells(start_row=r, start_column=c0, end_row=r, end_column=c1)

    # the department colour strip again, then the GPM line
    for name, members, colour in BANDS:
        for d in members:
            cell = ws.cell(32, col[d], d)
            cell.font = Font(name='Arial', bold=True, size=10,
                             color=INK if colour is None else WHITE)
            if colour:
                cell.fill = PatternFill('solid', fgColor=colour)
            cell.alignment = Alignment(horizontal='center')
            cell.border = box
            ws.merge_cells(start_row=32, start_column=col[d], end_row=32,
                           end_column=col[d] + 1)
    g = ws.cell(33, 1, 'GPM if Dep Employment Exp move to COS')
    g.font = Font(name='Arial', bold=True, size=9, color=INK)
    for d in DEPTS:
        c27, c26 = col[d], col[d] + 1
        ws.cell(33, c27, '=IFERROR((%s11-%s18)/%s4,"")' % (L(c27), L(c27), L(c27)))
        v = fy26['gpm'].get(d)
        if v is not None:
            ws.cell(33, c26, v)
        for cc in (c27, c26):
            cl = ws.cell(33, cc)
            if cl.value is None:
                continue
            cl.number_format = PCT
            cl.font = Font(name='Arial', bold=True, size=9, color=INK)
            cl.alignment = Alignment(horizontal='right')

    # employees and profitability per employee, per reporting group
    HEAD27 = {'Support': "='Overhead Split'!$B$6",
              'Production': "='Overhead Split'!$B$7+'Overhead Split'!$B$8",
              'Consulting': "='Overhead Split'!$B$9+'Overhead Split'!$B$10",
              'CTS': "='Overhead Split'!$B$12"}
    rows = [(35, 'Employees used, Q1 FY27', None), (36, 'Employees used, Q1 FY26', None),
            (37, 'Profitability Per Employee QTR, Q1 FY27', 0),
            (38, 'Profitability Per Employee QTR, Q1 FY26', 1)]
    for r, text, off in rows:
        lc = ws.cell(r, 1, text)
        lc.font = Font(name='Arial', bold=True, size=9, color=INK)
        for name, members, colour in BANDS:
            c0 = col[members[0]] if off is None else col[members[0]] + off
            c1 = col[members[-1]] + 1
            if r == 35:
                cell = ws.cell(r, c0, HEAD27[name])
                cell.number_format = '0.00'
            elif r == 36:
                cell = ws.cell(r, c0, fy26['employees'].get(name))
                cell.number_format = '0.00'
            else:
                src = 27 if off == 0 else 29
                hc = 35 if off == 0 else 36
                cell = ws.cell(r, c0, '=IFERROR(%s%d/%s%d,"")'
                               % (L(col[members[0]] + off), src,
                                  L(col[members[0]]), hc))
                cell.number_format = MONEY
            cell.font = Font(name='Arial', bold=True, size=9, color=INK)
            cell.alignment = Alignment(horizontal='center')
            cell.border = box
            if c1 > c0:
                ws.merge_cells(start_row=r, start_column=c0, end_row=r, end_column=c1)

    # red and green on every percentage and result line, her look
    rng = '%s:%s' % ('B5', '%s38' % L(last))
    for rule, fill, font in ((('greaterThan', ['0']), GREEN_F, GREEN_T),
                             (('lessThan', ['0']), RED_F, RED_T)):
        for r in (5, 12, 25, 28, 30, 33):
            ws.conditional_formatting.add(
                'B%d:%s%d' % (r, L(last), r),
                CellIsRule(operator=rule[0], formula=rule[1],
                           fill=PatternFill('solid', bgColor=fill),
                           font=Font(color=font)))

    note = ws.cell(40, 1,
                   'FY27 is live. It is read from the Engine sheet (Q1 is its columns E to G) '
                   'and the Overhead Split sheet, the same way P&L by Dept builds its overheads. '
                   'FY26 is a pasted block, taken from %s. It is the Q1 FY26 report as it was '
                   'published and it does not recalculate. Costs are shown positive and '
                   'subtracted, the way the report shows them. Two lines differ from the old '
                   'report: Credit Card fees paid has its own line here, because the pack holds '
                   'it separately and Total OpEx would not tie without it, and the GPM line '
                   'carries both years instead of one. Employees for FY27 come from the Overhead '
                   'Split headcount. The FY26 employee counts are the ones in the old report and '
                   'are on a different basis, so read the two per employee lines with care.'
                   % fy26['source'])
    note.font = Font(name='Arial', italic=True, size=8, color=INK)
    note.alignment = Alignment(wrap_text=True, vertical='top')
    ws.merge_cells(start_row=40, start_column=1, end_row=43, end_column=last)

    ws.column_dimensions['A'].width = 38
    for c in range(2, last + 1):
        ws.column_dimensions[L(c)].width = 12
    ws.freeze_panes = 'B4'
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # the grouping on the sheet she said would not collapse
    if 'Q1 Dept Summary' in pk.sheetnames:
        pk['Q1 Dept Summary'].sheet_properties.outlinePr.summaryBelow = False
        print("  'Q1 Dept Summary' outline set to summary above, so the buttons work")

    pk.save(PACK)
    print('  written: %r, columns B to %s' % (SHEET, L(last)))


if __name__ == '__main__':
    main()
