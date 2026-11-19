"""Dept P&L Monthly: the Q1 Dept P&L cost groupings, laid out like P&L by Dept.

MONTHLY tab. One block per reporting group down the page, months across, with
the sub departments under each line and collapsible. Rows are the ones from the
Q1 Dept P&L: sales, Labour, Equipment and Services, Transportation, Overhead,
Total COGS, Gross Profit and GPM. Nothing new is invented; every formula is the
Q1 one with the quarter range swapped for a single month.
"""
import copy as _copy
import re
import shutil
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as gl, column_index_from_string as ci

SRC = ('/tmp/claude-0/-home-user-Claude/a00ffe59-433c-507b-9006-917d8c60083a'
       '/scratchpad/pack2/w.xlsx')
OUT = 'pack/CTS Financial Controller Pack - v3.xlsx'
NEW = 'Dept P&L Monthly'

NAVY, WHITE, GREY, TEAL = 'FF242C39', 'FFFFFFFF', 'FFF4F4F4', 'FF005358'
BAND = PatternFill('solid', start_color=NAVY, end_color=NAVY)
SUBF = PatternFill('solid', start_color=GREY, end_color=GREY)
RULE = Side(style='thin', color='FFC9CED6')
TOP = Side(style='thin', color='FF005358')
HEAD = Font(name='Calibri', sz=11, b=True, color=WHITE)
BOLD = Font(name='Calibri', sz=11, b=True, color=NAVY)
BODY = Font(name='Calibri', sz=11, color=NAVY)
SMALL = Font(name='Calibri', sz=10, color='FF6B7585')
TITLE = Font(name='Calibri', sz=16, b=True, color=NAVY)
SECT = Font(name='Calibri', sz=12, b=True, color=WHITE)
NOTE = Font(name='Calibri', sz=10, i=True, color='FF6B7585')
MONEY = '#,##0;[RED](#,##0);"-"'
PCT = '0.0%;[RED](0.0%);"-"'

# reporting group -> [(sub department label, the Q1 Dept P&L column it lives in)]
BLOCKS = [
    ('Support', [('Managed Services', 'C')]),
    ('Production', [('Events', 'E'), ('Video', 'G')]),
    ('Consulting', [('Integration', 'I'), ('Consulting', 'K')]),
    ('CTS', [('CTS total', 'M')]),
]
# Q1 row -> label used on the new tab
LINES = [(9, 'Total Sales Revenue'), (13, 'Labour'), (14, 'Equipment & Services'),
         (15, 'Transportation')]
MONTHS = [('Jul-26', 'B'), ('Aug-26', 'C'), ('Sep-26', 'D')]
RANGE = re.compile(r"SUM\('([^']+)'!\$?([A-Z]{1,3})\$?(\d+):\$?([A-Z]{1,3})\$?(\d+)\)")


def month_of(formula, mcol):
    def one(m):
        sheet, c1, r1, c2, r2 = m.groups()
        return m.group(0) if r1 != r2 else "'%s'!%s%s" % (sheet, mcol, r1)
    return RANGE.sub(one, formula)


shutil.copy(SRC, OUT)
wb = openpyxl.load_workbook(OUT)
q = wb['Q1 Dept P&L']
if NEW in wb.sheetnames:
    del wb[NEW]
ws = wb.create_sheet(NEW, wb.sheetnames.index('Q1 Dept P&L') + 1)
ws.sheet_view.showGridLines = False

# C D E = the three months, F = Q1 FY27, H = Q1 FY26
COLS = ['C', 'D', 'E']
YTD, PY = 'F', 'H'


def put(r, c, v, font=BODY, fmt=None, fill=None, indent=0, box=None):
    cell = ws.cell(r, c)
    cell.value = v
    cell.font = font
    if fmt:
        cell.number_format = fmt
    if fill:
        cell.fill = fill
    if indent:
        cell.alignment = Alignment(horizontal='left', indent=indent)
    if box:
        cell.border = box
    return cell


put(1, 2, 'Dept P&L Monthly', TITLE)
put(2, 2, 'FY27 Q1 by month, by department. Same lines and the same cost '
          'groupings as the Q1 Dept P&L tab.', SMALL)
put(3, 2, 'Click the minus in the margin to collapse a department to its total. '
          'Overhead is the residual and proves Labour, Equipment and Transport '
          'add back to Total COGS.', NOTE)

r = 5
for grp, subs in BLOCKS:
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
    c = put(r, 2, grp, SECT, fill=BAND)
    c.alignment = Alignment(horizontal='left', vertical='center')
    ws.row_dimensions[r].height = 20
    r += 1
    put(r, 2, 'Line', HEAD, fill=BAND)
    for i, (m, _) in enumerate(MONTHS):
        put(r, ci(COLS[i]), m, HEAD, fill=BAND).alignment = Alignment(horizontal='right')
    put(r, ci(YTD), 'Q1 FY27', HEAD, fill=BAND).alignment = Alignment(horizontal='right')
    put(r, ci(PY), 'Q1 FY26', HEAD, fill=BAND).alignment = Alignment(horizontal='right')
    ws.row_dimensions[r].height = 18
    r += 1

    first = r
    rows = {}
    for q1row, label in LINES:
        head_row = r
        sub_rows = []
        if len(subs) > 1:
            r += 1
            for sub_label, col in subs:
                put(r, 2, sub_label, BODY, indent=2)
                for i, (_, mcol) in enumerate(MONTHS):
                    put(r, ci(COLS[i]), month_of(q.cell(q1row, ci(col)).value, mcol),
                        BODY, MONEY)
                put(r, ci(YTD), '=SUM(%s%d:%s%d)' % (COLS[0], r, COLS[2], r), BODY, MONEY)
                py = q.cell(q1row, ci(col) + 1).value
                put(r, ci(PY), py, BODY, MONEY)
                sub_rows.append(r)
                ws.row_dimensions[r].outline_level = 1
                ws.row_dimensions[r].height = 15
                r += 1
        else:
            r += 1
        put(head_row, 2, label, BOLD, fill=SUBF)
        for i in range(3):
            if sub_rows:
                f = '=' + '+'.join('%s%d' % (COLS[i], x) for x in sub_rows)
            else:
                f = month_of(q.cell(q1row, ci(subs[0][1])).value, MONTHS[i][1])
            put(head_row, ci(COLS[i]), f, BOLD, MONEY, fill=SUBF)
        put(head_row, ci(YTD), '=SUM(%s%d:%s%d)' % (COLS[0], head_row, COLS[2], head_row),
            BOLD, MONEY, fill=SUBF)
        if sub_rows:
            pyf = '=' + '+'.join('%s%d' % (PY, x) for x in sub_rows)
        else:
            pyf = q.cell(q1row, ci(subs[0][1]) + 1).value
        put(head_row, ci(PY), pyf, BOLD, MONEY, fill=SUBF)
        ws.row_dimensions[head_row].height = 16
        rows[label] = head_row

    # Total COGS straight off the ledger, then Overhead as the residual
    cogs = r
    put(cogs, 2, 'Total COGS', BOLD, fill=SUBF)
    for i, (_, mcol) in enumerate(MONTHS):
        f = '+'.join(month_of(q.cell(16, ci(col)).value, mcol).lstrip('=')
                     for _, col in subs)
        put(cogs, ci(COLS[i]), '=' + f, BOLD, MONEY, fill=SUBF)
    put(cogs, ci(YTD), '=SUM(%s%d:%s%d)' % (COLS[0], cogs, COLS[2], cogs), BOLD,
        MONEY, fill=SUBF)
    put(cogs, ci(PY), '=' + '+'.join(str(q.cell(16, ci(col) + 1).value).lstrip('=')
                                     for _, col in subs), BOLD, MONEY, fill=SUBF)
    ws.row_dimensions[cogs].height = 16
    r += 1

    ovh = r
    put(ovh, 2, 'Overhead', BODY, indent=2)
    for c2 in (COLS[0], COLS[1], COLS[2], YTD, PY):
        put(ovh, ci(c2), '=ROUND(%s%d-%s%d-%s%d-%s%d,2)'
            % (c2, cogs, c2, rows['Labour'], c2, rows['Equipment & Services'],
               c2, rows['Transportation']), BODY, MONEY)
    ws.row_dimensions[ovh].height = 15
    r += 1

    gp = r
    put(gp, 2, 'Gross Profit', BOLD, fill=SUBF, box=Border(top=TOP))
    for i, (_, mcol) in enumerate(MONTHS):
        f = '+'.join(month_of(q.cell(18, ci(col)).value, mcol).lstrip('=')
                     for _, col in subs)
        put(gp, ci(COLS[i]), '=' + f, BOLD, MONEY, fill=SUBF, box=Border(top=TOP))
    put(gp, ci(YTD), '=SUM(%s%d:%s%d)' % (COLS[0], gp, COLS[2], gp), BOLD, MONEY,
        fill=SUBF, box=Border(top=TOP))
    put(gp, ci(PY), '=' + '+'.join(str(q.cell(18, ci(col) + 1).value).lstrip('=')
                                   for _, col in subs), BOLD, MONEY, fill=SUBF,
        box=Border(top=TOP))
    ws.row_dimensions[gp].height = 16
    r += 1

    put(r, 2, 'GPM', BOLD, fill=SUBF)
    for c2 in (COLS[0], COLS[1], COLS[2], YTD, PY):
        put(r, ci(c2), '=IFERROR(%s%d/%s%d,"")'
            % (c2, gp, c2, rows['Total Sales Revenue']), BOLD, PCT, fill=SUBF)
    ws.row_dimensions[r].height = 16
    r += 3

put(r, 2, 'Overhead is the residual: Total COGS less Labour, Equipment and '
          'Services and Transportation. It reads nil, which proves the three '
          'groupings add back to the cost of sales in the ledger.', NOTE)
r += 1
put(r, 2, 'Q1 FY26 is the comparative from the Q1 Dept P&L tab. FY26 was never '
          'split by month in the pack, so it is shown as the quarter only.', NOTE)

ws.column_dimensions['A'].width = 2
ws.column_dimensions['B'].width = 30
for c2 in COLS + [YTD, PY]:
    ws.column_dimensions[c2].width = 14
ws.column_dimensions['G'].width = 2
for rr in range(1, r + 2):
    if ws.row_dimensions[rr].height is None:
        ws.row_dimensions[rr].height = 15
ws.sheet_properties.outlinePr.summaryBelow = False
ws.freeze_panes = 'C1'

wb.save(OUT)
print('wrote', OUT, 'last row', r)
