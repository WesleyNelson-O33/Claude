"""Monthly Dept P&L: the Q1 Dept P&L tab, same rows, same format, by month.

MONTHLY tab, not quarterly. Every quarter formula SUM('sheet'!B:D) becomes the
single month column B, C or D. Nothing else is changed, including the row set,
the cost of sales grouping and the reporting group split on the lower half.
"""
import copy as _copy
import re
import shutil
import openpyxl
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter as gl

SRC = ('/tmp/claude-0/-home-user-Claude/a00ffe59-433c-507b-9006-917d8c60083a'
       '/scratchpad/pack2/w.xlsx')
OUT = 'pack/CTS Financial Controller Pack - v2 monthly dept.xlsx'
Q1 = 'Q1 Dept P&L'
NEW = 'Monthly Dept P&L'

# sub department, reporting group, the Q1 column its FY27 formulas live in
GROUPS = [('Managed Services', 'Support', 'C'),
          ('Events', 'Production', 'E'),
          ('Video', 'Production', 'G'),
          ('Integration', 'Consulting', 'I'),
          ('Consulting', 'Consulting', 'K'),
          ('CTS', 'CTS', 'M')]
MONTHS = [('Jul 26', 'B'), ('Aug 26', 'C'), ('Sep 26', 'D')]

RANGE = re.compile(r"SUM\('([^']+)'!\$?([A-Z]{1,3})\$?(\d+):\$?([A-Z]{1,3})\$?(\d+)\)")


def to_month(formula, srccol, tgtcol, mcol):
    """Quarter formula -> one month, and same sheet refs -> the new column."""
    def one(m):
        sheet, c1, r1, c2, r2 = m.groups()
        if r1 != r2:
            return m.group(0)
        return "'%s'!%s%s" % (sheet, mcol, r1)
    out = RANGE.sub(one, formula)
    out = re.sub(r"(?<![!A-Z$:])%s(\d+)\b" % srccol,
                 lambda m: '%s%s' % (tgtcol, m.group(1)), out)
    return out


shutil.copy(SRC, OUT)
wb = openpyxl.load_workbook(OUT)
q = wb[Q1]
if NEW in wb.sheetnames:
    del wb[NEW]
ws = wb.create_sheet(NEW, wb.sheetnames.index(Q1) + 1)
ws.sheet_view.showGridLines = q.sheet_view.showGridLines

ws.column_dimensions['A'].width = q.column_dimensions['A'].width or 2
ws.column_dimensions['B'].width = q.column_dimensions['B'].width or 36
for i in range(len(GROUPS) * len(MONTHS)):
    ws.column_dimensions[gl(3 + i)].width = 12

for r in range(1, q.max_row + 1):
    ws.row_dimensions[r].height = q.row_dimensions[r].height
    src = q.cell(r, 2)
    dst = ws.cell(r, 2)
    dst.value = src.value
    dst._style = _copy.copy(src._style)

for g, (sub, grp, srccol) in enumerate(GROUPS):
    for j, (mlabel, mcol) in enumerate(MONTHS):
        tgt = 3 + g * len(MONTHS) + j
        tcol = gl(tgt)
        for r in range(1, q.max_row + 1):
            src = q.cell(r, openpyxl.utils.column_index_from_string(srccol))
            dst = ws.cell(r, tgt)
            dst._style = _copy.copy(src._style)
            v = src.value
            if isinstance(v, str) and v.startswith('='):
                dst.value = to_month(v, srccol, tcol, mcol)
            elif r > 6:
                dst.value = v
        # headers
        if j == 0:
            c = ws.cell(4, tgt)
            c.value = grp
            c._style = _copy.copy(q.cell(4, 3)._style)
            c2 = ws.cell(5, tgt)
            c2.value = sub
            c2._style = _copy.copy(q.cell(5, 3)._style)
        h = ws.cell(6, tgt)
        h.value = mlabel
        h._style = _copy.copy(q.cell(6, 3)._style)
        h.alignment = Alignment(horizontal='right', vertical='center')

ws['B1'] = 'MONTHLY DEPARTMENT P&L'
ws['B2'] = ('FY27 by month, July to September 2026, by department. '
            'Same rows and the same groupings as the Q1 Dept P&L tab.')
ws['B6'] = 'Month'
ws['B21'] = 'OVERHEADS AND RESULT, BY REPORTING GROUP, FY27 ONLY'
ws['B22'] = ('Overheads are split by reporting group, not by sub department, so '
             'they sit under Managed Services, Events and Integration. '
             'General Admin is inside the CTS column. '
             "The pack says the same at 'Overhead Split' row 179.")
ws['B40'] = 'Profitability Per Employee, month'

wb.save(OUT)
print('wrote', OUT)
