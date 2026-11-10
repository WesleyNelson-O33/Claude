"""The departmental result for the quarter, from the pack's own allocation.

Table 7 on Overhead Split gives each department its revenue, budget, gross
profit, allocated overhead and result, but only for the month on Setup. There
is no quarterly version, which is why the quarter slides could not carry a
result by department.

This runs that table for July, then August, then September, on copies of the
pack, and writes the three months and their total into a block at row 300.
Nothing is recalculated here and no method of mine is used. It is the pack's
allocation, run three times and added up.

Re-run it after any change to the ledger or the budget.
"""
import datetime, os, shutil, subprocess, sys
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

MONTHS = [(2026, 7), (2026, 8), (2026, 9)]
TOP = 300
DEPTS = ['Support', 'Events', 'Video', 'Integration', 'Consulting', 'Total']
COLS = [('Revenue', 2), ('Budget', 3), ('Gross profit', 5), ('All overhead', 7), ('Result', 8)]
INK, BAND = '242C39', 'EDEFF2'


def read_month(pack, y, m, work):
    tmp = os.path.join(work, 'm%d%02d.xlsx' % (y, m))
    shutil.copy(pack, tmp)
    wb = openpyxl.load_workbook(tmp)
    wb['Summary']['B3'] = datetime.datetime(y, m, 1)
    wb.save(tmp)
    subprocess.run(['/usr/bin/python3', '-I', '/home/user/Claude/tools/recalc.py', tmp],
                   check=True, capture_output=True)
    ws = openpyxl.load_workbook(tmp, data_only=True)['Overhead Split']
    got = {}
    for r in range(143, 149):
        name = ws.cell(r, 1).value
        if isinstance(name, str):
            got[name.strip()] = {k: ws.cell(r, c).value or 0 for k, c in COLS}
    os.remove(tmp)
    return got


def main(pack, work):
    os.makedirs(work, exist_ok=True)
    months = [read_month(pack, y, m, work) for y, m in MONTHS]
    for (y, m), d in zip(MONTHS, months):
        print('  %d-%02d read, CTS result %.2f' % (y, m, d['Total']['Result']))

    wb = openpyxl.load_workbook(pack)
    ws = wb['Overhead Split']
    ws.cell(TOP, 1).value = '9. DEPARTMENT RESULT, THE QUARTER (the month table run for each of the three)'
    head = ['Department'] + [k for k, _ in COLS]
    for c, h in enumerate(head, 1):
        cell = ws.cell(TOP + 1, c, h)
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = PatternFill('solid', fgColor=INK)
        cell.alignment = Alignment(horizontal='left' if c == 1 else 'right')

    for i, dept in enumerate(DEPTS):
        r = TOP + 2 + i
        ws.cell(r, 1).value = dept
        if dept == 'Total':
            ws.cell(r, 1).font = Font(bold=True)
        for c, (k, _) in enumerate(COLS, 2):
            v = sum(mth[dept][k] for mth in months)
            cell = ws.cell(r, c)
            cell.value = v
            cell.number_format = '#,##0.00'
            if dept == 'Total':
                cell.font = Font(bold=True)
                cell.fill = PatternFill('solid', fgColor=BAND)

    wb.save(pack)
    print('wrote the quarter block at rows %d to %d' % (TOP, TOP + 1 + len(DEPTS)))
    for i, dept in enumerate(DEPTS):
        r = TOP + 2 + i
        print('   %-12s %s' % (dept, '  '.join('%12.2f' % ws.cell(r, c).value
                                               for c in range(2, 2 + len(COLS)))))


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
