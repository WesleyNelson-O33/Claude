"""Work in progress at 30 September 2026, worked out from the job ledger.

The Account Transactions export carries a job number on every line of revenue
and cost of sales, so the position of each job can be built directly. That is
what WIP is: cost incurred on a job that has not been billed.

The export covers the profit and loss only, so account 11300 is not in it. What
this does is calculate what 11300 ought to say, so the two can be compared.

Run with the export as the only argument.
"""
import collections
import datetime
import sys

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as gl

OUT = 'recon/CTS work in progress - 30 September 2026.xlsx'
CUT = datetime.datetime(2026, 9, 30, 23, 59)
FRESH = datetime.datetime(2026, 4, 1)
BS_11300 = -133253.18          # account 11300 on the balance sheet at the cut
STANDING = ('9100', '9105', '9109', '9110', '2479')

NAVY, WHITE, GREY = 'FF242C39', 'FFFFFFFF', 'FFF4F4F4'
BAND = PatternFill('solid', start_color=NAVY, end_color=NAVY)
TOTF = PatternFill('solid', start_color=GREY, end_color=GREY)
RULE = Side(style='thin', color='FFC9CED6')
HEAD = Font(name='Calibri', sz=11, b=True, color=WHITE)
BOLD = Font(name='Calibri', sz=11, b=True, color=NAVY)
BODY = Font(name='Calibri', sz=11, color=NAVY)
NOTE = Font(name='Calibri', sz=10, i=True, color='FF6B7585')
MONEY = '#,##0.00;[RED](#,##0.00);"-"'


def bucket(name, last):
    head = name.split(' - ')[0].strip()
    if head in STANDING or 'quote' in name.lower():
        return 'Standing code, not a job'
    return ('Live' if last >= FRESH else 'Stale, last touched over six months ago')


def read(path):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    rows = wb['Account Transactions'].iter_rows(values_only=True)
    for _ in range(5):
        next(rows)
    job = collections.defaultdict(
        lambda: {'rev': 0.0, 'cos': 0.0, 'last': None, 'dept': '',
                 'cust': collections.Counter()})
    acct = collections.defaultdict(lambda: {'rev': 0.0, 'cos': 0.0})
    for r in rows:
        if r[0] is None or not hasattr(r[3], 'strftime') or r[3] > CUT:
            continue
        code = str(r[0]).strip()
        if not (code.startswith('4') or code.startswith('5')):
            continue
        amt = (r[5] or 0) - (r[6] or 0)              # debit positive
        j = str(r[7]).strip() if r[7] else '(no job number)'
        v = job[j]
        key = 'rev' if code.startswith('4') else 'cos'
        v[key] += -amt if key == 'rev' else amt
        v['last'] = r[3] if v['last'] is None else max(v['last'], r[3])
        if r[11]:
            v['dept'] = str(r[11])
        if r[4]:
            v['cust'][str(r[4])] += 1
    wb.close()
    return job


def band(ws, row, width, text):
    ws.cell(row, 1, text).font = HEAD
    for c in range(1, width + 1):
        ws.cell(row, c).fill = BAND


def header(ws, row, names, widths):
    for c, n in enumerate(names, 1):
        cell = ws.cell(row, c, n)
        cell.font = BOLD
        cell.border = Border(bottom=RULE)
        if c > 1:
            cell.alignment = Alignment(horizontal='right', wrap_text=True)
    for c, w in enumerate(widths, 1):
        ws.column_dimensions[gl(c)].width = w


def main(path):
    job = read(path)
    unbilled = {j: v for j, v in job.items()
                if v['rev'] <= 0.005 and v['cos'] > 0.005}
    short = {j: v for j, v in job.items()
             if v['rev'] > 0.005 and v['cos'] - v['rev'] > 0.005}
    groups = collections.defaultdict(list)
    for j, v in unbilled.items():
        groups[bucket(j, v['last'])].append((j, v))
    live = sum(v['cos'] for _, v in groups['Live'])

    wb = openpyxl.Workbook()

    # ------------------------------------------------------- 1. summary
    ws = wb.active
    ws.title = 'Summary'
    ws['A1'] = 'Work in progress at 30 September 2026'
    ws['A1'].font = Font(name='Calibri', sz=14, b=True, color=NAVY)
    for i, t in enumerate([
            'Built from the Account Transactions export, which carries a job '
            'number on every line of revenue and cost of sales.',
            'A job with cost against it and nothing billed is work in progress. '
            'That is the whole calculation.',
            'The export is profit and loss only, so account 11300 is not in it. '
            'This works out what 11300 ought to say.']):
        ws.cell(3 + i, 1, t).font = NOTE
    r = 8
    band(ws, r, 3, 'WHAT THE JOB LEDGER SAYS')
    r += 1
    rows_out = [
        ('Live jobs with cost and nothing billed', len(groups['Live']), live,
         'This is work in progress'),
        ('Standing codes, not jobs', len(groups['Standing code, not a job']),
         sum(v['cos'] for _, v in groups['Standing code, not a job']),
         'Department buckets. Belongs in overhead, not WIP'),
        ('Stale, last touched over six months ago',
         len(groups['Stale, last touched over six months ago']),
         sum(v['cos'] for _, v in groups['Stale, last touched over six months ago']),
         'Look like write-offs. Need a decision'),
    ]
    for name, n, amt, why in rows_out:
        ws.cell(r, 1, name).font = BODY
        ws.cell(r, 2, n).font = BODY
        cell = ws.cell(r, 3, round(amt, 2))
        cell.font = BODY
        cell.number_format = MONEY
        ws.cell(r, 4, why).font = NOTE
        r += 1
    ws.cell(r, 1, 'Total cost sitting on jobs with no revenue').font = BOLD
    ws.cell(r, 2, len(unbilled)).font = BOLD
    cell = ws.cell(r, 3, round(sum(v['cos'] for v in unbilled.values()), 2))
    cell.font = BOLD
    cell.number_format = MONEY
    for c in range(1, 4):
        ws.cell(r, c).fill = TOTF
    r += 2

    band(ws, r, 3, 'AGAINST THE BALANCE SHEET')
    r += 1
    for name, v in (('Work in progress the job ledger supports', live),
                    ('Account 11300 on the balance sheet', BS_11300),
                    ('Difference', live - BS_11300)):
        ws.cell(r, 1, name).font = BOLD if 'Difference' in name else BODY
        cell = ws.cell(r, 3, round(v, 2))
        cell.font = BOLD if 'Difference' in name else BODY
        cell.number_format = MONEY
        if 'Difference' in name:
            for c in range(1, 4):
                ws.cell(r, c).fill = TOTF
        r += 1
    r += 1
    for t in ['11300 is a credit. A work in progress asset is a debit. So the '
              'account is the wrong way round as well as the wrong size.',
              'Separately, 30 jobs have been billed but still cost more than '
              'they have earned, by '
              f'{sum(v["cos"] - v["rev"] for v in short.values()):,.0f} in total. '
              'Those are margin, not WIP, and are on the third tab.']:
        ws.cell(r, 1, t).font = NOTE
        r += 1
    for c, w in enumerate((48, 8, 18, 52), 1):
        ws.column_dimensions[gl(c)].width = w
    ws.sheet_view.showGridLines = False

    # --------------------------------------------- 2. the unbilled jobs
    ws = wb.create_sheet('Unbilled jobs')
    ws['A1'] = 'Jobs with cost against them and nothing billed, at 30 September 2026'
    ws['A1'].font = Font(name='Calibri', sz=14, b=True, color=NAVY)
    r = 3
    for group in ('Live', 'Standing code, not a job',
                  'Stale, last touched over six months ago'):
        items = sorted(groups[group], key=lambda x: -x[1]['cos'])
        band(ws, r, 4, group.upper())
        r += 1
        header(ws, r, ['Job', 'Cost carried', 'Last movement', 'Main supplier'],
               (46, 16, 16, 34))
        r += 1
        for j, v in items:
            ws.cell(r, 1, j).font = BODY
            cell = ws.cell(r, 2, round(v['cos'], 2))
            cell.font = BODY
            cell.number_format = MONEY
            ws.cell(r, 3, v['last'].date().isoformat()).font = BODY
            ws.cell(r, 4, v['cust'].most_common(1)[0][0] if v['cust'] else '').font = BODY
            r += 1
        ws.cell(r, 1, f'{group}, total').font = BOLD
        cell = ws.cell(r, 2, round(sum(v['cos'] for _, v in items), 2))
        cell.font = BOLD
        cell.number_format = MONEY
        for c in range(1, 5):
            ws.cell(r, c).fill = TOTF
        r += 2
    ws.sheet_view.showGridLines = False

    # ------------------------------- 3. billed but still cost-negative
    ws = wb.create_sheet('Cost over revenue')
    ws['A1'] = 'Jobs billed but still costing more than they have earned'
    ws['A1'].font = Font(name='Calibri', sz=14, b=True, color=NAVY)
    ws['A3'] = ('These are not work in progress. They have been invoiced. '
                'They are jobs that have lost money, or that have more to bill.')
    ws['A3'].font = NOTE
    header(ws, 5, ['Job', 'Billed', 'Cost', 'Shortfall', 'Last movement',
                   'Main contact'], (46, 15, 15, 15, 16, 32))
    r = 6
    for j, v in sorted(short.items(), key=lambda x: -(x[1]['cos'] - x[1]['rev'])):
        ws.cell(r, 1, j).font = BODY
        for c, val in ((2, v['rev']), (3, v['cos']), (4, v['cos'] - v['rev'])):
            cell = ws.cell(r, c, round(val, 2))
            cell.font = BODY
            cell.number_format = MONEY
        ws.cell(r, 5, v['last'].date().isoformat()).font = BODY
        ws.cell(r, 6, v['cust'].most_common(1)[0][0] if v['cust'] else '').font = BODY
        r += 1
    ws.cell(r, 1, 'Total').font = BOLD
    for c, val in ((2, sum(v['rev'] for v in short.values())),
                   (3, sum(v['cos'] for v in short.values())),
                   (4, sum(v['cos'] - v['rev'] for v in short.values()))):
        cell = ws.cell(r, c, round(val, 2))
        cell.font = BOLD
        cell.number_format = MONEY
    for c in range(1, 7):
        ws.cell(r, c).fill = TOTF
    ws.sheet_view.showGridLines = False

    wb.save(OUT)
    print('saved', OUT)
    print(f'  live work in progress        {live:>14,.2f}')
    print(f'  standing codes               '
          f'{sum(v["cos"] for _, v in groups["Standing code, not a job"]):>14,.2f}')
    print(f'  stale                        '
          f'{sum(v["cos"] for _, v in groups["Stale, last touched over six months ago"]):>14,.2f}')
    print(f'  balance sheet 11300          {BS_11300:>14,.2f}')
    print(f'  difference                   {live - BS_11300:>14,.2f}')


if __name__ == '__main__':
    main(sys.argv[1])
