"""What sits on the four standing job codes at 30 September 2026.

These four carry 85,531.16 of cost with no revenue against them. They came up
in the work in progress schedule as the largest block, and they are not work in
progress: a standing code is not a job, so nothing on it is ever going to be
billed out.

Tab 1  the four codes, by account and by supplier
Tab 2  every one of the lines behind them

Run with the Account Transactions export as the only argument.
"""
import collections
import datetime
import sys

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as gl

OUT = 'recon/CTS standing job codes - 30 September 2026.xlsx'
CUT = datetime.datetime(2026, 9, 30, 23, 59)
CODES = {'9100': '9100 PRODUCTION [PRD]',
         '9105': '9105 ONSITE (including PROD backup)',
         '2479': '2479 CONSULTING - Quotes [CONS]',
         '9109': '9109 ONSITE - Marketing [ONS]'}

NAVY, WHITE, GREY = 'FF242C39', 'FFFFFFFF', 'FFF4F4F4'
BAND = PatternFill('solid', start_color=NAVY, end_color=NAVY)
TOTF = PatternFill('solid', start_color=GREY, end_color=GREY)
RULE = Side(style='thin', color='FFC9CED6')
HEAD = Font(name='Calibri', sz=11, b=True, color=WHITE)
BOLD = Font(name='Calibri', sz=11, b=True, color=NAVY)
BODY = Font(name='Calibri', sz=11, color=NAVY)
NOTE = Font(name='Calibri', sz=10, i=True, color='FF6B7585')
MONEY = '#,##0.00;[RED](#,##0.00);"-"'


def read(path):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    rows = wb['Account Transactions'].iter_rows(values_only=True)
    for _ in range(5):
        next(rows)
    out = []
    for r in rows:
        if r[0] is None or not hasattr(r[3], 'strftime') or r[3] > CUT:
            continue
        job = str(r[7] or '').strip()
        head = job.split(' - ')[0].strip()
        if head not in CODES:
            continue
        code = str(r[0]).strip()
        if not (code.startswith('4') or code.startswith('5')):
            continue
        out.append({
            'head': head, 'job': job, 'date': r[3], 'code': code,
            'account': str(r[1]).strip(), 'source': str(r[2]),
            'contact': str(r[4] or '').strip() or '(journal, no contact)',
            'invoice': str(r[8] or '').strip(), 'reference': str(r[9] or '').strip(),
            'description': str(r[10] or '').strip(),
            'amount': (r[5] or 0) - (r[6] or 0)})
    wb.close()
    return out


def band(ws, row, width, text):
    ws.cell(row, 1, text).font = HEAD
    for c in range(1, width + 1):
        ws.cell(row, c).fill = BAND


def header(ws, row, names, widths, right=()):
    for c, n in enumerate(names, 1):
        cell = ws.cell(row, c, n)
        cell.font = BOLD
        cell.border = Border(bottom=RULE)
        if n in right:
            cell.alignment = Alignment(horizontal='right')
    for c, w in enumerate(widths, 1):
        ws.column_dimensions[gl(c)].width = w


def main(path):
    lines = read(path)
    total = sum(x['amount'] for x in lines)
    assert abs(total - 85531.16) < 0.02, total

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Summary'
    ws['A1'] = 'The four standing job codes at 30 September 2026'
    ws['A1'].font = Font(name='Calibri', sz=14, b=True, color=NAVY)
    for i, t in enumerate([
            'These four carry 85,531.16 of cost with nothing billed against '
            'them. They are not work in progress. A standing code is not a '
            'job, so none of this is ever going to be billed out.',
            'Cumulative, from the start of the export to 30 September 2026.']):
        ws.cell(3 + i, 1, t).font = NOTE

    r = 7
    for head in ('9100', '9105', '2479', '9109'):
        mine = [x for x in lines if x['head'] == head]
        band(ws, r, 5, f'{CODES[head]}   {sum(x["amount"] for x in mine):,.2f}')
        r += 1
        header(ws, r, ['Account', 'Supplier or journal', 'Lines', 'First', 'Amount'],
               (40, 42, 8, 13, 15), right=('Lines', 'Amount'))
        r += 1
        grp = collections.defaultdict(lambda: {'n': 0, 'amt': 0.0, 'd': []})
        for x in mine:
            g = grp[(f'{x["code"]} {x["account"]}', x['contact'])]
            g['n'] += 1
            g['amt'] += x['amount']
            g['d'].append(x['date'])
        for (acct, who), v in sorted(grp.items(), key=lambda x: -x[1]['amt']):
            ws.cell(r, 1, acct).font = BODY
            ws.cell(r, 2, who).font = BODY
            ws.cell(r, 3, v['n']).font = BODY
            ws.cell(r, 4, min(v['d']).date()).number_format = 'dd mmm yyyy'
            ws.cell(r, 4).font = BODY
            cell = ws.cell(r, 5, round(v['amt'], 2))
            cell.font = BODY
            cell.number_format = MONEY
            r += 1
        ws.cell(r, 1, f'{CODES[head]}, total').font = BOLD
        cell = ws.cell(r, 5, round(sum(x['amount'] for x in mine), 2))
        cell.font = BOLD
        cell.number_format = MONEY
        for c in range(1, 6):
            ws.cell(r, c).fill = TOTF
        r += 2

    ws.cell(r, 1, 'ALL FOUR').font = HEAD
    cell = ws.cell(r, 5, round(total, 2))
    cell.font = HEAD
    cell.number_format = MONEY
    for c in range(1, 6):
        ws.cell(r, c).fill = BAND
    r += 2
    for t in ['The largest single item is 70,153.33 on account 51400 Service '
              'hires, all of it on 9100. It is two recurring charges: a LiveU '
              'data plan at 1,350.00 a month over fifteen bills, and a Viocorp '
              'prepayment released at 4,536.67 a month over eleven journals.',
              'The rest is recurring running cost. Avangate software, Telstra '
              'and Optus mobile, Transport for NSW tolls, Goget car share.',
              'The two smallest codes are payroll only. 2479 is time spent '
              'quoting and 9109 is marketing time, both non-chargeable.',
              'None of this belongs in work in progress. It belongs in '
              'overhead, or on the jobs that used it.']:
        ws.cell(r, 1, t).font = NOTE
        r += 1
    ws.sheet_view.showGridLines = False

    # ------------------------------------------------------- the lines
    ws = wb.create_sheet('Every line')
    header(ws, 1, ['Job', 'Date', 'Code', 'Account', 'Source', 'Contact',
                   'Invoice number', 'Reference', 'Description', 'Amount'],
           (36, 12, 8, 36, 18, 36, 28, 20, 56, 14), right=('Amount',))
    for c in range(1, 11):
        ws.cell(1, c).font = HEAD
        ws.cell(1, c).fill = BAND
    r = 2
    for x in sorted(lines, key=lambda y: (y['head'], y['code'], y['date'])):
        for c, v in enumerate([x['job'], x['date'].date(), x['code'], x['account'],
                               x['source'], x['contact'], x['invoice'],
                               x['reference'], x['description'],
                               round(x['amount'], 2)], 1):
            cell = ws.cell(r, c, v)
            cell.font = BODY
            if c == 2:
                cell.number_format = 'dd mmm yyyy'
            if c == 10:
                cell.number_format = MONEY
        r += 1
    ws.cell(r, 1, 'Total').font = BOLD
    cell = ws.cell(r, 10, round(total, 2))
    cell.font = BOLD
    cell.number_format = MONEY
    for c in range(1, 11):
        ws.cell(r, c).fill = TOTF
    ws.auto_filter.ref = f'A1:J{r - 1}'
    ws.freeze_panes = 'A2'
    ws.sheet_view.showGridLines = False

    wb.save(OUT)
    print('saved', OUT)
    print(f'  {len(lines)} lines, {total:,.2f}')
    for head in ('9100', '9105', '2479', '9109'):
        mine = [x for x in lines if x['head'] == head]
        print(f'    {CODES[head]:<40}{len(mine):>5} lines'
              f'{sum(x["amount"] for x in mine):>13,.2f}')


if __name__ == '__main__':
    main(sys.argv[1])
