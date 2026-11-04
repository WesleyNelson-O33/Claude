"""Revenue against its own cost, July to September 2026, with the detail.

Tab 1  the five groups, revenue against cost, as before
Tab 2  every transaction behind them: date, contact, invoice number, job
Tab 3  the Aware Super job reconciled, the customer invoice against the bills

Built from the Account Transactions export, which carries an invoice number and
a job number on every line. Totals tie to the profit and loss to the cent.

Run with the export as the only argument.
"""
import collections
import datetime
import sys

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as gl

OUT = 'recon/CTS revenue against cost - Jul to Sep 2026.xlsx'
Q0, Q1 = datetime.datetime(2026, 7, 1), datetime.datetime(2026, 9, 30, 23, 59)
AWARE = '2604704'

NAVY, WHITE, GREY = 'FF242C39', 'FFFFFFFF', 'FFF4F4F4'
BAND = PatternFill('solid', start_color=NAVY, end_color=NAVY)
TOTF = PatternFill('solid', start_color=GREY, end_color=GREY)
RULE = Side(style='thin', color='FFC9CED6')
HEAD = Font(name='Calibri', sz=11, b=True, color=WHITE)
BOLD = Font(name='Calibri', sz=11, b=True, color=NAVY)
BODY = Font(name='Calibri', sz=11, color=NAVY)
SMALL = Font(name='Calibri', sz=10, color='FF6B7585')
NOTE = Font(name='Calibri', sz=10, i=True, color='FF6B7585')
MONEY = '#,##0.00;[RED](#,##0.00);"-"'
PCT = '0.0%;[RED](0.0%);"-"'

GROUPS = [
 ('Equipment sold', ['Equipment Sales'], ['Equipment - Purchase']),
 ('Equipment hired out', ['Equipment Hires'],
  ['Equipment - Hires', 'Service - hires', 'Damage Waiver']),
 ('Licences and subscriptions resold', ['Subscriptions & Licences - Income'],
  ['Subscriptions & Licences - Expense']),
 ('Labour',
  ['Contract Support Staff', 'Installation Labour', 'Production Labour',
   'Ad Hoc Service Labour', 'Conference Call Charge', 'Video Labour',
   'Tech Event Management', 'Tech Consulting/Project Manage'],
  ['Direct Salaries - ONS', 'Direct Wages - ONS', 'Direct Superannuation - ONS',
   'Sub-Contract Labour - ONS', 'Direct Salaries - PRD', 'Direct Wages - PRD',
   'Direct Superannuation - PRD', 'Sub-Contract Labour - PRD',
   'Direct Salaries - VID', 'Direct Wages - VID', 'Direct Superannuation - VID',
   'Direct Salaries - INTEGRATION', 'Direct Superannuation-INTEGRATION',
   'Sub-Contract Labour - INTEGRATION', 'Direct Salaries - CONS',
   'Direct Wages - CONS', 'Direct Superannuation - CONS']),
 ('Travel, freight and mobile recharged',
  ['Meal & Travel - ONS', 'Meal & Travel - PRD', 'Freight & Travel - PRD',
   'Freight & Travel - INTEGRATION', 'Freight & Travel - VID',
   'Mobile Phone Plans'],
  ['Direct Travel - ONS', 'Direct Travel - PRD', 'Direct Travel - CONS',
   'Direct Freight - ONS', 'Direct Freight - PRD', 'Direct Freight - INTEGRATION',
   'Direct Mobile - ONS', 'Direct Mobile - PRD', 'Credit Card Paid - ONS',
   'Credit Card Paid - PRD', 'Credit Card Paid - INTEGRATION',
   'Credit Card Paid - CONS']),
]
WHICH = {}
for g, revs, costs in GROUPS:
    for a in revs:
        WHICH[a] = (g, 'Revenue')
    for a in costs:
        WHICH[a] = (g, 'Cost')


def read(path):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    rows = wb['Account Transactions'].iter_rows(values_only=True)
    for _ in range(5):
        next(rows)
    lines, aware = [], []
    for r in rows:
        if r[0] is None or not hasattr(r[3], 'strftime'):
            continue
        code, name = str(r[0]).strip(), str(r[1]).strip()
        amt = (r[5] or 0) - (r[6] or 0)
        job = str(r[7] or '').strip()
        if job.startswith(AWARE) and (code.startswith('4') or code.startswith('5')):
            aware.append((r[3], code, name, str(r[8] or ''), str(r[9] or ''),
                          str(r[4] or ''), str(r[10] or ''),
                          -amt if code.startswith('4') else amt,
                          'Revenue' if code.startswith('4') else 'Cost'))
        if not (Q0 <= r[3] <= Q1) or name not in WHICH:
            continue
        g, kind = WHICH[name]
        lines.append((g, kind, code, name, r[3], str(r[4] or ''),
                      str(r[8] or ''), str(r[9] or ''), job,
                      str(r[11] or ''), -amt if kind == 'Revenue' else amt))
    wb.close()
    return lines, aware


def band(ws, row, width, text):
    ws.cell(row, 1, text).font = HEAD
    for c in range(1, width + 1):
        ws.cell(row, c).fill = BAND


def header(ws, row, names, widths):
    for c, n in enumerate(names, 1):
        cell = ws.cell(row, c, n)
        cell.font = BOLD
        cell.border = Border(bottom=RULE)
        if n in ('Amount', 'Revenue', 'Cost', 'Margin', 'Margin %'):
            cell.alignment = Alignment(horizontal='right')
    for c, w in enumerate(widths, 1):
        ws.column_dimensions[gl(c)].width = w


def main(path):
    lines, aware = read(path)
    agg = collections.defaultdict(float)
    for g, kind, code, name, *_rest, amt in lines:
        agg[(g, kind, code, name)] += amt
    tot = collections.defaultdict(float)
    for (g, kind, _c, _n), v in agg.items():
        tot[(g, kind)] += v
    rev_all = sum(v for (g, k), v in tot.items() if k == 'Revenue')
    cos_all = sum(v for (g, k), v in tot.items() if k == 'Cost')
    assert abs(rev_all - 1624459.78) < 0.02, rev_all
    assert abs(cos_all - 992160.28) < 0.02, cos_all

    wb = openpyxl.Workbook()

    # ------------------------------------------------------ 1. summary
    ws = wb.active
    ws.title = 'Summary'
    ws['A1'] = 'Revenue against its own cost, July to September 2026'
    ws['A1'].font = Font(name='Calibri', sz=14, b=True, color=NAVY)
    for i, t in enumerate([
            'Every income account in the quarter paired with the cost of sales '
            'accounts that carry its cost. Xero, accrual, whole of entity.',
            'Which cost belongs to which revenue is a judgement, set out '
            'account by account so it can be argued with. Nothing is '
            'reallocated and every account appears once.',
            'The Detail tab behind this carries every transaction, with its '
            'invoice number and job number.']):
        ws.cell(3 + i, 1, t).font = NOTE
    r = 7
    header(ws, r, ['Account', 'Revenue', 'Cost', 'Margin', 'Margin %'],
           (44, 17, 17, 17, 12))
    for c in range(1, 6):
        ws.cell(r, c).font = HEAD
        ws.cell(r, c).fill = BAND
    r += 1
    group_rows = []
    for g, revs, costs in GROUPS:
        head_row = r
        ws.cell(r, 1, g.upper()).font = BOLD
        for c in range(1, 6):
            ws.cell(r, c).fill = TOTF
            ws.cell(r, c).border = Border(top=RULE)
        r += 1
        for kind, accts, col in (('Revenue', revs, 2), ('Cost', costs, 3)):
            for a in sorted(accts,
                            key=lambda x: -abs(agg.get((g, kind, _code(agg, g, kind, x), x), 0))):
                keys = [k for k in agg if k[0] == g and k[1] == kind and k[3] == a]
                if not keys:
                    continue
                k = keys[0]
                ws.cell(r, 1, f'    {k[2]}  {a}').font = BODY if kind == 'Revenue' else SMALL
                cell = ws.cell(r, col, round(agg[k], 2))
                cell.font = BODY if kind == 'Revenue' else SMALL
                cell.number_format = MONEY
                r += 1
        ws.cell(head_row, 2, round(tot[(g, 'Revenue')], 2)).number_format = MONEY
        ws.cell(head_row, 3, round(tot[(g, 'Cost')], 2)).number_format = MONEY
        ws.cell(head_row, 4, f'=B{head_row}-C{head_row}').number_format = MONEY
        ws.cell(head_row, 5,
                f'=IF(B{head_row}=0,"",D{head_row}/B{head_row})').number_format = PCT
        for c in range(2, 6):
            ws.cell(head_row, c).font = BOLD
        group_rows.append(head_row)
        r += 1
    ws.cell(r, 1, 'GROSS PROFIT').font = HEAD
    for c in range(1, 6):
        ws.cell(r, c).fill = BAND
    for c in (2, 3):
        cell = ws.cell(r, c, '=' + '+'.join(f'{gl(c)}{x}' for x in group_rows))
        cell.font = HEAD
        cell.number_format = MONEY
    ws.cell(r, 4, f'=B{r}-C{r}').number_format = MONEY
    ws.cell(r, 5, f'=IF(B{r}=0,"",D{r}/B{r})').number_format = PCT
    for c in (4, 5):
        ws.cell(r, c).font = HEAD
    ws.sheet_view.showGridLines = False

    # ------------------------------------------------------- 2. detail
    ws = wb.create_sheet('Detail')
    header(ws, 1, ['Group', 'Revenue or cost', 'Code', 'Account', 'Date',
                   'Contact', 'Invoice number', 'Reference', 'Job', 'Cost centre',
                   'Amount'],
           (30, 15, 8, 34, 12, 32, 26, 22, 36, 14, 15))
    for c in range(1, 12):
        ws.cell(1, c).font = HEAD
        ws.cell(1, c).fill = BAND
    r = 2
    for row in sorted(lines, key=lambda x: (x[0], x[1] != 'Revenue', x[3], x[4])):
        for c, v in enumerate(row, 1):
            cell = ws.cell(r, c, v.date() if hasattr(v, 'date') else v)
            cell.font = BODY
            if c == 11:
                cell.number_format = MONEY
            if c == 5:
                cell.number_format = 'dd mmm yyyy'
        r += 1
    ws.auto_filter.ref = f'A1:K{r - 1}'
    ws.freeze_panes = 'A2'
    ws.sheet_view.showGridLines = False
    detail_rows = r - 2

    # -------------------------------------------- 3. the Aware Super job
    ws = wb.create_sheet('Aware Super job')
    ws['A1'] = 'Job 2604704, Aware Super Seminar Room Audio'
    ws['A1'].font = Font(name='Calibri', sz=14, b=True, color=NAVY)
    for i, t in enumerate([
            'Everything ever posted to the job, from the first accrual to the '
            'invoice. The customer invoice is INV-10596, raised 9 September.',
            'The supplier bill is INV41565 from Complete AV Solutions, marked '
            'Claim 2 of 3, so there is more of this job still to come.']):
        ws.cell(3 + i, 1, t).font = NOTE
    r = 6
    for kind, label in (('Revenue', 'WHAT WAS BILLED'), ('Cost', 'WHAT IT COST')):
        band(ws, r, 6, label)
        r += 1
        header(ws, r, ['Date', 'Code', 'Account', 'Invoice number', 'Contact',
                       'Amount'], (12, 8, 34, 28, 32, 15))
        r += 1
        sub = [x for x in aware if x[8] == kind]
        for d, code, name, num, ref, who, desc, amt, _k in sorted(sub):
            ws.cell(r, 1, d.date()).number_format = 'dd mmm yyyy'
            ws.cell(r, 2, code)
            ws.cell(r, 3, name)
            ws.cell(r, 4, num or ref)
            ws.cell(r, 5, who)
            cell = ws.cell(r, 6, round(amt, 2))
            cell.number_format = MONEY
            for c in range(1, 7):
                ws.cell(r, c).font = BODY
            r += 1
        ws.cell(r, 1, f'Total {kind.lower()}').font = BOLD
        cell = ws.cell(r, 6, round(sum(x[7] for x in sub), 2))
        cell.font = BOLD
        cell.number_format = MONEY
        for c in range(1, 7):
            ws.cell(r, c).fill = TOTF
        r += 2
    rv = sum(x[7] for x in aware if x[8] == 'Revenue')
    cv = sum(x[7] for x in aware if x[8] == 'Cost')
    band(ws, r, 6, 'THE JOB')
    r += 1
    for name, v, fmt in (('Billed', rv, MONEY), ('Cost', cv, MONEY),
                         ('Margin', rv - cv, MONEY),
                         ('Margin %', (rv - cv) / rv, PCT)):
        ws.cell(r, 1, name).font = BOLD
        cell = ws.cell(r, 6, round(v, 4))
        cell.font = BOLD
        cell.number_format = fmt
        r += 1
    r += 1
    band(ws, r, 6, 'HOW THE SPLIT GOES WRONG')
    r += 1
    eq_r = sum(x[7] for x in aware if x[8] == 'Revenue' and x[1] == '42600')
    eq_c = sum(x[7] for x in aware if x[8] == 'Cost' and x[1] == '51100')
    lb_r = sum(x[7] for x in aware if x[8] == 'Revenue' and x[1] in ('42300', '42800'))
    lb_c = sum(x[7] for x in aware if x[8] == 'Cost' and x[1] != '51100')
    header(ws, r, ['Line', '', '', '', 'Billed', 'Cost'], (12, 8, 34, 28, 32, 15))
    ws.cell(r, 1, 'Line')
    ws.cell(r, 5, 'Billed')
    ws.cell(r, 6, 'Cost')
    r += 1
    for name, rr, cc in (('Equipment', eq_r, eq_c), ('Labour and other', lb_r, lb_c)):
        ws.cell(r, 1, name).font = BODY
        for c, v in ((5, rr), (6, cc)):
            cell = ws.cell(r, c, round(v, 2))
            cell.font = BODY
            cell.number_format = MONEY
        r += 1
    r += 1
    for t in [f'Equipment cost {eq_c:,.2f} against equipment billed {eq_r:,.2f}. '
              f'The equipment line on this job reads as a loss of '
              f'{eq_r - eq_c:,.2f}.',
              f'Labour and other cost {lb_c:,.2f} against {lb_r:,.2f} billed, '
              f'which reads as a margin of {(lb_r - lb_c) / lb_r:.0%}.',
              'The job as a whole is fine. The split between the two lines is '
              'not, and it is what makes equipment look loss-making across '
              'the whole company.']:
        ws.cell(r, 1, t).font = NOTE
        r += 1
    ws.sheet_view.showGridLines = False

    wb.save(OUT)
    print('saved', OUT)
    print(f'  detail rows {detail_rows:,}')
    print(f'  Aware Super job: billed {rv:,.2f}  cost {cv:,.2f}  '
          f'margin {rv - cv:,.2f}  {(rv - cv) / rv:.1%}')
    print(f'    equipment billed {eq_r:,.2f} against cost {eq_c:,.2f}')
    print(f'    labour and other {lb_r:,.2f} against cost {lb_c:,.2f}')


def _code(agg, g, kind, acct):
    for k in agg:
        if k[0] == g and k[1] == kind and k[3] == acct:
            return k[2]
    return ''


if __name__ == '__main__':
    main(sys.argv[1])
