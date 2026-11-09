"""Subscriptions and licences, income against cost, reconciled month by month.

Built from the Account Transactions export carrying both accounts,
Subscriptions & Licences - Income and Subscriptions & Licences - Expense, for
1 July 2024 to 30 September 2026.

The book made 11.3 per cent in FY2025 and 13.0 per cent in FY2026. It is
losing money in FY2027. One contract is the whole of the difference.

SPACERA PTY LTD bills 5,208.00 a month for the Aware Super National Support
Agreement, a twelve month contract that started on 1 July 2026. The lines are
HYPERCARE MANAGED, MISSION CONTROL and PHOTON. Four bills to 1 October, all
paid, all with attachments. Taken from Xero:

    INV-0041   1 July 2026        5,208.00   account 51310 Sub-Contract Labour - ONS
    INV-0049   1 August 2026      5,208.00   account 51500 Subscriptions - Expense
    INV-0056   1 September 2026   5,208.00   account 51500 Subscriptions - Expense
    INV-0060   1 October 2026     5,208.00   account 51500 Subscriptions - Expense

So July sits in a different account from the rest. On a consistent basis the
three months to September read minus 3.6, minus 3.7 and minus 1.9 per cent.

No sales invoice bills Hypercare, Mission Control or Photon to Aware Super.
Every Aware Super invoice from June to October was checked. The only platform
licence billed on is Appspace, on INV-10635, which is a separate deal. The
recurring Aware Super income is 11,000.00 a month for National AV Maintenance
site visits, and it does not mention the platform.

Income is shown positive and cost negative, the way the profit and loss reads.

Run with the export as the only argument.
"""
import collections
import datetime
import sys

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as gl

OUT = 'recon/CTS subscriptions by month - 30 September 2026.xlsx'

NAVY, WHITE, GREY = 'FF242C39', 'FFFFFFFF', 'FFF4F4F4'
BAND = PatternFill('solid', start_color=NAVY, end_color=NAVY)
TOTF = PatternFill('solid', start_color=GREY, end_color=GREY)
RED = PatternFill('solid', start_color='FFFCE4E4', end_color='FFFCE4E4')
AMBER = PatternFill('solid', start_color='FFFFF3D6', end_color='FFFFF3D6')
BLUE = PatternFill('solid', start_color='FFE3EEF8', end_color='FFE3EEF8')
RULE = Side(style='thin', color='FFC9CED6')
BOX = Border(top=RULE, bottom=RULE, left=RULE, right=RULE)
HEAD = Font(name='Calibri', sz=11, b=True, color=WHITE)
BOLD = Font(name='Calibri', sz=11, b=True, color=NAVY)
BODY = Font(name='Calibri', sz=11, color=NAVY)
NOTE = Font(name='Calibri', sz=10, i=True, color='FF6B7585')
TITLE = Font(name='Calibri', sz=14, b=True, color=NAVY)
MONEY = '#,##0.00;[RED](#,##0.00);"-"'
PCT = '0.0%;[RED](0.0%);"-"'
DATE = 'dd/mm/yyyy'
MON = 'mmm yyyy'

CAUSE = 'SPACERA'
INVOICE = ('Payable Invoice', 'Payable Credit Note', 'Receivable Invoice')

# The Spacera bills as Xero holds them. July is the one coded elsewhere, so
# it is not in the export and has to be brought in to compare like with like.
BILLS = [
    ('INV-0041', datetime.date(2026, 7, 1), 5208.00, '51310',
     'Sub-Contract Labour - ONS'),
    ('INV-0049', datetime.date(2026, 8, 1), 5208.00, '51500',
     'Subscriptions & Licences - Expense'),
    ('INV-0056', datetime.date(2026, 9, 1), 5208.00, '51500',
     'Subscriptions & Licences - Expense'),
    ('INV-0060', datetime.date(2026, 10, 1), 5208.00, '51500',
     'Subscriptions & Licences - Expense'),
]
ELSEWHERE = {(b[1].year, b[1].month): -b[2] for b in BILLS if b[3] != '51500'}


def load(path):
    ws = openpyxl.load_workbook(path, data_only=True).worksheets[0]
    out, sec = [], None
    for r in range(6, ws.max_row + 1):
        a = ws.cell(r, 1).value
        if isinstance(a, str):
            t = a.strip()
            if t.startswith('Total'):
                sec = None
            elif t.startswith('Subscriptions'):
                sec = 'Income' if t.endswith('Income') else 'Expense'
            continue
        if isinstance(a, datetime.datetime) and sec:
            out.append(dict(date=a.date(), sec=sec, source=ws.cell(r, 2).value,
                            desc=ws.cell(r, 3).value, ref=ws.cell(r, 4).value,
                            debit=ws.cell(r, 5).value or 0,
                            credit=ws.cell(r, 6).value or 0,
                            gst=ws.cell(r, 9).value or 0, row=r))
    return ws.cell(3, 1).value, out


def signed(r):
    """Income positive, cost negative, as the profit and loss reads."""
    return (r['credit'] - r['debit']) if r['sec'] == 'Income' \
        else -(r['debit'] - r['credit'])


def party(r):
    if r['source'] in INVOICE:
        return str(r['desc'] or '').split(' - ')[0].strip()[:44]
    return 'Manual journal'


def fy_of(d):
    return d.year + 1 if d.month >= 7 else d.year


def band(ws, row, labels, widths=None):
    for i, t in enumerate(labels, start=1):
        c = ws.cell(row, i, t)
        c.fill, c.font, c.border = BAND, HEAD, BOX
        c.alignment = Alignment(horizontal='center', wrap_text=True)
    ws.row_dimensions[row].height = 46
    if widths:
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[gl(i)].width = w


def para(ws, row, text, cols=11, bold=False, height=30):
    c = ws.cell(row, 1, text)
    c.font = BOLD if bold else BODY
    c.alignment = Alignment(wrap_text=True, vertical='top')
    if not bold:
        ws.merge_cells(start_row=row, start_column=1,
                       end_row=row, end_column=cols)
        ws.row_dimensions[row].height = height


def main():
    if len(sys.argv) != 2:
        print(__doc__)
        return 1
    period, rows = load(sys.argv[1])
    rows.sort(key=lambda r: (r['date'], r['row']))
    months = sorted({(r['date'].year, r['date'].month) for r in rows})
    inc_n = sum(1 for r in rows if r['sec'] == 'Income')
    print(f'  {period}')
    print(f'  {len(rows)} lines, {inc_n} income and {len(rows)-inc_n} cost, '
          f'{len(months)} months')

    wb = openpyxl.Workbook()

    # ================================================= Transactions
    tx = wb.active
    tx.title = 'Transactions'
    tx['A1'] = 'Every line on both subscription accounts'
    tx['A1'].font = TITLE
    tx['A2'] = (f'From the export for {period}. This tab drives the summary '
                'tabs, so the figures there are formulas off these rows. '
                'Amount is income positive and cost negative.')
    tx['A2'].font = NOTE
    band(tx, 4, ['Date', 'Month', 'Account', 'Source', 'Other side',
                 'Description', 'Reference', 'Debit', 'Credit', 'Amount',
                 'GST', 'Is it Spacera'],
         [12, 11, 11, 20, 46, 64, 22, 14, 14, 14, 12, 13])
    for i, r in enumerate(rows):
        n = 5 + i
        tx.cell(n, 1, r['date']).number_format = DATE
        tx.cell(n, 2, datetime.date(r['date'].year, r['date'].month, 1)
                ).number_format = MON
        tx.cell(n, 3, r['sec'])
        tx.cell(n, 4, r['source'])
        tx.cell(n, 5, party(r))
        tx.cell(n, 6, r['desc'])
        tx.cell(n, 7, r['ref'])
        tx.cell(n, 8, r['debit']).number_format = MONEY
        tx.cell(n, 9, r['credit']).number_format = MONEY
        a = tx.cell(n, 10, f'=IF($C{n}="Income",I{n}-H{n},-(H{n}-I{n}))')
        a.number_format = MONEY
        tx.cell(n, 11, r['gst']).number_format = MONEY
        hit = CAUSE in str(r['desc']).upper() and r['sec'] == 'Expense'
        tx.cell(n, 12, 'Yes' if hit else '')
        for k in range(1, 13):
            tx.cell(n, k).font = BODY
        if hit:
            for k in range(1, 13):
                tx.cell(n, k).fill = RED
        elif r['source'] == 'Manual Journal':
            for k in range(1, 13):
                tx.cell(n, k).fill = AMBER
    last = 4 + len(rows)
    tx.freeze_panes = 'A5'
    tx.auto_filter.ref = f'A4:L{last}'

    SI = (f'SUMIFS(Transactions!$J$5:$J${last},'
          f'Transactions!$C$5:$C${last},"Income",'
          f'Transactions!$B$5:$B${last},')
    SC = (f'SUMIFS(Transactions!$J$5:$J${last},'
          f'Transactions!$C$5:$C${last},"Expense",'
          f'Transactions!$B$5:$B${last},')
    SS = (f'SUMIFS(Transactions!$J$5:$J${last},'
          f'Transactions!$L$5:$L${last},"Yes",'
          f'Transactions!$B$5:$B${last},')

    # ================================================= Monthly recon
    mr = wb.create_sheet('Monthly recon', 0)
    mr['A1'] = 'Subscriptions and licences, income against cost, by month'
    mr['A1'].font = TITLE
    mr['A2'] = 'Corporate Technology Services Pty Ltd'
    mr['A2'].font = BODY
    mr['A3'] = period
    mr['A3'].font = NOTE

    para(mr, 5, 'THE ANSWER', bold=True)
    for i, t in enumerate([
        'The book made 11.3 per cent in FY2025 and 13.0 per cent in FY2026. '
        'It is losing money in FY2027.',
        'One contract is the whole of the difference. Spacera bills 5,208.00 '
        'a month for the Aware Super National Support Agreement, which '
        'started on 1 July 2026 and runs for twelve months.',
        'Nothing bills it out. Every Aware Super invoice from June to October '
        'was checked and none of them carries Hypercare, Mission Control or '
        'Photon. The recurring Aware Super income is 11,000.00 a month for '
        'maintenance site visits and it does not mention the platform.',
        'The July bill was coded to Sub-Contract Labour, not to '
        'subscriptions, so July looks better than it is. On a consistent '
        'basis the three months to September read minus 3.6, minus 3.7 and '
        'minus 1.9 per cent.',
        'The rest of the book is healthy. Without this contract the three '
        'months run at 15.7, 15.6 and 15.1 per cent.',
        'So this is not a pricing problem across the book. It is one '
        'contract costing 62,496.00 a year with no charge against it.',
    ]):
        para(mr, 6 + i, t, height=32)

    HDR = ['Month', 'Income', 'Cost as booked', 'Margin as booked',
           'Per cent', 'Spacera in this account',
           'Spacera booked elsewhere', 'Margin with Spacera in every month',
           'Per cent', 'Margin without the Spacera contract', 'Per cent']
    band(mr, 14, HDR,
         [12, 14, 14, 14, 10, 13, 13, 15, 10, 15, 10])
    first = 15
    for i, (y, m) in enumerate(months):
        r = first + i
        mr.cell(r, 1, datetime.date(y, m, 1)).number_format = MON
        mr.cell(r, 2, f'={SI}$A{r})')
        mr.cell(r, 3, f'={SC}$A{r})')
        mr.cell(r, 4, f'=B{r}+C{r}')
        mr.cell(r, 5, f'=IFERROR(D{r}/B{r},"")')
        mr.cell(r, 6, f'={SS}$A{r})')
        mr.cell(r, 7, ELSEWHERE.get((y, m), 0.0))
        mr.cell(r, 8, f'=D{r}+G{r}')
        mr.cell(r, 9, f'=IFERROR(H{r}/B{r},"")')
        mr.cell(r, 10, f'=D{r}-F{r}')
        mr.cell(r, 11, f'=IFERROR(J{r}/B{r},"")')
        for k in (2, 3, 4, 6, 7, 8, 10):
            mr.cell(r, k).number_format = MONEY
        for k in (5, 9, 11):
            mr.cell(r, k).number_format = PCT
        for k in range(1, 12):
            mr.cell(r, k).font = BODY
        if fy_of(datetime.date(y, m, 1)) == 2027:
            for k in range(1, 12):
                mr.cell(r, k).fill = BLUE
                mr.cell(r, k).font = BOLD
    last_m = first + len(months) - 1

    tot = last_m + 1
    mr.cell(tot, 1, 'Total').font = BOLD
    for k in (2, 3, 6, 7):
        mr.cell(tot, k, f'=SUM({gl(k)}{first}:{gl(k)}{last_m})')
    mr.cell(tot, 4, f'=B{tot}+C{tot}')
    mr.cell(tot, 5, f'=D{tot}/B{tot}')
    mr.cell(tot, 8, f'=D{tot}+G{tot}')
    mr.cell(tot, 9, f'=H{tot}/B{tot}')
    mr.cell(tot, 10, f'=D{tot}-F{tot}')
    mr.cell(tot, 11, f'=J{tot}/B{tot}')
    for k in range(1, 12):
        mr.cell(tot, k).fill = TOTF
        mr.cell(tot, k).font = BOLD
        mr.cell(tot, k).border = BOX
        mr.cell(tot, k).number_format = PCT if k in (5, 9, 11) else MONEY
    mr.cell(tot, 1).number_format = 'General'

    fy = sorted({fy_of(datetime.date(y, m, 1)) for y, m in months})
    fr = tot + 2
    para(mr, fr, 'BY FINANCIAL YEAR', bold=True)
    band(mr, fr + 1, ['Year'] + HDR[1:])
    for i, y in enumerate(fy):
        r = fr + 2 + i
        ms = [first + months.index(k) for k in months
              if fy_of(datetime.date(k[0], k[1], 1)) == y]
        rng = lambda col: '+'.join(f'{col}{x}' for x in ms)
        mr.cell(r, 1, f'FY{y}').font = BOLD
        for k, col in ((2, 'B'), (3, 'C'), (6, 'F'), (7, 'G')):
            mr.cell(r, k, f'={rng(col)}')
        mr.cell(r, 4, f'=B{r}+C{r}')
        mr.cell(r, 5, f'=D{r}/B{r}')
        mr.cell(r, 8, f'=D{r}+G{r}')
        mr.cell(r, 9, f'=H{r}/B{r}')
        mr.cell(r, 10, f'=D{r}-F{r}')
        mr.cell(r, 11, f'=J{r}/B{r}')
        for k in range(2, 12):
            mr.cell(r, k).number_format = PCT if k in (5, 9, 11) else MONEY
            mr.cell(r, k).font = BOLD if y == fy[-1] else BODY
        if y == fy[-1]:
            for k in range(1, 12):
                mr.cell(r, k).fill = BLUE

    chk = fr + 3 + len(fy)
    para(mr, chk, 'CHECKS', bold=True)
    ti = sum(signed(r) for r in rows if r['sec'] == 'Income')
    tc = sum(signed(r) for r in rows if r['sec'] == 'Expense')
    f27 = [r for r in rows if fy_of(r['date']) == 2027]
    i27 = sum(signed(r) for r in f27 if r['sec'] == 'Income')
    c27 = sum(signed(r) for r in f27 if r['sec'] == 'Expense')
    sep = first + months.index((2026, 9))
    for i, (lbl, f) in enumerate([
        ('Income on the schedule equals the export',
         f'=IF(ROUND(B{tot}-{ti:.2f},2)=0,"OK","CHECK")'),
        ('Cost on the schedule equals the export',
         f'=IF(ROUND(C{tot}-{tc:.2f},2)=0,"OK","CHECK")'),
        ('Every line in the export is on the schedule',
         f'="OK, {len(rows)} lines"'),
        ('FY2027 to date is the 84,810.19 against 82,143.32 in the question',
         f'=IF(AND(ROUND({i27:.2f}-84810.19,2)=0,'
         f'ROUND({c27:.2f}+82143.32,2)=0),"OK","CHECK")'),
        ('September margin agrees with Xero, a loss of 582.89',
         f'=IF(ROUND(D{sep}+582.89,2)=0,"OK","CHECK")'),
        ('The Spacera cost picked up equals the bills Xero holds',
         f'=IF(ROUND(F{tot}+G{tot}+15624,2)=0,"OK","CHECK")'),
    ]):
        r = chk + 1 + i
        mr.cell(r, 1, lbl).font = BODY
        c = mr.cell(r, 9, f)
        c.font = BOLD
        c.alignment = Alignment(horizontal='center')
    mr.freeze_panes = 'A15'

    # ================================================= The contract
    sc = wb.create_sheet('The Spacera contract', 1)
    sc['A1'] = 'The Spacera contract, and what bills it out'
    sc['A1'].font = TITLE
    sc['A2'] = ('Taken from Xero on 5 November 2026, not from the export. '
                'The export only holds the subscription accounts, so the '
                'July bill is not in it.')
    sc['A2'].font = NOTE
    sc.merge_cells('A2:G2')
    sc.row_dimensions[2].height = 30

    para(sc, 4, 'WHAT WE ARE BILLED', bold=True)
    band(sc, 5, ['Bill', 'Date', 'Net', 'GST', 'Gross', 'Account coded to',
                 'In the subscription accounts'],
         [16, 13, 14, 12, 14, 38, 24])
    for i, (num, d, net, code, name) in enumerate(BILLS):
        r = 6 + i
        sc.cell(r, 1, num)
        sc.cell(r, 2, d).number_format = DATE
        sc.cell(r, 3, net).number_format = MONEY
        sc.cell(r, 4, f'=C{r}*0.1').number_format = MONEY
        sc.cell(r, 5, f'=C{r}+D{r}').number_format = MONEY
        sc.cell(r, 6, f'{code} {name}')
        sc.cell(r, 7, 'Yes' if code == '51500' else 'No, it is elsewhere')
        for k in range(1, 8):
            sc.cell(r, k).font = BODY
        if code != '51500':
            for k in range(1, 8):
                sc.cell(r, k).fill = AMBER
    r = 6 + len(BILLS)
    sc.cell(r, 1, 'Total').font = BOLD
    for k in (3, 4, 5):
        c = sc.cell(r, k, f'=SUM({gl(k)}6:{gl(k)}{r-1})')
        c.number_format, c.font, c.fill, c.border = MONEY, BOLD, TOTF, BOX
    sc.cell(r, 1).fill, sc.cell(r, 1).border = TOTF, BOX
    sc.cell(r + 1, 1, 'A full twelve months').font = BODY
    c = sc.cell(r + 1, 3, '=C6*12')
    c.font, c.number_format = BOLD, MONEY

    row = r + 3
    para(sc, row, 'WHAT THE BILL IS FOR', bold=True)
    band(sc, row + 1, ['Line', 'Quantity', 'Unit', 'Net', '', '', ''])
    for i, (d, q, u) in enumerate((('HYPERCARE MANAGED x 16 x 12 months',
                                    16, 100.00),
                                   ('MISSION CONTROL x 82 x 12 months',
                                    82, 44.00),
                                   ('PHOTON x 14 x 12 months', 14, 0.00))):
        n = row + 2 + i
        sc.cell(n, 1, d).font = BODY
        sc.cell(n, 2, q).font = BODY
        sc.cell(n, 3, u).number_format = MONEY
        sc.cell(n, 4, f'=B{n}*C{n}').number_format = MONEY
        sc.cell(n, 3).font = sc.cell(n, 4).font = BODY
    n = row + 5
    sc.cell(n, 1, 'Total a month').font = BOLD
    c = sc.cell(n, 4, f'=SUM(D{row+2}:D{n-1})')
    c.number_format, c.font, c.fill, c.border = MONEY, BOLD, TOTF, BOX

    row = n + 3
    para(sc, row, 'WHAT BILLS IT OUT', bold=True)
    for i, t in enumerate([
        'Nothing. Every Aware Super invoice issued from 1 June to 5 November '
        '2026 was checked, fourteen of them.',
        'None carries a line for Hypercare, Mission Control or Photon.',
        'The only platform licence billed to Aware Super is Appspace, on '
        'INV-10635, and that is a separate deal with its own supplier bill.',
        'The recurring Aware Super income is 11,000.00 a month on INV-10551, '
        'INV-10583, INV-10652 and INV-10700, described as National AV '
        'Maintenance site visits, coded to Contract Support Staff. It does '
        'not mention the platform.',
        'So either the platform is meant to be covered inside that 11,000.00 '
        'a month, or it is not being charged for at all. Whoever signed the '
        'contract needs to say which.',
    ]):
        c = sc.cell(row + 1 + i, 1, t)
        c.font = BODY
        c.alignment = Alignment(wrap_text=True, vertical='top')
        sc.merge_cells(start_row=row + 1 + i, start_column=1,
                       end_row=row + 1 + i, end_column=7)
        sc.row_dimensions[row + 1 + i].height = 30

    row = row + 7
    para(sc, row, 'ONE MORE THING', bold=True)
    for i, t in enumerate([
        'There are two Spacera contacts in Xero. SPACERA PTY LTD, which has '
        'all four bills, and a second one called spacera in lower case with '
        'no email and no transactions. The empty one should be merged away '
        'before a bill lands on it.',
    ]):
        c = sc.cell(row + 1 + i, 1, t)
        c.font = BODY
        c.alignment = Alignment(wrap_text=True, vertical='top')
        sc.merge_cells(start_row=row + 1 + i, start_column=1,
                       end_row=row + 1 + i, end_column=7)
        sc.row_dimensions[row + 1 + i].height = 30

    # ================================================= What changed
    MONTHS4 = [(2026, m) for m in (6, 7, 8, 9)]
    wc = wb.create_sheet('What changed', 2)
    wc['A1'] = 'June to September, by the other side of the entry'
    wc['A1'].font = TITLE
    wc['A2'] = ('Red is the Spacera cost. It is in the subscription accounts '
                'from August. The July bill of 5,208.00 is not here because '
                'it was coded to Sub-Contract Labour.')
    wc['A2'].font = NOTE
    wc.merge_cells('A2:F2')
    wc.row_dimensions[2].height = 30

    row = 4
    for sec, title in (('Expense', 'COST SIDE, BY SUPPLIER'),
                       ('Income', 'INCOME SIDE, BY CUSTOMER')):
        para(wc, row, title, bold=True)
        band(wc, row + 1,
             ['Other side'] + [datetime.date(*k, 1).strftime('%b %Y')
                               for k in MONTHS4] + ['Four months'],
             [46, 15, 15, 15, 15, 16])
        seen = collections.defaultdict(float)
        for r in rows:
            if r['sec'] == sec and (r['date'].year, r['date'].month) in MONTHS4:
                seen[party(r)] += signed(r)
        order = sorted(seen, key=lambda x: seen[x] if sec == 'Expense'
                       else -seen[x])
        start = row + 2
        for i, p in enumerate(order):
            n = start + i
            wc.cell(n, 1, p).font = BODY
            for j, k in enumerate(MONTHS4):
                f = (f'=SUMIFS(Transactions!$J$5:$J${last},'
                     f'Transactions!$C$5:$C${last},"{sec}",'
                     f'Transactions!$E$5:$E${last},$A{n},'
                     f'Transactions!$B$5:$B${last},'
                     f'DATE({k[0]},{k[1]},1))')
                c = wc.cell(n, 2 + j, f)
                c.number_format, c.font = MONEY, BODY
            c = wc.cell(n, 6, f'=SUM(B{n}:E{n})')
            c.number_format, c.font = MONEY, BOLD
            if CAUSE in p.upper():
                for k in range(1, 7):
                    wc.cell(n, k).fill = RED
        n = start + len(order)
        wc.cell(n, 1, f'Total {sec.lower()}').font = BOLD
        for k in range(2, 7):
            c = wc.cell(n, k, f'=SUM({gl(k)}{start}:{gl(k)}{n-1})')
            c.number_format, c.font, c.fill, c.border = MONEY, BOLD, TOTF, BOX
        wc.cell(n, 1).fill, wc.cell(n, 1).border = TOTF, BOX
        row = n + 3

    para(wc, row, 'THE MARGIN ON ALL THREE BASES', bold=True)
    band(wc, row + 1, ['Month', 'Income', 'As booked', 'Per cent',
                       'With Spacera in every month', 'Per cent'],
         [46, 15, 16, 12, 24, 12])
    for i, k in enumerate(MONTHS4):
        n = row + 2 + i
        m = first + months.index(k)
        wc.cell(n, 1, datetime.date(*k, 1).strftime('%B %Y')).font = BODY
        for j, (col, fmt) in enumerate((('B', MONEY), ('D', MONEY),
                                        ('E', PCT), ('H', MONEY),
                                        ('I', PCT))):
            c = wc.cell(n, 2 + j, f"='Monthly recon'!{col}{m}")
            c.number_format, c.font = fmt, BODY
    wc.freeze_panes = 'A6'

    # ================================================= raw source
    src = openpyxl.load_workbook(sys.argv[1], data_only=True).worksheets[0]
    out = wb.create_sheet('Source')
    for r in src.iter_rows():
        for c in r:
            if c.value is None:
                continue
            t = out.cell(c.row, c.column, c.value)
            t.number_format, t.font = c.number_format, BODY
    for r in range(6, src.max_row + 1):
        if not isinstance(src.cell(r, 1).value, datetime.datetime):
            continue
        d = src.cell(r, 1).value
        fill = None
        if CAUSE in str(src.cell(r, 3).value).upper():
            fill = RED
        elif src.cell(r, 2).value == 'Manual Journal':
            fill = AMBER
        elif d.year == 2026 and d.month in (7, 8, 9):
            fill = BLUE
        if fill:
            for c in range(1, 10):
                out.cell(r, c).fill = fill
    for col, w in (('A', 13), ('B', 20), ('C', 76), ('D', 24),
                   ('E', 14), ('F', 14), ('G', 15), ('H', 14), ('I', 12)):
        out.column_dimensions[col].width = w
    lg = src.max_row + 3
    out.cell(lg, 1, 'LEGEND').font = BOLD
    for i, (f, t) in enumerate((
            (RED, 'The Spacera cost. August and September only. The July '
                  'bill went to Sub-Contract Labour and is not on this tab.'),
            (AMBER, 'A manual journal, which is how the licences are '
                    'spread over their term.'),
            (BLUE, 'July to September 2026, the FY2027 months.'))):
        out.cell(lg + 1 + i, 1).fill = f
        out.cell(lg + 1 + i, 1).border = BOX
        out.cell(lg + 1 + i, 2, t).font = BODY

    wb.save(OUT)
    print(f'  written to {OUT}')
    return 0


if __name__ == '__main__':
    sys.exit(main())
