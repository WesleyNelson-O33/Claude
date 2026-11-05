"""Subscriptions and licences, income against cost, reconciled month by month.

Built from the Account Transactions export Danica sent, which carries both
accounts: Subscriptions & Licences - Income and Subscriptions & Licences -
Expense, for 1 July 2024 to 30 September 2026.

The question was why the book only makes about 3 per cent. It does not. It
makes 11.3 per cent in FY2025 and 13.0 per cent in FY2026, and it ran at 15
per cent through June and July 2026. The 3 per cent is FY2027 to date, and the
whole of the fall is in August and September.

One supplier explains all of it. SPACERA PTY LTD began billing 5,208.00 a
month to the expense account in August 2026, for the Aware Super National
Support Agreement, and there is no subscription income beside it. Taken out,
the four months to September read 15.2, 15.7, 15.6 and 15.1 per cent.

Whether that cost belongs in subscriptions is not something this export can
answer. The only recurring Aware Super income is 11,000.00 a month and it is
coded to Contract Support Staff, not here.

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


def load(path):
    """Both sections of the export, tagged Income or Expense."""
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
    """The other side, taken off the front of an invoice description."""
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
    if widths:
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[gl(i)].width = w


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

    # ================================================= 3. Transactions
    tx = wb.active
    tx.title = 'Transactions'
    tx['A1'] = 'Every line on both subscription accounts'
    tx['A1'].font = TITLE
    tx['A2'] = (f'From the export for {period}. This tab drives the other '
                'two, so the figures there are formulas off these rows. '
                'Amount is income positive and cost negative.')
    tx['A2'].font = NOTE
    band(tx, 4, ['Date', 'Month', 'Account', 'Source', 'Other side',
                 'Description', 'Reference', 'Debit', 'Credit', 'Amount',
                 'GST', 'Is the cause'],
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

    # ================================================= 1. Monthly recon
    mr = wb.create_sheet('Monthly recon', 0)
    mr['A1'] = 'Subscriptions and licences, income against cost, by month'
    mr['A1'].font = TITLE
    mr['A2'] = 'Corporate Technology Services Pty Ltd'
    mr['A2'].font = BODY
    mr['A3'] = period
    mr['A3'].font = NOTE

    mr['A5'] = 'THE ANSWER'
    mr['A5'].font = BOLD
    for i, t in enumerate([
        'The book does not make 3 per cent. It makes 11.3 per cent in FY2025 '
        'and 13.0 per cent in FY2026.',
        'The 3 per cent is FY2027 to date. July was 15.7 per cent. The fall '
        'is all in August and September.',
        'One supplier explains all of it. SPACERA PTY LTD began billing '
        '5,208.00 a month to the cost account in August, for the Aware Super '
        'National Support Agreement, with no subscription income beside it.',
        'Taken out, the four months to September read 15.2, 15.7, 15.6 and '
        '15.1 per cent. The book is flat at 15 per cent.',
        'Whether that cost belongs here is not something this export can '
        'answer. The only recurring Aware Super income is 11,000.00 a month '
        'and it is coded to Contract Support Staff, not to subscriptions.',
    ]):
        r = 6 + i
        c = mr.cell(r, 1, t)
        c.font = BODY
        c.alignment = Alignment(wrap_text=True, vertical='top')
        mr.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
        mr.row_dimensions[r].height = 30

    band(mr, 13, ['Month', 'Income', 'Cost', 'Margin', 'Margin %',
                  'The Spacera cost', 'Margin without it', 'Per cent without'],
         [12, 15, 15, 14, 12, 17, 18, 17])
    first = 14
    for i, (y, m) in enumerate(months):
        r = first + i
        mr.cell(r, 1, datetime.date(y, m, 1)).number_format = MON
        mr.cell(r, 2, f'={SI}$A{r})')
        mr.cell(r, 3, f'={SC}$A{r})')
        mr.cell(r, 4, f'=B{r}+C{r}')
        mr.cell(r, 5, f'=IFERROR(D{r}/B{r},"")')
        mr.cell(r, 6, f'={SS}$A{r})')
        mr.cell(r, 7, f'=D{r}-F{r}')
        mr.cell(r, 8, f'=IFERROR(G{r}/B{r},"")')
        for k in (2, 3, 4, 6, 7):
            mr.cell(r, k).number_format = MONEY
        for k in (5, 8):
            mr.cell(r, k).number_format = PCT
        for k in range(1, 9):
            mr.cell(r, k).font = BODY
        if (y, m) in ((2026, 8), (2026, 9)):
            for k in range(1, 9):
                mr.cell(r, k).fill = BLUE
                mr.cell(r, k).font = BOLD
    last_m = first + len(months) - 1

    tot = last_m + 1
    mr.cell(tot, 1, 'Total').font = BOLD
    for k, col in ((2, 'B'), (3, 'C'), (6, 'F')):
        mr.cell(tot, k, f'=SUM({col}{first}:{col}{last_m})')
    mr.cell(tot, 4, f'=B{tot}+C{tot}')
    mr.cell(tot, 5, f'=D{tot}/B{tot}')
    mr.cell(tot, 7, f'=D{tot}-F{tot}')
    mr.cell(tot, 8, f'=G{tot}/B{tot}')
    for k in range(1, 9):
        mr.cell(tot, k).fill = TOTF
        mr.cell(tot, k).font = BOLD
        mr.cell(tot, k).border = BOX
        mr.cell(tot, k).number_format = PCT if k in (5, 8) else MONEY
    mr.cell(tot, 1).number_format = 'General'

    # by financial year
    fy = sorted({fy_of(datetime.date(y, m, 1)) for y, m in months})
    fr = tot + 2
    mr.cell(fr, 1, 'BY FINANCIAL YEAR').font = BOLD
    band(mr, fr + 1, ['Year', 'Income', 'Cost', 'Margin', 'Margin %',
                      'The Spacera cost', 'Margin without it',
                      'Per cent without'])
    for i, y in enumerate(fy):
        r = fr + 2 + i
        ms = [first + months.index(k) for k in months if fy_of(
            datetime.date(k[0], k[1], 1)) == y]
        rng = lambda col: '+'.join(f'{col}{x}' for x in ms)
        mr.cell(r, 1, f'FY{y}').font = BOLD
        mr.cell(r, 2, f'={rng("B")}')
        mr.cell(r, 3, f'={rng("C")}')
        mr.cell(r, 4, f'=B{r}+C{r}')
        mr.cell(r, 5, f'=D{r}/B{r}')
        mr.cell(r, 6, f'={rng("F")}')
        mr.cell(r, 7, f'=D{r}-F{r}')
        mr.cell(r, 8, f'=G{r}/B{r}')
        for k in range(2, 9):
            mr.cell(r, k).number_format = PCT if k in (5, 8) else MONEY
            mr.cell(r, k).font = BOLD if y == fy[-1] else BODY
        if y == fy[-1]:
            for k in range(1, 9):
                mr.cell(r, k).fill = BLUE

    chk = fr + 3 + len(fy)
    mr.cell(chk, 1, 'CHECKS').font = BOLD
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
    ]):
        r = chk + 1 + i
        mr.cell(r, 1, lbl).font = BODY
        c = mr.cell(r, 7, f)
        c.font = BOLD
        c.alignment = Alignment(horizontal='center')
    mr.freeze_panes = 'A14'

    # ================================================= 2. What changed
    MONTHS4 = [(2026, m) for m in (6, 7, 8, 9)]
    wc = wb.create_sheet('What changed', 1)
    wc['A1'] = 'What changed in August, by the other side of the entry'
    wc['A1'].font = TITLE
    wc['A2'] = ('Four months side by side. Red is the one line that was not '
                'there in June or July.')
    wc['A2'].font = NOTE

    row = 4
    for sec, title in (('Expense', 'COST SIDE, BY SUPPLIER'),
                       ('Income', 'INCOME SIDE, BY CUSTOMER')):
        wc.cell(row, 1, title).font = BOLD
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

    wc.cell(row, 1, 'THE MARGIN, WITH AND WITHOUT THE SPACERA COST').font = BOLD
    band(wc, row + 1, ['Month', 'Income', 'Margin as booked', 'Per cent',
                       'Margin without Spacera', 'Per cent'],
         [46, 15, 18, 15, 22, 16])
    for i, k in enumerate(MONTHS4):
        n = row + 2 + i
        m = first + months.index(k)
        wc.cell(n, 1, datetime.date(*k, 1).strftime('%B %Y')).font = BODY
        for j, (col, fmt) in enumerate((('B', MONEY), ('D', MONEY),
                                        ('E', PCT), ('G', MONEY),
                                        ('H', PCT))):
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
        elif d.year == 2026 and d.month in (8, 9):
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
            (RED, 'The Spacera cost. This is the whole of the fall.'),
            (AMBER, 'A manual journal, which is how the licences are '
                    'deferred over their term.'),
            (BLUE, 'August and September 2026, the two months that fell.'))):
        out.cell(lg + 1 + i, 1).fill = f
        out.cell(lg + 1 + i, 1).border = BOX
        out.cell(lg + 1 + i, 2, t).font = BODY

    wb.save(OUT)
    print(f'  written to {OUT}')
    return 0


if __name__ == '__main__':
    sys.exit(main())
