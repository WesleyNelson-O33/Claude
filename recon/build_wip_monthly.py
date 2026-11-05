"""Account 11300 Work in Progress, reconciled month by month.

Built from the two Account Transactions exports Danica sent. The long one runs
from 1 July 2024, the short one from 1 July 2025, and they agree line for line
over the overlap, so the long one is used and the short one is kept as a check.

The schedule runs opening, debits, credits, movement, closing for every month.
It ties from a credit of 141,095.27 at 1 July 2024 to a credit of 133,253.18 at
30 September 2026, which is what the balance sheet says, and on to a credit of
125,405.18 at 31 October 2026.

The second tab nets every item in the account down to what is still carrying a
balance, because the monthly WIP journals are never reversed, so the closing
balance is the sum of every accrual ever posted.

Nothing here proposes a correction. It reports the account as it stands.

Run with the long export then the short export as arguments.
"""
import datetime
import re
import sys
import collections

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as gl

OUT = 'recon/CTS work in progress by month - 30 September 2026.xlsx'

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
DATE = 'dd/mm/yyyy'
MON = 'mmm yyyy'

CUT = datetime.date(2026, 9, 30)
LIC = ('licence', 'license', 'proxyclick', 'eptura', 'xio', 'subscription',
       'desk booking', 'appspace', 'visitor mg')


def load(path):
    """The WIP side of every line, plus the opening and closing the export
    states, so the schedule can be checked against the export itself."""
    ws = openpyxl.load_workbook(path, data_only=True).worksheets[0]
    period, opening, closing, out = ws.cell(3, 1).value, None, None, []
    for r in range(6, ws.max_row + 1):
        a = ws.cell(r, 1).value
        d, c = ws.cell(r, 5).value or 0, ws.cell(r, 6).value or 0
        if isinstance(a, str):
            if a.strip() == 'Opening Balance':
                opening = d - c
            elif a.strip() == 'Closing Balance':
                closing = d - c
        elif isinstance(a, datetime.datetime):
            out.append(dict(date=a.date(), source=ws.cell(r, 2).value,
                            desc=ws.cell(r, 3).value, ref=ws.cell(r, 4).value,
                            debit=d, credit=c, row=r))
    return period, opening, closing, out


def item_of(desc):
    """The same accrual and its reversal share one item, so they net off."""
    return re.sub(r'^Reversal:\s*', '', str(desc or '')).strip()


def cat_of(item):
    low = item.lower()
    if any(w in low for w in LIC):
        return 'Licences and subscriptions'
    if low.startswith(('wip journal', 'wipjournal')):
        return 'Monthly WIP journal'
    if low.startswith('wip reclass'):
        return 'WIP reclass'
    return 'Other accruals'


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
    if len(sys.argv) != 3:
        print(__doc__)
        return 1
    plong, pshort = sys.argv[1], sys.argv[2]
    perL, opL, clL, rowsL = load(plong)
    perS, opS, clS, rowsS = load(pshort)

    # ---- the two exports have to agree before either is used
    key = lambda r: (r['date'], r['source'], r['desc'], r['ref'],
                     r['debit'], r['credit'])
    start_short = min(r['date'] for r in rowsS)
    pre = sum(r['debit'] - r['credit'] for r in rowsL if r['date'] < start_short)
    agree_open = abs(opL + pre - opS) < 0.005
    overlap = sorted(key(r) for r in rowsL if r['date'] >= start_short)
    agree_lines = overlap == sorted(key(r) for r in rowsS)
    agree_close = abs(clL - clS) < 0.005
    print(f'  long export  {perL}')
    print(f'  short export {perS}')
    print(f'  openings agree at the handover: {agree_open}')
    print(f'  {len(overlap)} overlapping lines identical: {agree_lines}')
    print(f'  closing balances agree: {agree_close}')
    if not (agree_open and agree_lines and agree_close):
        print('  the two exports disagree, stopping')
        return 1

    rows = sorted(rowsL, key=lambda r: (r['date'], r['row']))
    months = sorted({(r['date'].year, r['date'].month) for r in rows})

    wb = openpyxl.Workbook()

    # ================================================= 3. Transactions
    tx = wb.active
    tx.title = 'Transactions'
    tx['A1'] = 'Every line on account 11300'
    tx['A1'].font = TITLE
    tx['A2'] = (f'From the export for {perL}. This tab drives the other two, '
                'so the figures there are formulas off these rows.')
    tx['A2'].font = NOTE
    band(tx, 4, ['Date', 'Month', 'Source', 'Description', 'Item',
                 'Category', 'Reference', 'Debit', 'Credit', 'Movement'],
         [12, 11, 16, 62, 52, 26, 12, 14, 14, 14])
    for i, r in enumerate(rows):
        n = 5 + i
        it = item_of(r['desc'])
        tx.cell(n, 1, r['date']).number_format = DATE
        tx.cell(n, 2, datetime.date(r['date'].year, r['date'].month, 1)
                ).number_format = MON
        tx.cell(n, 3, r['source'])
        tx.cell(n, 4, r['desc'])
        tx.cell(n, 5, it)
        tx.cell(n, 6, cat_of(it))
        tx.cell(n, 7, r['ref'])
        tx.cell(n, 8, r['debit']).number_format = MONEY
        tx.cell(n, 9, r['credit']).number_format = MONEY
        c = tx.cell(n, 10, f'=H{n}-I{n}')
        c.number_format = MONEY
        for k in range(1, 11):
            tx.cell(n, k).font = BODY
    last_tx = 4 + len(rows)
    tx.freeze_panes = 'A5'
    tx.auto_filter.ref = f'A4:J{last_tx}'

    # ================================================= 1. Monthly recon
    mr = wb.create_sheet('Monthly recon', 0)
    mr['A1'] = 'Work in progress, account 11300, reconciled by month'
    mr['A1'].font = TITLE
    mr['A2'] = 'Corporate Technology Services Pty Ltd'
    mr['A2'].font = BODY
    mr['A3'] = f'{perL}. Both exports agree, so the long one is used.'
    mr['A3'].font = NOTE

    mr['A5'] = 'WHERE IT LANDS'
    mr['A5'].font = BOLD
    checks = [
        ('Balance at 1 July 2024, per the export', opL),
        ('Balance at 30 September 2026', None),
        ('Balance at 31 October 2026, per the export', clL),
        ('Months in the schedule', len(months)),
        ('Lines in the schedule', len(rows)),
    ]
    for i, (lbl, val) in enumerate(checks):
        r = 6 + i
        mr.cell(r, 1, lbl).font = BODY
        if val is not None:
            c = mr.cell(r, 4, val)
            c.font = BOLD
            c.number_format = MONEY if isinstance(val, float) else '#,##0'

    band(mr, 14, ['Month', 'Opening', 'Debits', 'Credits', 'Movement',
                  'Closing', 'Lines'], [12, 16, 16, 16, 16, 16, 9])
    first = 15
    for i, (y, m) in enumerate(months):
        r = first + i
        mr.cell(r, 1, datetime.date(y, m, 1)).number_format = MON
        if i == 0:
            mr.cell(r, 2, opL)
        else:
            mr.cell(r, 2, f'=F{r-1}')
        mr.cell(r, 3, f'=SUMIFS(Transactions!$H$5:$H${last_tx},'
                      f'Transactions!$B$5:$B${last_tx},$A{r})')
        mr.cell(r, 4, f'=SUMIFS(Transactions!$I$5:$I${last_tx},'
                      f'Transactions!$B$5:$B${last_tx},$A{r})')
        mr.cell(r, 5, f'=C{r}-D{r}')
        mr.cell(r, 6, f'=B{r}+E{r}')
        mr.cell(r, 7, f'=COUNTIFS(Transactions!$B$5:$B${last_tx},$A{r})')
        for k in range(2, 7):
            mr.cell(r, k).number_format = MONEY
            mr.cell(r, k).font = BODY
        mr.cell(r, 1).font = BODY
        mr.cell(r, 7).font = BODY
        mr.cell(r, 7).number_format = '#,##0'
        if (y, m) == (2026, 9):
            for k in range(1, 8):
                mr.cell(r, k).fill = BLUE
                mr.cell(r, k).font = BOLD
    last_m = first + len(months) - 1

    tot = last_m + 1
    mr.cell(tot, 1, 'Total').font = BOLD
    mr.cell(tot, 2, f'=B{first}')
    mr.cell(tot, 3, f'=SUM(C{first}:C{last_m})')
    mr.cell(tot, 4, f'=SUM(D{first}:D{last_m})')
    mr.cell(tot, 5, f'=C{tot}-D{tot}')
    mr.cell(tot, 6, f'=B{tot}+E{tot}')
    mr.cell(tot, 7, f'=SUM(G{first}:G{last_m})')
    for k in range(1, 8):
        mr.cell(tot, k).fill = TOTF
        mr.cell(tot, k).font = BOLD
        mr.cell(tot, k).border = BOX
        if 2 <= k <= 6:
            mr.cell(tot, k).number_format = MONEY
    mr.cell(tot, 7).number_format = '#,##0'

    sep_row = first + months.index((2026, 9))
    mr['D7'] = f'=F{sep_row}'
    mr['D7'].font, mr['D7'].number_format = BOLD, MONEY

    chk = tot + 2
    mr.cell(chk, 1, 'CHECKS').font = BOLD
    for i, (lbl, f) in enumerate([
        ('Closing on the schedule equals the closing the export states',
         f'=IF(ROUND(F{tot}-{clL},2)=0,"OK","CHECK")'),
        ('Every line in the export is on the schedule',
         f'=IF(G{tot}={len(rows)},"OK","CHECK")'),
        ('The two exports agree line for line over the overlap',
         f'="OK, {len(overlap)} lines"'),
        ('30 September agrees with the balance sheet, a credit of 133,253.18',
         f'=IF(ROUND(F{sep_row}+133253.18,2)=0,"OK","CHECK")'),
    ]):
        r = chk + 1 + i
        mr.cell(r, 1, lbl).font = BODY
        c = mr.cell(r, 6, f)
        c.font = BOLD
        c.alignment = Alignment(horizontal='center')

    note = chk + 6
    for i, t in enumerate([
        'WHAT THE SCHEDULE SHOWS',
        'The account is an asset. It is a credit in 26 of the 28 months, '
        'which is the wrong way round.',
        'The monthly WIP journals are not reversed. 343 of them were posted to '
        '30 September 2026 and none has a reversal, so the closing balance is '
        'the sum of every accrual ever put through.',
        'Accounts 44100 Opening Work in Progress and 44200 Closing Work in '
        'Progress have no entries at all in the June to September general '
        'ledger, so the balance sheet moves without the profit and loss WIP '
        'accounts being used.',
        'The second tab nets the account down to what is still carrying a '
        'balance, and splits it by what the item is.',
    ]):
        c = mr.cell(note + i, 1, t)
        c.font = BOLD if i == 0 else BODY
        c.alignment = Alignment(wrap_text=True, vertical='top')
        if i:
            mr.merge_cells(start_row=note + i, start_column=1,
                           end_row=note + i, end_column=7)
            mr.row_dimensions[note + i].height = 32
    mr.freeze_panes = 'A15'

    # ================================================= 2. What is in it
    item = collections.defaultdict(
        lambda: {'net': 0.0, 'first': None, 'n': 0})
    for r in rows:
        if r['date'] > CUT:
            continue
        it = item[item_of(r['desc'])]
        it['net'] += r['debit'] - r['credit']
        it['n'] += 1
        it['first'] = (r['date'] if it['first'] is None
                       else min(it['first'], r['date']))
    live = {k: v for k, v in item.items() if abs(v['net']) > 0.005}

    wi = wb.create_sheet('What is in it', 1)
    wi['A1'] = 'What is sitting in the 133,253.18 credit at 30 September 2026'
    wi['A1'].font = TITLE
    wi['A2'] = ('Every item netted against its own reversal. An item that '
                'fully reversed is left out. The opening balance at 1 July '
                '2024 is shown on its own, because the export does not say '
                'what is in it.')
    wi['A2'].font = NOTE
    wi.merge_cells('A2:F2')
    wi.row_dimensions[2].height = 30

    wi['A4'] = 'BY WHAT THE ITEM IS'
    wi['A4'].font = BOLD
    band(wi, 5, ['Category', 'Items', 'Balance', '', '', ''],
         [40, 10, 18, 14, 16, 14])
    cats = ['Licences and subscriptions', 'Monthly WIP journal',
            'WIP reclass', 'Other accruals']
    for i, c in enumerate(cats):
        r = 6 + i
        wi.cell(r, 1, c).font = BODY
        wi.cell(r, 2, f'=COUNTIFS($B$20:$B$9999,$A{r})').font = BODY
        x = wi.cell(r, 3, f'=SUMIFS($E$20:$E$9999,$B$20:$B$9999,$A{r})')
        x.font, x.number_format = BODY, MONEY
        if c == cats[0]:
            for k in (1, 2, 3):
                wi.cell(r, k).fill = RED
    r = 6 + len(cats)
    wi.cell(r, 1, 'Items still carrying a balance').font = BOLD
    wi.cell(r, 2, f'=SUM(B6:B{r-1})').font = BOLD
    wi.cell(r, 3, f'=SUM(C6:C{r-1})').font = BOLD
    wi.cell(r, 3).number_format = MONEY
    wi.cell(r + 1, 1, 'Opening balance at 1 July 2024').font = BODY
    y = wi.cell(r + 1, 3, opL)
    y.font, y.number_format = BODY, MONEY
    wi.cell(r + 2, 1, 'Balance at 30 September 2026').font = BOLD
    z = wi.cell(r + 2, 3, f'=C{r}+C{r+1}')
    z.font, z.number_format = BOLD, MONEY
    for k in (1, 2, 3):
        wi.cell(r + 2, k).fill = TOTF
        wi.cell(r + 2, k).border = BOX
    wi.cell(r + 3, 6, f'=IF(ROUND(C{r+2}+133253.18,2)=0,"OK","CHECK")'
            ).font = BOLD
    wi.cell(r + 3, 1, 'Agrees with the balance sheet').font = BODY

    wi['A17'] = 'EVERY ITEM STILL CARRYING A BALANCE'
    wi['A17'].font = BOLD
    wi['A18'] = ('Sorted biggest credit first. Red is a licence or a '
                 'subscription, which is not work in progress.')
    wi['A18'].font = NOTE
    band(wi, 19, ['Item', 'Category', 'First posted', 'Financial year',
                  'Balance', 'Lines'], [78, 28, 14, 14, 16, 9])
    for i, (k, v) in enumerate(sorted(live.items(), key=lambda x: x[1]['net'])):
        n = 20 + i
        wi.cell(n, 1, k)
        wi.cell(n, 2, cat_of(k))
        wi.cell(n, 3, v['first']).number_format = DATE
        wi.cell(n, 4, f'FY{fy_of(v["first"])}')
        wi.cell(n, 5, v['net']).number_format = MONEY
        wi.cell(n, 6, v['n'])
        for c in range(1, 7):
            wi.cell(n, c).font = BODY
        if cat_of(k) == 'Licences and subscriptions':
            for c in range(1, 7):
                wi.cell(n, c).fill = RED
        elif k.lower().startswith(('wip journal', 'wipjournal')):
            for c in range(1, 7):
                wi.cell(n, c).fill = AMBER
    wi.freeze_panes = 'A20'
    wi.auto_filter.ref = f'A19:F{19 + len(live)}'

    # ================================================= raw sources
    for path, name, per in ((plong, 'Source - from Jul 2024', perL),
                            (pshort, 'Source - from Jul 2025', perS)):
        src = openpyxl.load_workbook(path, data_only=True).worksheets[0]
        out = wb.create_sheet(name)
        for row in src.iter_rows():
            for c in row:
                if c.value is None:
                    continue
                t = out.cell(c.row, c.column, c.value)
                t.number_format = c.number_format
                t.font = BODY
        for r in range(6, src.max_row + 1):
            d = src.cell(r, 1).value
            if not isinstance(d, datetime.datetime):
                continue
            it = item_of(src.cell(r, 3).value)
            fill = None
            if cat_of(it) == 'Licences and subscriptions':
                fill = RED
            elif it.lower().startswith(('wip journal', 'wipjournal')):
                fill = AMBER
            if d.year == 2026 and d.month == 9:
                fill = BLUE
            if fill:
                for c in range(1, 10):
                    out.cell(r, c).fill = fill
        for col, w in (('A', 13), ('B', 17), ('C', 70), ('D', 12),
                       ('E', 14), ('F', 14), ('G', 15), ('H', 14), ('I', 12)):
            out.column_dimensions[col].width = w
        lg = src.max_row + 3
        out.cell(lg, 1, 'LEGEND').font = BOLD
        for i, (f, t) in enumerate((
                (RED, 'A licence or a subscription. Not work in progress.'),
                (AMBER, 'A monthly WIP journal. None of these is reversed.'),
                (BLUE, 'September 2026, the month being reconciled.'))):
            out.cell(lg + 1 + i, 1).fill = f
            out.cell(lg + 1 + i, 1).border = BOX
            out.cell(lg + 1 + i, 2, t).font = BODY

    wb.save(OUT)
    print(f'  {len(rows)} lines, {len(months)} months, '
          f'{len(live)} items still carrying a balance')
    print(f'  written to {OUT}')
    return 0


if __name__ == '__main__':
    sys.exit(main())
