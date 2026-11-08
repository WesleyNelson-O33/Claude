"""Licence margins by month, with the term dates.

Reads a full Account Transactions export and reports the subscriptions and
licences book: account 41175 income against account 51500 cost.

A licence is sold as a 12 month term on one invoice, then released to the
P&L a twelfth at a time by a monthly manual journal. Income and cost are
released in the same journal and carry the same reference. So the margin on
a licence is read off one of those monthly journals, and the term dates come
out of the journal description.

Usage: build_licence_margins.py <Account Transactions.xlsx> <out.xlsx>
"""
import sys, re, datetime
from collections import defaultdict
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

INC, EXP = '41175', '51500'

MONTHS = {m.lower(): i for i, m in enumerate(
    ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
     'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'], 1)}
LONG = {m.lower(): i for i, m in enumerate(
    ['January', 'February', 'March', 'April', 'May', 'June', 'July',
     'August', 'September', 'October', 'November', 'December'], 1)}
MONTHS.update(LONG)
MONTHS['sept'] = 9

MONTH_RE = re.compile(
    r'\b(' + '|'.join(sorted(MONTHS, key=len, reverse=True)) + r')'
    r'\s*\.?\s*((?:19|20)\d{2}|\d{2})?\b', re.I)


def read(path):
    """Every 41175 and 51500 line in the export."""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb.worksheets[0]
    out = []
    for r in ws.iter_rows(min_row=6, values_only=True):
        code = str(r[0]) if r[0] is not None else ''
        if code not in (INC, EXP):
            continue
        if not isinstance(r[3], datetime.datetime):
            continue
        out.append(dict(
            code=code, account=r[1], source=r[2], date=r[3].date(),
            contact=r[4], debit=float(r[5] or 0), credit=float(r[6] or 0),
            job=r[7], inv=r[8], ref=r[9], desc=r[10], cc=r[11]))
    wb.close()
    return out


def signed(x):
    """Income positive, cost negative, so the two add up to the margin."""
    if x['code'] == INC:
        return x['credit'] - x['debit']
    return -(x['debit'] - x['credit'])


def norm(d):
    """The description with the Revenue / Expense tag and the padding gone."""
    d = re.sub(r'\s+', ' ', str(d or '')).strip()
    d = re.sub(r'\((Revenue|Expense|Income|Cost)\)', '', d, flags=re.I)
    d = d.replace('–', '-').replace('—', '-')
    return re.sub(r'\s+', ' ', d).strip(' -|')


def year(tok, near):
    """A two digit year read against the century of the dates around it."""
    if tok is None:
        return None
    n = int(tok)
    return n if n > 100 else 2000 + n


def term(desc):
    """The term start and end read out of the description.

    Handles Jul2025-Jun2026, Sep 2025 - Aug 2026, (November 2025 - October
    2026), Jul26-Jun27, Jan-Dec 2025 and (January - December 2026), where
    only one of the two months carries a year.
    """
    hits = [(m.group(1).lower(), m.group(2), m.start())
            for m in MONTH_RE.finditer(desc)]
    if len(hits) < 2:
        return None, None, 'no term dates in the description'
    for a, b in zip(hits, hits[1:]):
        gap = desc[a[2]:b[2]]
        if not re.search(r'[-]|\bto\b', gap, re.I):
            continue
        m1, y1 = MONTHS[a[0]], year(a[1], None)
        m2, y2 = MONTHS[b[0]], year(b[1], None)
        if y1 is None and y2 is None:
            continue
        if y1 is None:                       # Jan-Dec 2025
            y1 = y2 if m1 <= m2 else y2 - 1
        if y2 is None:                       # (November - October 2026)
            y2 = y1 if m2 >= m1 else y1 + 1
        start = datetime.date(y1, m1, 1)
        end = datetime.date(y2, m2, 1)
        note = ''
        if end < start:
            # Nov 2025 - Oct 2025 in the source is a typo for Oct 2026.
            end = datetime.date(y2 + 1, m2, 1)
            note = 'end year corrected, source reads an earlier year'
        months = (end.year - start.year) * 12 + end.month - start.month + 1
        stated = re.search(r'(\d{1,2})\s*month', desc, re.I)
        if stated and int(stated.group(1)) != months:
            note = (note + '; ' if note else '') + (
                f'description says {stated.group(1)} months but the dates '
                f'in it span {months}')
        return start, end, note or f'{months} month term'
    return None, None, ''


def month_end(d):
    nxt = datetime.date(d.year + (d.month == 12), d.month % 12 + 1, 1)
    return nxt - datetime.timedelta(days=1)


def customer(desc):
    """The customer as the description leads with it."""
    head = re.split(r'\s[-|]\s|\s\|\s', desc)[0]
    head = re.sub(r'\b(Year|Term)\s*\d.*$', '', head, flags=re.I)
    head = MONTH_RE.sub('', head)
    head = re.sub(r'\s+', ' ', head).strip(' -|(),')
    words = head.split()
    for n in (3, 2, 1):
        if len(words) >= n:
            return ' '.join(words[:n])
    return head


def build(rows):
    """Split the book into releases, deferrals and lines that never pair."""
    byref = defaultdict(lambda: defaultdict(list))
    loose = []
    for r in rows:
        if r['source'] == 'Manual Journal' and r['ref']:
            byref[str(r['ref'])][r['code']].append(r)
        else:
            loose.append(r)

    releases, deferrals = defaultdict(list), []
    for ref, v in byref.items():
        if not (v[INC] and v[EXP]):
            loose.extend(v[INC] + v[EXP])
            continue
        inc = sum(signed(x) for x in v[INC])
        cost = sum(signed(x) for x in v[EXP])
        d = norm(v[INC][0]['desc'])
        date = min(x['date'] for x in v[INC])
        if inc > 0:
            releases[d].append(dict(ref=ref, date=date, inc=inc, cost=cost))
        else:
            deferrals.append(dict(ref=ref, date=date, inc=inc, cost=cost,
                                  desc=d))
    return releases, deferrals, loose


# ---------------------------------------------------------------- formatting

HEAD = Font(bold=True, color='FFFFFF', size=10)
HFILL = PatternFill('solid', fgColor='1F3864')
SUB = Font(bold=True, size=10)
SFILL = PatternFill('solid', fgColor='D9E2F3')
TITLE = Font(bold=True, size=13)
RED = PatternFill('solid', fgColor='FFC7CE')
AMBER = PatternFill('solid', fgColor='FFEB9C')
GREEN = PatternFill('solid', fgColor='C6EFCE')
THIN = Border(*[Side(style='thin', color='BFBFBF')] * 4)
MONEY = '#,##0.00;[Red]-#,##0.00'
PCT = '0.0%;[Red]-0.0%'
DATE = 'dd/mm/yyyy'


def head_row(ws, r, labels, widths=None, freeze=True):
    for c, lab in enumerate(labels, 1):
        cell = ws.cell(r, c, lab)
        cell.font, cell.fill, cell.border = HEAD, HFILL, THIN
        cell.alignment = Alignment(wrap_text=True, vertical='top')
    if widths:
        for c, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(c)].width = w
    # Only freeze where the sheet is one long list. Setting freeze_panes and
    # then clearing it leaves a selection pointing at a pane that is no
    # longer there, and Excel reports the workbook as damaged.
    if freeze:
        ws.freeze_panes = ws.cell(r + 1, 1)


def put(ws, r, vals, fmts):
    for c, (v, f) in enumerate(zip(vals, fmts), 1):
        cell = ws.cell(r, c, v)
        cell.border = THIN
        if f:
            cell.number_format = f
        if isinstance(v, str):
            cell.alignment = Alignment(wrap_text=True, vertical='top')
    return r + 1


def main(src, out):
    rows = read(src)
    releases, deferrals, loose = build(rows)
    wb = openpyxl.Workbook()

    # ---------------------------------------------------- 0. read this first
    ws = wb.active
    ws.title = 'Read This First'
    ws.column_dimensions['A'].width = 46
    for col in 'BCDE':
        ws.column_dimensions[col].width = 15
    ws['A1'] = 'Licence margins, the short answer'
    ws['A1'].font = TITLE

    def line(r, text, bold=False, fill=None):
        c = ws.cell(r, 1, text)
        c.alignment = Alignment(wrap_text=True, vertical='top')
        if bold:
            c.font = SUB
        if fill:
            c.fill = fill
        return r + 1

    rel_i = rel_c = 0.0
    for v in releases.values():
        n = len(v)
        rel_i += round(sum(y['inc'] for y in v) / n, 2)
        rel_c += round(sum(y['cost'] for y in v) / n, 2)

    r = 3
    r = line(r, 'Every licence, priced the way it was sold, makes money. '
                'The average margin across the 40 licences is '
                f'{(rel_i + rel_c) / rel_i * 100:.1f} per cent.', True)
    r += 1
    r = line(r, 'The P&L shows a much thinner margin because licence cost '
                'is going through 51500 with nothing billed against it.')
    r += 1

    r = line(r, 'Margin per year, as booked', True)
    head_row(ws, r, ['Year', 'Income', 'Cost', 'Margin', 'Margin %'], freeze=False)
    r += 1
    years = (('FY2025', (2024, 7), (2025, 6)),
             ('FY2026', (2025, 7), (2026, 6)),
             ('FY2027 to 30 Sep 2026', (2026, 7), (2026, 9)))
    for label, lo, hi in years:
        sel = [x for x in rows
               if lo <= (x['date'].year, x['date'].month) <= hi]
        i = sum(signed(x) for x in sel if x['code'] == INC)
        c = sum(signed(x) for x in sel if x['code'] == EXP)
        r = put(ws, r, [label, round(i, 2), round(c, 2), round(i + c, 2),
                        (i + c) / i if i else None],
                [None, MONEY, MONEY, MONEY, PCT])
    r += 1

    sel = [x for x in rows if (2026, 7) <= (x['date'].year, x['date'].month)
           <= (2026, 9)]
    li = sum(signed(x) for x in loose
             if (2026, 7) <= (x['date'].year, x['date'].month) <= (2026, 9)
             and x['code'] == INC)
    lc = sum(signed(x) for x in loose
             if (2026, 7) <= (x['date'].year, x['date'].month) <= (2026, 9)
             and x['code'] == EXP)
    ti = sum(signed(x) for x in sel if x['code'] == INC)
    tc = sum(signed(x) for x in sel if x['code'] == EXP)

    r = line(r, 'Why FY2027 reads 3.1 per cent', True)
    head_row(ws, r, ['', 'Income', 'Cost', 'Margin', 'Margin %'], freeze=False)
    r += 1
    for label, i, c in (
            ('Licences released by journal', ti - li, tc - lc),
            ('Cost and income with no release journal', li, lc),
            ('Total, as the P&L shows it', ti, tc)):
        r = put(ws, r, [label, round(i, 2), round(c, 2), round(i + c, 2),
                        (i + c) / i if i else None],
                [None, MONEY, MONEY, MONEY, PCT])
        if label.startswith('Cost and income'):
            for col in range(1, 6):
                ws.cell(r - 1, col).fill = RED
    r += 1

    r = line(r, 'The unbilled cost in FY2027', True)
    head_row(ws, r, ['Supplier', 'Bills', 'Cost'], freeze=False)
    r += 1
    agg = {}
    for x in loose:
        if x['code'] != EXP:
            continue
        if not ((2026, 7) <= (x['date'].year, x['date'].month) <= (2026, 9)):
            continue
        a = agg.setdefault(x['contact'] or 'No supplier on the line',
                           [0, 0.0])
        a[0] += 1
        a[1] += signed(x)
    for k in sorted(agg, key=lambda y: agg[y][1]):
        r = put(ws, r, [str(k), agg[k][0], round(agg[k][1], 2)],
                [None, '0', MONEY])
    r += 1

    r = line(r, 'What to look at', True)
    for txt in ('Licence Margins: every licence, its term dates and its '
                'margin per month. Worst first.',
                'Licence by Month: the exact month each release was '
                'posted, so you can see a gap.',
                'Unpaired Cost and Income: the bills with no release '
                'journal. This is where the loss is.',
                'Deferrals to WIP: the journals that take a 12 month '
                'invoice out of the P&L.',
                'By Month and Raw: the monthly totals and every line '
                'behind them.'):
        r = line(r, txt)

    # ------------------------------------------------ 1. licence margins
    ws = wb.create_sheet('Licence Margins')
    ws['A1'] = 'Subscriptions and licences, margin per licence per month'
    ws['A1'].font = TITLE
    ws['A2'] = ('Account 41175 income against account 51500 cost. One row '
                'per licence. The monthly figures are what the monthly '
                'release journal puts through the P&L.')
    ws['A3'] = f'Source: {src.split("/")[-1]}. Worst margin first.'
    ws['A3'].font = Font(italic=True, size=9)

    cols = ['Licence (journal description)', 'Customer', 'Term start',
            'Term end', 'Months in term', 'Months released',
            'First release', 'Last release', 'Income per month',
            'Cost per month', 'Margin per month', 'Margin %',
            'Income released to date', 'Cost released to date',
            'Margin released to date', 'Months still to release',
            'Income still to come', 'Note']
    head_row(ws, 5, cols,
             [62, 20, 11, 11, 9, 9, 11, 11, 13, 13, 13, 9, 14, 14, 14, 9,
              14, 34])
    fmts = [None, None, DATE, DATE, '0', '0', DATE, DATE, MONEY, MONEY,
            MONEY, PCT, MONEY, MONEY, MONEY, '0', MONEY, None]

    table = []
    for desc, items in releases.items():
        items.sort(key=lambda x: x['date'])
        n = len(items)
        inc_t = sum(x['inc'] for x in items)
        cost_t = sum(x['cost'] for x in items)
        m_inc = round(inc_t / n, 2)
        m_cost = round(cost_t / n, 2)
        start, end, note = term(desc)
        tm = ((end.year - start.year) * 12 + end.month - start.month + 1
              if start and end else None)
        left = (tm - n) if tm else None
        table.append(dict(
            desc=desc, cust=customer(desc), start=start, end=end, tm=tm,
            n=n, first=items[0]['date'], last=items[-1]['date'],
            m_inc=m_inc, m_cost=m_cost, m_marg=round(m_inc + m_cost, 2),
            pct=(m_inc + m_cost) / m_inc if m_inc else None,
            inc_t=round(inc_t, 2), cost_t=round(cost_t, 2),
            marg_t=round(inc_t + cost_t, 2),
            left=left,
            to_come=round(m_inc * left, 2) if left and left > 0 else None,
            note=note))
    table.sort(key=lambda x: (x['pct'] if x['pct'] is not None else 99))

    r = 6
    for t in table:
        r2 = put(ws, r, [t['desc'], t['cust'], t['start'], t['end'], t['tm'],
                         t['n'], t['first'], t['last'], t['m_inc'],
                         t['m_cost'], t['m_marg'], t['pct'], t['inc_t'],
                         t['cost_t'], t['marg_t'], t['left'], t['to_come'],
                         t['note']], fmts)
        p = t['pct']
        if p is not None:
            fill = RED if p < 0.05 else AMBER if p < 0.12 else GREEN
            for c in (11, 12):
                ws.cell(r, c).fill = fill
        if t['tm'] and t['n'] > t['tm']:
            ws.cell(r, 6).fill = RED
        r = r2

    tot = r + 1
    ws.cell(tot, 1, 'Total, monthly releases').font = SUB
    for c, key in ((9, 'm_inc'), (10, 'm_cost'), (11, 'm_marg'),
                   (13, 'inc_t'), (14, 'cost_t'), (15, 'marg_t')):
        cell = ws.cell(tot, c, round(sum(t[key] for t in table), 2))
        cell.number_format, cell.font, cell.fill = MONEY, SUB, SFILL
    ti = sum(t['m_inc'] for t in table)
    cell = ws.cell(tot, 12, (ti + sum(t['m_cost'] for t in table)) / ti)
    cell.number_format, cell.font, cell.fill = PCT, SUB, SFILL
    for c in (1, 9, 10, 11, 13, 14, 15):
        ws.cell(tot, c).fill = SFILL

    leg = tot + 2
    ws.cell(leg, 1, 'Colour on the margin columns').font = SUB
    for i, (f, txt) in enumerate((
            (RED, 'RED: margin under 5 per cent. The licence barely '
                  'covers its own cost.'),
            (AMBER, 'AMBER: margin 5 to 12 per cent.'),
            (GREEN, 'GREEN: margin 12 per cent or better.'),
            (RED, 'RED on Months released: released more months than the '
                  'term runs for.')), 1):
        ws.cell(leg + i, 1, txt).fill = f

    # ------------------------------------------------------- 2. by month
    ws = wb.create_sheet('By Month')
    ws['A1'] = 'Subscriptions and licences by month, as booked'
    ws['A1'].font = TITLE
    ws['A2'] = ('Everything on 41175 and 41500 in the month, not only the '
                'release journals. This is what the P&L shows.')
    head_row(ws, 4, ['Month', 'Income 41175', 'Cost 51500', 'Margin',
                     'Margin %', 'Income lines', 'Cost lines'],
             [12, 15, 15, 15, 10, 12, 12])
    bym = defaultdict(lambda: dict(i=0.0, c=0.0, ni=0, nc=0))
    for x in rows:
        k = (x['date'].year, x['date'].month)
        s = signed(x)
        if x['code'] == INC:
            bym[k]['i'] += s
            bym[k]['ni'] += 1
        else:
            bym[k]['c'] += s
            bym[k]['nc'] += 1
    r = 5
    for k in sorted(bym):
        v = bym[k]
        m = v['i'] + v['c']
        r = put(ws, r, [datetime.date(k[0], k[1], 1), round(v['i'], 2),
                        round(v['c'], 2), round(m, 2),
                        (m / v['i'] if v['i'] else None), v['ni'], v['nc']],
                ['mmm yyyy', MONEY, MONEY, MONEY, PCT, '0', '0'])
    ti = sum(v['i'] for v in bym.values())
    tc = sum(v['c'] for v in bym.values())
    r += 1
    ws.cell(r, 1, 'Total').font = SUB
    for c, v, f in ((2, round(ti, 2), MONEY), (3, round(tc, 2), MONEY),
                    (4, round(ti + tc, 2), MONEY),
                    (5, (ti + tc) / ti, PCT)):
        cell = ws.cell(r, c, v)
        cell.number_format, cell.font, cell.fill = f, SUB, SFILL
    ws.cell(r, 1).fill = SFILL

    # --------------------------------------------- 3. licence by month grid
    ws = wb.create_sheet('Licence by Month')
    ws['A1'] = 'Margin per licence per month, release journals only'
    ws['A1'].font = TITLE
    ws['A2'] = ('The exact month each release was posted. Blank means no '
                'release journal that month.')
    months = sorted({x['date'].replace(day=1)
                     for v in releases.values() for x in v})
    head_row(ws, 4, ['Licence'] + [m.strftime('%b %y') for m in months]
             + ['Total margin'], [62] + [11] * len(months) + [14])
    r = 5
    for t in table:
        ws.cell(r, 1, t['desc']).border = THIN
        ws.cell(r, 1).alignment = Alignment(wrap_text=True, vertical='top')
        by = {x['date'].replace(day=1): x['inc'] + x['cost']
              for x in releases[t['desc']]}
        for c, m in enumerate(months, 2):
            cell = ws.cell(r, c, round(by[m], 2) if m in by else None)
            cell.number_format, cell.border = MONEY, THIN
        cell = ws.cell(r, len(months) + 2, t['marg_t'])
        cell.number_format, cell.font, cell.border = MONEY, SUB, THIN
        r += 1

    # ------------------------------------------------ 4. lines that never pair
    ws = wb.create_sheet('Unpaired Cost and Income')
    ws['A1'] = 'Lines with no matching other side'
    ws['A1'].font = TITLE
    ws['A2'] = ('Invoices, bills, and journals where only one of income or '
                'cost was posted. Unbilled licence cost sits here.')
    head_row(ws, 4, ['Side', 'Date', 'Source', 'Contact', 'Amount',
                     'Job', 'Invoice', 'Reference', 'Description'],
             [9, 11, 16, 26, 13, 10, 14, 20, 60])
    loose.sort(key=lambda x: (x['code'], x['date']))
    r = 5
    for x in loose:
        r = put(ws, r, ['Income' if x['code'] == INC else 'Cost', x['date'],
                        x['source'], x['contact'], round(signed(x), 2),
                        x['job'], x['inv'], x['ref'], norm(x['desc'])],
                [None, DATE, None, None, MONEY, None, None, None, None])
    r += 1
    ws.cell(r, 1, 'Total').font = SUB
    cell = ws.cell(r, 5, round(sum(signed(x) for x in loose), 2))
    cell.number_format, cell.font, cell.fill = MONEY, SUB, SFILL

    # --------------------------------------------------- 5. deferrals
    ws = wb.create_sheet('Deferrals to WIP')
    ws['A1'] = 'Journals that push a licence out of the P&L to WIP'
    ws['A1'].font = TITLE
    ws['A2'] = ('A 12 month invoice is taken out in full, then released a '
                'month at a time. Income negative, cost positive.')
    head_row(ws, 4, ['Date', 'Reference', 'Income taken out',
                     'Cost taken out', 'Net', 'Description'],
             [11, 20, 15, 15, 15, 70])
    deferrals.sort(key=lambda x: x['date'])
    r = 5
    for d in deferrals:
        r = put(ws, r, [d['date'], d['ref'], round(d['inc'], 2),
                        round(d['cost'], 2), round(d['inc'] + d['cost'], 2),
                        d['desc']],
                [DATE, None, MONEY, MONEY, MONEY, None])
    r += 1
    ws.cell(r, 1, 'Total').font = SUB
    for c, key in ((3, 'inc'), (4, 'cost')):
        cell = ws.cell(r, c, round(sum(d[key] for d in deferrals), 2))
        cell.number_format, cell.font, cell.fill = MONEY, SUB, SFILL

    # --------------------------------------------------------- 6. raw
    ws = wb.create_sheet('Raw 41175 and 51500')
    head_row(ws, 1, ['Account code', 'Account', 'Source', 'Date', 'Contact',
                     'Debit', 'Credit', 'Job', 'Invoice', 'Reference',
                     'Description', 'Cost centre'],
             [12, 26, 16, 11, 26, 12, 12, 10, 14, 20, 60, 20])
    r = 2
    for x in sorted(rows, key=lambda y: (y['date'], y['code'])):
        r = put(ws, r, [x['code'], x['account'], x['source'], x['date'],
                        x['contact'], x['debit'] or None, x['credit'] or None,
                        x['job'], x['inv'], x['ref'], norm(x['desc']),
                        x['cc']],
                [None, None, None, DATE, None, MONEY, MONEY, None, None,
                 None, None, None])

    wb.save(out)
    print(f'  licences            {len(table)}')
    print(f'  release journals    {sum(t["n"] for t in table)}')
    print(f'  deferral journals   {len(deferrals)}')
    print(f'  unpaired lines      {len(loose)}')
    print(f'  no term date parsed {sum(1 for t in table if not t["start"])}')
    print(f'  monthly margin      {sum(t["m_marg"] for t in table):,.2f}')
    print(f'  book margin         {round(ti + tc, 2):,.2f} on '
          f'{round(ti, 2):,.2f} income')
    print(f'  written             {out}')


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
