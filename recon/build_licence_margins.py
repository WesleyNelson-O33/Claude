"""Licence margins by month, with the term dates, built on live formulas.

Reads a full Account Transactions export and reports the subscriptions and
licences book: account 41175 income against account 51500 cost.

A licence is sold as a 12 month term on one invoice, then released to the
P&L a twelfth at a time by a monthly manual journal. Income and cost are
released in the same journal and carry the same reference. So the margin on
a licence is read off one of those monthly journals, and the term dates come
out of the journal description.

Every number on every tab is a formula over the Raw Data tab. Python only
does the part a formula cannot do: pairing each journal's two sides by
reference, and reading the term dates out of the description. Those land on
Raw Data as the Pairing and Licence ID columns, and everything else is
SUMIFS and COUNTIFS over them. Change a figure on Raw Data and the whole
workbook moves with it.

Usage: build_licence_margins.py <Account Transactions.xlsx> <out.xlsx>
"""
import sys, re, datetime
from collections import defaultdict
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.utils import get_column_letter

INC, EXP = 41175, 51500
RAW = "'Raw Data'"

MONTHS = {m.lower(): i for i, m in enumerate(
    ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
     'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'], 1)}
MONTHS.update({m.lower(): i for i, m in enumerate(
    ['January', 'February', 'March', 'April', 'May', 'June', 'July',
     'August', 'September', 'October', 'November', 'December'], 1)})
MONTHS['sept'] = 9
MONTH_RE = re.compile(
    r'\b(' + '|'.join(sorted(MONTHS, key=len, reverse=True)) +
    r')\s*\.?\s*((?:19|20)\d{2}|\d{2})?\b', re.I)


# ------------------------------------------------------------------ reading

def read(path):
    """Every 41175 and 51500 line in the export."""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb.worksheets[0]
    out = []
    for r in ws.iter_rows(min_row=6, values_only=True):
        try:
            code = int(r[0])
        except (TypeError, ValueError):
            continue
        if code not in (INC, EXP) or not isinstance(r[3], datetime.datetime):
            continue
        out.append(dict(
            code=code, account=r[1], source=r[2], date=r[3].date(),
            contact=r[4], debit=float(r[5] or 0), credit=float(r[6] or 0),
            job=r[7], inv=r[8], ref=r[9], desc=r[10], cc=r[11]))
    wb.close()
    return out


def signed(x):
    """Credit less debit. Income comes out positive, cost negative, so the
    two sides of a journal add up to the margin."""
    return x['credit'] - x['debit']


def norm(d):
    """The description with the Revenue / Expense tag and the padding gone."""
    d = re.sub(r'\s+', ' ', str(d or '')).strip()
    d = re.sub(r'\((Revenue|Expense|Income|Cost)\)', '', d, flags=re.I)
    d = d.replace('–', '-').replace('—', '-')
    return re.sub(r'\s+', ' ', d).strip(' -|')


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
        if not re.search(r'-|\bto\b', desc[a[2]:b[2]], re.I):
            continue
        m1, m2 = MONTHS[a[0]], MONTHS[b[0]]
        y1 = int(a[1]) if a[1] else None
        y2 = int(b[1]) if b[1] else None
        y1 = y1 + 2000 if y1 and y1 < 100 else y1
        y2 = y2 + 2000 if y2 and y2 < 100 else y2
        if y1 is None and y2 is None:
            continue
        if y1 is None:
            y1 = y2 if m1 <= m2 else y2 - 1
        if y2 is None:
            y2 = y1 if m2 >= m1 else y1 + 1
        start, end = datetime.date(y1, m1, 1), datetime.date(y2, m2, 1)
        note = ''
        if end < start:
            end = datetime.date(y2 + 1, m2, 1)
            note = 'end year corrected, source reads an earlier year'
        months = (end.year - start.year) * 12 + end.month - start.month + 1
        stated = re.search(r'(\d{1,2})\s*month', desc, re.I)
        if stated and int(stated.group(1)) != months:
            note = (note + '; ' if note else '') + (
                f'description says {stated.group(1)} months but the dates '
                f'in it span {months}')
        return start, end, note or f'{months} month term'
    return None, None, 'no term dates in the description'


def customer(desc):
    """The customer as the description leads with it."""
    head = re.split(r'\s[-|]\s|\s\|\s', desc)[0]
    head = re.sub(r'\b(Year|Term)\s*\d.*$', '', head, flags=re.I)
    head = re.sub(r'\s+', ' ', MONTH_RE.sub('', head)).strip(' -|(),')
    return ' '.join(head.split()[:3])


def classify(rows):
    """Tag every line as part of a release, a deferral, or neither.

    A release journal has both an income and a cost side under one reference
    and takes income into the P&L. A deferral is the same shape the other
    way round: it lifts a whole year's invoice back out to WIP.
    """
    byref = defaultdict(lambda: defaultdict(list))
    for r in rows:
        r['pairing'] = 'Unpaired'
        r['licence'] = None
        r['marker'] = None
        if r['source'] == 'Manual Journal' and r['ref']:
            byref[str(r['ref'])][r['code']].append(r)

    releases, deferrals = defaultdict(list), []
    for ref, v in byref.items():
        if not (v[INC] and v[EXP]):
            continue
        inc = sum(signed(x) for x in v[INC])
        if inc > 0:
            key = norm(v[INC][0]['desc'])
            releases[key].append(dict(
                ref=ref, date=min(x['date'] for x in v[INC]), inc=inc,
                cost=sum(signed(x) for x in v[EXP]), lines=v[INC] + v[EXP]))
        else:
            deferrals.append(dict(
                ref=ref, date=min(x['date'] for x in v[INC]), inc=inc,
                cost=sum(signed(x) for x in v[EXP]),
                desc=norm(v[INC][0]['desc']), lines=v[INC] + v[EXP]))
            for x in v[INC] + v[EXP]:
                x['pairing'] = 'Deferral'
    return releases, deferrals


# --------------------------------------------------------------- formatting

TITLE = Font(bold=True, size=13)
HEAD = Font(bold=True, color='FFFFFF', size=10)
HFILL = PatternFill('solid', fgColor='1F3864')
SUB = Font(bold=True, size=10)
SFILL = PatternFill('solid', fgColor='D9E2F3')
RED = PatternFill('solid', fgColor='FFC7CE')
AMBER = PatternFill('solid', fgColor='FFEB9C')
GREEN = PatternFill('solid', fgColor='C6EFCE')
THIN = Border(*[Side(style='thin', color='BFBFBF')] * 4)
MONEY = '#,##0.00;[Red]-#,##0.00'
PCT = '0.0%;[Red]-0.0%'
DATE = 'dd/mm/yyyy'
WRAP = Alignment(wrap_text=True, vertical='top')


def head_row(ws, r, labels, widths=None, freeze=True):
    for c, lab in enumerate(labels, 1):
        cell = ws.cell(r, c, lab)
        cell.font, cell.fill, cell.border, cell.alignment = (
            HEAD, HFILL, THIN, WRAP)
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
        if isinstance(v, str) and not v.startswith('='):
            cell.alignment = WRAP
    return r + 1


def bold_row(ws, r, cols):
    for c in cols:
        cell = ws.cell(r, c)
        cell.font, cell.fill = SUB, SFILL


# -------------------------------------------------------------------- build

def main(src, out):
    rows = read(src)
    releases, deferrals = classify(rows)

    # one licence id per licence, so the formulas stay short and no
    # description with a stray asterisk can act as a SUMIFS wildcard
    table = []
    for desc, items in releases.items():
        items.sort(key=lambda x: x['date'])
        n = len(items)
        inc = sum(x['inc'] for x in items)
        start, end, note = term(desc)
        table.append(dict(desc=desc, items=items, n=n,
                          pct=(inc + sum(x['cost'] for x in items)) / inc,
                          cust=customer(desc), start=start, end=end,
                          note=note))
    table.sort(key=lambda t: t['pct'])
    for i, t in enumerate(table, 1):
        t['id'] = f'L{i:02d}'
        for rel in t['items']:
            for x in rel['lines']:
                x['pairing'] = 'Release'
                x['licence'] = t['id']
            # one marker per journal, so a month can be counted
            next(x for x in rel['lines'] if x['code'] == INC)['marker'] = 1

    raw = sorted(rows, key=lambda x: (x['date'], x['code'], str(x['ref'])))
    first, last = 2, len(raw) + 1
    M = f'{RAW}!$M${first}:$M${last}'     # amount
    A = f'{RAW}!$A${first}:$A${last}'     # account code
    D = f'{RAW}!$D${first}:$D${last}'     # date
    E = f'{RAW}!$E${first}:$E${last}'     # contact
    N = f'{RAW}!$N${first}:$N${last}'     # month
    O = f'{RAW}!$O${first}:$O${last}'     # pairing
    P = f'{RAW}!$P${first}:$P${last}'     # licence id
    Q = f'{RAW}!$Q${first}:$Q${last}'     # month marker
    J = f'{RAW}!$J${first}:$J${last}'     # reference

    wb = openpyxl.Workbook()
    s_sum = wb.active
    s_sum.title = 'Read This First'
    s_lic = wb.create_sheet('Licence Margins')
    s_mon = wb.create_sheet('By Month')
    s_grid = wb.create_sheet('Licence by Month')
    s_unp = wb.create_sheet('Unpaired Cost and Income')
    s_def = wb.create_sheet('Deferrals to WIP')
    s_raw = wb.create_sheet('Raw Data')

    # ------------------------------------------------------- 6. raw data
    ws = s_raw
    ws['A1'] = ('Every line on 41175 Subscriptions and Licences Income and '
                '51500 Subscriptions and Licences Expense')
    ws['A1'].font = TITLE
    ws['A2'] = ('Columns M to Q are added. Everything on every other tab is '
                'a formula over them. Amount is credit less debit, so income '
                'is positive and cost is negative.')
    head_row(ws, 4, ['Account code', 'Account', 'Source', 'Date', 'Contact',
                     'Debit', 'Credit', 'Job', 'Invoice', 'Reference',
                     'Description', 'Cost centre', 'Amount', 'Month',
                     'Pairing', 'Licence ID', 'Release month'],
             [12, 24, 16, 11, 26, 12, 12, 10, 13, 13, 58, 18, 13, 11, 11,
              11, 12])
    r = 5
    for x in raw:
        r = put(ws, r, [
            x['code'], x['account'], x['source'], x['date'], x['contact'],
            x['debit'] or None, x['credit'] or None, x['job'], x['inv'],
            x['ref'], norm(x['desc']), x['cc'],
            f'=G{r}-F{r}', f'=DATE(YEAR(D{r}),MONTH(D{r}),1)',
            x['pairing'], x['licence'], x['marker']],
            ['0', None, None, DATE, None, MONEY, MONEY, None, None, None,
             None, None, MONEY, 'mmm yyyy', None, None, '0'])
    first, last = 5, r - 1
    M = f'{RAW}!$M${first}:$M${last}'
    A = f'{RAW}!$A${first}:$A${last}'
    D = f'{RAW}!$D${first}:$D${last}'
    E = f'{RAW}!$E${first}:$E${last}'
    N = f'{RAW}!$N${first}:$N${last}'
    O = f'{RAW}!$O${first}:$O${last}'
    P = f'{RAW}!$P${first}:$P${last}'
    Q = f'{RAW}!$Q${first}:$Q${last}'
    J = f'{RAW}!$J${first}:$J${last}'

    # ------------------------------------------------- 1. licence margins
    ws = s_lic
    ws['A1'] = 'Subscriptions and licences, margin per licence per month'
    ws['A1'].font = TITLE
    ws['A2'] = ('One row per licence. The monthly figures are what the '
                'monthly release journal puts through the P&L. Worst margin '
                'first. Every figure is a formula over Raw Data.')
    cols = ['ID', 'Licence (journal description)', 'Customer', 'Term start',
            'Term end', 'Months in term', 'Months released', 'First release',
            'Last release', 'Income per month', 'Cost per month',
            'Margin per month', 'Margin %', 'Income released to date',
            'Cost released to date', 'Margin released to date',
            'Months still to release', 'Income still to come', 'Note']
    head_row(ws, 5, cols,
             [6, 58, 20, 11, 11, 9, 9, 11, 11, 13, 13, 13, 9, 14, 14, 14, 9,
              14, 34])
    fmts = [None, None, None, DATE, DATE, '0', '0', DATE, DATE, MONEY, MONEY,
            MONEY, PCT, MONEY, MONEY, MONEY, '0', MONEY, None]
    r = 6
    for t in table:
        i = r
        r = put(ws, r, [
            t['id'], t['desc'], t['cust'], t['start'], t['end'],
            f'=IF(OR($D{i}="",$E{i}=""),"",(YEAR($E{i})-YEAR($D{i}))*12'
            f'+MONTH($E{i})-MONTH($D{i})+1)',
            f'=SUMIFS({Q},{P},$A{i})',
            # Raw Data is sorted by date and both sides of a journal
            # carry the same date, so the first and last match give the
            # first and last release. MINIFS is not available everywhere.
            f'=IFERROR(INDEX({D},MATCH($A{i},{P},0)),"")',
            f'=IFERROR(LOOKUP(2,1/({P}=$A{i}),{D}),"")',
            f'=IFERROR($N{i}/$G{i},"")',
            f'=IFERROR($O{i}/$G{i},"")',
            f'=IF($J{i}="","",$J{i}+$K{i})',
            f'=IFERROR($L{i}/$J{i},"")',
            f'=SUMIFS({M},{P},$A{i},{A},{INC})',
            f'=SUMIFS({M},{P},$A{i},{A},{EXP})',
            f'=$N{i}+$O{i}',
            f'=IF($F{i}="","",$F{i}-$G{i})',
            f'=IF(OR($Q{i}="",$Q{i}<=0),"",$J{i}*$Q{i})',
            t['note']], fmts)
    body_end = r - 1
    tot = r + 1
    ws.cell(tot, 2, 'Total, monthly run rate').font = SUB
    for c in (10, 11, 12, 14, 15, 16):
        L = get_column_letter(c)
        cell = ws.cell(tot, c, f'=SUM({L}6:{L}{body_end})')
        cell.number_format = MONEY
    ws.cell(tot, 13, f'=IFERROR(L{tot}/J{tot},"")').number_format = PCT
    bold_row(ws, tot, [2, 10, 11, 12, 13, 14, 15, 16])

    rng = f'M6:M{body_end}'
    ws.conditional_formatting.add(rng, CellIsRule(
        operator='lessThan', formula=['0.05'], fill=RED, stopIfTrue=True))
    ws.conditional_formatting.add(rng, CellIsRule(
        operator='lessThan', formula=['0.12'], fill=AMBER, stopIfTrue=True))
    ws.conditional_formatting.add(rng, CellIsRule(
        operator='greaterThanOrEqual', formula=['0.12'], fill=GREEN))
    ws.conditional_formatting.add(f'G6:G{body_end}', FormulaRule(
        formula=[f'AND($F6<>"",$G6>$F6)'], fill=RED))

    leg = tot + 2
    ws.cell(leg, 2, 'Colour on Margin %').font = SUB
    for i, (f, txt) in enumerate((
            (RED, 'RED: under 5 per cent. The licence barely covers its '
                  'own cost.'),
            (AMBER, 'AMBER: 5 to 12 per cent.'),
            (GREEN, 'GREEN: 12 per cent or better.'),
            (RED, 'RED on Months released: released more months than the '
                  'term runs for.')), 1):
        ws.cell(leg + i, 2, txt).fill = f

    # ------------------------------------------------------- 2. by month
    ws = s_mon
    ws['A1'] = 'Subscriptions and licences by month, as booked'
    ws['A1'].font = TITLE
    ws['A2'] = ('Everything on 41175 and 51500 in the month, not only the '
                'release journals. This is what the P&L shows.')
    head_row(ws, 4, ['Month', 'Income 41175', 'Cost 51500', 'Margin',
                     'Margin %', 'Income lines', 'Cost lines'],
             [12, 15, 15, 15, 10, 12, 12])
    months = sorted({x['date'].replace(day=1) for x in rows})
    r = 5
    for m in months:
        i = r
        r = put(ws, r, [m,
                        f'=SUMIFS({M},{N},$A{i},{A},{INC})',
                        f'=SUMIFS({M},{N},$A{i},{A},{EXP})',
                        f'=$B{i}+$C{i}', f'=IFERROR($D{i}/$B{i},"")',
                        f'=COUNTIFS({N},$A{i},{A},{INC})',
                        f'=COUNTIFS({N},$A{i},{A},{EXP})'],
                ['mmm yyyy', MONEY, MONEY, MONEY, PCT, '0', '0'])
    me = r - 1
    r += 1
    ws.cell(r, 1, 'Total').font = SUB
    for c in (2, 3, 4, 6, 7):
        L = get_column_letter(c)
        cell = ws.cell(r, c, f'=SUM({L}5:{L}{me})')
        cell.number_format = MONEY if c < 5 else '0'
    ws.cell(r, 5, f'=IFERROR(D{r}/B{r},"")').number_format = PCT
    bold_row(ws, r, range(1, 8))

    # ------------------------------------------------ 3. licence by month
    ws = s_grid
    ws['A1'] = 'Margin per licence per month, release journals only'
    ws['A1'].font = TITLE
    ws['A2'] = ('The exact month each release was posted. Blank means no '
                'release journal that month.')
    rel_months = sorted({x['date'].replace(day=1)
                         for t in table for rel in t['items']
                         for x in rel['lines']})
    head_row(ws, 4, ['ID', 'Licence'] + [''] * len(rel_months)
             + ['Total margin'],
             [6, 54] + [11] * len(rel_months) + [14])
    for c, m in enumerate(rel_months, 3):
        cell = ws.cell(4, c, m)
        cell.number_format = 'mmm yy'
        cell.font, cell.fill, cell.border, cell.alignment = (
            HEAD, HFILL, THIN, Alignment(horizontal='center'))
    lastcol = get_column_letter(len(rel_months) + 2)
    r = 5
    for t in table:
        i = r
        ws.cell(r, 1, t['id']).border = THIN
        cell = ws.cell(r, 2, t['desc'])
        cell.border, cell.alignment = THIN, WRAP
        for c in range(3, len(rel_months) + 3):
            L = get_column_letter(c)
            cell = ws.cell(r, c,
                           f'=IF(SUMIFS({Q},{P},$A{i},{N},{L}$4)=0,"",'
                           f'SUMIFS({M},{P},$A{i},{N},{L}$4))')
            cell.number_format, cell.border = MONEY, THIN
        cell = ws.cell(r, len(rel_months) + 3, f'=SUM(C{i}:{lastcol}{i})')
        cell.number_format, cell.font, cell.border = MONEY, SUB, THIN
        r += 1

    # -------------------------------------------- 4. lines that never pair
    ws = s_unp
    ws['A1'] = 'Lines with no matching other side'
    ws['A1'].font = TITLE
    ws['A2'] = ('Invoices, bills, and journals where only one of income or '
                'cost was posted. Unbilled licence cost sits here.')
    head_row(ws, 4, ['Side', 'Date', 'Source', 'Contact', 'Debit', 'Credit',
                     'Amount', 'Job', 'Invoice', 'Reference', 'Description'],
             [9, 11, 16, 26, 12, 12, 13, 10, 13, 13, 58])
    loose = [x for x in raw if x['pairing'] == 'Unpaired']
    loose.sort(key=lambda x: (x['code'], x['date']))
    r = 5
    for x in loose:
        i = r
        r = put(ws, r, ['Income' if x['code'] == INC else 'Cost', x['date'],
                        x['source'], x['contact'], x['debit'] or None,
                        x['credit'] or None, f'=F{i}-E{i}', x['job'],
                        x['inv'], x['ref'], norm(x['desc'])],
                [None, DATE, None, None, MONEY, MONEY, MONEY, None, None,
                 None, None])
    ue = r - 1
    r += 1
    ws.cell(r, 1, 'Total').font = SUB
    ws.cell(r, 7, f'=SUM(G5:G{ue})').number_format = MONEY
    bold_row(ws, r, [1, 7])

    # -------------------------------------------------------- 5. deferrals
    ws = s_def
    ws['A1'] = 'Journals that push a licence out of the P&L to WIP'
    ws['A1'].font = TITLE
    ws['A2'] = ('A 12 month invoice is taken out in full, then released a '
                'month at a time. Income negative, cost positive.')
    head_row(ws, 4, ['Date', 'Reference', 'Income taken out',
                     'Cost taken out', 'Net', 'Description'],
             [11, 16, 16, 16, 16, 70])
    deferrals.sort(key=lambda x: x['date'])
    r = 5
    for d in deferrals:
        i = r
        r = put(ws, r, [d['date'], d['ref'],
                        f'=SUMIFS({M},{J},$B{i},{A},{INC})',
                        f'=SUMIFS({M},{J},$B{i},{A},{EXP})',
                        f'=$C{i}+$D{i}', d['desc']],
                [DATE, None, MONEY, MONEY, MONEY, None])
    de = r - 1
    r += 1
    ws.cell(r, 1, 'Total').font = SUB
    for c in (3, 4, 5):
        L = get_column_letter(c)
        ws.cell(r, c, f'=SUM({L}5:{L}{de})').number_format = MONEY
    bold_row(ws, r, [1, 3, 4, 5])

    # ------------------------------------------------- 0. read this first
    ws = s_sum
    ws.column_dimensions['A'].width = 46
    for col in 'BCDE':
        ws.column_dimensions[col].width = 15
    ws['A1'] = 'Licence margins, the short answer'
    ws['A1'].font = TITLE

    def line(r, text, bold=False, fill=None):
        c = ws.cell(r, 1, text)
        if not str(text).startswith('='):
            c.alignment = WRAP
        if bold:
            c.font = SUB
        if fill:
            c.fill = fill
        return r + 1

    r = 3
    r = line(r, f'="Every licence, priced the way it was sold, makes money. '
                f'The average margin across the "'
                f'&COUNTA(\'Licence Margins\'!$B$6:$B${body_end})'
                f'&" licences is "'
                f'&TEXT(\'Licence Margins\'!$M${tot}*100,"0.0")'
                f'&" per cent."', True)
    r += 1
    r = line(r, 'The P&L shows a much thinner margin because licence cost '
                'is going through 51500 with nothing billed against it.')
    r += 1

    r = line(r, 'Margin per year, as booked', True)
    head_row(ws, r, ['Year', 'Income', 'Cost', 'Margin', 'Margin %'],
             freeze=False)
    r += 1
    for label, lo, hi in (('FY2025', (2024, 7, 1), (2025, 6, 30)),
                          ('FY2026', (2025, 7, 1), (2026, 6, 30)),
                          ('FY2027 to 30 Sep 2026', (2026, 7, 1),
                           (2026, 9, 30))):
        i = r
        win = (f'{D},">="&DATE({lo[0]},{lo[1]},{lo[2]}),'
               f'{D},"<="&DATE({hi[0]},{hi[1]},{hi[2]})')
        r = put(ws, r, [label,
                        f'=SUMIFS({M},{A},{INC},{win})',
                        f'=SUMIFS({M},{A},{EXP},{win})',
                        f'=$B{i}+$C{i}', f'=IFERROR($D{i}/$B{i},"")'],
                [None, MONEY, MONEY, MONEY, PCT])
    r += 1

    win27 = (f'{D},">="&DATE(2026,7,1),{D},"<="&DATE(2026,9,30)')
    r = line(r, 'Why FY2027 is so thin', True)
    head_row(ws, r, ['', 'Income', 'Cost', 'Margin', 'Margin %'],
             freeze=False)
    r += 1
    for label, pair in (('Licences released by journal', 'Release'),
                        ('Deferred out to WIP', 'Deferral'),
                        ('Cost and income with no release journal',
                         'Unpaired')):
        i = r
        r = put(ws, r, [label,
                        f'=SUMIFS({M},{A},{INC},{O},"{pair}",{win27})',
                        f'=SUMIFS({M},{A},{EXP},{O},"{pair}",{win27})',
                        f'=$B{i}+$C{i}', f'=IFERROR($D{i}/$B{i},"")'],
                [None, MONEY, MONEY, MONEY, PCT])
        if pair == 'Unpaired':
            for c in range(1, 6):
                ws.cell(i, c).fill = RED
    i = r
    r = put(ws, r, ['Total, as the P&L shows it',
                    f'=SUM(B{i-3}:B{i-1})', f'=SUM(C{i-3}:C{i-1})',
                    f'=$B{i}+$C{i}', f'=IFERROR($D{i}/$B{i},"")'],
            [None, MONEY, MONEY, MONEY, PCT])
    bold_row(ws, i, range(1, 6))
    r += 1

    r = line(r, 'The unbilled cost in FY2027', True)
    head_row(ws, r, ['Supplier', 'Bills', 'Cost'], freeze=False)
    r += 1
    agg = defaultdict(float)
    for x in loose:
        if x['code'] == EXP and datetime.date(2026, 7, 1) <= x['date'] \
                <= datetime.date(2026, 9, 30):
            agg[x['contact']] += signed(x)
    for k in sorted(agg, key=lambda y: agg[y]):
        crit = f'"{k}"' if k else '"="'
        label = str(k) if k else 'No supplier on the line'
        i = r
        r = put(ws, r, [label,
                        f'=COUNTIFS({A},{EXP},{O},"Unpaired",{E},{crit},'
                        f'{win27})',
                        f'=SUMIFS({M},{A},{EXP},{O},"Unpaired",{E},{crit},'
                        f'{win27})'],
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
                'By Month: the monthly totals.',
                'Raw Data: every line, and the columns every formula '
                'reads. Change a figure here and the workbook follows.'):
        r = line(r, txt)
    r += 1
    r = line(r, 'Note: the term dates and the Pairing and Licence ID '
                'columns on Raw Data are read out of the journal '
                'descriptions. Everything else is a formula.')

    wb.save(out)
    print(f'  licences            {len(table)}')
    print(f'  release journals    {sum(t["n"] for t in table)}')
    print(f'  deferral journals   {len(deferrals)}')
    print(f'  unpaired lines      {len(loose)}')
    print(f'  raw lines           {len(raw)}')
    print(f'  no term date        {sum(1 for t in table if not t["start"])}')
    print(f'  written             {out}')


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
