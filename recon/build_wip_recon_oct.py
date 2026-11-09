"""Work in Progress reconciliation to 31 October 2026, account 11300.

Same shape as the September file, with two tabs added:

  Where it went        the other side of each WIP journal, by account
  Accruals vs reversals  whether each month's accrual actually reversed

Those two are the correction. Reading the WIP account on its own makes a
released accrual look like free profit. It is not. The accrual comes out
because the real supplier bill has gone in.

Categories and the item netting are taken from build_wip_monthly, so this
file is cut the same way as last month's.

Usage: build_wip_recon_oct.py <WIP export> <GL Detail> <P&L Analysis> <out>
"""
import sys, re, datetime
from collections import defaultdict
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

sys.path.insert(0, __file__.rsplit('/', 1)[0])
from build_wip_monthly import load, item_of, cat_of

WIPNAME = 'Work in Progress'
CATS = ['Licences and subscriptions', 'Monthly WIP journal', 'WIP reclass',
        'Other accruals']
INCOME_WORDS = ('Income', 'Sales', 'Revenue')


def read_gl(path):
    """General Ledger Detail. Accounts are section headers, so the account
    is carried down from the last header seen."""
    rows = list(openpyxl.load_workbook(path, data_only=True, read_only=True)
                .worksheets[0].iter_rows(values_only=True))
    out, acct = [], None
    for r in rows[4:]:
        a = r[0]
        if a is None:
            continue
        if isinstance(a, datetime.datetime):
            out.append(dict(account=acct, date=a.date(),
                            source=str(r[1] or ''),
                            desc=re.sub(r'\s+', ' ', str(r[2] or '')).strip(),
                            ref=str(r[3] or ''), debit=float(r[4] or 0),
                            credit=float(r[5] or 0)))
            continue
        s = str(a)
        if s.startswith(('Total ', 'Net movement', 'Date')):
            continue
        if not [c for c in r[1:] if c not in (None, '')]:
            acct = s
    return out


def read_pl(path):
    """The P&L Analysis export. This is where the job numbers and the
    invoice numbers live. It covers profit and loss accounts only, so
    account 11300 is not in it."""
    rows = list(openpyxl.load_workbook(path, data_only=True, read_only=True)
                .worksheets[0].iter_rows(values_only=True))
    out = []
    for r in rows[4:]:
        try:
            code = int(r[0])
        except (TypeError, ValueError):
            continue
        if not isinstance(r[3], datetime.datetime):
            continue
        out.append(dict(
            code=code, account=str(r[1] or ''), source=str(r[2] or ''),
            date=r[3].date(), net=float(r[6] or 0) - float(r[5] or 0),
            job=str(r[7] or ''), inv=str(r[8] or ''), ref=str(r[9] or ''),
            desc=re.sub(r'\s+', ' ', str(r[10] or '')).strip()))
    return out


def match_jobs(wip, pl):
    """Put a job number on each WIP line.

    A WIP line and its other side share a journal reference. Usually they
    also share the figure, one as a debit and one as a credit, and that
    pair identifies the job on its own. Where the WIP side is one line and
    the other side is split across several, the figures do not pair, so
    the description is used instead. The two agree wherever both resolve.
    Anything that still points at more than one job is left blank rather
    than guessed at.
    """
    by_amt, by_desc = defaultdict(list), defaultdict(list)
    for x in pl:
        by_amt[(x['ref'], round(abs(x['net']), 2))].append(x)
        by_desc[(x['ref'], desckey(x['desc']))].append(x)

    def one(cands):
        jobs = {jobkey(y['job']) for y in cands if jobkey(y['job'])}
        if len(jobs) != 1:
            return None, None, len(jobs)
        j = jobs.pop()
        return j, next(y['job'] for y in cands if jobkey(y['job']) == j), 1

    for x in wip:
        ref = str(x['ref'])
        j, name, n = one(by_amt.get(
            (ref, round(abs(x['debit'] - x['credit']), 2)), []))
        how = 'matched on the figure'
        if j is None:
            j, name, n2 = one(by_desc.get((ref, desckey(x['desc'])), []))
            how = 'matched on the description'
            n = max(n, n2)
        if j:
            x['job'], x['jobname'], x['match'] = j, name, how
        else:
            x['job'], x['jobname'] = '', ''
            x['match'] = ('several jobs on this journal line' if n > 1
                          else 'no job found')
    return wip


def invoices_by_job(pl):
    """Every sales invoice raised on each job, with its date."""
    out = defaultdict(dict)
    for x in pl:
        if x['code'] < 50000 and x['inv'].startswith('INV-'):
            j = jobkey(x['job'])
            if j:
                out[j][x['inv']] = x['date']
    return out


def jobkey(j):
    m = re.match(r'\s*([0-9]{3,}[A-Za-z]?)', j or '')
    return m.group(1) if m else ''


def desckey(d):
    """An accrual and its reversal describe the same thing, so the word
    Reversal comes off before the two are compared."""
    d = re.sub(r'^Reversal:\s*', '', str(d or '')).strip()
    return re.sub(r'\s+', ' ', d).lower()[:70]


def kind(desc):
    return 'Reversal' if str(desc).lower().startswith('reversal') else 'Posting'


TITLE = Font(bold=True, size=13)
HEAD = Font(bold=True, color='FFFFFF', size=10)
HFILL = PatternFill('solid', fgColor='1F3864')
SUB = Font(bold=True, size=10)
SFILL = PatternFill('solid', fgColor='D9E2F3')
IN = PatternFill('solid', fgColor='FFF2CC')
RED = PatternFill('solid', fgColor='FFC7CE')
AMBER = PatternFill('solid', fgColor='FFEB9C')
GREEN = PatternFill('solid', fgColor='C6EFCE')
THIN = Border(*[Side(style='thin', color='BFBFBF')] * 4)
MONEY = '#,##0.00;[Red]-#,##0.00'
DATE = 'dd/mm/yyyy'
WRAP = Alignment(wrap_text=True, vertical='top')


def band(ws, r, labels, widths=None, freeze=True):
    for c, lab in enumerate(labels, 1):
        cell = ws.cell(r, c, lab)
        cell.font, cell.fill, cell.border, cell.alignment = (
            HEAD, HFILL, THIN, WRAP)
    if widths:
        for c, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(c)].width = w
    # Freezing and then clearing leaves a selection naming a pane that is
    # gone, and Excel calls the workbook damaged.
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


def bold(ws, r, cols):
    for c in cols:
        ws.cell(r, c).font = SUB
        ws.cell(r, c).fill = SFILL


def main(wip_path, gl_path, pl_path, out):
    period, opening, closing, wip = load(wip_path)
    gl = read_gl(gl_path)
    pl = read_pl(pl_path)
    refs = {str(x['ref']) for x in wip}
    other = [x for x in gl if x['ref'] in refs and x['account'] != WIPNAME]
    byref_job = defaultdict(set)
    for x in pl:
        if x['ref'] and jobkey(x['job']):
            byref_job[x['ref']].add(jobkey(x['job']))
    match_jobs(wip, pl)
    inv_by_job = invoices_by_job(pl)
    pl_from = min(x['date'] for x in pl)
    for x in wip:
        if x['date'] < pl_from:
            x['match'] = 'before the job export starts'

    wip.sort(key=lambda x: (x['date'], str(x['ref'])))
    other.sort(key=lambda x: (x['date'], x['account']))
    months = sorted({x['date'].replace(day=1) for x in wip})
    omonths = sorted({x['date'].replace(day=1) for x in other})

    wb = openpyxl.Workbook()
    s_rec = wb.active
    s_rec.title = 'Monthly recon'
    s_in = wb.create_sheet('What is in it')
    s_job = wb.create_sheet('By job and invoice')
    s_went = wb.create_sheet('Where it went')
    s_rev = wb.create_sheet('Accruals vs reversals')
    s_tx = wb.create_sheet('Transactions')
    s_src = wb.create_sheet('Source - from Jul 2025')
    s_osrc = wb.create_sheet('Source - other side')

    # ----------------------------------------------------- Transactions
    ws = s_tx
    ws['A1'] = 'Every line on account 11300'
    ws['A1'].font = TITLE
    ws['A2'] = f'From the export for {period}'
    band(ws, 4, ['Date', 'Month', 'Source', 'Description', 'Item',
                 'Category', 'Type', 'Reference', 'Debit', 'Credit', 'Net',
                 'Job number', 'Job', 'Invoiced', 'Invoice numbers',
                 'First invoice date', 'Last invoice date', 'Job match'],
         [11, 11, 15, 56, 50, 26, 11, 12, 13, 13, 13, 12, 42, 11, 30, 13,
          13, 26])
    r = 5
    for x in wip:
        i = r
        it = item_of(x['desc'])
        invs = inv_by_job.get(x['job'], {}) if x['job'] else {}
        dates = sorted(invs.values())
        r = put(ws, r, [x['date'], f'=DATE(YEAR(A{i}),MONTH(A{i}),1)',
                        x['source'], str(x['desc']), it, cat_of(it),
                        kind(x['desc']), str(x['ref']),
                        x['debit'] or None, x['credit'] or None,
                        f'=I{i}-J{i}',
                        x['job'] or None, x['jobname'] or None,
                        ('yes' if invs else
                         ('no' if x['job'] else None)),
                        ', '.join(sorted(invs)) or None,
                        dates[0] if dates else None,
                        dates[-1] if dates else None,
                        x['match']],
                [DATE, 'mmm yyyy', None, None, None, None, None, None,
                 MONEY, MONEY, MONEY, None, None, None, None, DATE, DATE,
                 None])
        if cat_of(it) == 'Licences and subscriptions':
            ws.cell(i, 6).fill = RED
        elif cat_of(it) == 'Monthly WIP journal':
            ws.cell(i, 6).fill = AMBER
        if x['job'] and not invs:
            ws.cell(i, 14).fill = AMBER
    t1, t2 = 5, r - 1
    T = "'Transactions'!"
    TB = f'{T}$B${t1}:$B${t2}'      # month
    TE = f'{T}$E${t1}:$E${t2}'      # item
    TF = f'{T}$F${t1}:$F${t2}'      # category
    TG = f'{T}$G${t1}:$G${t2}'      # type
    TI = f'{T}$I${t1}:$I${t2}'      # debit
    TJ = f'{T}$J${t1}:$J${t2}'      # credit
    TK = f'{T}$K${t1}:$K${t2}'      # net
    TL = f'{T}$L${t1}:$L${t2}'      # job number
    TN = f'{T}$N${t1}:$N${t2}'      # invoiced

    # ------------------------------------------------- Source - other side
    ws = s_osrc
    ws['A1'] = 'The other side of every WIP journal'
    ws['A1'].font = TITLE
    ws['A2'] = ('From the General Ledger Detail export. Lines sharing a '
                'reference with a WIP line, with WIP itself left out. The '
                'export starts 1 June 2026, so earlier months are not here.')
    band(ws, 4, ['Date', 'Month', 'Account', 'Income or Cost', 'Source',
                 'Description', 'Reference', 'Jobs', 'Debit', 'Credit',
                 'Net'],
         [11, 11, 32, 14, 15, 60, 12, 24, 13, 13, 13])
    r = 5
    for x in other:
        i = r
        side = ('Income' if any(w in x['account'] for w in INCOME_WORDS)
                else 'Cost')
        r = put(ws, r, [x['date'], f'=DATE(YEAR(A{i}),MONTH(A{i}),1)',
                        x['account'], side, x['source'], x['desc'],
                        x['ref'],
                        ', '.join(sorted(byref_job.get(x['ref'], ()))[:3]),
                        x['debit'] or None, x['credit'] or None,
                        f'=I{i}-J{i}'],
                [DATE, 'mmm yyyy', None, None, None, None, None, None,
                 MONEY, MONEY, MONEY])
    o1, o2 = 5, r - 1
    O = "'Source - other side'!"
    OB = f'{O}$B${o1}:$B${o2}'
    OC = f'{O}$C${o1}:$C${o2}'
    OD = f'{O}$D${o1}:$D${o2}'
    OK = f'{O}$K${o1}:$K${o2}'

    # ------------------------------------------------ Source - from Jul 2025
    ws = s_src
    ws['A1'] = 'Work in Progress Transactions'
    ws['A1'].font = TITLE
    ws['A2'] = 'Corporate Technology Services Pty Ltd'
    ws['A3'] = period
    band(ws, 4, ['Date', 'Source', 'Description', 'Reference', 'Debit',
                 'Credit'], [11, 15, 74, 12, 13, 13])
    ws.cell(5, 1, 'Opening Balance').font = SUB
    c = ws.cell(5, 6, round(-opening, 2))
    c.number_format, c.fill, c.border = MONEY, IN, THIN
    r = 6
    for x in wip:
        r = put(ws, r, [x['date'], x['source'], str(x['desc']),
                        str(x['ref']), x['debit'] or None,
                        x['credit'] or None],
                [DATE, None, None, None, MONEY, MONEY])
    ws.cell(r + 1, 1, 'Closing Balance').font = SUB
    c = ws.cell(r + 1, 6, round(-closing, 2))
    c.number_format, c.fill, c.border = MONEY, IN, THIN

    # ------------------------------------------------------ Monthly recon
    ws = s_rec
    ws['A1'] = 'Work in progress, account 11300'
    ws['A1'].font = TITLE
    ws['A2'] = 'Corporate Technology Services Pty Ltd'
    ws['A3'] = period
    ws.column_dimensions['A'].width = 40
    ws['A5'] = 'WHERE IT LANDS'
    ws['A5'].font = SUB
    rows_ = [('Balance at 1 July 2025, per the export', round(opening, 2),
              True),
             ('Balance at 30 September 2026, Q1 FY27 close', None, False),
             ('Balance at 31 October 2026, per the export',
              round(closing, 2), True),
             ('Months in the schedule', len(months), False),
             ('Lines in the schedule', len(wip), False)]
    r = 6
    sep_row = None
    for lab, val, is_input in rows_:
        ws.cell(r, 1, lab).alignment = WRAP
        if lab.startswith('Balance at 30 September'):
            sep_row = r
        else:
            c = ws.cell(r, 4, val)
            c.border = THIN
            c.number_format = MONEY if isinstance(val, float) else '0'
            if is_input:
                c.fill = IN
        r += 1
    head = r + 1
    band(ws, head, ['Month', 'Opening', 'Debits', 'Credits', 'Movement',
                    'Closing', 'Lines'],
         [40, 15, 15, 15, 15, 15, 8])
    r = head + 1
    first_data = r
    for n, m in enumerate(months):
        i = r
        prev = '$D$6' if n == 0 else f'$F{i-1}'
        r = put(ws, r, [m, f'={prev}', f'=SUMIFS({TI},{TB},$A{i})',
                        f'=SUMIFS({TJ},{TB},$A{i})', f'=$C{i}-$D{i}',
                        f'=$B{i}+$E{i}', f'=COUNTIFS({TB},$A{i})'],
                ['mmm yyyy', MONEY, MONEY, MONEY, MONEY, MONEY, '0'])
        if m == datetime.date(2026, 9, 1):
            sep_closing = i
    last_data = r - 1
    r += 1
    ws.cell(r, 1, 'Total and closing').font = SUB
    for col, f in ((3, f'=SUM(C{first_data}:C{last_data})'),
                   (4, f'=SUM(D{first_data}:D{last_data})'),
                   (5, f'=SUM(E{first_data}:E{last_data})'),
                   (6, f'=F{last_data}')):
        ws.cell(r, col, f).number_format = MONEY
    bold(ws, r, range(1, 8))
    tot_row = r
    ws.cell(sep_row, 4, f'=F{sep_closing}').number_format = MONEY
    ws.cell(sep_row, 4).border = THIN
    r += 2
    ws.cell(r, 1, 'Agrees with the export').font = SUB
    ws.cell(r, 4, f'=IF(ROUND(F{tot_row}-D8,2)=0,"OK","CHECK")').font = SUB
    r += 1
    ws.cell(r, 1, 'Difference').alignment = WRAP
    ws.cell(r, 4, f'=F{tot_row}-D8').number_format = MONEY

    # ------------------------------------------------------ What is in it
    ws = s_in
    ws['A1'] = 'What is sitting in the balance'
    ws['A1'].font = TITLE
    ws['A2'] = ('Every item netted against its own reversal, so an accrual '
                'that reversed nets to nil and drops out.')
    ws['A4'] = 'BY WHAT THE ITEM IS'
    ws['A4'].font = SUB
    band(ws, 5, ['Category', 'Items', 'Still carrying a balance',
                 'Balance'], [40, 10, 22, 18])
    items = defaultdict(float)
    for x in wip:
        items[item_of(x['desc'])] += x['debit'] - x['credit']
    r = 6
    for cat in CATS:
        i = r
        r = put(ws, r, [cat,
                        f'=COUNTIFS($A$30:$A$9999,$A{i})',
                        f'=COUNTIFS($A$30:$A$9999,$A{i},$C$30:$C$9999,'
                        f'"<>0")',
                        f'=SUMIFS($C$30:$C$9999,$A$30:$A$9999,$A{i})'],
                [None, '0', '0', MONEY])
        if cat == 'Licences and subscriptions':
            ws.cell(i, 1).fill = RED
        elif cat == 'Monthly WIP journal':
            ws.cell(i, 1).fill = AMBER
    i = r
    r = put(ws, r, ['Total activity', f'=SUM(B6:B{i-1})',
                    f'=SUM(C6:C{i-1})', f'=SUM(D6:D{i-1})'],
            [None, '0', '0', MONEY])
    bold(ws, i, range(1, 5))
    act_row = i
    r += 1
    j = r
    r = put(ws, r, ['Opening balance at 1 July 2025', '', '',
                    "='Monthly recon'!$D$6"], [None, None, None, MONEY])
    r = put(ws, r, ['Closing balance at 31 October 2026', '', '',
                    f'=$D${j}+$D${act_row}'], [None, None, None, MONEY])
    bold(ws, r - 1, range(1, 5))
    r = put(ws, r, ['Agrees with the export', '', '',
                    f"=IF(ROUND($D{r-1}-'Monthly recon'!$D$8,2)=0,"
                    f'"OK","CHECK")'], [None, None, None, None])
    r += 1
    ws.cell(r, 1, 'The point').font = SUB
    r += 1
    for txt in ('The opening balance is almost the whole closing balance. '
                'Sixteen months of activity nets to under a thousand '
                'dollars.',
                'So the balance did not come from anything in this file. '
                'It was already there at 1 July 2025.'):
        c = ws.cell(r, 1, txt)
        c.alignment = WRAP
        ws.merge_cells(start_row=r, start_column=1, end_row=r,
                       end_column=4)
        r += 1

    ws.cell(29, 1, 'EVERY ITEM').font = SUB
    band(ws, 29, ['Category', 'Item', 'Balance', 'Lines'],
         None, freeze=False)
    r = 30
    for it in sorted(items, key=lambda y: items[y]):
        cat = cat_of(it)
        i = r
        r = put(ws, r, [cat, it, round(items[it], 2),
                        f'=COUNTIFS({TE},$B{i})'],
                [None, None, MONEY, '0'])

    # -------------------------------------------------- By job and invoice
    ws = s_job
    ws['A1'] = 'WIP balance by job, against the invoices on that job'
    ws['A1'].font = TITLE
    ws['A2'] = ('This is the test of whether an item still belongs in WIP. '
                'Cost sitting in WIP is right while the job is not '
                'invoiced. Once the job is invoiced the cost should have '
                'come out.')
    ws['A3'] = ('Job numbers come from matching each WIP line to its other '
                'side on the journal. The job export starts 1 June 2026, '
                'so lines before that carry no job.')
    ws['A3'].font = Font(italic=True, size=9)
    band(ws, 5, ['Job number', 'Job', 'WIP balance', 'Lines', 'Invoiced',
                 'Invoice numbers', 'First invoice date',
                 'Last invoice date', 'Days since last invoice',
                 'Look at this one'],
         [12, 44, 15, 8, 10, 32, 13, 13, 13, 40])
    jobs = {}
    for x in wip:
        if x['job']:
            jobs.setdefault(x['job'], x['jobname'])
    order = sorted(jobs, key=lambda j: sum(
        y['debit'] - y['credit'] for y in wip if y['job'] == j))
    r = 6
    for j in order:
        i = r
        invs = inv_by_job.get(j, {})
        dates = sorted(invs.values())
        r = put(ws, r, [
            j, jobs[j],
            f'=SUMIFS({TK},{TL},$A{i})',
            f'=COUNTIFS({TL},$A{i})',
            'yes' if invs else 'no',
            ', '.join(sorted(invs)) or None,
            dates[0] if dates else None,
            dates[-1] if dates else None,
            f'=IF($H{i}="","",DATE(2026,10,31)-$H{i})',
            f'=IF(ROUND($C{i},2)=0,"",'
            f'IF($E{i}="no","not invoiced yet, so WIP is correct",'
            f'IF($I{i}>60,"invoiced over 60 days ago and still in WIP",'
            f'"invoiced recently, check it has cleared")))'],
            [None, None, MONEY, '0', None, None, DATE, DATE, '0', None])
    job_last = r - 1
    r += 1
    ws.cell(r, 2, 'Total with a job number').font = SUB
    ws.cell(r, 3, f'=SUM(C6:C{job_last})').number_format = MONEY
    ws.cell(r, 4, f'=SUM(D6:D{job_last})').number_format = '0'
    bold(ws, r, range(1, 11))
    r += 1
    ws.cell(r, 2, 'Lines with no job number').font = SUB
    ws.cell(r, 3, f'=SUM({TK})-C{r-1}').number_format = MONEY
    ws.cell(r, 4, f'=COUNTA({TK})-D{r-1}').number_format = '0'
    bold(ws, r, range(1, 11))
    r += 1
    ws.cell(r, 2, 'Total activity, ties to the recon').font = SUB
    ws.cell(r, 3, f'=C{r-2}+C{r-1}').number_format = MONEY
    bold(ws, r, range(1, 11))
    from openpyxl.formatting.rule import FormulaRule as _FR
    ws.conditional_formatting.add(f'A6:J{job_last}', _FR(
        formula=['$I6>60'], fill=RED))

    # ------------------------------------------------------- Where it went
    ws = s_went
    ws['A1'] = 'Where the WIP movement went'
    ws['A1'].font = TITLE
    ws['A2'] = ('The other side of each WIP journal, by account, June to '
                'October 2026. A credit here means cost came out of the '
                'P&L or income went in.')
    ws['A3'] = ('Read it with the supplier bills in mind. In Q1 the big '
                'credits are accruals releasing because the real bill '
                'arrived in the same quarter, not profit appearing.')
    ws['A3'].font = Font(italic=True, size=9)
    accounts = sorted({x['account'] for x in other})
    band(ws, 5, ['Account', 'Income or Cost'] + [''] * len(omonths)
         + ['Total'], [34, 14] + [13] * len(omonths) + [14])
    for c, m in enumerate(omonths, 3):
        cell = ws.cell(5, c, m)
        cell.number_format = 'mmm yy'
        cell.font, cell.fill, cell.border = HEAD, HFILL, THIN
        cell.alignment = Alignment(horizontal='center')
    lastcol = get_column_letter(len(omonths) + 2)
    totcol = len(omonths) + 3
    r = 6
    for a in accounts:
        i = r
        ws.cell(r, 1, a).border = THIN
        ws.cell(r, 1).alignment = WRAP
        ws.cell(r, 2, 'Income' if any(w in a for w in INCOME_WORDS)
                else 'Cost').border = THIN
        for c in range(3, totcol):
            L = get_column_letter(c)
            cell = ws.cell(r, c, f'=SUMIFS({OK},{OC},$A{i},{OB},{L}$5)')
            cell.number_format, cell.border = MONEY, THIN
        cell = ws.cell(r, totcol, f'=SUM(C{i}:{lastcol}{i})')
        cell.number_format, cell.font, cell.border = MONEY, SUB, THIN
        r += 1
    r += 1
    for label, crit in (('Income accounts', 'Income'),
                        ('Cost accounts', 'Cost'), ('All accounts', None)):
        ws.cell(r, 1, label).font = SUB
        for c in range(3, totcol):
            L = get_column_letter(c)
            f = (f'=SUMIFS({OK},{OB},{L}$5)' if crit is None else
                 f'=SUMIFS({OK},{OB},{L}$5,{OD},"{crit}")')
            ws.cell(r, c, f).number_format = MONEY
        ws.cell(r, totcol, f'=SUM(C{r}:{lastcol}{r})').number_format = MONEY
        bold(ws, r, [1, 2] + list(range(3, totcol + 1)))
        r += 1

    # ------------------------------------------------ Accruals vs reversals
    ws = s_rev
    ws['A1'] = 'Accruals posted against reversals'
    ws['A1'].font = TITLE
    ws['A2'] = ('An accrual should reverse the month after it is posted. '
                'This shows whether it did. A reversal is a line whose '
                'description starts with the word Reversal.')
    band(ws, 4, ['Month', 'Posted', 'Reversed', 'Net', 'Posting lines',
                 'Reversal lines', 'Any reversals'],
         [14, 16, 16, 16, 14, 14, 14])
    r = 5
    for m in months:
        i = r
        r = put(ws, r, [m,
                        f'=SUMIFS({TK},{TB},$A{i},{TG},"Posting")',
                        f'=SUMIFS({TK},{TB},$A{i},{TG},"Reversal")',
                        f'=$B{i}+$C{i}',
                        f'=COUNTIFS({TB},$A{i},{TG},"Posting")',
                        f'=COUNTIFS({TB},$A{i},{TG},"Reversal")',
                        f'=IF($F{i}=0,"none","yes")'],
                ['mmm yyyy', MONEY, MONEY, MONEY, '0', '0', None])
    rl = r - 1
    r += 1
    ws.cell(r, 1, 'Total').font = SUB
    for c in (2, 3, 4, 5, 6):
        L = get_column_letter(c)
        ws.cell(r, c, f'=SUM({L}5:{L}{rl})').number_format = (
            MONEY if c < 5 else '0')
    bold(ws, r, range(1, 8))
    from openpyxl.formatting.rule import FormulaRule
    ws.conditional_formatting.add(f'A5:G{rl}', FormulaRule(
        formula=[f'$F5=0'], fill=RED))

    wb.save(out)
    print(f'  period           {period}')
    print(f'  WIP lines        {len(wip)}')
    print(f'  other side lines {len(other)}')
    print(f'  items            {len(items)}')
    print(f'  opening {opening:,.2f}   closing {closing:,.2f}')
    print(f'  written          {out}')


if __name__ == '__main__':
    main(*sys.argv[1:5])
