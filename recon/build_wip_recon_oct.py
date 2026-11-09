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
    for n, r in enumerate(rows[4:], 5):
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
            desc=re.sub(r'\s+', ' ', str(r[10] or '')).strip(),
            row=n))
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


def paste_source(wb, title, path):
    """Copy a Xero export into the workbook cell for cell.

    Every row, every column, in the order the report prints them, with the
    report's own title rows, its Opening and Closing Balance lines and its
    number formats. Nothing is renamed, reordered or left out, so the tab
    can be checked straight against the report in Xero.
    """
    src = openpyxl.load_workbook(path, data_only=True).worksheets[0]
    ws = wb.create_sheet(title)
    for row in src.iter_rows():
        for cell in row:
            if cell.value is None:
                continue
            out = ws.cell(cell.row, cell.column, cell.value)
            out.number_format = cell.number_format
    for letter, dim in src.column_dimensions.items():
        if dim.width:
            ws.column_dimensions[letter].width = dim.width
    for c in range(1, src.max_column + 1):
        letter = get_column_letter(c)
        if letter not in ws.column_dimensions or \
                not ws.column_dimensions[letter].width:
            ws.column_dimensions[letter].width = 18
    ws.freeze_panes = 'A6'
    return ws


def wip_balance_rows(path):
    """Which rows of the WIP export hold Opening Balance and Closing
    Balance, so the recon can read them rather than carry a typed copy."""
    ws = openpyxl.load_workbook(path, data_only=True,
                                read_only=True).worksheets[0]
    opening = closing = None
    for n, r in enumerate(ws.iter_rows(values_only=True), 1):
        label = str(r[0]).strip() if r[0] is not None else ''
        if label == 'Opening Balance':
            opening = n
        elif label == 'Closing Balance':
            closing = n
    if not (opening and closing):
        raise SystemExit('the export has no Opening or Closing Balance row')
    return opening, closing


def jobkey(j):
    m = re.match(r'\s*([0-9]{3,}[A-Za-z]?)', j or '')
    return m.group(1) if m else ''


def desckey(d):
    """An accrual and its reversal describe the same thing, so the word
    Reversal comes off before the two are compared."""
    d = re.sub(r'^Reversal:\s*', '', str(d or '')).strip()
    return re.sub(r'\s+', ' ', d).lower()[:70]


def xl_item(desc):
    """The item exactly as the Item column on Transactions works it out.

    The formula there strips a leading "Reversal:" and then TRIMs, and
    Excel's TRIM also squeezes runs of spaces down to one. Two source
    descriptions that differ only by a stray double space are therefore
    the same item on the sheet, so they have to be the same item here.
    """
    d = str(desc or '')
    if d[:9].lower() == 'reversal:':
        d = d[9:]
    # Xero descriptions carry line breaks. Excel's TRIM does not remove
    # them, so they are turned into spaces first, on both sides.
    d = d.replace('\r', ' ').replace('\n', ' ')
    return re.sub(r' +', ' ', d).strip()


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
MONEY0 = '#,##0.00;[Red]-#,##0.00;""'
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
    open_row, close_row = wip_balance_rows(wip_path)
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
    s_life = wb.create_sheet('Job life cycle')
    s_went = wb.create_sheet('Where it went')
    s_rev = wb.create_sheet('Accruals vs reversals')
    s_tx = wb.create_sheet('Transactions')
    s_keys = wb.create_sheet('Job keys')
    s_inv = wb.create_sheet('Job invoices')
    s_osrc = wb.create_sheet('Source - other side')

    # --------------------------------------------------------- Job keys
    # One row for every line on the P&L analysis export, in its own order,
    # every cell a formula pointing back at it. Two keys are built here:
    # one on the journal reference and the figure, one on the journal
    # reference and the description. Both are what the Transactions tab
    # looks a job number up on.
    PX = "'Xero - P&L analysis'!"
    WX = "'Xero - WIP transactions'!"
    KY = "'Job keys'!"
    JI = "'Job invoices'!"
    ws = s_keys
    ws['A1'] = 'Job keys, built from the P&L analysis export'
    ws['A1'].font = TITLE
    ws['A2'] = ('Nothing is typed here. Every cell reads the row named in '
                'column A. The two key columns are how a WIP line finds '
                'its job.')
    band(ws, 4, ['Where it comes from', 'Key on the figure',
                 'Key on the description', 'Job number', 'Job', 'Account',
                 'Income or cost', 'Date', 'Invoice', 'Reference',
                 'Source', 'Amount'],
         [26, 30, 56, 12, 42, 32, 13, 11, 16, 13, 16, 13])
    r = 5
    for x in pl:
        n = x['row']
        i = r
        r = put(ws, r, [
            f'Xero - P&L analysis row {n}',
            f'={PX}J{n}&"|"&TEXT(ROUND(ABS({PX}G{n}-{PX}F{n}),2),"0.00")',
            f'={PX}J{n}&"|"&LOWER(LEFT(TRIM(SUBSTITUTE(SUBSTITUTE('
            f'IF(LEFT({PX}K{n},9)="Reversal:",MID({PX}K{n},10,500),'
            f'{PX}K{n}),CHAR(10)," "),CHAR(13)," ")),70))',
            f'=TRIM(LEFT({PX}H{n},FIND(" ",{PX}H{n}&" ")-1))',
            f'={PX}H{n}', f'={PX}A{n}&" "&{PX}B{n}',
            f'=IF({PX}A{n}<50000,"Income","Cost")',
            f'={PX}D{n}', f'={PX}I{n}', f'={PX}J{n}', f'={PX}C{n}',
            f'={PX}G{n}-{PX}F{n}'],
            [None, None, None, None, None, None, None, DATE, None, None,
             None, MONEY])
        ws.cell(i, 1).font = Font(italic=True, size=8)
    k1, k2 = 5, r - 1
    KB = f'{KY}$B${k1}:$B${k2}'
    KC = f'{KY}$C${k1}:$C${k2}'
    KD = f'{KY}$D${k1}:$D${k2}'
    KE = f'{KY}$E${k1}:$E${k2}'
    KH = f'{KY}$H${k1}:$H${k2}'
    KL = f'{KY}$L${k1}:$L${k2}'

    # ----------------------------------------------------- Job invoices
    # Every sales invoice line, job by job and oldest first, so a job's
    # first invoice is its first match and its last is its last.
    ws = s_inv
    ws['A1'] = 'Sales invoices raised on each job'
    ws['A1'].font = TITLE
    ws['A2'] = ('Taken from the Job keys tab, sorted by job and then by '
                'date. The last column builds up the list of invoices on '
                'the job as you go down.')
    band(ws, 4, ['Job number', 'Invoice', 'Date',
                 'Invoices on this job so far', 'Where it comes from'],
         [12, 18, 12, 40, 26])
    rowmap = {x['row']: n for n, x in enumerate(pl, 5)}
    invlines = [x for x in pl
                if x['code'] < 50000 and x['inv'].startswith('INV-')
                and jobkey(x['job'])]
    invlines.sort(key=lambda x: (jobkey(x['job']), x['date'], x['inv']))
    r = 5
    for x in invlines:
        kr = rowmap[x['row']]
        i = r
        prev = i - 1
        lst = (f'=IF($A{i}<>$A{prev},$B{i},'
               f'IF(ISNUMBER(SEARCH($B{i},$D{prev})),$D{prev},'
               f'$D{prev}&", "&$B{i}))') if i > 5 else f'=$B{i}'
        r = put(ws, r, [f'={KY}D{kr}', f'={KY}I{kr}', f'={KY}H{kr}', lst,
                        f'Job keys row {kr}'],
                [None, None, DATE, None, None])
        ws.cell(i, 5).font = Font(italic=True, size=8)
    j1, j2 = 5, r - 1
    JA = f'{JI}$A${j1}:$A${j2}'
    JB = f'{JI}$B${j1}:$B${j2}'
    JC = f'{JI}$C${j1}:$C${j2}'
    JD = f'{JI}$D${j1}:$D${j2}'

    # ----------------------------------------------------- Transactions
    ws = s_tx
    ws['A1'] = 'Every line on account 11300'
    ws['A1'].font = TITLE
    ws['A2'] = (f'From the export for {period}. Nothing is typed on this '
                f'tab. Every cell is a formula, and the last column names '
                f'the Xero row it reads.')
    band(ws, 4, ['Date', 'Month', 'Source', 'Description', 'Item',
                 'Category', 'Type', 'Reference', 'Debit', 'Credit', 'Net',
                 'Key on the figure', 'Job by the figure',
                 'Key on the description', 'Job by the description',
                 'Job number', 'Job', 'Invoiced', 'Invoice numbers',
                 'First invoice date', 'Last invoice date',
                 'Where it comes from',
                 'Item key, used by the formulas'],
         [11, 11, 15, 54, 48, 24, 11, 12, 13, 13, 13, 26, 13, 44, 13, 12,
          40, 10, 30, 13, 13, 26, 30])
    LICW = ','.join(f'"{w}"' for w in
                    ('licence', 'license', 'proxyclick', 'eptura', 'xio',
                     'subscription', 'desk booking', 'appspace',
                     'visitor mg'))
    r = 5
    for x in wip:
        i = r
        n = x['row']
        uniq = lambda key, kr: (
            f'=IF(COUNTIF({kr},${key}{i})=0,"",'
            f'IF(COUNTIF({kr},${key}{i})=COUNTIFS({kr},${key}{i},{KD},'
            f'INDEX({KD},MATCH(${key}{i},{kr},0))),'
            f'INDEX({KD},MATCH(${key}{i},{kr},0)),""))')
        r = put(ws, r, [
            f'={WX}A{n}', f'=DATE(YEAR(A{i}),MONTH(A{i}),1)',
            f'={WX}B{n}', f'={WX}C{n}',
            f'=TRIM(SUBSTITUTE(SUBSTITUTE(IF(LEFT(D{i},9)="Reversal:",'
            f'MID(D{i},10,500),D{i}),CHAR(10)," "),CHAR(13)," "))',
            f'=IF(SUMPRODUCT(--ISNUMBER(SEARCH({{{LICW}}},E{i})))>0,'
            f'"Licences and subscriptions",'
            f'IF(OR(LEFT(E{i},11)="WIP Journal",LEFT(E{i},10)='
            f'"WIPJournal"),"Monthly WIP journal",'
            f'IF(LEFT(E{i},11)="WIP Reclass","WIP reclass",'
            f'"Other accruals")))',
            f'=IF(LEFT(D{i},8)="Reversal","Reversal","Posting")',
            f'={WX}D{n}', f'={WX}E{n}', f'={WX}F{n}', f'=I{i}-J{i}',
            f'=H{i}&"|"&TEXT(ROUND(ABS(K{i}),2),"0.00")',
            uniq('L', KB),
            f'=H{i}&"|"&LOWER(LEFT(E{i},70))',
            uniq('N', KC),
            f'=IF(M{i}<>"",M{i},O{i})',
            f'=IF(P{i}="","",IFERROR(INDEX({KE},MATCH(P{i},{KD},0)),""))',
            f'=IF(P{i}="","",IF(COUNTIF({JA},P{i})>0,"yes","no"))',
            f'=IF(P{i}="","",IFERROR(LOOKUP(2,1/({JA}=P{i}),{JD}),""))',
            f'=IF(P{i}="","",IFERROR(INDEX({JC},MATCH(P{i},{JA},0)),""))',
            f'=IF(P{i}="","",IFERROR(LOOKUP(2,1/({JA}=P{i}),{JC}),""))',
            f'Xero - WIP transactions row {n}',
            f'=SUBSTITUTE(SUBSTITUTE(SUBSTITUTE($E{i},"~","~~"),'
            f'"*","~*"),"?","~?")'],
            [DATE, 'mmm yyyy', None, None, None, None, None, None, MONEY,
             MONEY, MONEY, None, None, None, None, None, None, None, None,
             DATE, DATE, None, None])
        ws.cell(i, 22).font = Font(italic=True, size=8)
        for c in (12, 13, 14, 15):
            ws.cell(i, c).font = Font(italic=True, size=8)
    t1, t2 = 5, r - 1
    T = "'Transactions'!"
    TA = f'{T}$A${t1}:$A${t2}'      # date
    TB = f'{T}$B${t1}:$B${t2}'      # month
    TE = f'{T}$E${t1}:$E${t2}'      # item
    TF = f'{T}$F${t1}:$F${t2}'      # category
    TG = f'{T}$G${t1}:$G${t2}'      # type
    TI = f'{T}$I${t1}:$I${t2}'      # debit
    TJ = f'{T}$J${t1}:$J${t2}'      # credit
    TK = f'{T}$K${t1}:$K${t2}'      # net
    TL = f'{T}$P${t1}:$P${t2}'      # job number
    TN = f'{T}$R${t1}:$R${t2}'      # invoiced
    TW = f'{T}$W${t1}:$W${t2}'      # item key

    # ------------------------------------------------- Source - other side
    # The other side of each WIP journal, taken from the P&L analysis
    # export so every cell can be a formula. The ledger export also holds
    # Tracking Transfers on these journals, which nets to nil and is a
    # clearing account, not profit and loss.
    ws = s_osrc
    ws['A1'] = 'The other side of every WIP journal'
    ws['A1'].font = TITLE
    ws['A2'] = ('Lines on the P&L analysis export that share a journal '
                'reference with a WIP line. It starts 1 June 2026, so '
                'earlier months are not here. Nothing is typed.')
    band(ws, 4, ['Date', 'Month', 'Account', 'Income or Cost', 'Source',
                 'Description', 'Reference', 'Job', 'Debit', 'Credit',
                 'Net', 'Where it comes from'],
         [11, 11, 32, 14, 15, 58, 12, 24, 13, 13, 13, 26])
    other = [x for x in pl if x['ref'] in refs]
    other.sort(key=lambda x: (x['date'], x['account']))
    r = 5
    for x in other:
        i = r
        n = x['row']
        kr = rowmap[n]
        r = put(ws, r, [
            f'={PX}D{n}', f'=DATE(YEAR(A{i}),MONTH(A{i}),1)',
            f'={KY}F{kr}', f'={KY}G{kr}', f'={PX}C{n}', f'={PX}K{n}',
            f'={PX}J{n}', f'={PX}H{n}', f'={PX}F{n}', f'={PX}G{n}',
            f'=J{i}-I{i}', f'Xero - P&L analysis row {n}'],
            [DATE, 'mmm yyyy', None, None, None, None, None, None, MONEY,
             MONEY, MONEY, None])
        ws.cell(i, 12).font = Font(italic=True, size=8)
    o1, o2 = 5, r - 1
    O = "'Source - other side'!"
    OA = f'{O}$A${o1}:$A${o2}'
    OB = f'{O}$B${o1}:$B${o2}'
    OC = f'{O}$C${o1}:$C${o2}'
    OD = f'{O}$D${o1}:$D${o2}'
    OK = f'{O}$K${o1}:$K${o2}'

    # ------------------------------------------------------ Monthly recon
    ws = s_rec
    ws['A1'] = 'Work in progress, account 11300'
    ws['A1'].font = TITLE
    ws['A2'] = 'Corporate Technology Services Pty Ltd'
    ws['A3'] = period
    ws.column_dimensions['A'].width = 40
    ws['A5'] = 'WHERE IT LANDS'
    ws['A5'].font = SUB
    # The opening and closing balances are read off the export's own
    # Opening Balance and Closing Balance rows, not typed in.
    rows_ = [('Balance at 1 July 2025, per the export',
              f'={WX}E{open_row}-{WX}F{open_row}', MONEY, True),
             ('Balance at 30 September 2026, Q1 FY27 close', None, MONEY,
              False),
             ('Balance at 31 October 2026, per the export',
              f'={WX}E{close_row}-{WX}F{close_row}', MONEY, True),
             ('Months in the schedule', None, '0', False),
             ('Lines in the schedule', f'=COUNT({TK})', '0', False)]
    r = 6
    sep_row = months_row = None
    for lab, val, fmt, is_input in rows_:
        ws.cell(r, 1, lab).alignment = WRAP
        if lab.startswith('Balance at 30 September'):
            sep_row = r
        elif lab.startswith('Months in'):
            months_row = r
        else:
            c = ws.cell(r, 4, val)
            c.border = THIN
            c.number_format = fmt
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
        month = (f'=DATE(YEAR(MIN({TA})),MONTH(MIN({TA})),1)' if n == 0
                 else f'=EDATE($A{i-1},1)')
        r = put(ws, r, [month, f'={prev}', f'=SUMIFS({TI},{TB},$A{i})',
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
    c = ws.cell(months_row, 4, f'=COUNT($A${first_data}:$A${last_data})')
    c.number_format, c.border = '0', THIN
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
        items[xl_item(x['desc'])] += x['debit'] - x['credit']
    r = 6
    for cat in CATS:
        i = r
        r = put(ws, r, [cat,
                        f'=COUNTIFS($A$30:$A$9999,$A{i})',
                        f'=COUNTIFS($A$30:$A$9999,$A{i},$C$30:$C$9999,'
                        f'"<>0")',
                        f'=SUMIFS({TK},{TF},$A{i})'],
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
    band(ws, 29, ['Category', 'Item', 'Balance', 'Lines',
                  'Item key, used by the formulas',
                  'Where it comes from'], None, freeze=False)
    # The first Transactions row each item appears on. Reading the item
    # from there means the text on both sides of the SUMIFS is the same
    # cell, so nothing can drift between them.
    firstrow = {}
    for n, x in enumerate(wip, t1):
        firstrow.setdefault(xl_item(x['desc']), n)
    r = 30
    for it in sorted(items, key=lambda y: items[y]):
        i = r
        n = firstrow[it]
        r = put(ws, r, [f"='Transactions'!$F{n}", f"='Transactions'!$E{n}",
                        f'=SUMIFS({TK},{TW},$E{i})',
                        f'=COUNTIFS({TW},$E{i})',
                        f"='Transactions'!$W{n}",
                        f'Transactions row {n}'],
                [None, None, MONEY, '0', None, None])
        ws.cell(i, 5).font = Font(italic=True, size=8)
        ws.cell(i, 6).font = Font(italic=True, size=8)

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
            f'=IF(COUNTIF({JA},$A{i})>0,"yes","no")',
            f'=IFERROR(LOOKUP(2,1/({JA}=$A{i}),{JD}),"")',
            f'=IFERROR(INDEX({JC},MATCH($A{i},{JA},0)),"")',
            f'=IFERROR(LOOKUP(2,1/({JA}=$A{i}),{JC}),"")',
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

    # ------------------------------------------------------ Job life cycle
    # For each job, every event in date order: when cost or revenue goes
    # into WIP, when it comes back out, what hits the P&L, and which
    # invoice. The running balance is the WIP balance on that job.
    ws = s_life
    ws['A1'] = 'The workings, job by job, in date order'
    ws['A1'].font = TITLE
    ws['A2'] = ('Read each block downwards. You can see the day something '
                'goes into WIP, the day it comes out, what lands in the '
                'P&L, and which invoice it was.')
    ws['A3'] = ('Jobs run from 1 June 2026, because that is where the job '
                'numbers start. Jobs still holding a WIP balance come '
                'first.')
    ws['A3'].font = Font(italic=True, size=9)
    for col, w in zip('ABCDEFGHI',
                      (11, 44, 30, 22, 13, 13, 13, 15, 28)):
        ws.column_dimensions[col].width = w
    WX = "'Xero - WIP transactions'!"
    PX = "'Xero - P&L analysis'!"

    plw_refs = refs
    pl_by_job = defaultdict(list)
    for x in pl:
        j = jobkey(x['job'])
        if j:
            pl_by_job[j].append(x)

    jobs = {}
    for x in wip:
        if x['job']:
            jobs.setdefault(x['job'], x['jobname'])
    bal_of = {j: round(sum(y['debit'] - y['credit']
                           for y in wip if y['job'] == j), 2)
              for j in jobs}
    # Q1 FY27 is the quarter being reported on, so the tab leads with the
    # jobs that moved through WIP in it. The block still shows the whole
    # life of the job, because a Q1 release usually starts in June.
    Q1A, Q1B = datetime.date(2026, 7, 1), datetime.date(2026, 9, 30)
    q1_wip = {j: round(sum(y['debit'] - y['credit'] for y in wip
                           if y['job'] == j and Q1A <= y['date'] <= Q1B), 2)
              for j in jobs}
    q1_pl = {j: round(sum(x['net'] for x in pl
                          if jobkey(x['job']) == j
                          and Q1A <= x['date'] <= Q1B), 2)
             for j in jobs}
    order = [j for j in jobs if q1_wip[j] != 0]
    order.sort(key=lambda j: -abs(q1_wip[j]))
    rest = [j for j in jobs if q1_wip[j] == 0]
    rest.sort(key=lambda j: (bal_of[j] == 0, -abs(bal_of[j])))
    order += rest

    def is_income(acct_or_code):
        return acct_or_code < 50000

    def wip_event(x, side):
        """Plain English for a WIP line, given what the other side was.

        A debit to WIP is money sitting in WIP. Whether that is revenue
        going into the P&L or cost coming out of it depends on which
        account the other side hit, so the wording says both.
        """
        into = x['debit'] > 0
        rev = kind(x['desc']) == 'Reversal'
        if side == 'Income':
            what = ('Revenue into the P&L, the debit sits in WIP' if into
                    else 'Revenue out of the P&L, the credit sits in WIP')
        else:
            what = ('Cost out of the P&L, the debit sits in WIP' if into
                    else 'Cost into the P&L, the credit sits in WIP')
        return ('Reversal journal. ' + what) if rev else what

    # what the other side of each WIP journal reference was, per job
    side_of = {}
    for x in pl:
        if x['ref'] in plw_refs and jobkey(x['job']):
            k = (x['ref'], jobkey(x['job']),
                 round(abs(x['net']), 2))
            side_of[k] = ('Income' if is_income(x['code']) else 'Cost')

    q1win_t = (f'{TB},">="&DATE(2026,7,1),'
               f'{TB},"<="&DATE(2026,9,1)')
    contents = []
    r = 5
    for j in order:
        start = r
        invs = inv_by_job.get(j, {})
        dates = sorted(invs.values())
        contents.append((j, jobs[j], q1_wip[j], q1_pl[j], bal_of[j],
                         'yes' if invs else 'no', start))
        for c in range(1, 10):
            ws.cell(r, c).fill = HFILL
            ws.cell(r, c).border = THIN
        ws.cell(r, 1, f'JOB {j}   {jobs[j]}').font = Font(
            bold=True, color='FFFFFF', size=11)
        tag = ('Moved through WIP this quarter' if q1_wip[j]
               else 'No WIP movement this quarter')
        ws.cell(r, 7, tag).font = Font(bold=True, color='FFFFFF', size=9)
        r += 1
        for c, lab in enumerate(
                ['Date', 'What happened', 'Account', 'Invoice',
                 'Into WIP', 'Out of WIP', 'Into the P&L',
                 'WIP balance', 'Where it comes from'], 1):
            cell = ws.cell(r, c, lab)
            cell.font, cell.fill, cell.border, cell.alignment = (
                SUB, SFILL, THIN, WRAP)
        r += 1
        first = r

        events = []
        for x in wip:
            if x['job'] != j:
                continue
            amt = round(abs(x['debit'] - x['credit']), 2)
            side = side_of.get((str(x['ref']), j, amt), 'Cost')
            events.append(dict(
                date=x['date'], sort=0, what=wip_event(x, side),
                src='wip', row=x['row']))
        for x in pl_by_job.get(j, []):
            if x['ref'] in plw_refs and x['source'] == 'Manual Journal':
                what = ('The other side of that journal, in the P&L')
                inv = x['ref']
            elif x['source'].startswith('Receivable'):
                what = 'Invoiced to the customer, revenue into the P&L'
                inv = x['inv']
            elif x['source'].startswith('Payable'):
                what = 'Supplier bill, cost into the P&L'
                inv = x['inv']
            else:
                what = f'{x["source"]}, into the P&L'
                inv = x['inv'] or x['ref']
            events.append(dict(
                date=x['date'], sort=1, what=what,
                src='pl', row=x['row']))
        events.sort(key=lambda e: (e['date'], e['sort'], e['what']))

        for e in events:
            i = r
            prev = '0' if i == first else f'$H{i-1}'
            n = e['row']
            if e['src'] == 'wip':
                vals = [f'={WX}A{n}', e['what'], '11300 Work in Progress',
                        f'={WX}D{n}', f'={WX}E{n}', f'={WX}F{n}', None,
                        f'={prev}+N(E{i})-N(F{i})',
                        f'Xero - WIP transactions row {n}']
            else:
                vals = [f'={PX}D{n}', e['what'],
                        f'={PX}A{n}&" "&{PX}B{n}',
                        f'=IF({PX}I{n}="",{PX}J{n},{PX}I{n})',
                        None, None, f'={PX}G{n}-{PX}F{n}',
                        f'={prev}+N(E{i})-N(F{i})',
                        f'Xero - P&L analysis row {n}']
            r = put(ws, r, vals,
                    [DATE, None, None, None, MONEY0, MONEY0, MONEY0, MONEY,
                     None])
            ws.cell(i, 9).font = Font(italic=True, size=8)
            if e['src'] == 'wip':
                ws.cell(i, 2).fill = AMBER
            elif e['what'].startswith('Invoiced'):
                ws.cell(i, 2).fill = GREEN
        last = r - 1

        i = r
        r = put(ws, r, ['', 'Totals', '', '',
                        f'=SUM(E{first}:E{last})',
                        f'=SUM(F{first}:F{last})',
                        f'=SUM(G{first}:G{last})',
                        f'=E{i}-F{i}', ''],
                [None, None, None, None, MONEY, MONEY, MONEY, MONEY,
                 None])
        bold(ws, i, range(1, 10))
        tot = i
        i = r
        r = put(ws, r, ['', 'Left in WIP at 31 October', '', '', '', '',
                        '', f'=H{tot}', ''],
                [None, None, None, None, None, None, None, MONEY, None])
        ws.cell(i, 3, f'=IF(ROUND(H{i},2)=0,"nothing left",'
                      f'IF(H{i}<0,"credit","debit"))')
        bold(ws, i, range(1, 10))
        i = r
        r = put(ws, r, ['', 'Margin on the job so far', '', '', '', '',
                        f'=G{tot}', '', ''],
                [None, None, None, None, None, None, MONEY, None, None])
        bold(ws, i, range(1, 10))
        # Both Q1 figures are summed from this job's own rows above, so
        # the working is on screen and nothing depends on matching the
        # job number a second time.
        win = (f'$A${first}:$A${last},">="&DATE(2026,7,1),'
               f'$A${first}:$A${last},"<="&DATE(2026,9,30)')
        i = r
        r = put(ws, r, ['', 'Moved through WIP in Q1, July to September',
                        '', '', '', '', '',
                        f'=SUMIFS($E${first}:$E${last},{win})'
                        f'-SUMIFS($F${first}:$F${last},{win})', ''],
                [None, None, None, None, None, None, None, MONEY, None])
        bold(ws, i, range(1, 10))
        i = r
        r = put(ws, r, ['', 'Hit the P&L in Q1, July to September', '', '',
                        '', '', f'=SUMIFS($G${first}:$G${last},{win})',
                        '', ''],
                [None, None, None, None, None, None, MONEY, None, None])
        bold(ws, i, range(1, 10))

        days = (datetime.date(2026, 10, 31) - dates[-1]).days if dates else 0
        if round(bal_of[j], 2) == 0:
            note = 'Nothing left in WIP. It went in and it came out.'
            fill = GREEN
        elif not invs:
            note = ('Not invoiced yet, so cost sitting in WIP is correct.')
            fill = GREEN
        elif days > 60:
            note = (f'Invoiced on {dates[-1].strftime("%d/%m/%Y")}, which '
                    f'is {days} days before 31 October, and it is still '
                    f'holding a balance. Look at this one.')
            fill = RED
        else:
            note = (f'Invoiced on {dates[-1].strftime("%d/%m/%Y")}. Still '
                    f'holding a balance. Check it clears next month.')
            fill = AMBER
        cell = ws.cell(r, 1, note)
        cell.alignment = WRAP
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=9)
        for c in range(1, 10):
            ws.cell(r, c).fill = fill
        r += 2

    ws2 = wb.create_sheet('Contents', 2)
    ws2['A1'] = 'Jobs on the Job life cycle tab'
    ws2['A1'].font = TITLE
    ws2['A2'] = ('Go to the row shown. Jobs still holding a WIP balance '
                 'are listed first.')
    band(ws2, 4, ['Job number', 'Job', 'Moved through WIP in Q1',
                  'Hit the P&L in Q1', 'Left in WIP at 31 Oct', 'Invoiced',
                  'Row on Job life cycle'], [12, 44, 18, 16, 18, 10, 20])
    rr = 5
    LC = "'Job life cycle'!"
    for jn, name, q1w, q1p, bal, inv_, row in contents:
        i = rr
        rr = put(ws2, rr, [
            jn, name,
            f'=SUMIFS({TI},{TL},$A{i},{q1win_t})'
            f'-SUMIFS({TJ},{TL},$A{i},{q1win_t})',
            f'=SUMIFS({KL},{KD},$A{i},{KH},">="&DATE(2026,7,1),'
            f'{KH},"<="&DATE(2026,9,30))',
            f'=SUMIFS({TK},{TL},$A{i})',
            f'=IF(COUNTIF({JA},$A{i})>0,"yes","no")',
            f'=MATCH("JOB "&$A{i}&" *",{LC}$A:$A,0)'],
            [None, None, MONEY, MONEY, MONEY, None, '0'])
        if q1w:
            ws2.cell(i, 3).fill = AMBER
    rr += 1
    last = rr - 2
    ws2.cell(rr, 2, 'Total for the jobs listed above').font = SUB
    for c in (3, 4, 5):
        L = get_column_letter(c)
        ws2.cell(rr, c, f'=SUM({L}5:{L}{last})').number_format = MONEY
    bold(ws2, rr, range(1, 8))
    listed = rr
    q1win = (f'{TB},">="&DATE(2026,7,1),{TB},"<="&DATE(2026,9,1)')
    rr += 1
    ws2.cell(rr, 2, 'Lines with no job number').font = SUB
    ws2.cell(rr, 3,
             f'=SUMIFS({TK},{q1win})-C{listed}').number_format = MONEY
    ws2.cell(rr, 5, f'=SUM({TK})-E{listed}').number_format = MONEY
    bold(ws2, rr, range(1, 8))
    rr += 1
    ws2.cell(rr, 2, 'Whole account, ties to the Monthly recon').font = SUB
    ws2.cell(rr, 3, f'=SUMIFS({TK},{q1win})').number_format = MONEY
    ws2.cell(rr, 5, f'=SUM({TK})').number_format = MONEY
    bold(ws2, rr, range(1, 8))
    rr += 2
    ws2.cell(rr, 2, 'Moved through WIP in Q1 is the WIP side. Hit the P&L '
                    'in Q1 is everything that job put through the profit '
                    'and loss in the quarter, invoices and bills '
                    'included.').alignment = WRAP
    ws2.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=7)

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
        first_m = (f'=DATE(YEAR(MIN({OA})),MONTH(MIN({OA})),1)' if c == 3
                   else f'=EDATE({get_column_letter(c-1)}$5,1)')
        cell = ws.cell(5, c, first_m)
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
    for n, m in enumerate(months):
        i = r
        month = (f"='Monthly recon'!$A${first_data}" if n == 0
                 else f'=EDATE($A{i-1},1)')
        r = put(ws, r, [month,
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

    # ------------------------------------------- the Xero reports, verbatim
    paste_source(wb, 'Xero - WIP transactions', wip_path)
    paste_source(wb, 'Xero - GL detail', gl_path)
    paste_source(wb, 'Xero - P&L analysis', pl_path)

    wb.save(out)
    print(f'  period           {period}')
    print(f'  WIP lines        {len(wip)}')
    print(f'  other side lines {len(other)}')
    print(f'  items            {len(items)}')
    print(f'  opening {opening:,.2f}   closing {closing:,.2f}')
    print(f'  written          {out}')


if __name__ == '__main__':
    main(*sys.argv[1:5])
