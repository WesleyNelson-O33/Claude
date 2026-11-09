"""Spacera bills against the Aware Super maintenance income, job 26022301.

The cost of the Aware Super National Support Agreement is billed monthly
by Spacera. The income for the same contract is billed monthly to Aware
Super. This matches the two and shows which account each side landed in.

The income side is a formula over the P&L analysis export pasted into the
file. The bill side comes from Xero itself, because the export does not
carry the bill's own line coding, and every one of those cells is marked
as read from Xero on the date shown.

Usage: build_spacera_match.py <P&L Analysis export> <bills.json> <out.xlsx>
"""
import sys, re, datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

sys.path.insert(0, __file__.rsplit('/', 1)[0])
from build_wip_recon_oct import read_pl, paste_source, jobkey

JOB = None          # all four set from the bills file
PX = "'Xero - P&L analysis'!"
READ_ON = None
BILLS = []
LINES = []


def load_bills(path):
    """The bill coding, read from Xero and kept in a file of its own.

    The export does not carry the account each bill line was coded to, so
    that has to come from Xero. It lives outside this script because it is
    supplier and contract detail, not code.
    """
    import json
    d = json.load(open(path))
    bills = [(b[0], datetime.date(*map(int, b[1].split('-'))), b[2], b[3],
              b[4]) for b in d['bills']]
    lines = [tuple(x) for x in d['lines']]
    return d['job'], d['read_on'], bills, lines


TITLE = Font(bold=True, size=13)
H = Font(bold=True, color='FFFFFF', size=10)
HF = PatternFill('solid', fgColor='1F3864')
SUB = Font(bold=True, size=10)
SF = PatternFill('solid', fgColor='D9E2F3')
RED = PatternFill('solid', fgColor='FFC7CE')
AMBER = PatternFill('solid', fgColor='FFEB9C')
XERO = PatternFill('solid', fgColor='FFF2CC')
THIN = Border(*[Side(style='thin', color='BFBFBF')] * 4)
MONEY = '#,##0.00;[Red]-#,##0.00'
DATE = 'dd/mm/yyyy'
WRAP = Alignment(wrap_text=True, vertical='top')


def band(ws, r, labels, widths=None):
    for c, lab in enumerate(labels, 1):
        cell = ws.cell(r, c, lab)
        cell.font, cell.fill, cell.border, cell.alignment = H, HF, THIN, WRAP
    if widths:
        for c, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(c)].width = w


def put(ws, r, vals, fmts):
    for c, (v, f) in enumerate(zip(vals, fmts), 1):
        cell = ws.cell(r, c, v)
        cell.border = THIN
        if f:
            cell.number_format = f
        if isinstance(v, str) and not v.startswith('='):
            cell.alignment = WRAP
    return r + 1


def main(pl_path, bills_path, out):
    global JOB, READ_ON, BILLS, LINES
    JOB, READ_ON, BILLS, LINES = load_bills(bills_path)
    pl = read_pl(pl_path)
    job = [x for x in pl if jobkey(x['job']) == JOB]
    income = sorted((x for x in job if x['code'] < 50000
                     and x['source'].startswith('Receivable')),
                    key=lambda y: y['date'])

    wb = openpyxl.Workbook()

    # ------------------------------------------------- read this first
    ws = wb.active
    ws.title = 'Read this first'
    ws.column_dimensions['A'].width = 100
    ws['A1'] = 'Spacera against Aware Super, job 26022301'
    ws['A1'].font = TITLE
    r = 3
    for txt, bold in (
        ('There is no 2024 history. The contract started on 1 July 2026 '
         'and only four Spacera bills exist.', True),
        ('', False),
        ('The contract is the Aware Super National Support Agreement, '
         'twelve months, quote QU-0003 against PO-00016885. Spacera bill '
         '5,208.00 a month and CTS bill Aware Super about 11,000.00 a '
         'month for the same job.', False),
        ('', False),
        ('The cost has been coded to two different accounts.', True),
        ('July went to 51310 Sub-Contract Labour. August, September and '
         'October went to 51500 Subscriptions and Licences.', False),
        ('', False),
        ('July and August were keyed by hand. September and October came '
         'from a repeating bill, and that template is set to 51500. So '
         'every bill from here on will go to 51500 unless the template is '
         'changed.', False),
        ('', False),
        ('Why it matters', True),
        ('The income for this contract sits on 41100 Contract Support '
         'Staff. So every month the cost goes to 51500, the licence '
         'accounts carry cost with no matching income, and the licence '
         'margin reads lower than it is.', False),
        ('', False),
        ('In Q1 that is 10,416.00, being August and September. July is not '
         'in it because it went to 51310.', False),
        ('', False),
        ('The question to settle', True),
        ('Which account is right for this cost. Once that is decided, fix '
         'the repeating bill template so it stays consistent, and decide '
         'whether July or August and September need moving.', False),
        ('', False),
        ('Where the figures come from', True),
        ('The income is a formula over the P&L analysis export pasted into '
         'this file. The bill coding is read from Xero itself on '
         f'{READ_ON}, because the export does not carry the coding on the '
         'bill. Those cells are shaded cream.', False),
    ):
        c = ws.cell(r, 1, txt)
        c.alignment = WRAP
        if bold:
            c.font = SUB
        r += 1

    # ------------------------------------------------ bills vs income
    ws = wb.create_sheet('Bills against income')
    ws['A1'] = 'Month by month, the Spacera bill against the Aware Super invoice'
    ws['A1'].font = TITLE
    ws['A2'] = (f'Cream cells are read from Xero on {READ_ON}. Everything '
                f'else is a formula over the pasted export.')
    ws['A2'].font = Font(italic=True, size=9)
    band(ws, 4, ['Month', 'Spacera bill', 'Bill date', 'Cost',
                 'Coded to', 'Account name', 'How it was entered',
                 'Aware Super invoice', 'Invoice date', 'Income',
                 'Margin', 'Where the income comes from'],
         [11, 12, 11, 12, 10, 32, 22, 16, 11, 12, 12, 26])
    r = 5
    for n, (num, bdate, code, acct, how) in enumerate(BILLS):
        i = r
        inc = income[n] if n < len(income) else None
        r = put(ws, r, [
            bdate.strftime('%b %Y'), num, bdate, 5208.00, code, acct, how,
            inc['inv'] if inc else None,
            f'={PX}D{inc["row"]}' if inc else None,
            f'={PX}G{inc["row"]}-{PX}F{inc["row"]}' if inc else None,
            f'=IF(J{i}="","",J{i}-D{i})',
            f'Xero - P&L analysis row {inc["row"]}' if inc else None],
            [None, None, DATE, MONEY, None, None, None, None, DATE, MONEY,
             MONEY, None])
        for c in (2, 3, 4, 5, 6, 7):
            ws.cell(i, c).fill = XERO
        ws.cell(i, 5).fill = RED if code == '51310' else AMBER
        ws.cell(i, 12).font = Font(italic=True, size=8)
    last = r - 1
    i = r
    r = put(ws, r, ['Total', '', '', f'=SUM(D5:D{last})', '', '', '', '',
                    '', f'=SUM(J5:J{last})', f'=SUM(K5:K{last})', ''],
            [None, None, None, MONEY, None, None, None, None, None, MONEY,
             MONEY, None])
    for c in range(1, 13):
        ws.cell(i, c).fill = SF
        ws.cell(i, c).font = SUB
    r += 1
    i = r
    r = put(ws, r, ['In Q1 only, July to September', '', '',
                    f'=SUM(D5:D7)', '', '', '', '', '', f'=SUM(J5:J7)',
                    f'=SUM(K5:K7)', ''],
            [None, None, None, MONEY, None, None, None, None, None, MONEY,
             MONEY, None])
    for c in range(1, 13):
        ws.cell(i, c).font = SUB
    r += 1
    i = r
    r = put(ws, r, ['Of that, sitting in 51500', '', '', f'=D6+D7', '', '',
                    'August and September only', '', '', '', '', ''],
            [None, None, None, MONEY, None, None, None, None, None, None,
             None, None])
    for c in range(1, 13):
        ws.cell(i, c).fill = RED
        ws.cell(i, c).font = SUB

    # ----------------------------------------------- the bill line items
    ws = wb.create_sheet('Spacera bill lines')
    ws['A1'] = 'What is on each Spacera bill'
    ws['A1'].font = TITLE
    ws['A2'] = (f'Read from Xero on {READ_ON}. Every bill carries the same '
                f'three lines.')
    ws['A2'].font = Font(italic=True, size=9)
    band(ws, 4, ['Bill', 'Bill date', 'Description', 'Quantity',
                 'Unit price', 'Amount', 'Coded to'],
         [12, 11, 44, 10, 12, 12, 10])
    r = 5
    for num, bdate, code, acct, how in BILLS:
        for desc, qty, unit, amt in LINES:
            i = r
            r = put(ws, r, [num, bdate, desc, qty, unit, amt,
                            code if amt else ''],
                    [None, DATE, None, '0', MONEY, MONEY, None])
            for c in range(1, 8):
                ws.cell(i, c).fill = XERO

    # ------------------------------------------------ every job line
    ws = wb.create_sheet('Job 26022301')
    ws['A1'] = 'Every line on job 26022301 in the export'
    ws['A1'].font = TITLE
    band(ws, 3, ['Date', 'Account code', 'Account', 'Source', 'Contact',
                 'Invoice', 'Amount', 'Where it comes from'],
         [11, 12, 32, 18, 26, 20, 13, 26])
    r = 4
    for x in sorted(job, key=lambda y: (y['date'], y['code'])):
        i = r
        r = put(ws, r, [f'={PX}D{x["row"]}', f'={PX}A{x["row"]}',
                        f'={PX}B{x["row"]}', f'={PX}C{x["row"]}',
                        f'={PX}E{x["row"]}', f'={PX}I{x["row"]}',
                        f'={PX}G{x["row"]}-{PX}F{x["row"]}',
                        f'Xero - P&L analysis row {x["row"]}'],
                [DATE, None, None, None, None, None, MONEY, None])
        ws.cell(i, 8).font = Font(italic=True, size=8)
    jl = r - 1
    i = r
    put(ws, r, ['Total', '', '', '', '', '', f'=SUM(G4:G{jl})', ''],
        [None, None, None, None, None, None, MONEY, None])
    for c in range(1, 9):
        ws.cell(i, c).fill = SF
        ws.cell(i, c).font = SUB

    paste_source(wb, 'Xero - P&L analysis', pl_path)

    wb.save(out)
    print(f'  Spacera bills      {len(BILLS)}')
    print(f'  Aware Super invoices {len(income)}')
    print(f'  job lines          {len(job)}')
    print(f'  written            {out}')


if __name__ == '__main__':
    main(*sys.argv[1:4])
