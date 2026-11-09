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
              b[4], b[5] if len(b) > 5 else b[2]) for b in d['bills']]
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
    ws['A3'] = ('STATUS: the correction below is not yet done. The Coded '
                'to column is what Xero has now.')
    ws['A3'].font = SUB
    ws['A3'].fill = AMBER
    band(ws, 5, ['Month', 'Spacera bill', 'Bill date', 'Cost',
                 'Coded to now', 'Should be', 'How it was entered',
                 'Done', 'Aware Super invoice', 'Invoice date', 'Income',
                 'Margin', 'Where the income comes from'],
         [11, 12, 11, 12, 13, 11, 22, 7, 16, 11, 12, 12, 26])
    r = 6
    for n, (num, bdate, code, acct, how, should) in enumerate(BILLS):
        i = r
        inc = income[n] if n < len(income) else None
        r = put(ws, r, [
            bdate.strftime('%b %Y'), num, bdate, 5208.00, code, should,
            how, '',
            inc['inv'] if inc else None,
            f'={PX}D{inc["row"]}' if inc else None,
            f'={PX}G{inc["row"]}-{PX}F{inc["row"]}' if inc else None,
            f'=IF(K{i}="","",K{i}-D{i})',
            f'Xero - P&L analysis row {inc["row"]}' if inc else None],
            [None, None, DATE, MONEY, None, None, None, None, None, DATE,
             MONEY, MONEY, None])
        for c in (2, 3, 4, 5, 7):
            ws.cell(i, c).fill = XERO
        if code != should:
            ws.cell(i, 5).fill = RED
            ws.cell(i, 6).fill = PatternFill('solid', fgColor='DDEBF7')
            ws.cell(i, 8).fill = AMBER
        ws.cell(i, 13).font = Font(italic=True, size=8)
    last = r - 1
    first = 6
    i = r
    r = put(ws, r, ['Total', '', '', f'=SUM(D{first}:D{last})', '', '', '',
                    '', '', '', f'=SUM(K{first}:K{last})',
                    f'=SUM(L{first}:L{last})', ''],
            [None, None, None, MONEY, None, None, None, None, None, None,
             MONEY, MONEY, None])
    for c in range(1, 14):
        ws.cell(i, c).fill = SF
        ws.cell(i, c).font = SUB
    r += 1
    i = r
    r = put(ws, r, ['In Q1, July to September', '', '',
                    f'=SUM(D{first}:D{first+2})', '', '', '', '', '', '',
                    f'=SUM(K{first}:K{first+2})',
                    f'=SUM(L{first}:L{first+2})', ''],
            [None, None, None, MONEY, None, None, None, None, None, None,
             MONEY, MONEY, None])
    for c in range(1, 14):
        ws.cell(i, c).font = SUB
    r += 1
    i = r
    r = put(ws, r, ['In 51500 as Xero has it now', '', '',
                    f'=D{first+1}+D{first+2}', '', '',
                    'August and September', '', '', '', '', '', ''],
            [None, None, None, MONEY, None, None, None, None, None, None,
             None, None, None])
    for c in range(1, 14):
        ws.cell(i, c).fill = RED
        ws.cell(i, c).font = SUB
    r += 1
    i = r
    r = put(ws, r, ['In 51500 once July is corrected', '', '',
                    f'=SUM(D{first}:D{first+2})', '', '',
                    'July, August and September', '', '', '', '', '', ''],
            [None, None, None, MONEY, None, None, None, None, None, None,
             None, None, None])
    for c in range(1, 14):
        ws.cell(i, c).fill = PatternFill('solid', fgColor='DDEBF7')
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
    for num, bdate, code, acct, how, should in BILLS:
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

    # ----------------------------------- effect on the licence accounts
    ws = wb.create_sheet('Effect on the licence accounts')
    for col, w in zip('ABCD', (52, 16, 16, 40)):
        ws.column_dimensions[col].width = w
    ws['A1'] = 'What the correction does to the licence accounts, Q1 FY27'
    ws['A1'].font = TITLE
    ws['A2'] = ('STATUS: not done yet. The after column is what it will '
                'read once July is moved to 51500.')
    ws['A2'].font = SUB
    ws['A2'].fill = AMBER
    Q = (f'{PX}$D:$D,">="&DATE(2026,7,1),{PX}$D:$D,"<="&DATE(2026,9,30)')
    band(ws, 4, ['', 'As Xero has it now', 'Once July is moved',
                 'Note'])
    jul = 5208.00
    rows_ = [
        ('Licence income, 41175',
         f'=SUMIFS({PX}$G:$G,{PX}$A:$A,"41175",{Q})'
         f'-SUMIFS({PX}$F:$F,{PX}$A:$A,"41175",{Q})',
         '=B5', 'unchanged'),
        ('Licence cost, 51500',
         f'=SUMIFS({PX}$G:$G,{PX}$A:$A,"51500",{Q})'
         f'-SUMIFS({PX}$F:$F,{PX}$A:$A,"51500",{Q})',
         f'=B6-{jul}', 'July Spacera comes in'),
        ('Margin', '=B5+B6', '=C5+C6', ''),
        ('Margin per cent', '=IFERROR(B7/B5,"")',
         '=IFERROR(C7/C5,"")', ''),
    ]
    r = 5
    for lab, b, c, note in rows_:
        i = r
        r = put(ws, r, [lab, b, c, note],
                [None,
                 '0.0%;[Red]-0.0%' if 'per cent' in lab else MONEY,
                 '0.0%;[Red]-0.0%' if 'per cent' in lab else MONEY,
                 None])
        if lab == 'Margin' or 'per cent' in lab:
            for cc in range(1, 5):
                ws.cell(i, cc).fill = SF
                ws.cell(i, cc).font = SUB
    r += 1
    ws.cell(r, 1, 'Spacera in 51500 in Q1').font = SUB
    ws.cell(r, 2, f'=D{9}+D{10}').number_format = MONEY
    ws.cell(r, 2, "='Bills against income'!D7+'Bills against income'!D8")
    ws.cell(r, 2).number_format = MONEY
    ws.cell(r, 3, "=SUM('Bills against income'!D6:D8)")
    ws.cell(r, 3).number_format = MONEY
    ws.cell(r, 4, 'August and September now, all three months after')
    r += 2
    for txt in ('Correcting July makes the licence margin worse, not '
                'better. That is expected. It puts the whole quarter of '
                'Aware Super maintenance cost into the licence account.',
                '',
                'The mismatch itself does not go away. The cost for this '
                'contract sits in Subscriptions and Licences while the '
                'income sits in Contract Support Staff. Correcting July '
                'only makes the coding consistent.'):
        c = ws.cell(r, 1, txt)
        c.alignment = WRAP
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
        r += 1

    paste_source(wb, 'Xero - P&L analysis', pl_path)

    wb.save(out)
    print(f'  Spacera bills      {len(BILLS)}')
    print(f'  Aware Super invoices {len(income)}')
    print(f'  job lines          {len(job)}')
    print(f'  written            {out}')


if __name__ == '__main__':
    main(*sys.argv[1:4])
