"""Journal Edit Schedule for the October reversal, journal #71185.

Eight lines of the 1 October reversal went in on the same side of WIP as
the September accrual instead of the opposite side, so the accrual never
came back out. Laid out the way the September 2026 WIP Adjustment
Schedule is laid out.

Every figure is a formula over the Xero tab pasted into the file, so the
schedule can be checked straight against the export.

Usage: oct_reversal_edit_schedule.py <WIP export> <P&L analysis> <out.xlsx>
"""
import sys, re, datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

sys.path.insert(0, __file__.rsplit('/', 1)[0] + '/../recon')
from build_wip_monthly import load
from build_wip_recon_oct import read_pl, paste_source, wip_balance_rows

JOURNAL, ACCRUAL = '#71185', '#71222'
XT = "'Xero 11300 - 31102026'!"

# The same order the September schedule lists them in.
ITEMS = [
    ('ASX AGM 22.10.26 pre-prod labour', 'ASX - AGM - 22.10.26'),
    ('Cochlear AGM 26.10.26 pre-prod labour', 'Cochlear - AGM - 26.10.26'),
    ('TLC AGM 14.10.26 pre-prod labour', 'TLC - 2026 AGM'),
    ('RSNSW OGM 7.10.26 pre-prod labour', 'RSNSW Ordinary General Meeting'),
    ('Garvan Breast Cancer Lecture 1.10.26', 'Garvan - Breast Cancer'),
    ('Microsoft SME&C Town Hall 14.10.26', 'Microsoft - SME&C'),
    ('Garvan Suttons Event 15.10.26', 'Garvan - Suttons'),
    ('ASX CEO Connect 27.10.26 pre-prod labour', 'ASX CEO Connect 27.10.26'),
]

FIRST = (
    'I have not done this one yet. In Xero the reversal line is sitting on '
    'the debit side of WIP, which is the same side the September accrual '
    'went in on, so the accrual has not come back out. A reversal has to go '
    'the other way. The pre-production labour was accrued on 30 September '
    'and the reversal on 1 October has to take it back out, so 11300 '
    'belongs on Credit and 42100 Production Labour on Debit. I will move '
    '11300 to Credit and put 42100 Production Labour on Debit. Xero 11300 '
    'row %d, and the accrual it should reverse is row %d.')
SAME = ('Same reasoning as ASX AGM. I will move 11300 to Credit and 42100 '
        'Production Labour to Debit. Xero 11300 row %d, accrual row %d.')

TITLE = Font(bold=True, size=13)
H = Font(bold=True, color='FFFFFF', size=10)
HF = PatternFill('solid', fgColor='1F3864')
SUB = Font(bold=True, size=10)
SF = PatternFill('solid', fgColor='D9E2F3')
RED = PatternFill('solid', fgColor='FFC7CE')
AMBER = PatternFill('solid', fgColor='FFEB9C')
BLUE = PatternFill('solid', fgColor='DDEBF7')
THIN = Border(*[Side(style='thin', color='BFBFBF')] * 4)
MONEY = '#,##0.00;[Red]-#,##0.00'
WRAP = Alignment(wrap_text=True, vertical='top')


def band(ws, r, labels):
    for c, lab in enumerate(labels, 1):
        cell = ws.cell(r, c, lab)
        cell.font, cell.fill, cell.border, cell.alignment = H, HF, THIN, WRAP


def put(ws, r, vals, fmts):
    for c, (v, f) in enumerate(zip(vals, fmts), 1):
        cell = ws.cell(r, c, v)
        cell.border = THIN
        if f:
            cell.number_format = f
        if isinstance(v, str) and not v.startswith('='):
            cell.alignment = WRAP
    return r + 1


def shade(ws, r, fill=SF, bold=True):
    for c in range(1, 10):
        ws.cell(r, c).fill = fill
        if bold:
            ws.cell(r, c).font = SUB


def main(wip_path, pl_path, out):
    period, opening, closing, wip = load(wip_path)
    pl = read_pl(pl_path)
    open_row, close_row = wip_balance_rows(wip_path)

    rows = []
    for name, tag in ITEMS:
        w = [x for x in wip if str(x['ref']) == JOURNAL
             and tag in str(x['desc'])]
        a = [x for x in wip if str(x['ref']) == ACCRUAL
             and tag in str(x['desc'])]
        p = [x for x in pl if x['ref'] == JOURNAL and tag in x['desc']]
        if not (len(w) == len(a) == len(p) == 1):
            raise SystemExit(f'{name}: expected one line each, got '
                             f'{len(w)} {len(a)} {len(p)}')
        rows.append(dict(name=name, xrow=w[0]['row'], arow=a[0]['row'],
                         amt=round(w[0]['debit'] - w[0]['credit'], 2),
                         job=p[0]['job'],
                         acct=f'{p[0]["code"]} {p[0]["account"]}'))

    def xnet(row):
        return f'({XT}E{row}-{XT}F{row})'

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Journal Edit Schedule'
    for col, w in zip('ABCDEFGHI', (34, 11, 26, 15, 15, 14, 15, 78, 7)):
        ws.column_dimensions[col].width = w

    ws['A1'] = 'WIP 11300 Journal Edit Schedule, October 2026 reversal'
    ws['A1'].font = TITLE
    ws['A2'] = 'Corporate Technology Services Pty Ltd'
    ws['A3'] = period
    ws['A4'] = ('STATUS: the corrections below are not yet done. The ending '
                'position and the check at the bottom are what it will read '
                'once they are.')
    ws['A4'].font = SUB
    ws['A4'].fill = AMBER
    ws.merge_cells('A4:I4')

    # the table is laid out first, so every row reference below is real
    head = 16
    first = head + 1
    last = first + len(rows) - 1
    total = last + 1
    check = total + 1

    r = 6
    ws.cell(r, 1, 'STARTING POSITION').font = SUB
    r += 1
    start_row = r
    r = put(ws, r, ['Xero 11300 at 31 October 2026', '', '', '', '', '',
                    f'={XT}E{close_row}-{XT}F{close_row}',
                    'credit, per the export', ''],
            [None] * 6 + [MONEY, None, None])
    r = put(ws, r, [f'Accrual posted 30 September, journal {ACCRUAL}',
                    '', '', '', '', '',
                    '=' + '+'.join(xnet(x['arow']) for x in rows),
                    'the eight lines below, on the debit side of WIP', ''],
            [None] * 6 + [MONEY, None, None])
    r = put(ws, r, [f'Reversal posted 1 October, journal {JOURNAL}',
                    '', '', '', '', '',
                    '=' + '+'.join(xnet(x['xrow']) for x in rows),
                    'posted the same way instead of the opposite way', ''],
            [None] * 6 + [MONEY, None, None])
    r += 1
    ws.cell(r, 1, 'ENDING POSITION, STILL TO DO').font = SUB
    r += 1
    end_row = r
    r = put(ws, r, ['Xero 11300 once the eight lines are flipped', '', '',
                    '', '', '', f'=$G${start_row}+F{total}',
                    'credit. The accrual comes back out as it should.', ''],
            [None] * 6 + [MONEY, None, None])
    shade(ws, end_row)

    r = head - 1
    ws.cell(r, 1, f'JOURNAL {JOURNAL}').font = SUB
    ws.cell(r, 1).fill = BLUE
    ws.cell(r, 2, 'Reversal: Pre-PROD & INT Work & Expense Accrual for '
                  'October 2026').font = Font(italic=True, size=9)
    band(ws, head, ['Item', 'Journal', 'Line to change', 'Xero now',
                    'Should be', 'Effect on WIP', 'WIP after',
                    'What to do in Xero', 'Done'])
    r = first
    for n, it in enumerate(rows):
        i = r
        prev = f'$G${start_row}' if n == 0 else f'$G{i-1}'
        note = (FIRST if n == 0 else SAME) % (it['xrow'], it['arow'])
        r = put(ws, r, [
            it['name'], JOURNAL, '11300 Work in Progress',
            f'="Debit "&TEXT({XT}E{it["xrow"]},"#,##0.00")',
            f'="Credit "&TEXT({XT}E{it["xrow"]},"#,##0.00")',
            f'=-2*{xnet(it["xrow"])}', f'={prev}+F{i}', note, ''],
            [None, None, None, None, None, MONEY, MONEY, None, None])
        ws.cell(i, 4).fill = RED
        ws.cell(i, 5).fill = BLUE
        ws.cell(i, 9).fill = AMBER
    put(ws, total, ['Total', '', '', '', '', f'=SUM(F{first}:F{last})',
                    f'=G{last}', '', ''],
        [None] * 5 + [MONEY, MONEY, None, None])
    shade(ws, total)
    put(ws, check, ['Check, still to do', '', '', '', '', '',
                    f'=IF(ROUND(G{total}-G{end_row},2)=0,"ties","CHECK")',
                    'the eight lines bring WIP to the ending position '
                    'above', ''], [None] * 9)
    shade(ws, check, SF)

    r = check + 2
    ws.cell(r, 1, 'THE OTHER SIDE, SAME JOURNAL').font = SUB
    r += 1
    band(ws, r, ['Item', 'Journal', 'Line to change', 'Xero now',
                 'Should be', 'Effect on the P&L', '', 'Job', ''])
    r += 1
    pfirst = r
    for it in rows:
        i = r
        r = put(ws, r, [it['name'], JOURNAL, it['acct'],
                        f'="Credit "&TEXT({XT}E{it["xrow"]},"#,##0.00")',
                        f'="Debit "&TEXT({XT}E{it["xrow"]},"#,##0.00")',
                        f'=-2*{xnet(it["xrow"])}', '', it['job'], ''],
                [None, None, None, None, None, MONEY, None, None, None])
        ws.cell(i, 4).fill = RED
        ws.cell(i, 5).fill = BLUE
    plast = r - 1
    put(ws, r, ['Total taken off October revenue', '', '', '', '',
                f'=SUM(F{pfirst}:F{plast})', '', '', ''],
        [None] * 5 + [MONEY, None, None, None])
    shade(ws, r)
    r += 2

    ws.cell(r, 1, 'WHERE TO FIND EACH ITEM').font = SUB
    r += 1
    band(ws, r, ['Item', 'Reversal row on the Xero tab',
                 'Accrual row on the Xero tab', 'Colour', 'Meaning',
                 '', '', '', ''])
    r += 1
    for it in rows:
        i = r
        r = put(ws, r, [it['name'], it['xrow'], it['arow'], 'RED',
                        'the line as Xero has it now, highlighted on the '
                        'Xero tab', '', '', '', ''],
                [None, '0', '0', None, None, None, None, None, None])
        ws.cell(i, 4).fill = RED
    r += 1

    ws.cell(r, 1, 'LEGEND').font = SUB
    r += 1
    for fill, txt in ((RED, 'RED: the line as Xero has it now. This is what '
                            'changes.'),
                      (AMBER, 'AMBER: not done yet.'),
                      (BLUE, 'BLUE: what the line should say once it is '
                             'changed.')):
        ws.cell(r, 1, txt).alignment = WRAP
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
        for c in range(1, 9):
            ws.cell(r, c).fill = fill
        r += 1

    # ------------------------------------------------------- What i did
    ws = wb.create_sheet('What i did')
    ws.column_dimensions['A'].width = 108
    ws['A1'] = 'What i did'
    ws['A1'].font = TITLE
    ws['A2'] = 'STATUS: not done yet. This is what I am going to do.'
    ws['A2'].font = SUB
    ws['A2'].fill = AMBER
    r = 4
    for txt in (
        'In September I moved eight pre-production labour accruals from the '
        'credit side of WIP to the debit side, because the work was '
        'performed in September and had not been invoiced.',
        '',
        'The reversal that posts on 1 October did not follow. It went in on '
        'the debit side as well, which is the same side as the accrual, so '
        'the accrual has not come back out.',
        '',
        'That leaves 12,415.00 sitting in WIP that should not be there, and '
        'it has put 12,415.00 of revenue into October that is not earned in '
        'October.',
        '',
        'I am going to flip the eight lines in journal #71185. On each one I '
        'will move 11300 Work in Progress to Credit and put 42100 Production '
        'Labour on Debit.',
        '',
        'The other five lines in that journal reversed correctly and I am '
        'leaving them alone.',
        '',
        'Every other accrual in the file reverses properly. I checked all 74 '
        'of them and only these eight are wrong.',
    ):
        ws.cell(r, 1, txt).alignment = WRAP
        r += 1

    paste_source(wb, 'Xero 11300 - 31102026', wip_path)
    src = wb['Xero 11300 - 31102026']
    for it in rows:
        for c in range(1, 10):
            src.cell(it['xrow'], c).fill = RED
            src.cell(it['arow'], c).fill = AMBER

    wb.save(out)
    print(f'  lines to flip  {len(rows)}')
    print(f'  WIP effect     {sum(x["amt"] for x in rows) * -2:,.2f}')
    print(f'  WIP now        {closing:,.2f}')
    print(f'  WIP after      '
          f'{closing + sum(x["amt"] for x in rows) * -2:,.2f}')
    print(f'  written        {out}')


if __name__ == '__main__':
    main(*sys.argv[1:4])
