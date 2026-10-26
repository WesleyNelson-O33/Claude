"""Forecast pack FY27: her own workbook, with July to September replaced by actuals.

Built from her file rather than rebuilt from scratch, so the raw data block on
the left of each sheet, the line labels, the layout and the row numbers all stay
exactly where she has them.

What changes:
  - July, August and September on the revenue, cost of sales and overhead lines
    are overwritten with actuals from the controller pack, which is the Xero GL.
  - Every formula that pointed at the linked workbooks (sales forecast, fixed
    costs, casual pays, FY27 budget detailed) is frozen to the value it last
    held, so the pack opens and reads correctly without those files.
  - Internal formulas are left alone, so the totals, gross profit and net profit
    still calculate.

What does not change: the raw data columns, the own OpEx block on each sheet,
and October to June on every line.
"""
import collections
import shutil
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter as gl

PACK = 'pack/CTS Financial Controller Pack - v20.xlsx'
SRC = ('/tmp/claude-0/-home-user-Claude/a00ffe59-433c-507b-9006-917d8c60083a'
       '/scratchpad/fc/fc.xlsx')
OUT = 'forecast/CTS Forecast Pack FY27 - actuals to September.xlsx'

MON3 = ('Jul', 'Aug', 'Sep')
ACT = PatternFill('solid', start_color='FFE8F4F3', end_color='FFE8F4F3')
NOTE = Font(name='Calibri', sz=10, i=True, color='FF6B7585')

# group -> (Cleanup department tags, first month column, total column)
GROUPS = {'ONS': (['ONSITE'], 6, 18),
          'PRD': (['PRODUCTION', 'VIDEO'], 8, 20),
          'CONS': (['INTEGRATION', 'CONSULTING'], 7, 19)}

EQUIP = ['Equipment - Purchase', 'Equipment - Hires', 'Service - hires',
         'Equipment -Maintenance Support', 'Stamp Duty', 'Damage Waiver',
         'Subscriptions & Licences - Expense', 'Credit Card Paid - ADMIN']
TRANSPORT = ['Direct Travel', 'Direct Freight', 'Direct Mobile']


def tag(names, suffix):
    return [f'{n} - {suffix}' for n in names]


# row on her sheet -> (kind, accounts).  None = nothing in the GL identifies it.
# 'BALANCE' = whatever is left of the category once the named lines are taken.
LINES = {
 'ONS': {
    2:  ('rev', ['Mobile Phone Plans']),
    3:  ('rev', None),
    4:  ('rev', None),
    5:  ('rev', 'BALANCE'),
    7:  ('cos', ['Direct Salaries - ONS']),
    8:  ('cos', ['Direct Superannuation - ONS']),
    9:  ('cos', None),
    10: ('cos', ['Direct Wages - ONS']),
    11: ('cos', None),
    12: ('cos', EQUIP + ['Sub-Contract Labour - ONS', 'Credit Card Paid - ONS',
                         'Supplier Per Diems - ONS']),
    13: ('cos', tag(TRANSPORT, 'ONS')),
 },
 'PRD': {
    2:  ('rev', 'BALANCE'),
    4:  ('cos', ['Direct Salaries - PRD']),
    5:  ('cos', ['Direct Salaries - VID']),
    6:  ('cos', ['Direct Superannuation - PRD']),
    7:  ('cos', ['Direct Superannuation - VID']),
    8:  ('cos', None),
    9:  ('cos', ['Direct Wages - PRD']),
    10: ('cos', ['Direct Wages - VID']),
    11: ('cos', ('PRODUCTION', EQUIP + ['Sub-Contract Labour - PRD',
                                        'Credit Card Paid - PRD',
                                        'Supplier Per Diems - PRD'])),
    12: ('cos', ('VIDEO', EQUIP + ['Sub-Contract Labour - VID',
                                   'Credit Card Paid - VID',
                                   'Supplier Per Diems - VID'])),
    13: ('cos', tag(TRANSPORT, 'PRD')),
    14: ('cos', tag(TRANSPORT, 'VID')),
 },
 'CONS': {
    2:  ('rev', 'BALANCE'),
    5:  ('cos', ['Direct Salaries - INTEGRATION', 'Direct Salaries - CONS',
                 'Direct Wages - INTEGRATION', 'Direct Wages - CONS']),
    6:  ('cos', ['Direct Superannuation-INTEGRATION',
                 'Direct Superannuation - CONS']),
    7:  ('cos', EQUIP + ['Sub-Contract Labour - INTEGRATION',
                         'Sub-Contract Labour - CONS',
                         'Credit Card Paid - INTEGRATION',
                         'Credit Card Paid - CONS',
                         'Supplier Per Diems - INTEGRATION',
                         'Supplier Per Diems - CONS']),
    8:  ('cos', None),
    9:  ('cos', tag(TRANSPORT, 'INTEGRATION') + tag(TRANSPORT, 'CONS')),
 },
}

# the four overhead lines, in sheet row order
OHROWS = {'ONS': {35: 'Business Insurances', 36: 'Employment Expenses',
                  37: 'General & Administrative Exp', 38: 'Operating Expenses'},
          'PRD': {30: 'Business Insurances', 31: 'Employment Expenses',
                  32: 'General & Administrative Exp', 33: 'Operating Expenses'},
          'CONS': {28: 'Business Insurances', 29: 'Employment Expenses',
                   30: 'General & Administrative Exp', 31: 'Operating Expenses'}}

# the total and result rows: these must keep calculating
KEEP = {'ONS': {6, 14, 23, 24, 25, 26, 27, 30, 40},
        'PRD': {3, 15, 24, 25, 26, 27, 28, 38},
        'CONS': {4, 10, 20, 21, 22, 23, 24, 26, 33, 42, 43}}

NOTE_CELL = {'ONS': 19, 'PRD': 21, 'CONS': 20}

print('reading the GL from the pack')
wb = openpyxl.load_workbook(PACK, data_only=True, read_only=True)
glr = collections.defaultdict(float)                 # (dept, mon, cat, account)
for row in wb['Cleanup'].iter_rows(min_row=5, max_row=20004, max_col=10,
                                   values_only=True):
    if row[0] in (None, ''):
        continue
    mo = row[1]
    if not hasattr(mo, 'strftime'):
        continue
    glr[(str(row[3]).strip(), mo.strftime('%b'), str(row[4]).strip(),
         str(row[8]).strip())] += row[0]
wb.close()


def gl_sum(tags, mon, cat, accounts):
    if accounts in (None, 'BALANCE'):
        return 0.0
    if isinstance(accounts, tuple):
        tags, accounts = [accounts[0]], accounts[1]
    return sum(glr.get((t, mon, cat, a), 0.0) for t in tags for a in accounts)


def gl_cat_total(tags, mon, cat):
    return sum(v for (d, m, c, _), v in glr.items()
               if d in tags and m == mon and c == cat)


# the overhead block: the pack's own split, folded into her four categories
pv = openpyxl.load_workbook(PACK, data_only=True)
osh, eng = pv['Overhead Split'], pv['Engine']
ecol = {eng.cell(4, c).value.strftime('%b'): c for c in range(5, 18)
        if hasattr(eng.cell(4, c).value, 'strftime')}
FOLD = {'Business Insurances': ['Business Insurances'],
        'Employment Expenses': ['Employment Expenses'],
        'General & Administrative Exp': ['General & Administrative Exp', 'Other',
                                         'Credit Card fees paid'],
        'Operating Expenses': ['Operating Expenses',
                               'Advertising & Promotion Exp']}
ENGOFF = {'General & Administrative Exp': 0, 'Business Insurances': 1,
          'Advertising & Promotion Exp': 2, 'Operating Expenses': 3,
          'Employment Expenses': 4, 'Credit Card fees paid': 5}
ENGBASE = {'ONS': [28], 'PRD': [59, 90], 'CONS': [121, 152]}
KEYCOL = {'ONS': [16], 'PRD': [17, 18], 'CONS': [19, 20]}


def overhead(group, mon):
    oc = {'Jul': 3, 'Aug': 4, 'Sep': 5}[mon]
    ec = ecol[mon]
    out = {}
    for lbl, subs in FOLD.items():
        t = 0.0
        for s in subs:
            if s in ENGOFF:
                t += sum(eng.cell(b + ENGOFF[s], ec).value or 0
                         for b in ENGBASE[group])
            t += sum((osh.cell(r, oc).value or 0)
                     * sum(osh.cell(r, k).value or 0 for k in KEYCOL[group])
                     for r in range(23, 87)
                     if str(osh.cell(r, 21).value or '') == s)
        out[lbl] = t
    return out


print('copying her workbook')
shutil.copyfile(SRC, OUT)
wf = openpyxl.load_workbook(OUT, data_only=False)
wv = openpyxl.load_workbook(SRC, data_only=True)

frozen = written = 0
for group, (tags, first, totcol) in GROUPS.items():
    ws, vs = wf[group], wv[group]

    # 1. freeze every linked formula to the value it last held
    for row in ws.iter_rows():
        for c in row:
            v = c.value
            if isinstance(v, str) and v.startswith('=') and '[1]' in v:
                cached = vs.cell(c.row, c.column).value
                c.value = cached if isinstance(cached, (int, float)) else 0
                frozen += 1

    # 2. freeze October to June on every detail line.
    #    Some lines copy the month to their left (PRD row 32 is =J32), so an
    #    actual written into September would otherwise drag itself across the
    #    rest of the year. Freezing holds her budget where she put it.
    for r in range(1, ws.max_row + 1):
        if r in KEEP[group]:
            continue
        for i in range(3, 12):
            c = ws.cell(r, first + i)
            if isinstance(c.value, str) and c.value.startswith('='):
                cached = vs.cell(r, first + i).value
                c.value = cached if isinstance(cached, (int, float)) else 0
                frozen += 1

    # 3. write the actual months
    for r, (kind, accounts) in LINES[group].items():
        cat = 'Income' if kind == 'rev' else 'Cost of Sales'
        for i, mon in enumerate(MON3):
            if accounts == 'BALANCE':
                named = sum(gl_sum(tags, mon, cat, a)
                            for k, a in LINES[group].values()
                            if k == kind and a not in (None, 'BALANCE'))
                val = gl_cat_total(tags, mon, cat) - named
            elif accounts is None:
                val = 0.0
            else:
                val = gl_sum(tags, mon, cat, accounts)
            cell = ws.cell(r, first + i)
            cell.value = round(val, 2)
            cell.fill = ACT
            written += 1

    for i, mon in enumerate(MON3):
        oh = overhead(group, mon)
        for r, lbl in OHROWS[group].items():
            cell = ws.cell(r, first + i)
            # her block is a positive cost, subtracted from gross profit
            cell.value = round(abs(oh[lbl]), 2)
            cell.fill = ACT
            written += 1
        ws.cell(1, first + i).fill = ACT

    # 4. the note
    n = ws.cell(1, NOTE_CELL[group])
    n.value = ('July, August and September are actual, from the controller pack. '
               'October to June are the budget as it stood. The own OpEx block is '
               'budget in every month; the GL does not map to those line labels. '
               'It does not feed gross profit or net profit.')
    n.font = NOTE

print(f'  linked formulas frozen: {frozen}')
print(f'  actual cells written:   {written}')
wf.save(OUT)
print('saved', OUT)

# --------------------------------------------------------- the October tab
# Xero, accrual, 1 to 26 October 2026, company level. The report data was last
# refreshed 23 October and the month's invoicing is not raised, so this is a
# position and not a result. It cannot be split ONS / PRD / CONS: the largest
# lines, on both sides, carry no department in the account name.
OCT = {'period': '1 to 26 October 2026', 'refreshed': '23 October 2026',
       'income': 96489.15, 'cos': -267594.63, 'overhead': -126105.28,
       'other_income': 430.17, 'gross_profit': -171105.48, 'net': -296780.59}
OCT_ACC = [
 ('Income', [('Equipment Sales', 107500.00), ('Production Labour', 14108.75),
             ('Equipment Hires', 4803.00), ('Video Labour', 1905.00),
             ('Freight & Travel - PRD', 600.00),
             ('Ad Hoc Service Labour', -7650.00),
             ('Contract Support Staff', -24777.60)]),
 ('Cost of sales', [('Equipment - Purchase', -86973.33),
             ('Direct Salaries - ONS', -45676.45),
             ('Equipment - Hires', -31331.33),
             ('Subscriptions & Licences - Expense', -20963.99),
             ('Direct Wages - ONS', -14100.99),
             ('Sub-Contract Labour - PRD', -12898.80),
             ('Direct Salaries - PRD', -9419.20),
             ('Sub-Contract Labour - INTEGRATION', -9090.00),
             ('Service - hires', -7280.14), ('Direct Wages - PRD', -6415.53),
             ('Direct Superannuation - ONS', -5709.18),
             ('Direct Freight - PRD', -3828.86),
             ('Direct Freight - INTEGRATION', -3158.00),
             ('Direct Travel - ONS', -2775.66),
             ('Direct Travel - PRD', -1798.07),
             ('Direct Salaries - CONS', -1774.77),
             ('Direct Superannuation - PRD', -1660.16),
             ('All other cost of sales accounts', -2740.17)]),
 ('Overheads', [('Indirect Salaries - ADMIN', -27110.38),
             ('Indirect Salaries - PRD', -23609.20), ('Rent', -20681.16),
             ('IT Network Service & Support', -9178.55),
             ('Superannuation', -8053.25), ('Rental Outgoings', -5811.90),
             ('Indirect Wages - PRD', -4997.28), ('Advertising', -4250.00),
             ('Indirect Salaries - INTEGRATION', -3931.17),
             ('Dues & Subscriptions', -3822.57),
             ('Indirect Salaries - ONS', -3311.34),
             ('Indirect Salaries - VID', -2944.02),
             ('Indirect Salaries - CONS', -1957.80),
             ('Indirect Wages - ONS', -1348.40),
             ('All other overhead accounts', -5098.26)]),
 ('Other income', [('Interest Received', 430.17)]),
]
for head, lines in OCT_ACC:
    tot = sum(v for _, v in lines)
    key = {'Income': 'income', 'Cost of sales': 'cos',
           'Overheads': 'overhead', 'Other income': 'other_income'}[head]
    assert abs(tot - OCT[key]) < 0.005, (head, tot, OCT[key])
assert abs(OCT['income'] + OCT['cos'] - OCT['gross_profit']) < 0.005
assert abs(OCT['gross_profit'] + OCT['overhead'] + OCT['other_income']
           - OCT['net']) < 0.005

NAVY, WHITE, GREYL = 'FF242C39', 'FFFFFFFF', 'FFF4F4F4'
from openpyxl.styles import Border, Side, Alignment
BAND = PatternFill('solid', start_color=NAVY, end_color=NAVY)
TOTF = PatternFill('solid', start_color=GREYL, end_color=GREYL)
MONEY = '#,##0;[RED]\\(#,##0\\);\\-'
BODY = Font(name='Calibri', sz=11, color=NAVY)
BOLD = Font(name='Calibri', sz=11, b=True, color=NAVY)
HEADF = Font(name='Calibri', sz=11, b=True, color=WHITE)

ws = wf.create_sheet('October Xero')
ws['A1'] = 'October 2026, straight from Xero'
ws['A1'].font = Font(name='Calibri', sz=14, b=True, color=NAVY)
for i, t in enumerate([
    f'Period {OCT["period"]}, accrual basis. Xero report data last refreshed '
    f'{OCT["refreshed"]}.',
    'October is not closed and the month\'s invoicing is not raised, so this is '
    'a position and not a result.',
    'Company level only. It cannot be split ONS, PRD and CONS: the largest '
    'lines on both sides carry no department in the account name, and the Xero '
    'tracking categories are not available through this connection.',
    'October on the ONS, PRD and CONS sheets is still the budget. This tab does '
    'not feed those sheets.']):
    ws.cell(3 + i, 1, t).font = NOTE

R = 9
ws.cell(R, 1, 'COMPANY RESULT').font = HEADF
for c in range(1, 4):
    ws.cell(R, c).fill = BAND
for i, (lbl, key, bold) in enumerate([
        ('Income', 'income', False), ('Cost of sales', 'cos', False),
        ('Gross profit', 'gross_profit', True),
        ('Overheads', 'overhead', False),
        ('Other income', 'other_income', False),
        ('Net profit', 'net', True)]):
    r = R + 1 + i
    ws.cell(r, 1, lbl).font = BOLD if bold else BODY
    cell = ws.cell(r, 3, OCT[key])
    cell.font = BOLD if bold else BODY
    cell.number_format = MONEY
    if bold:
        ws.cell(r, 1).fill = TOTF
        cell.fill = TOTF
        ws.cell(r, 2).fill = TOTF

r = R + 9
for head, lines in OCT_ACC:
    ws.cell(r, 1, head.upper()).font = HEADF
    for c in range(1, 4):
        ws.cell(r, c).fill = BAND
    r += 1
    for name, v in lines:
        ws.cell(r, 1, name).font = BODY
        cell = ws.cell(r, 3, v)
        cell.font = BODY
        cell.number_format = MONEY
        r += 1
    ws.cell(r, 1, f'Total {head.lower()}').font = BOLD
    cell = ws.cell(r, 3, sum(v for _, v in lines))
    cell.font = BOLD
    cell.number_format = MONEY
    for c in range(1, 4):
        ws.cell(r, c).fill = TOTF
    r += 2

ws.column_dimensions['A'].width = 44
ws.column_dimensions['B'].width = 3
ws.column_dimensions['C'].width = 16
ws.sheet_view.showGridLines = False
wf.save(OUT)
print('added the October Xero tab')
