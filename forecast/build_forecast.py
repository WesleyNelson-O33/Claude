"""Forecast FY27: actuals to September, budget from October.

Her own forecast workbook is driven by links to another file (ONS Sales
Forecast, Fixed Costs, Casual Pays, FY27 Budget - detailed). Writing into it
would strip the cached results of those links, so this builds a self contained
copy instead: the same three sheets, the same line structure, actuals for July
to September and her budget untouched from October on.

Actuals come from the controller pack, which is the Xero GL. The overhead block
uses the pack's departmental split, the same three keys that drive the ONS, PRD
and CONS management pages, so the figures agree with those pages.

Three lines carry nothing in the actual months because nothing in the GL
identifies them: PWC IT Nat. Pool of Hrs, Discretionary Bonuses and Salaries FT
Fixed Term on PRD, and Aware Super on CONS. Their amounts sit in the line that
does carry the cost, so every total is right.
"""
import collections
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter as gl

PACK = 'pack/CTS Financial Controller Pack - v20.xlsx'
SRC = '/tmp/claude-0/-home-user-Claude/a00ffe59-433c-507b-9006-917d8c60083a/scratchpad/fc/fc.xlsx'
OUT = 'forecast/CTS Forecast FY27 - actuals to September.xlsx'

NAVY, WHITE, GREYL, TEAL = 'FF242C39', 'FFFFFFFF', 'FFF4F4F4', 'FF005358'
BAND = PatternFill('solid', start_color=NAVY, end_color=NAVY)
ACT = PatternFill('solid', start_color='FFE8F4F3', end_color='FFE8F4F3')
TOTF = PatternFill('solid', start_color=GREYL, end_color=GREYL)
RULE = Side(style='medium', color=NAVY)
MONEY = '#,##0;[RED]\\(#,##0\\);\\-'
PCT = '0.0%;[RED]\\(0.0%\\);\\-'
BODY = Font(name='Calibri', sz=11, color=NAVY)
BOLD = Font(name='Calibri', sz=11, b=True, color=NAVY)
HEADF = Font(name='Calibri', sz=11, b=True, color=WHITE)
NOTE = Font(name='Calibri', sz=10, color='FF6B7585')

MONTHS = ['July', 'August', 'September', 'October', 'November', 'December',
          'January', 'February', 'March', 'April', 'May', 'June']
ACTUAL = 3                                   # July, August, September

# group -> (Cleanup department tags, source sheet, first month column on it)
GROUPS = {
    'ONS':  (['ONSITE'], 'ONS', 6),
    'PRD':  (['PRODUCTION', 'VIDEO'], 'PRD', 8),
    'CONS': (['INTEGRATION', 'CONSULTING'], 'CONS', 7),
}
EQUIP = ['Equipment - Purchase', 'Equipment - Hires', 'Service - hires',
         'Equipment -Maintenance Support', 'Stamp Duty', 'Damage Waiver',
         'Subscriptions & Licences - Expense', 'Credit Card Paid - ADMIN']
TRANSPORT = ['Direct Travel', 'Direct Freight', 'Direct Mobile']


def tagged(names, suffix):
    return [f'{n} - {suffix}' for n in names]


# line, label, the GL accounts (None = nothing in the GL identifies it)
LINES = {
 'ONS': [
    ('rev', 'Mobile', ['Mobile Phone Plans']),
    ('rev', 'Hotline', None),
    ('rev', 'Spotify', None),
    ('rev', 'Revenue', 'BALANCE'),
    ('tot', 'Total Revenue', None),
    ('cos', 'Direct Salaries', ['Direct Salaries - ONS']),
    ('cos', 'Direct Superannuation', ['Direct Superannuation - ONS']),
    ('cos', 'PWC IT Nat. Pool of Hrs COS', None),
    ('cos', 'Salaries (Casual + Super + Spacera)', ['Direct Wages - ONS']),
    ('cos', 'Discretionary Bonuses', None),
    ('cos', 'Equipment & Services', EQUIP + ['Sub-Contract Labour - ONS',
                                             'Credit Card Paid - ONS',
                                             'Supplier Per Diems - ONS']),
    ('cos', 'Transport', tagged(TRANSPORT, 'ONS')),
    ('tot', 'Total COGs', None),
 ],
 'PRD': [
    ('rev', 'Revenue', 'BALANCE'),
    ('tot', 'Total Revenue', None),
    ('cos', 'Direct Salaries (Production)', ['Direct Salaries - PRD']),
    ('cos', 'Direct Salaries (Video)', ['Direct Salaries - VID']),
    ('cos', 'Direct Superannuation (Production)', ['Direct Superannuation - PRD']),
    ('cos', 'Direct Superannuation (Video)', ['Direct Superannuation - VID']),
    ('cos', 'Salaries FT Fixed Term (Gross + Super)', None),
    ('cos', 'Salaries Casual Production (Gross + Super)', ['Direct Wages - PRD']),
    ('cos', 'Salaries Casual Video (Gross + Super)', ['Direct Wages - VID']),
    ('cos', 'Equipment (Production)', ('PRODUCTION', EQUIP +
                                       ['Sub-Contract Labour - PRD',
                                        'Credit Card Paid - PRD',
                                        'Supplier Per Diems - PRD'])),
    ('cos', 'Equipment (Video)', ('VIDEO', EQUIP + ['Sub-Contract Labour - VID',
                                                    'Credit Card Paid - VID',
                                                    'Supplier Per Diems - VID'])),
    ('cos', 'Transportation (Production)', tagged(TRANSPORT, 'PRD')),
    ('cos', 'Transportation (Video)', tagged(TRANSPORT, 'VID')),
    ('tot', 'Total COGs', None),
 ],
 'CONS': [
    ('rev', 'Revenue', 'BALANCE'),
    ('tot', 'Total Revenue', None),
    ('cos', 'Direct Salaries', ['Direct Salaries - INTEGRATION', 'Direct Salaries - CONS',
                                'Direct Wages - INTEGRATION', 'Direct Wages - CONS']),
    ('cos', 'Direct Superannuation', ['Direct Superannuation-INTEGRATION',
                                      'Direct Superannuation - CONS']),
    ('cos', 'INT- Equip/Sub-Contract/Licenses', EQUIP +
            ['Sub-Contract Labour - INTEGRATION', 'Sub-Contract Labour - CONS',
             'Credit Card Paid - INTEGRATION', 'Credit Card Paid - CONS',
             'Supplier Per Diems - INTEGRATION', 'Supplier Per Diems - CONS']),
    ('cos', 'Aware Super (COS)', None),
    ('cos', 'Transport (INT & CONS)', tagged(TRANSPORT, 'INTEGRATION') +
            tagged(TRANSPORT, 'CONS')),
    ('tot', 'Total COGs', None),
 ],
}
OVERHEAD = ['Business Insurances', 'Employment Expenses',
            'General & Administrative Exp', 'Operating Expenses']
print('reading the GL from the pack')

# ------------------------------------------------- the actuals, from the GL
wb = openpyxl.load_workbook(PACK, data_only=True, read_only=True)
gl_rows = collections.defaultdict(float)          # (dept, month, cat, account)
for row in wb['Cleanup'].iter_rows(min_row=5, max_row=20004, max_col=10,
                                   values_only=True):
    if row[0] in (None, ''):
        continue
    mo = row[1]
    if not hasattr(mo, 'strftime'):
        continue
    gl_rows[(str(row[3]).strip(), mo.strftime('%b'), str(row[4]).strip(),
             str(row[8]).strip())] += row[0]
wb.close()
MON3 = ('Jul', 'Aug', 'Sep')


def gl_sum(tags, mon, cat, accounts):
    """Actual for a set of accounts. accounts may be (tag, [names]) to pin a tag."""
    if accounts is None or accounts == 'BALANCE':
        return 0.0
    if isinstance(accounts, tuple):
        tags, accounts = [accounts[0]], accounts[1]
    return sum(gl_rows.get((t, mon, cat, a), 0.0) for t in tags for a in accounts)


def gl_cat_total(tags, mon, cat):
    return sum(v for (d, m, c, _), v in gl_rows.items()
               if d in tags and m == mon and c == cat)


# the overhead block: the pack's split, folded into her four categories
pv = openpyxl.load_workbook(PACK, data_only=True)
osh, eng = pv['Overhead Split'], pv['Engine']
ecol = {eng.cell(4, c).value.strftime('%b'): c for c in range(5, 18)
        if hasattr(eng.cell(4, c).value, 'strftime')}
FOLD = {'Business Insurances': ['Business Insurances'],
        'Employment Expenses': ['Employment Expenses'],
        'General & Administrative Exp': ['General & Administrative Exp', 'Other',
                                         'Credit Card fees paid'],
        'Operating Expenses': ['Operating Expenses', 'Advertising & Promotion Exp']}
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


# ------------------------------------------------- her budget, from her file
src = openpyxl.load_workbook(SRC, data_only=True)
BUDGET_ROWS = {                      # group -> {line label: row on her sheet}
 'ONS':  {'Mobile': 2, 'Hotline': 3, 'Spotify': 4, 'Revenue': 5,
          'Direct Salaries': 7, 'Direct Superannuation': 8,
          'PWC IT Nat. Pool of Hrs COS': 9,
          'Salaries (Casual + Super + Spacera)': 10, 'Discretionary Bonuses': 11,
          'Equipment & Services': 12, 'Transport': 13,
          'Business Insurances': 35, 'Employment Expenses': 36,
          'General & Administrative Exp': 37, 'Operating Expenses': 38},
 'PRD':  {'Revenue': 2, 'Direct Salaries (Production)': 4,
          'Direct Salaries (Video)': 5,
          'Direct Superannuation (Production)': 6,
          'Direct Superannuation (Video)': 7,
          'Salaries FT Fixed Term (Gross + Super)': 8,
          'Salaries Casual Production (Gross + Super)': 9,
          'Salaries Casual Video (Gross + Super)': 10,
          'Equipment (Production)': 11, 'Equipment (Video)': 12,
          'Transportation (Production)': 13, 'Transportation (Video)': 14,
          'Business Insurances': 30, 'Employment Expenses': 31,
          'General & Administrative Exp': 32, 'Operating Expenses': 33},
 'CONS': {'Revenue': 2, 'Direct Salaries': 5, 'Direct Superannuation': 6,
          'INT- Equip/Sub-Contract/Licenses': 7, 'Aware Super (COS)': 8,
          'Transport (INT & CONS)': 9,
          'Business Insurances': 28, 'Employment Expenses': 29,
          'General & Administrative Exp': 30, 'Operating Expenses': 31},
}


def budget(group, label, i):
    ws = src[GROUPS[group][1]]
    r = BUDGET_ROWS[group].get(label)
    if r is None:
        return 0.0
    v = ws.cell(r, GROUPS[group][2] + i).value
    return float(v) if isinstance(v, (int, float)) else 0.0

print('building the workbook')

# Xero, accrual, 1 to 26 October 2026, company level. The Xero report data was
# last refreshed 23 October, and the month's invoicing is not raised, so this is
# a position snapshot and nothing more.
OCT = {'period': '1 to 26 October 2026', 'refreshed': '23 October 2026',
       'income': 96489.15, 'cos': -267594.63, 'overhead': -126105.28,
       'other_income': 430.17, 'gross_profit': -171105.48, 'net': -296780.59}

out = openpyxl.Workbook()
out.remove(out.active)


def band(ws, row, cols, text):
    ws.cell(row, 1, text).font = HEADF
    for c in range(1, cols + 1):
        ws.cell(row, c).fill = BAND


def build(group):
    ws = out.create_sheet(group)
    tags = GROUPS[group][0]
    ws['A1'] = f'{group}  forecast FY27'
    ws['A1'].font = Font(name='Calibri', sz=14, b=True, color=NAVY)
    ws['A2'] = ('July to September are actual, from the controller pack. '
                'October to June are her budget, unchanged.')
    ws['A2'].font = NOTE
    HEAD = 4
    ws.cell(HEAD, 1, 'Line').font = BOLD
    for i, m in enumerate(MONTHS):
        c = ws.cell(HEAD, 2 + i, m)
        c.font = BOLD
        c.alignment = Alignment(horizontal='right')
        if i < ACTUAL:
            c.fill = ACT
    ws.cell(HEAD, 14, 'Total').font = BOLD
    ws.cell(HEAD, 15, 'YTD actual').font = BOLD
    for c in range(1, 16):
        ws.cell(HEAD, c).border = Border(bottom=RULE)

    r = HEAD + 1
    rows = {}
    rev_rows, cos_rows = [], []
    for kind, label, accounts in LINES[group]:
        if kind == 'tot':
            rows[label] = r
            ws.cell(r, 1, label).font = BOLD
            src_rows = rev_rows if 'Revenue' in label else cos_rows
            for i in range(12):
                col = gl(2 + i)
                cell = ws.cell(r, 2 + i,
                               '=' + '+'.join(f'{col}{x}' for x in src_rows))
                cell.font = BOLD
                cell.number_format = MONEY
                cell.fill = TOTF
            r += 1
            continue
        rows[label] = r
        ws.cell(r, 1, label).font = BODY
        (rev_rows if kind == 'rev' else cos_rows).append(r)
        for i, mon in enumerate(MONTHS):
            cell = ws.cell(r, 2 + i)
            if i < ACTUAL:
                m3 = MON3[i]
                cat = 'Income' if kind == 'rev' else 'Cost of Sales'
                if accounts == 'BALANCE':
                    named = sum(gl_sum(tags, m3, cat, a)
                                for k, l, a in LINES[group]
                                if k == kind and a not in (None, 'BALANCE'))
                    val = gl_cat_total(tags, m3, cat) - named
                else:
                    val = gl_sum(tags, m3, cat, accounts)
                cell.value = round(val, 2)
                cell.fill = ACT
            else:
                cell.value = round(budget(group, label, i), 2)
            cell.font = BODY
            cell.number_format = MONEY
        r += 1

    # the overhead block
    r += 1
    band(ws, r, 15, 'OVERHEAD CHARGED TO THE DEPARTMENT')
    r += 1
    oh_rows = []
    for label in OVERHEAD:
        rows[label] = r
        oh_rows.append(r)
        ws.cell(r, 1, label).font = BODY
        for i in range(12):
            cell = ws.cell(r, 2 + i)
            if i < ACTUAL:
                cell.value = round(overhead(group, MON3[i])[label], 2)
                cell.fill = ACT
            else:
                cell.value = round(-abs(budget(group, label, i)), 2)
            cell.font = BODY
            cell.number_format = MONEY
        r += 1
    rows['Total overhead'] = r
    ws.cell(r, 1, 'Total overhead').font = BOLD
    for i in range(12):
        col = gl(2 + i)
        cell = ws.cell(r, 2 + i, '=' + '+'.join(f'{col}{x}' for x in oh_rows))
        cell.font = BOLD
        cell.number_format = MONEY
        cell.fill = TOTF
    r += 2

    # the result
    band(ws, r, 15, 'RESULT')
    r += 1
    tr, tc, to = rows['Total Revenue'], rows['Total COGs'], rows['Total overhead']
    for label, formula, fmt in (
            ('Gross Profit', lambda c: f'={c}{tr}+{c}{tc}', MONEY),
            ('Gross Profit Margin', lambda c: f'=IFERROR(({c}{tr}+{c}{tc})/{c}{tr},"")', PCT),
            ('Net Profit', lambda c: f'={c}{tr}+{c}{tc}+{c}{to}', MONEY),
            ('Net Profit Margin',
             lambda c: f'=IFERROR(({c}{tr}+{c}{tc}+{c}{to})/{c}{tr},"")', PCT)):
        rows[label] = r
        ws.cell(r, 1, label).font = BOLD
        for i in range(12):
            cell = ws.cell(r, 2 + i, formula(gl(2 + i)))
            cell.font = BOLD
            cell.number_format = fmt
            if i < ACTUAL:
                cell.fill = ACT
        r += 1

    # Total and YTD actual columns
    for label, rr in rows.items():
        if 'Margin' in label:
            ws.cell(rr, 14, f'=IFERROR(N{rows["Net Profit"] if "Net" in label else rows["Gross Profit"]}'
                            f'/N{rows["Total Revenue"]},"")').number_format = PCT
            ws.cell(rr, 15, f'=IFERROR(O{rows["Net Profit"] if "Net" in label else rows["Gross Profit"]}'
                            f'/O{rows["Total Revenue"]},"")').number_format = PCT
        else:
            ws.cell(rr, 14, f'=SUM(B{rr}:M{rr})').number_format = MONEY
            ws.cell(rr, 15, f'=SUM(B{rr}:D{rr})').number_format = MONEY
        ws.cell(rr, 14).font = BOLD
        ws.cell(rr, 15).font = BOLD

    ws.column_dimensions['A'].width = 42
    for c in range(2, 16):
        ws.column_dimensions[gl(c)].width = 13
    ws.freeze_panes = ws[f'B{HEAD + 1}']
    ws.sheet_view.showGridLines = False
    return rows


built = {g: build(g) for g in GROUPS}
print('  sheets built:', ', '.join(built))

# ------------------------------------------------------------- the summary
ws = out.create_sheet('Summary', 0)
ws['A1'] = 'CTS forecast FY27'
ws['A1'].font = Font(name='Calibri', sz=14, b=True, color=NAVY)
for i, t in enumerate([
        'July to September are actual, taken from the controller pack, which is the Xero GL.',
        'The overhead charged to each department uses the controller pack\'s split, the same '
        'three keys that drive the ONS, PRD and CONS management pages.',
        'October to June are her budget, carried over unchanged from the forecast workbook.',
        'October has not been closed. There is a Xero snapshot of it further down, at company '
        'level only.']):
    ws.cell(3 + i, 1, t).font = NOTE

R = 8
band(ws, R, 15, '1. THE THREE GROUPS, BY MONTH')
ws.cell(R + 1, 1, 'Group').font = BOLD
for i, m in enumerate(MONTHS):
    c = ws.cell(R + 1, 2 + i, m)
    c.font = BOLD
    c.alignment = Alignment(horizontal='right')
    if i < ACTUAL:
        c.fill = ACT
ws.cell(R + 1, 14, 'Total').font = BOLD
ws.cell(R + 1, 15, 'YTD actual').font = BOLD
for c in range(1, 16):
    ws.cell(R + 1, c).border = Border(bottom=RULE)

r = R + 2
for label, key in (('Total Revenue', 'Total Revenue'), ('Total COGs', 'Total COGs'),
                   ('Total overhead', 'Total overhead'), ('Net Profit', 'Net Profit')):
    ws.cell(r, 1, label).font = HEADF
    for c in range(1, 16):
        ws.cell(r, c).fill = BAND
    r += 1
    for g in GROUPS:
        ws.cell(r, 1, f'   {g}').font = BODY
        for i in range(12):
            cell = ws.cell(r, 2 + i, f"='{g}'!{gl(2 + i)}{built[g][key]}")
            cell.font = BODY
            cell.number_format = MONEY
            if i < ACTUAL:
                cell.fill = ACT
        ws.cell(r, 14, f'=SUM(B{r}:M{r})').number_format = MONEY
        ws.cell(r, 15, f'=SUM(B{r}:D{r})').number_format = MONEY
        r += 1
    ws.cell(r, 1, '   CTS').font = BOLD
    for i in range(12):
        col = gl(2 + i)
        cell = ws.cell(r, 2 + i, f'=SUM({col}{r - 3}:{col}{r - 1})')
        cell.font = BOLD
        cell.number_format = MONEY
        cell.fill = TOTF
    for c in (14, 15):
        ws.cell(r, c, f'=SUM({gl(c)}{r - 3}:{gl(c)}{r - 1})').number_format = MONEY
        ws.cell(r, c).font = BOLD
    r += 2

# the October snapshot
band(ws, r, 6, f'2. OCTOBER SNAPSHOT FROM XERO, {OCT["period"].upper()}')
r += 1
ws.cell(r, 1, 'This is not a closed month and not a forecast input.').font = NOTE
ws.cell(r + 1, 1, f'The month\'s invoicing is not raised, so revenue is well short. The Xero '
                  f'report data was last refreshed {OCT["refreshed"]}.').font = NOTE
r += 3
ws.cell(r, 1, 'Line').font = BOLD
ws.cell(r, 2, 'Per Xero').font = BOLD
ws.cell(r, 3, 'October budget').font = BOLD
ws.cell(r, 4, 'Difference').font = BOLD
for c in range(1, 5):
    ws.cell(r, c).border = Border(bottom=RULE)
    ws.cell(r, c).alignment = Alignment(horizontal='right' if c > 1 else 'left')
r += 1
OCT_LINES = [('Total income', 'income', 'Total Revenue'),
             ('Cost of sales', 'cos', 'Total COGs'),
             ('Overheads', 'overhead', 'Total overhead'),
             ('Other income', 'other_income', None),
             ('Net profit', 'net', 'Net Profit')]
for label, k, key in OCT_LINES:
    ws.cell(r, 1, label).font = BODY if label != 'Net profit' else BOLD
    ws.cell(r, 2, OCT[k]).number_format = MONEY
    ws.cell(r, 2).font = BODY if label != 'Net profit' else BOLD
    if key:
        ws.cell(r, 3, '=' + '+'.join(f"'{g}'!E{built[g][key]}" for g in GROUPS))
        ws.cell(r, 3).number_format = MONEY
        ws.cell(r, 4, f'=B{r}-C{r}').number_format = MONEY
        ws.cell(r, 3).font = BODY
        ws.cell(r, 4).font = BODY
    r += 1

r += 2
band(ws, r, 6, '3. WHAT IS NOT IN THE ACTUAL MONTHS')
r += 1
for t in ['Four forecast lines carry nothing in July to September, because nothing in the '
          'GL identifies them on its own:',
          '   ONS, PWC IT Nat. Pool of Hrs COS. It is a job number, not an account.',
          '   ONS, Discretionary Bonuses.',
          '   PRD, Salaries FT Fixed Term.',
          '   CONS, Aware Super (COS).',
          'The cost itself is in the line that carries it, so every total is right. Tell me '
          'which accounts or jobs feed those four and I will split them out.',
          'October by department needs the October GL export, the same file that goes into '
          'the controller pack. Xero on its own cannot split it, because the largest October '
          'accounts carry no department in the name.']:
    ws.cell(r, 1, t).font = NOTE
    r += 1

ws.column_dimensions['A'].width = 46
for c in range(2, 16):
    ws.column_dimensions[gl(c)].width = 13
ws.sheet_view.showGridLines = False

out.save(OUT)
print(f'saved {OUT}')
