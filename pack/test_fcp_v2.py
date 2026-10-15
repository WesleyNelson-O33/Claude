"""Load synthetic data through all three paste zones, recalculate, check it ties."""
import json, random, datetime as dt, subprocess, sys, shutil
import openpyxl
HERE = __file__.rsplit('/', 1)[0]
SRC = f'{HERE}/CTS Financial Controller Pack v2.xlsx'
TST = f'{HERE}/_test.xlsx'
MASTER = json.load(open(f'{HERE}/data/lists_master.json'))
BSACC = json.load(open(f'{HERE}/data/bs_accounts.json'))
random.seed(7)

CC = {'ONSITE': 'ONSITE', 'PRODUCTION': 'PRODUCTION', 'VIDEO': 'VIDEO',
      'INTEGRATION': 'INTEGRATION', 'CONSULTING': 'CONSULTING', 'ADMIN': 'CTS'}
MONTHS = [dt.date(2026, 7, 15), dt.date(2026, 8, 15), dt.date(2026, 9, 15)]
pl_accts = [m for m in MASTER if m['cat'] in ('Income', 'Cost of Sales', 'Expenses',
                                              'Other Income', 'Other Expenses')]
rows = []          # the GL lines we will paste
exp_sub = {}       # (dept, cat, sub, period) -> amount
exp_add = {}       # (dept, addback, period) -> amount
def add(date, code, name, dr, cr, cc, job=''):
    rows.append([len(rows) + 1, code, name, 'Invoice', date, 'Test Contact',
                 dr or None, cr or None, job, '', '', 'test line', cc])

for p, d in enumerate(MONTHS, start=1):
    for dept, tag in CC.items():
        for m in random.sample(pl_accts, 26):
            amt = round(random.uniform(200, 9000), 2)
            income = m['cat'] in ('Income', 'Other Income')
            dr, cr = (0, amt) if income else (amt, 0)
            add(d, m['code'], m['name'], dr, cr, tag)
            signed = cr - dr
            exp_sub[(dept, m['cat'], m['sub'], p)] = \
                round(exp_sub.get((dept, m['cat'], m['sub'], p), 0) + signed, 2)
            if m['addback']:
                exp_add[(dept, m['addback'], p)] = \
                    round(exp_add.get((dept, m['addback'], p), 0) + signed, 2)
# edge cases
byc = {m['code']: m for m in MASTER}
def one(code, dept, p, dr, cr, cc, job=''):
    m = byc[code]; add(MONTHS[p - 1], code, m['name'], dr, cr, cc, job)
    s = cr - dr
    exp_sub[(dept, m['cat'], m['sub'], p)] = round(exp_sub.get((dept, m['cat'], m['sub'], p), 0) + s, 2)
    if m['addback']:
        exp_add[(dept, m['addback'], p)] = round(exp_add.get((dept, m['addback'], p), 0) + s, 2)
one('61300', 'ADMIN', 3, 12000, 0, 'CTS')                  # depreciation
one('61730', 'ADMIN', 3, 71153, 0, 'CTS')                  # instant asset write off
one('91000', 'ADMIN', 3, 4500, 0, 'CTS')                   # interest expense
one('81000', 'ADMIN', 3, 0, 900, 'CTS')                    # interest income
one('92000', 'ADMIN', 3, 30000, 0, 'CTS')                  # income tax
one('42100', 'PRODUCTION', 3, 0, 5000, '', 'J0099 [PRD]')  # blank cost centre, bracket tag
# things that must be ignored
add(dt.date(2026, 6, 20), '41100', 'Contract Support Staff', 0, 999999, 'ONSITE')  # prior FY
add(MONTHS[2], '99999', 'Not A Real Account', 0, 55555, 'ONSITE')                  # unknown
add(MONTHS[2], '41100', 'Contract Support Staff', 0, 777, 'NOWHERE')               # bad tag

shutil.copy(SRC, TST)
wb = openpyxl.load_workbook(TST)
gp = wb['GL_Paste']
gp.cell(1, 1, 'Corporate Technology Services Pty Ltd')
gp.cell(2, 1, 'Account Transactions for P&L analysis   1 Jul 2026 to 30 Sep 2026')
for i, h in enumerate(['Index', 'Account Code', 'Account Name', 'Source', 'Date', 'Contact',
                       'Debit', 'Credit', 'Job Numbers', 'Invoice Number', 'Reference',
                       'Description', 'Cost Centres']):
    gp.cell(3, 1 + i, h)
for i, row in enumerate(rows):
    for c, v in enumerate(row, start=1):
        gp.cell(4 + i, c, v)
print('GL lines pasted:', len(rows))

# ---- the P&L export, built from the same numbers so the control must agree --
CATS = ['Income', 'Cost of Sales', 'Expenses', 'Other Income', 'Other Expenses']
sep = {c: round(sum(v for (d, cat, s, p), v in exp_sub.items() if cat == c and p == 3), 2)
       for c in CATS}
pp = wb['PL_Paste']
pp.cell(1, 1, 'Account'); pp.cell(1, 2, 'Sep 2026'); pp.cell(1, 3, 'Aug 2026')
LBL = {'Income': 'Total Income', 'Cost of Sales': 'Total Cost of Sales',
       'Expenses': 'Total Operating Expenses', 'Other Income': 'Total Other Income',
       'Other Expenses': 'Total Other Expenses'}
for i, c in enumerate(CATS):
    pp.cell(3 + i, 1, LBL[c]); pp.cell(3 + i, 2, sep[c])
print('Sep per the GL:', {k: f'{v:,.2f}' for k, v in sep.items()})

# ---- the balance sheet export, made to balance ------------------------------
bp = wb['BS_Paste']
bp.cell(1, 1, 'Account'); bp.cell(1, 2, '30 Sep 2026')
bp.cell(1, 3, '31 Aug 2026'); bp.cell(1, 4, '1 Jul 2026')
bal = {}
r = 3
for a in BSACC:
    if a['group'] == 'Equity': continue
    v = round(random.uniform(1000, 90000), 2)
    bal[a['code']] = v
    bp.cell(r, 1, a['name'])
    for c, k in ((2, 1.0), (3, 0.97), (4, 0.9)): bp.cell(r, c, round(v * k, 2))
    r += 1
assets = sum(v for c, v in bal.items() if next(a for a in BSACC if a['code'] == c)['group'].endswith('Assets'))
liabs = sum(v for c, v in bal.items() if 'Liabilities' in next(a for a in BSACC if a['code'] == c)['group'])
eq = round(assets - liabs, 2)
bp.cell(r, 1, 'Retained Earnings')
for c, k in ((2, 1.0), (3, 0.97), (4, 0.9)): bp.cell(r, c, round(eq * k, 2))
print(f'BS: assets {assets:,.2f}  liabilities {liabs:,.2f}  equity plug {eq:,.2f}')
wb.save(TST)
json.dump({'sub': {'|'.join(map(str, k)): v for k, v in exp_sub.items()},
           'add': {'|'.join(map(str, k)): v for k, v in exp_add.items()},
           'sep': sep, 'assets': assets, 'liabs': liabs, 'equity': eq},
          open(f'{HERE}/_expected.json', 'w'))
print('wrote', TST)
