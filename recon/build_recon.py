"""Revenue against its own cost, line by line, July to September 2026.

Every income account in the quarter paired with the cost of sales accounts that
carry its cost. The five groups add back to the P&L: revenue 1,624,459.78, cost
of sales 992,160.28, gross profit 632,299.50.

The pairing is a judgement about which cost belongs to which revenue. It is set
out account by account so it can be argued with. Nothing is reallocated: every
account appears exactly once and the totals tie to Xero.
"""
import collections
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as gl

PACK = 'pack/CTS Financial Controller Pack - v20.xlsx'
OUT = 'recon/CTS revenue against cost - Jul to Sep 2026.xlsx'
MONTHS = ('Jul 26', 'Aug 26', 'Sep 26')

NAVY, WHITE, GREY = 'FF242C39', 'FFFFFFFF', 'FFF4F4F4'
BAND = PatternFill('solid', start_color=NAVY, end_color=NAVY)
TOTF = PatternFill('solid', start_color=GREY, end_color=GREY)
RULE = Side(style='thin', color='FFC9CED6')
HEAD = Font(name='Calibri', sz=11, b=True, color=WHITE)
BOLD = Font(name='Calibri', sz=11, b=True, color=NAVY)
BODY = Font(name='Calibri', sz=11, color=NAVY)
SMALL = Font(name='Calibri', sz=10, color='FF6B7585')
NOTE = Font(name='Calibri', sz=10, i=True, color='FF6B7585')
MONEY = '#,##0.00;[RED](#,##0.00);"-"'
PCT = '0.0%;[RED](0.0%);"-"'

# group -> (revenue accounts, cost accounts)
GROUPS = [
 ('Equipment sold',
  ['Equipment Sales'],
  ['Equipment - Purchase']),
 ('Equipment hired out',
  ['Equipment Hires'],
  ['Equipment - Hires', 'Service - hires', 'Damage Waiver']),
 ('Licences and subscriptions resold',
  ['Subscriptions & Licences - Income'],
  ['Subscriptions & Licences - Expense']),
 ('Labour',
  ['Contract Support Staff', 'Installation Labour', 'Production Labour',
   'Ad Hoc Service Labour', 'Conference Call Charge', 'Video Labour',
   'Tech Event Management', 'Tech Consulting/Project Manage'],
  ['Direct Salaries - ONS', 'Direct Wages - ONS', 'Direct Superannuation - ONS',
   'Sub-Contract Labour - ONS', 'Direct Salaries - PRD', 'Direct Wages - PRD',
   'Direct Superannuation - PRD', 'Sub-Contract Labour - PRD',
   'Direct Salaries - VID', 'Direct Wages - VID', 'Direct Superannuation - VID',
   'Direct Salaries - INTEGRATION', 'Direct Superannuation-INTEGRATION',
   'Sub-Contract Labour - INTEGRATION', 'Direct Salaries - CONS',
   'Direct Wages - CONS', 'Direct Superannuation - CONS']),
 ('Travel, freight and mobile recharged',
  ['Meal & Travel - ONS', 'Meal & Travel - PRD', 'Freight & Travel - PRD',
   'Freight & Travel - INTEGRATION', 'Freight & Travel - VID',
   'Mobile Phone Plans'],
  ['Direct Travel - ONS', 'Direct Travel - PRD', 'Direct Travel - CONS',
   'Direct Freight - ONS', 'Direct Freight - PRD', 'Direct Freight - INTEGRATION',
   'Direct Mobile - ONS', 'Direct Mobile - PRD', 'Credit Card Paid - ONS',
   'Credit Card Paid - PRD', 'Credit Card Paid - INTEGRATION',
   'Credit Card Paid - CONS']),
]

print('reading the ledger')
wb = openpyxl.load_workbook(PACK, data_only=True, read_only=True)
inc, cos = collections.defaultdict(float), collections.defaultdict(float)
for row in wb['Cleanup'].iter_rows(min_row=5, max_row=20004, max_col=10,
                                   values_only=True):
    if row[0] in (None, ''):
        continue
    mo = row[1].strftime('%b %y') if hasattr(row[1], 'strftime') else '?'
    if mo not in MONTHS:
        continue
    cat, acct = str(row[4]).strip(), str(row[8]).strip()
    if cat == 'Income':
        inc[acct] += row[0]
    elif cat == 'Cost of Sales':
        cos[acct] += row[0]
wb.close()

# every account must appear exactly once, or the recon does not tie
used_i = [a for _, r, _ in GROUPS for a in r]
used_c = [a for _, _, c in GROUPS for a in c]
assert len(used_i) == len(set(used_i)), 'a revenue account is in two groups'
assert len(used_c) == len(set(used_c)), 'a cost account is in two groups'
missing = (set(inc) - set(used_i)) | (set(cos) - set(used_c))
assert not missing, f'not in any group: {sorted(missing)}'
extra = (set(used_i) - set(inc)) | (set(used_c) - set(cos))
if extra:
    print('  listed but nil this quarter:', sorted(extra))

out = openpyxl.Workbook()
ws = out.active
ws.title = 'Revenue against cost'

ws['A1'] = 'Revenue against its own cost, July to September 2026'
ws['A1'].font = Font(name='Calibri', sz=14, b=True, color=NAVY)
for i, t in enumerate([
        'Every income account in the quarter, paired with the cost of sales '
        'accounts that carry its cost. Xero, accrual, whole of entity.',
        'Which cost belongs to which revenue is a judgement. It is set out '
        'account by account below so it can be argued with.',
        'Nothing is reallocated. Every account appears once and the totals tie '
        'to the profit and loss.']):
    ws.cell(3 + i, 1, t).font = NOTE

r = 7
for c, h in enumerate(['Account', 'Revenue', 'Cost', 'Margin', 'Margin %'], 1):
    cell = ws.cell(r, c, h)
    cell.font = HEAD
    cell.fill = BAND
    if c > 1:
        cell.alignment = Alignment(horizontal='right')
r += 1

group_rows = []
for name, rev_accts, cost_accts in GROUPS:
    rv = sum(inc.get(a, 0.0) for a in rev_accts)
    cv = sum(cos.get(a, 0.0) for a in cost_accts)
    ws.cell(r, 1, name.upper()).font = BOLD
    for c in range(1, 6):
        ws.cell(r, c).fill = TOTF
        ws.cell(r, c).border = Border(top=RULE)
    head_row = r
    r += 1
    for a in sorted(rev_accts, key=lambda x: -inc.get(x, 0.0)):
        if a not in inc:
            continue
        ws.cell(r, 1, '    ' + a).font = BODY
        cell = ws.cell(r, 2, round(inc[a], 2))
        cell.font = BODY
        cell.number_format = MONEY
        r += 1
    for a in sorted(cost_accts, key=lambda x: cos.get(x, 0.0)):
        if a not in cos:
            continue
        ws.cell(r, 1, '    ' + a).font = SMALL
        cell = ws.cell(r, 3, round(abs(cos[a]), 2))
        cell.font = SMALL
        cell.number_format = MONEY
        r += 1
    ws.cell(head_row, 2, round(rv, 2)).number_format = MONEY
    ws.cell(head_row, 3, round(abs(cv), 2)).number_format = MONEY
    ws.cell(head_row, 4, f'=B{head_row}-C{head_row}').number_format = MONEY
    ws.cell(head_row, 5,
            f'=IF(B{head_row}=0,"",D{head_row}/B{head_row})').number_format = PCT
    for c in range(2, 6):
        ws.cell(head_row, c).font = BOLD
    group_rows.append(head_row)
    r += 1

ws.cell(r, 1, 'GROSS PROFIT').font = HEAD
for c in range(1, 6):
    ws.cell(r, c).fill = BAND
for c in (2, 3):
    cell = ws.cell(r, c, '=' + '+'.join(f'{gl(c)}{x}' for x in group_rows))
    cell.font = HEAD
    cell.number_format = MONEY
ws.cell(r, 4, f'=B{r}-C{r}').number_format = MONEY
ws.cell(r, 5, f'=IF(B{r}=0,"",D{r}/B{r})').number_format = PCT
for c in (4, 5):
    ws.cell(r, c).font = HEAD
total_row = r

r += 3
ws.cell(r, 1, 'WHAT THE EQUIPMENT LINE WOULD LOOK LIKE CODED CONSISTENTLY').font = HEAD
for c in range(1, 6):
    ws.cell(r, c).fill = BAND
r += 1
for t in ['On the Aware Super job in September, 124,635.75 of equipment was '
          'bought and 63,357.17 was billed as equipment. The rest of the job '
          'was billed as installation labour.',
          'Moving 61,278.58 of that revenue onto the equipment line is what it '
          'takes for the job\'s equipment to break even. It is an illustration '
          'of the size of the mis-coding, not a proposed journal.',
          'Gross profit does not change. Only the split between two lines does.']:
    ws.cell(r, 1, t).font = NOTE
    r += 1
r += 1
SHIFT = 61278.58
for c, h in enumerate(['Line', 'Revenue as coded', 'Revenue restated',
                       'Cost', 'Margin % restated'], 1):
    cell = ws.cell(r, c, h)
    cell.font = BOLD
    cell.border = Border(bottom=RULE)
    if c > 1:
        cell.alignment = Alignment(horizontal='right', wrap_text=True)
r += 1
eq_rev, eq_cost = inc['Equipment Sales'], abs(cos['Equipment - Purchase'])
lab_rev = sum(inc.get(a, 0.0) for a in GROUPS[3][1])
lab_cost = sum(abs(cos.get(a, 0.0)) for a in GROUPS[3][2])
for name, rev, cost, shift in (('Equipment sold', eq_rev, eq_cost, SHIFT),
                               ('Labour', lab_rev, lab_cost, -SHIFT)):
    ws.cell(r, 1, name).font = BODY
    for c, v in ((2, rev), (3, rev + shift), (4, cost)):
        cell = ws.cell(r, c, round(v, 2))
        cell.font = BODY
        cell.number_format = MONEY
    ws.cell(r, 5, f'=(C{r}-D{r})/C{r}').number_format = PCT
    ws.cell(r, 5).font = BOLD
    r += 1

ws.column_dimensions['A'].width = 44
for c in range(2, 6):
    ws.column_dimensions[gl(c)].width = 18
ws.freeze_panes = 'A8'
ws.sheet_view.showGridLines = False
out.save(OUT)
print('saved', OUT)
