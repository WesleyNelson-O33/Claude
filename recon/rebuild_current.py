"""Rebuild the recon on Xero as it stands, and add the GST.

The Aware Super invoice has been recoded since the transaction export was
taken. 92,601.72 moved from Installation Labour to Equipment Sales, so the
ledger now matches the invoice line for line, and equipment is no longer
loss-making.

  Summary          the five groups on the current ledger, with the old
                   equipment and labour figures beside them
  GST              the Aware Super invoices and the Complete AV bills,
                   net, GST and gross
  Aware invoice    the job before and after the recode, with GST
  Detail           left as it was, and flagged as the pre-recode export

Figures are Xero, accrual, whole of entity, July to September 2026, read
4 November 2026.
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as gl

OUT = 'recon/CTS revenue against cost - Jul to Sep 2026.xlsx'
PULLED = '4 November 2026'

NAVY, WHITE, GREY = 'FF242C39', 'FFFFFFFF', 'FFF4F4F4'
MINT, AMBER = 'FFE8F4F3', 'FFFFF4D6'
BAND = PatternFill('solid', start_color=NAVY, end_color=NAVY)
TOTF = PatternFill('solid', start_color=GREY, end_color=GREY)
WAS = PatternFill('solid', start_color=AMBER, end_color=AMBER)
NOW = PatternFill('solid', start_color=MINT, end_color=MINT)
RULE = Side(style='thin', color='FFC9CED6')
THICK = Side(style='medium', color=NAVY)
HEAD = Font(name='Calibri', sz=11, b=True, color=WHITE)
BOLD = Font(name='Calibri', sz=11, b=True, color=NAVY)
BODY = Font(name='Calibri', sz=11, color=NAVY)
SMALL = Font(name='Calibri', sz=10, color='FF6B7585')
NOTE = Font(name='Calibri', sz=10, i=True, color='FF6B7585')
MONEY = '#,##0.00;[RED](#,##0.00);"-"'
PCT = '0.0%;[RED](0.0%);"-"'

# Xero, July to September 2026, read 4 November 2026
NOW_ACC = {
 'Equipment Sales': 274912.89, 'Equipment Hires': 172852.77,
 'Subscriptions & Licences - Income': 84810.19,
 'Contract Support Staff': 606466.07, 'Installation Labour': 87649.13,
 'Production Labour': 172607.75, 'Ad Hoc Service Labour': 94737.72,
 'Conference Call Charge': 64004.50, 'Video Labour': 22346.25,
 'Tech Event Management': 9545.00, 'Tech Consulting/Project Manage': 1067.00,
 'Meal & Travel - ONS': 14215.01, 'Meal & Travel - PRD': 4955.00,
 'Freight & Travel - PRD': 9245.00, 'Freight & Travel - INTEGRATION': 1633.00,
 'Freight & Travel - VID': 300.00, 'Mobile Phone Plans': 3112.50,
 'Equipment - Purchase': 225129.96, 'Equipment - Hires': 39800.64,
 'Service - hires': 23704.26, 'Damage Waiver': 135.60,
 'Subscriptions & Licences - Expense': 82143.32,
 'Direct Salaries - ONS': 252435.35, 'Direct Wages - ONS': 88053.15,
 'Direct Superannuation - ONS': 40487.00, 'Sub-Contract Labour - ONS': 5208.00,
 'Direct Salaries - PRD': 45007.69, 'Direct Wages - PRD': 19833.38,
 'Direct Superannuation - PRD': 10019.47, 'Sub-Contract Labour - PRD': 21762.30,
 'Direct Salaries - VID': 4701.40, 'Direct Wages - VID': 2895.79,
 'Direct Superannuation - VID': 1817.27,
 'Direct Salaries - INTEGRATION': 8088.02,
 'Direct Superannuation-INTEGRATION': 984.35,
 'Sub-Contract Labour - INTEGRATION': 90041.70,
 'Direct Salaries - CONS': 4530.30, 'Direct Wages - CONS': 235.71,
 'Direct Superannuation - CONS': 931.51,
 'Direct Travel - ONS': 7776.53, 'Direct Travel - PRD': 6673.11,
 'Direct Travel - CONS': 219.03, 'Direct Freight - ONS': 464.00,
 'Direct Freight - PRD': 6154.10, 'Direct Freight - INTEGRATION': 1099.51,
 'Direct Mobile - ONS': 1359.76, 'Direct Mobile - PRD': 370.92,
 'Credit Card Paid - ONS': 39.46, 'Credit Card Paid - PRD': 56.48,
 'Credit Card Paid - INTEGRATION': 0.66, 'Credit Card Paid - CONS': 0.55,
}
# the two accounts the recode moved, as the export had them
WAS_ACC = {'Equipment Sales': 182311.17, 'Installation Labour': 180250.85}
MOVED = NOW_ACC['Equipment Sales'] - WAS_ACC['Equipment Sales']

GROUPS = [
 ('Equipment sold', ['Equipment Sales'], ['Equipment - Purchase']),
 ('Equipment hired out', ['Equipment Hires'],
  ['Equipment - Hires', 'Service - hires', 'Damage Waiver']),
 ('Licences and subscriptions resold', ['Subscriptions & Licences - Income'],
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
# invoice / bill, net, GST
SALES = [('INV-10596', 'Seminar Room AV Upgrade', 205352.17, 20535.23),
         ('INV-10635', 'Appspace licensing, 12 months', 21197.37, 2119.74),
         ('INV-10652', 'National maintenance, September', 11000.00, 1100.00),
         ('INV-10646', 'Perth BYOD USB-C replacement', 2957.00, 295.70),
         ('INV-10619', 'UC engine remediation', 671.00, 67.10)]
BILLS = [('INV41565', 'Claim 2 of 3, Aware Super Seminar Room', 161926.60, 16192.67),
         ('INV41554', 'Labour hire, YSG Studio', 3659.00, 365.90)]
OWN_STAFF = 2026.21

assert abs(sum(NOW_ACC[a] for g, r, c in GROUPS for a in r) - 1624459.78) < 0.02
assert abs(sum(NOW_ACC[a] for g, r, c in GROUPS for a in c) - 992160.28) < 0.02


def band(ws, row, cols, text):
    ws.cell(row, 1, text).font = HEAD
    for c in range(1, cols + 1):
        ws.cell(row, c).fill = BAND


def put(ws, r, c, v, bold=False, fill=None, fmt=MONEY, font=None):
    cell = ws.cell(r, c, v)
    cell.font = font or (BOLD if bold else BODY)
    cell.number_format = fmt
    if fill:
        cell.fill = fill
    return cell


wb = openpyxl.load_workbook(OUT)

# ---------------------------------------------------------- 1. summary
ws = wb['Summary']
ws.delete_rows(1, ws.max_row)
ws['A1'] = 'Revenue against its own cost, July to September 2026'
ws['A1'].font = Font(name='Calibri', sz=14, b=True, color=NAVY)
for i, t in enumerate([
        f'Xero, accrual, whole of entity, read {PULLED}.',
        'The Aware Super invoice has been recoded since the transaction export '
        f'was taken. {MOVED:,.2f} moved from Installation Labour to Equipment '
        'Sales, so the ledger now matches the invoice line for line.',
        'Equipment is no longer loss-making. The two lines the recode touched '
        'are shown with their old figures beside them.']):
    ws.cell(3 + i, 1, t).font = NOTE
r = 7
for c, h in ((1, 'Account'), (2, 'Revenue'), (3, 'Cost'), (4, 'Margin'),
             (5, 'Margin %'), (7, 'Was, pre-recode')):
    cell = ws.cell(r, c, h)
    cell.font = HEAD
    cell.fill = BAND
    if c > 1:
        cell.alignment = Alignment(horizontal='right')
for c in (6,):
    ws.cell(r, c).fill = BAND
r += 1
group_rows = []
for name, revs, costs in GROUPS:
    head = r
    ws.cell(r, 1, name.upper()).font = BOLD
    for c in range(1, 6):
        ws.cell(r, c).fill = TOTF
        ws.cell(r, c).border = Border(top=RULE)
    r += 1
    for kind, accts, col in (('rev', revs, 2), ('cost', costs, 3)):
        for a in sorted(accts, key=lambda x: -NOW_ACC.get(x, 0)):
            if a not in NOW_ACC:
                continue
            ws.cell(r, 1, '    ' + a).font = BODY if kind == 'rev' else SMALL
            put(ws, r, col, round(NOW_ACC[a], 2),
                font=BODY if kind == 'rev' else SMALL)
            if a in WAS_ACC:
                put(ws, r, 7, round(WAS_ACC[a], 2), fill=WAS, font=SMALL)
            r += 1
    rv = sum(NOW_ACC[a] for a in revs if a in NOW_ACC)
    cv = sum(NOW_ACC[a] for a in costs if a in NOW_ACC)
    put(ws, head, 2, round(rv, 2), bold=True, fill=TOTF)
    put(ws, head, 3, round(cv, 2), bold=True, fill=TOTF)
    put(ws, head, 4, f'=B{head}-C{head}', bold=True, fill=TOTF)
    put(ws, head, 5, f'=IF(B{head}=0,"",D{head}/B{head})', bold=True,
        fill=TOTF, fmt=PCT)
    group_rows.append(head)
    r += 1
ws.cell(r, 1, 'GROSS PROFIT').font = HEAD
for c in range(1, 6):
    ws.cell(r, c).fill = BAND
for c in (2, 3):
    put(ws, r, c, '=' + '+'.join(f'{gl(c)}{x}' for x in group_rows), font=HEAD)
put(ws, r, 4, f'=B{r}-C{r}', font=HEAD)
put(ws, r, 5, f'=IF(B{r}=0,"",D{r}/B{r})', font=HEAD, fmt=PCT)
r += 2
for t in [f'Equipment sold was {WAS_ACC["Equipment Sales"]:,.2f} against '
          f'{NOW_ACC["Equipment - Purchase"]:,.2f} of cost before the recode, '
          f'a loss of {WAS_ACC["Equipment Sales"] - NOW_ACC["Equipment - Purchase"]:,.2f}. '
          f'It is now {NOW_ACC["Equipment Sales"]:,.2f} against the same cost.',
          'Total revenue, gross profit and net profit are the same either way. '
          'Only the split between two revenue lines moved.',
          'Licences and subscriptions is unchanged and is the one line still '
          'worth looking at.']:
    ws.cell(r, 1, t).font = NOTE
    r += 1
for c, w in ((1, 44), (2, 17), (3, 17), (4, 17), (5, 12), (6, 3), (7, 18)):
    ws.column_dimensions[gl(c)].width = w
ws.sheet_view.showGridLines = False

# -------------------------------------------------------------- 2. GST
if 'GST' in wb.sheetnames:
    del wb['GST']
ws = wb.create_sheet('GST', 1)
ws['A1'] = 'Aware Super and Complete AV Solutions, September 2026'
ws['A1'].font = Font(name='Calibri', sz=14, b=True, color=NAVY)
ws['A2'] = (f'Straight off the invoices and bills in Xero, read {PULLED}. '
            'All GST exclusive at 10 per cent. All paid.')
ws['A2'].font = NOTE
r = 4
for title, items, who in (('INVOICED TO AWARE SUPER', SALES, 'Aware Super'),
                          ('BILLED BY COMPLETE AV SOLUTIONS', BILLS,
                           'Complete AV Solutions Pty Ltd')):
    band(ws, r, 5, title)
    r += 1
    for c, h in ((1, 'Number'), (2, 'What for'), (3, 'Net'), (4, 'GST'),
                 (5, 'Gross')):
        cell = ws.cell(r, c, h)
        cell.font = BOLD
        cell.border = Border(bottom=RULE)
        if c > 2:
            cell.alignment = Alignment(horizontal='right')
    r += 1
    first = r
    for num, what, net, gst in items:
        ws.cell(r, 1, num).font = BODY
        ws.cell(r, 2, what).font = BODY
        put(ws, r, 3, round(net, 2))
        put(ws, r, 4, round(gst, 2))
        put(ws, r, 5, f'=C{r}+D{r}')
        r += 1
    ws.cell(r, 1, 'Total').font = BOLD
    for c in (3, 4, 5):
        put(ws, r, c, f'=SUM({gl(c)}{first}:{gl(c)}{r - 1})', bold=True)
    for c in range(1, 6):
        ws.cell(r, c).fill = TOTF
    if title.startswith('INVOICED'):
        sales_tot = r
    else:
        bills_tot = r
    r += 2

band(ws, r, 5, 'THE AWARE SUPER JOB, 2604704')
r += 1
for label, net, gst, fill in (
        ('Invoiced, INV-10596', SALES[0][2], SALES[0][3], NOW),
        ('Cost, INV41565', BILLS[0][2], BILLS[0][3], WAS),
        ('Cost, own staff, no GST', OWN_STAFF, 0.0, WAS)):
    ws.cell(r, 1, label).font = BODY
    put(ws, r, 3, round(net, 2), fill=fill)
    put(ws, r, 4, round(gst, 2), fill=fill)
    put(ws, r, 5, f'=C{r}+D{r}', fill=fill)
    r += 1
ws.cell(r, 1, 'Margin, and the GST payable on the job').font = BOLD
put(ws, r, 3, f'=C{r - 3}-C{r - 2}-C{r - 1}', bold=True, fill=TOTF)
put(ws, r, 4, f'=D{r - 3}-D{r - 2}-D{r - 1}', bold=True, fill=TOTF)
put(ws, r, 5, f'=C{r}+D{r}', bold=True, fill=TOTF)
for c in range(1, 6):
    ws.cell(r, c).fill = TOTF
r += 1
put(ws, r, 3, f'=C{r - 1}/C{r - 4}', bold=True, fmt=PCT)
ws.cell(r, 1, 'Margin on the job, net of GST').font = BOLD
r += 2
for t in ['The GST column is what is collected on the invoice and claimed on '
          'the bill. The difference is what the job contributes to the BAS.',
          'Own staff carries no GST, because it is payroll.',
          'The margin line is net of GST on both sides, so it is the same '
          '41,399.36 that shows on the job.']:
    ws.cell(r, 1, t).font = NOTE
    r += 1
for c, w in ((1, 36), (2, 44), (3, 16), (4, 15), (5, 16)):
    ws.column_dimensions[gl(c)].width = w
ws.sheet_view.showGridLines = False

# ------------------------------------- 3. flag the pre-recode tabs
# The flag goes in a spare cell to the right. Inserting a row would not move
# the formula references on these tabs, and would break them.
FLAG = Font(name='Calibri', sz=10, i=True, color='FFC00000')
MSG = ('PRE-RECODE. These lines are the invoice as it was first coded. Xero '
       f'has since been corrected: {MOVED:,.2f} moved from Installation '
       'Labour to Equipment Sales, so equipment on this job is now '
       '153,904.89 invoiced against 124,635.75 of cost, a 19.0% margin, and '
       'labour is 51,447.28 against 39,317.06, a 23.6% margin. The job total '
       'does not change. See the Summary and GST tabs for the current '
       'position.')
for tab, cellref in (('Detail', 'M1'),
                     ('Aware invoice, was and now', 'A4'),
                     ('Aware Super job', 'H1')):
    ws = wb[tab]
    ws[cellref] = MSG
    ws[cellref].font = FLAG
    print(f'  flagged {tab} at {cellref}')

wb.save(OUT)
print('saved', OUT)
print('  tabs:', wb.sheetnames)
print(f'  moved from Installation Labour to Equipment Sales: {MOVED:,.2f}')
