"""Turnover for the VIC Labour Hire Authority portal, four quarters to 30 Sep 2026.

Turnover is defined in clause 3 of the Labour Hire Licensing Regulations 2018
(Vic) as the total ordinary income derived in the course of running a business.
LHA says it covers every part of the business, not just labour hire, and every
jurisdiction, not just Victoria. It excludes GST and is rounded down to the
nearest dollar.

Figures are accrual, straight from Xero, whole of entity, no tracking filter.
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter as gl

OUT = 'lha/CTS turnover - 4 quarters to 30 September 2026.xlsx'
PULLED = 'Xero report data last refreshed 30 October 2026'

NAVY, WHITE, GREY = 'FF242C39', 'FFFFFFFF', 'FFF4F4F4'
BAND = PatternFill('solid', start_color=NAVY, end_color=NAVY)
TOT = PatternFill('solid', start_color=GREY, end_color=GREY)
RULE = Side(style='medium', color=NAVY)
MONEY = '#,##0.00;[RED]\\(#,##0.00\\);\\-'
WHOLE = '#,##0;[RED]\\(#,##0\\);\\-'
BODY = Font(name='Calibri', sz=11, color=NAVY)
BOLD = Font(name='Calibri', sz=11, b=True, color=NAVY)
HEAD = Font(name='Calibri', sz=11, b=True, color=WHITE)
NOTE = Font(name='Calibri', sz=10, color='FF6B7585')

# quarter, BAS period, trading income, interest received, other income
Q = [('Oct to Dec 2025', 'December 2025 quarter', 1894997.35,  502.76, 481.12),
     ('Jan to Mar 2026', 'March 2026 quarter',    1303788.88,  982.89, 531.47),
     ('Apr to Jun 2026', 'June 2026 quarter',     1685817.62, 9454.03, 334.24),
     ('Jul to Sep 2026', 'September 2026 quarter',1624459.78, 2924.66, 475.54)]

wb = openpyxl.Workbook()
ws = wb.active
ws.title = 'Turnover'

ws['A1'] = 'Turnover for the four quarters to 30 September 2026'
ws['A1'].font = Font(name='Calibri', sz=14, b=True, color=NAVY)
ws['A2'] = 'Corporate Technology Services Pty Ltd'
ws['A2'].font = BOLD
for i, t in enumerate([
        'For the Victorian Labour Hire Authority portal. Turnover is the total '
        'ordinary income derived in the course of running the business, under '
        'clause 3 of the Labour Hire Licensing Regulations 2018 (Vic).',
        'Whole of entity. Every part of the business, not just labour hire. '
        'Every jurisdiction, not just Victoria. Accrual basis, excludes GST.',
        f'Source: Xero profit and loss, no tracking filter. {PULLED}.']):
    ws.cell(4 + i, 1, t).font = NOTE

R = 8
for c, h in enumerate(['Quarter', 'BAS period', 'Trading income',
                       'Interest received', 'Other income',
                       'Total ordinary income'], start=1):
    cell = ws.cell(R, c, h)
    cell.font = HEAD
    cell.fill = BAND
    cell.alignment = Alignment(horizontal='right' if c > 2 else 'left',
                               wrap_text=True)

r = R + 1
for name, bas, trade, interest, other in Q:
    ws.cell(r, 1, name).font = BODY
    ws.cell(r, 2, bas).font = BODY
    for c, v in ((3, trade), (4, interest), (5, other)):
        cell = ws.cell(r, c, v)
        cell.font = BODY
        cell.number_format = MONEY
    cell = ws.cell(r, 6, f'=SUM(C{r}:E{r})')
    cell.font = BOLD
    cell.number_format = MONEY
    r += 1

ws.cell(r, 1, 'Four quarters').font = BOLD
for c in range(1, 7):
    ws.cell(r, c).fill = TOT
    ws.cell(r, c).border = Border(top=RULE)
for c in range(3, 7):
    cell = ws.cell(r, c, f'=SUM({gl(c)}{R + 1}:{gl(c)}{r - 1})')
    cell.font = BOLD
    cell.number_format = MONEY
total_row = r

r += 2
ws.cell(r, 1, 'THE FIGURE TO ENTER ON THE PORTAL').font = HEAD
for c in range(1, 7):
    ws.cell(r, c).fill = BAND
r += 1
ws.cell(r, 1, 'Total ordinary income, rounded down to the nearest dollar').font = BOLD
cell = ws.cell(r, 6, f'=ROUNDDOWN(F{total_row},0)')
cell.font = Font(name='Calibri', sz=14, b=True, color=NAVY)
cell.number_format = WHOLE
cell.fill = TOT
answer_row = r

r += 2
ws.cell(r, 1, 'If the Authority wants the figure to tie to the BAS instead').font = BOLD
r += 1
ws.cell(r, 1, 'Trading income only, rounded down. Interest is not a taxable '
              'supply, so it is not in G1 on the BAS.').font = BODY
cell = ws.cell(r, 6, f'=ROUNDDOWN(C{total_row},0)')
cell.font = BOLD
cell.number_format = WHOLE

r += 2
for t in ['Licence tier: both figures sit in Tier 2, turnover between '
          '$2,000,001 and $10,000,000, so the fee is the same either way.',
          'Check before lodging: if CTS reports GST on a cash basis, the BAS '
          'will not agree to these accrual figures and the turnover should be '
          'restated on the same basis as the BAS.',
          'The four BAS to upload are the December 2025, March 2026, June 2026 '
          'and September 2026 quarters.']:
    ws.cell(r, 1, t).font = NOTE
    r += 1

ws.column_dimensions['A'].width = 56
ws.column_dimensions['B'].width = 24
for c in range(3, 7):
    ws.column_dimensions[gl(c)].width = 20
ws.row_dimensions[R].height = 30
ws.sheet_view.showGridLines = False
wb.save(OUT)
print('saved', OUT)
