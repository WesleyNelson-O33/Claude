"""
Three working blocks so the last three deck graphs can live in the pack itself.

The deck is built from snips of this workbook, so every chart on a slide has to
exist here. These three did not:

  Cost Analysis   revenue by department, the month against budget and last year
  Utilisation     chargeable, non-chargeable and leave hours by department
  Overheads MoM   employment expense split, office against departmental staff

Each block reads the month on Setup, so all three follow the dropdown.
"""
import openpyxl, sys
from openpyxl.styles import Font, Alignment

WB = sys.argv[1] if len(sys.argv) > 1 else 'CTS Financial Controller Pack - v20.xlsx'
DEPTS = [('Support', 41), ('Events', 8), ('Video', 16), ('Integration', 25), ('Consulting', 33)]
BOLD = Font(bold=True, size=10)
HEADNOTE = Font(bold=True, size=10)

wb = openpyxl.load_workbook(WB)

# ---- A. Cost Analysis: revenue by department, actual, budget, last year ----
c = wb['Cost Analysis']
c['A190'] = 'REVENUE BY DEPARTMENT, THE MONTH AGAINST BUDGET AND LAST YEAR'
c['A190'].font = HEADNOTE
c['A191'], c['B191'], c['C191'], c['D191'] = 'Department', 'Actual', 'Budget', 'Last year'
for cell in ('A191', 'B191', 'C191', 'D191'):
    c[cell].font = BOLD
    c[cell].alignment = Alignment(horizontal='right' if cell[0] != 'A' else 'left')
for i, (name, r) in enumerate(DEPTS):
    rr = 192 + i
    c[f'A{rr}'] = name
    c[f'B{rr}'] = f"=SUMIF('P&L FY27'!$B$5:$M$5,Setup!$B$5,'P&L FY27'!$B{r}:$M{r})"
    c[f'C{rr}'] = f"=SUMIF('Budget FY27'!$B$5:$M$5,Setup!$B$5,'Budget FY27'!$B{r}:$M{r})"
    c[f'D{rr}'] = f"=SUMIF('Actual FY26'!$B$5:$M$5,Setup!$B$5,'Actual FY26'!$B{r}:$M{r})"
    for col in 'BCD':
        c[f'{col}{rr}'].number_format = '#,##0;[Red](#,##0);-'
c['A197'] = 'Total'
c['A197'].font = BOLD
for col in 'BCD':
    c[f'{col}197'] = f'=SUM({col}192:{col}196)'
    c[f'{col}197'].number_format = '#,##0;[Red](#,##0);-'
    c[f'{col}197'].font = BOLD

# ---- B. Utilisation: hours by department for the month on Setup ----
u = wb['Utilisation']
u['A66'] = 'HOURS BY DEPARTMENT, THE MONTH ON SETUP'
u['A66'].font = HEADNOTE
u['A67'], u['B67'], u['C67'], u['D67'] = 'Department', 'Chargeable', 'Non-chargeable', 'Leave'
for cell in ('A67', 'B67', 'C67', 'D67'):
    u[cell].font = BOLD
    u[cell].alignment = Alignment(horizontal='right' if cell[0] != 'A' else 'left')
for i, (name, row) in enumerate([('Support', 9), ('Events', 15), ('Video', 21), ('Integration', 27), ('Consulting', 33)]):
    rr = 68 + i
    u[f'A{rr}'] = name
    for j, col in enumerate('BCD'):
        u[f'{col}{rr}'] = f'=INDEX($D{row + j}:$O{row + j},MATCH(Setup!$B$5,$D$4:$O$4,0))'
        u[f'{col}{rr}'].number_format = '#,##0.00'
u['A73'] = 'Total'
u['A73'].font = BOLD
for col in 'BCD':
    u[f'{col}73'] = f'=SUM({col}68:{col}72)'
    u[f'{col}73'].number_format = '#,##0.00'
    u[f'{col}73'].font = BOLD

# ---- C. Overheads MoM: employment expense split for the month on Setup ----
m = wb['Overheads MoM']
m['A160'] = 'EMPLOYMENT EXPENSE SPLIT, THE MONTH ON SETUP'
m['A160'].font = HEADNOTE
m['A161'], m['B161'], m['C161'] = 'Department', 'Departmental staff', 'Office staff'
for cell in ('A161', 'B161', 'C161'):
    m[cell].font = BOLD
    m[cell].alignment = Alignment(horizontal='right' if cell[0] != 'A' else 'left')
for i, (name, dr, orow) in enumerate([('Support', 133, 134), ('Events', 137, 138), ('Video', 141, 142),
                                      ('Integration', 145, 146), ('Consulting', 149, 150)]):
    rr = 162 + i
    m[f'A{rr}'] = name
    m[f'B{rr}'] = f'=-INDEX($B{dr}:$M{dr},MATCH(Setup!$B$5,$B$132:$M$132,0))'
    m[f'C{rr}'] = f'=-INDEX($B{orow}:$M{orow},MATCH(Setup!$B$5,$B$132:$M$132,0))'
    for col in 'BC':
        m[f'{col}{rr}'].number_format = '#,##0;[Red](#,##0);-'
m['A167'] = 'Total'
m['A167'].font = BOLD
for col in 'BC':
    m[f'{col}167'] = f'=SUM({col}162:{col}166)'
    m[f'{col}167'].number_format = '#,##0;[Red](#,##0);-'
    m[f'{col}167'].font = BOLD

wb.save(WB)
print('three chart blocks added: Cost Analysis 190, Utilisation 66, Overheads MoM 160')
