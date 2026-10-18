"""
Section 7 on Overhead Split: the department result table the board sees.

Revenue, budget, variance, gross profit, gross margin, all overhead and the
result, one row per department. Everything was already on this tab in pieces;
this puts it in the single block the deck snips.

The five department results add back to net profit, which the check row proves.
"""
import openpyxl, sys
from openpyxl.styles import Font, Alignment

WB = sys.argv[1] if len(sys.argv) > 1 else 'CTS Financial Controller Pack - v20.xlsx'
# department, Summary income row, Summary cost of sales row, own overhead row, allocated row
DEPTS = [('Support', 41, 130, 101, 90), ('Events', 8, 61, 102, 91), ('Video', 16, 78, 103, 92),
         ('Integration', 25, 96, 104, 93), ('Consulting', 33, 113, 105, 94)]
HEADS = ['Department', 'Revenue', 'Budget', 'Var', 'Gross profit', 'GM %', 'All overhead', 'Result']
BOLD = Font(bold=True, size=10)
MONEY = '#,##0;[Red](#,##0);-'

wb = openpyxl.load_workbook(WB)
o = wb['Overhead Split']
o['A140'] = '7.  DEPARTMENT RESULT, THE MONTH ON SETUP'
o['A140'].font = Font(bold=True, size=10)
o['A141'] = ('Revenue less direct costs less all overhead. All overhead is the department’s own '
             'overhead from section 5 plus its share of admin from section 4.')
for i, h in enumerate(HEADS):
    cell = o.cell(142, 1 + i)
    cell.value = h
    cell.font = BOLD
    cell.alignment = Alignment(horizontal='left' if i == 0 else 'right')

for i, (name, ri, rc, ro, ra) in enumerate(DEPTS):
    r = 143 + i
    o.cell(r, 1).value = name
    o.cell(r, 2).value = f"=SUMIF('P&L FY27'!$B$5:$M$5,Setup!$B$5,'P&L FY27'!$B{ri}:$M{ri})"
    o.cell(r, 3).value = f"=SUMIF('Budget FY27'!$B$5:$M$5,Setup!$B$5,'Budget FY27'!$B{ri}:$M{ri})"
    o.cell(r, 4).value = f'=B{r}-C{r}'
    o.cell(r, 5).value = (f"=B{r}+SUMIF('P&L FY27'!$B$5:$M$5,Setup!$B$5,'P&L FY27'!$B{rc}:$M{rc})")
    o.cell(r, 6).value = f'=IFERROR(E{r}/B{r},0)'
    o.cell(r, 7).value = f'=INDEX($C{ro}:$N{ro},MATCH(Setup!$B$5,$C$100:$N$100,0))+INDEX($C{ra}:$N{ra},MATCH(Setup!$B$5,$C$89:$N$89,0))'
    o.cell(r, 8).value = f'=E{r}+G{r}'
    for col in (2, 3, 4, 5, 7, 8):
        o.cell(r, col).number_format = MONEY
    o.cell(r, 6).number_format = '0.0%'

o.cell(148, 1).value = 'Total'
o.cell(148, 1).font = BOLD
for col in (2, 3, 4, 5, 7, 8):
    c = o.cell(148, col)
    c.value = f'=SUM({chr(64 + col)}143:{chr(64 + col)}147)'
    c.number_format = MONEY
    c.font = BOLD
o.cell(148, 6).value = '=IFERROR(E148/B148,0)'
o.cell(148, 6).number_format = '0.0%'
o.cell(148, 6).font = BOLD

o.cell(150, 1).value = 'Check, the five results against net profit, should be nil'
o.cell(150, 2).value = "=H148-SUMIF('P&L FY27'!$B$5:$M$5,Setup!$B$5,'P&L FY27'!$B180:$M180)"
o.cell(150, 2).number_format = MONEY

wb.save(WB)
print('Overhead Split section 7 added at row 140')
