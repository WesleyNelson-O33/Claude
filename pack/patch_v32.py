"""Four charts she asked for, with their workings.

  Top 10 Clients per department   on the Top 10 Clients sheet, the ten biggest
                                  customers split across the departments
  Other operating expenses        the accounts inside Operating Expenses
  Operating cost distribution     where every dollar of operating cost goes
  Labour mix                      direct against indirect labour, line by line

The labour split is her own: the Type column on Lists marks an account Labour,
and the Category says whether it is a direct cost or an overhead. Nothing here
invents a grouping.
"""
import os, json, copy
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as gl

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = f'{HERE}/CTS Financial Controller Pack - v20.xlsx'
B = json.load(open(f'{HERE}/data/brand.json'))
INK, WHITE, GREY, GREYL = B['charcoal'], B['white'], B['grey'], B['grey_lt']
AQUA = B['bright_teal']
MONEY = '\\$#,##0;[Red]"($"#,##0\\);\\-'
COST = '[Red]"($"#,##0\\);\\$#,##0;\\-'
PCT = '0.0%;[Red](0.0%);\\-'

CA, CB, CD = ('Cleanup!$A$5:$A$20004', 'Cleanup!$B$5:$B$20004', 'Cleanup!$D$5:$D$20004')
CE, CF, CI, CJ = ('Cleanup!$E$5:$E$20004', 'Cleanup!$F$5:$F$20004',
                  'Cleanup!$I$5:$I$20004', 'Cleanup!$J$5:$J$20004')
YTD = (f'{CB},">="&DATE(Setup!$B$7-1,Setup!$B$6,1),{CB},"<="&Setup!$B$5')

DEPTS = [('Support', 'ONSITE'), ('Events', 'PRODUCTION'), ('Video', 'VIDEO'),
         ('Integration', 'INTEGRATION'), ('Consulting', 'CONSULTING'),
         ('General Admin', 'ADMIN')]
COSBUCKET = {
 'Labour': ['Sub-Contract Labour', 'Service - hires', 'Direct Wages & Salaries',
            'Direct Salaries', 'Direct Wages', 'Direct Superannuation',
            'Direct Workers Comp'],
 'Equipment & Services': ['Equipment - Purchase', 'Equipment - Hires',
                          'Subscriptions & Licences', 'Equipment -Maintenance Support',
                          'Stamp Duty', 'Damage Waiver', 'Credit Card fees paid - COGS'],
 'Transportation': ['Supplier Per Diems'],
 'Other direct': ['Other Direct Expenses'],
}
INDIRECT = ['Payroll Tax', 'Superannuation', 'Wages & Salaries', 'Indirect Salaries*',
            'Indirect Wages*', "Workers' Compensation", 'Leave expense*']
OHHEAD = {'General & Administrative Exp': 7, 'Business Insurances': 25,
          'Advertising & Promotion Exp': 30, 'Operating Expenses': 32,
          'Employment Expenses': 57, 'Credit Card fees paid': 83}
OPS_FIRST, OPS_LAST = 33, 56           # the Operating Expenses accounts

wb = openpyxl.load_workbook(OUT)
PL = wb['P&L FY27']
thin = Side(style='thin', color=GREY)
hair = Side(style='medium', color=AQUA)

def banner(ws, r, text, ncol):
    ws.cell(r, 1, text)
    for c in range(1, ncol + 1):
        ws.cell(r, c).fill = PatternFill('solid', fgColor=INK)
        ws.cell(r, c).font = Font(name='Calibri', size=12, bold=True, color=WHITE)

def header(ws, r, labels):
    for i, t in enumerate(labels, start=1):
        c = ws.cell(r, i, t)
        c.font = Font(name='Calibri', size=10, bold=True, color=WHITE)
        c.fill = PatternFill('solid', fgColor=INK)
        c.alignment = Alignment(horizontal='right' if i > 1 else 'left', wrap_text=True)
        c.border = Border(bottom=hair)

def plain(ws, r, ncol, bold=False, fill=None):
    for c in range(1, ncol + 1):
        cell = ws.cell(r, c)
        cell.font = Font(name='Calibri', size=11, bold=bold, color=INK)
        cell.border = Border(bottom=thin)
        if fill:
            cell.fill = PatternFill('solid', fgColor=fill)
        if c > 1:
            cell.alignment = Alignment(horizontal='right')

# =============================== 1. top ten by department ==================
ws = wb['Top 10 Clients']
S = 105
banner(ws, S, 'TOP TEN BY DEPARTMENT, YEAR TO DATE', 8)
header(ws, S + 1, ['Customer'] + [d for d, _ in DEPTS] + ['Total'])
for i in range(10):
    r, src = S + 2 + i, 7 + i
    ws.cell(r, 1, f'=IF($B{src}="","",$B{src})')
    for j, (_, tag) in enumerate(DEPTS):
        ws.cell(r, 2 + j, f'=IF($A{r}="","",SUMIFS({CA},{CJ},$A{r},{CD},"{tag}",'
                          f'{CE},"Income",{YTD}))').number_format = MONEY
    ws.cell(r, 8, f'=IF($A{r}="","",SUM($B{r}:$G{r}))').number_format = MONEY
    plain(ws, r, 8)
TOT = S + 12
ws.cell(TOT, 1, 'Top ten together')
for c in range(2, 9):
    ws.cell(TOT, c, f'=SUM({gl(c)}{S+2}:{gl(c)}{S+11})').number_format = MONEY
plain(ws, TOT, 8, bold=True, fill=GREYL)
ws.cell(TOT + 1, 1, 'Check, against the top ten total above')
ws.cell(TOT + 1, 8, f'=ROUND($H{TOT}-$D$17,2)').number_format = MONEY
plain(ws, TOT + 1, 8, bold=True)
print(f'Top 10 Clients: by department block rows {S} to {TOT+1}')

# =============================== 2, 3, 4 on one new sheet ==================
if 'Cost Analysis' in wb.sheetnames:
    del wb['Cost Analysis']
ws = wb.create_sheet('Cost Analysis', wb.sheetnames.index('Overheads MoM') + 1)
ws.sheet_properties.tabColor = INK
ws['A1'] = 'Cost analysis'
ws['A1'].font = Font(name='Calibri', size=16, bold=True, color=INK)
ws['A2'] = '=TEXT(Setup!$B$5,"mmm yyyy")&"   |   year to date"'
ws['A2'].font = Font(name='Calibri', size=10, bold=True, color=INK)
OH = "'Overheads MoM'"

# ---- other operating expenses, ranked -------------------------------------
A0 = 4
banner(ws, A0, 'OTHER OPERATING EXPENSES, YEAR TO DATE', 4)
header(ws, A0 + 1, ['Account', 'Year to date', 'Share of operating expenses', ''])
N = OPS_LAST - OPS_FIRST + 1
W0 = A0 + N + 6                       # the workings sit below
for i in range(N):
    r = A0 + 2 + i
    key = f'MATCH(LARGE(${gl(3)}${W0+1}:${gl(3)}${W0+N},{i+1}),${gl(3)}${W0+1}:${gl(3)}${W0+N},0)'
    ws.cell(r, 1, f'=INDEX(${gl(1)}${W0+1}:${gl(1)}${W0+N},{key})')
    ws.cell(r, 2, f'=INDEX(${gl(2)}${W0+1}:${gl(2)}${W0+N},{key})').number_format = COST
    ws.cell(r, 3, f'=IFERROR($B{r}/{OH}!$N${OHHEAD["Operating Expenses"]},0)')
    ws.cell(r, 3).number_format = PCT
    plain(ws, r, 3)
OT = A0 + N + 2
ws.cell(OT, 1, 'Operating expenses in total')
ws.cell(OT, 2, f'={OH}!$N${OHHEAD["Operating Expenses"]}').number_format = COST
ws.cell(OT, 3, f'=IF($B{OT}=0,"",SUM($B${A0+2}:$B${A0+1+N})/$B{OT})').number_format = PCT
plain(ws, OT, 3, bold=True, fill=GREYL)

banner(ws, W0 - 1, 'THE WORKINGS, STRAIGHT OFF THE OVERHEADS MONTH BY MONTH TAB', 4)
header(ws, W0, ['Account', 'Year to date', 'Sort key', ''])
for i in range(N):
    r, src = W0 + 1 + i, OPS_FIRST + i
    ws.cell(r, 1, f'=TRIM({OH}!$A${src})')
    ws.cell(r, 2, f'={OH}!$N${src}').number_format = COST
    ws.cell(r, 3, f'=-$B{r}+ROW()/1000000')
    plain(ws, r, 3)
print(f'operating expenses: chart rows {A0+2} to {A0+1+N}, workings {W0+1} to {W0+N}')

# ---- operating cost distribution -----------------------------------------
C0 = W0 + N + 3
banner(ws, C0, 'OPERATING COST DISTRIBUTION, YEAR TO DATE', 4)
header(ws, C0 + 1, ['Cost bucket', 'Year to date', 'Share of operating cost', ''])
rows = []
r = C0 + 2
for name, subs in COSBUCKET.items():
    arr = ';'.join(f'"{x}"' for x in subs)
    ws.cell(r, 1, 'Cost of sales, ' + name.lower())
    ws.cell(r, 2, f'=SUM(SUMIFS({CA},{CF},{{{arr}}},{CE},"Cost of Sales",{YTD}))')
    ws.cell(r, 2).number_format = COST
    rows.append(r); r += 1
for name, hr in OHHEAD.items():
    ws.cell(r, 1, 'Overhead, ' + name)
    ws.cell(r, 2, f'={OH}!$N${hr}').number_format = COST
    rows.append(r); r += 1
CT = r
ws.cell(CT, 1, 'Operating cost in total')
ws.cell(CT, 2, f'=SUM($B${rows[0]}:$B${rows[-1]})').number_format = COST
for rr in rows:
    ws.cell(rr, 3, f'=IFERROR($B{rr}/$B${CT},0)').number_format = PCT
    plain(ws, rr, 3)
ws.cell(CT, 3, 1).number_format = PCT
plain(ws, CT, 3, bold=True, fill=GREYL)
ws.cell(CT + 1, 1, 'Check, cost of sales plus overheads per the P&L')
ws.cell(CT + 1, 2, "=ROUND($B$%d-('P&L FY27'!$O$59+'P&L FY27'!$O$164),2)" % CT)
ws.cell(CT + 1, 2).number_format = COST
plain(ws, CT + 1, 3, bold=True)
print(f'cost distribution: rows {rows[0]} to {CT}')

# ---- labour mix -----------------------------------------------------------
L0 = CT + 3
banner(ws, L0, 'LABOUR MIX, YEAR TO DATE', 4)
header(ws, L0 + 1, ['Labour line', 'Direct', 'Indirect', 'Share of all labour'])
lrows = []
r = L0 + 2
for sub in COSBUCKET['Labour']:
    ws.cell(r, 1, sub)
    ws.cell(r, 2, f'=SUMIFS({CA},{CF},"{sub}",{CE},"Cost of Sales",{YTD})')
    ws.cell(r, 2).number_format = COST
    ws.cell(r, 3, 0).number_format = COST
    lrows.append(r); r += 1
for acc in INDIRECT:
    ws.cell(r, 1, acc.replace('*', ''))
    ws.cell(r, 2, 0).number_format = COST
    ws.cell(r, 3, f'=SUMIFS({CA},{CI},"{acc}",{CE},"Expenses",{YTD})')
    ws.cell(r, 3).number_format = COST
    lrows.append(r); r += 1
LT = r
ws.cell(LT, 1, 'All labour')
for c in (2, 3):
    ws.cell(LT, c, f'=SUM({gl(c)}${lrows[0]}:{gl(c)}${lrows[-1]})').number_format = COST
for rr in lrows:
    ws.cell(rr, 4, f'=IFERROR(($B{rr}+$C{rr})/($B${LT}+$C${LT}),0)').number_format = PCT
    plain(ws, rr, 4)
ws.cell(LT, 4, 1).number_format = PCT
plain(ws, LT, 4, bold=True, fill=GREYL)
# Not every account under Employment Expenses is labour. Staff amenities and
# staff entertainment sit there too, and her Type column does not call them
# labour, so they are not in the mix above. This line is that balance, not a
# check that should come to nil.
ws.cell(LT + 1, 1, 'Employment Expenses that is not labour, staff amenities and the like')
ws.cell(LT + 1, 3, f'={OH}!$N${OHHEAD["Employment Expenses"]}-$C${LT}')
ws.cell(LT + 1, 3).number_format = COST
plain(ws, LT + 1, 4, bold=True)
print(f'labour mix: rows {lrows[0]} to {LT}')

ws.column_dimensions['A'].width = 38
for c, w in (('B', 16), ('C', 16), ('D', 18)):
    ws.column_dimensions[c].width = w

json.dump({'ops_first': A0 + 2, 'ops_n': N, 'ops_total': OT,
           'dist_first': rows[0], 'dist_last': rows[-1], 'dist_total': CT,
           'lab_first': lrows[0], 'lab_last': lrows[-1], 'lab_total': LT,
           'top10_dept': S}, open(f'{HERE}/data/costcharts.json', 'w'), indent=1)
wb.save(OUT)
print('saved')
