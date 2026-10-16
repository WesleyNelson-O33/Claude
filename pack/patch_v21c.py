"""Dept P&L: one department at a time, against last year and against budget.

Her layout. Pick the department from the dropdown and the whole card moves.

  This month   straight off the P&L by Dept tab, so it ties to the P&L
  Last year    off the Actual FY26 tab
  Budget       off the Budget FY27 tab

COGS is grouped the way her account master groups it, by the Type column on
Lists: Labour, Equipment & Services, Transport. Other Direct Expenses spans
more than one Type, so it keeps its own line rather than being forced into one.
"""
import os, json, copy
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter as gl

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = f'{HERE}/CTS Financial Controller Pack - v20.xlsx'
B = json.load(open(f'{HERE}/data/brand.json'))
D = json.load(open(f'{HERE}/data/dept_colours.json'))
BUD = json.load(open(f'{HERE}/data/budget_fy27.json'))

TEAL, AQUA, BLUE, NAVY = B['dark_teal'], B['bright_teal'], B['primary_blue'], B['deep_blue']
RED, INK, WHITE, GREY, GREYL = (B['accent_red'], B['charcoal'], B['white'],
                                B['grey'], B['grey_lt'])
MONEY = '\\$#,##0;[Red]"($"#,##0\\);\\-'
COST  = '[Red]"($"#,##0\\);\\$#,##0;\\-'          # costs read positive
PCT   = '0.0%;[Red](0.0%);\\-'
VAR   = '0.0%;[Red](0.0%);\\-'

# department -> tag, FY26/Budget income total row, COS total row, COS first row,
#               column on P&L by Dept, column on the Overhead Split share block
# tag, FY26/Budget income total row, COS total row, COS first row,
# first row of the department's block on the Engine, row on the overhead split
DEPT = {
 'Support':       ('ONSITE',      41, 130, 131,   5, 1),
 'Events':        ('PRODUCTION',   8,  61,  62,  36, 2),
 'Video':         ('VIDEO',       16,  78,  79,  67, 3),
 'Integration':   ('INTEGRATION', 25,  96,  97,  98, 4),
 'Consulting':    ('CONSULTING',  33, 113, 114, 129, 5),
 'General Admin': ('ADMIN',       49, 147, 148, 160, 0),
}
ORDER = list(DEPT)
# the sixteen cost of sales subcategories, in the order every block holds them
COS = ['Equipment - Purchase','Equipment - Hires','Sub-Contract Labour','Service - hires',
       'Subscriptions & Licences','Equipment -Maintenance Support','Stamp Duty',
       'Damage Waiver','Supplier Per Diems','Direct Wages & Salaries','Direct Salaries',
       'Direct Wages','Direct Superannuation','Direct Workers Comp','Other Direct Expenses',
       'Credit Card fees paid - COGS']
TYPE = {'Labour': [3,4,10,11,12,13,14],
        'Equipment & Services': [1,2,5,6,7,8,16],
        'Transportation': [9],
        'Other Direct Expenses': [15]}
# her expense lines, and the row each sits on in the P&L by Dept month block
EXP = [('Business Insurances', 35), ('Employment Expenses', 38),
       ('General & Administrative Exp', 34), ('Advertising & Promotion Exp', 36),
       ('Operating Expenses', 37), ('Credit Card fees paid', 39)]
# where each overhead heading sits inside the department's six Engine rows
EXPOFF = {'General & Administrative Exp': 1, 'Business Insurances': 2,
          'Advertising & Promotion Exp': 3, 'Operating Expenses': 4,
          'Employment Expenses': 5, 'Credit Card fees paid': 6}

wb = openpyxl.load_workbook(OUT)
if 'Dept P&L' in wb.sheetnames:
    del wb['Dept P&L']
ws = wb.create_sheet('Dept P&L', wb.sheetnames.index('P&L by Dept') + 1)
ws.sheet_properties.tabColor = TEAL

# ---------------------------------------------------- 1. budget expenses ---
# The Budget FY27 tab holds overheads at company level only, but the budget
# itself is set by department. Put that detail on the budget tab, where it
# belongs, and read it from here.
bud = wb['Budget FY27']
BR = 195
bud.cell(BR, 1, 'Expenses by department, as the budget was set')
bud.cell(BR, 1).font = Font(name='Calibri', size=11, bold=True, color=WHITE)
for c in range(1, 15):
    bud.cell(BR, c).fill = PatternFill('solid', fgColor=TEAL)
bud.cell(BR + 1, 1, 'Department')
bud.cell(BR + 1, 2, 'Line')
for i in range(12):
    bud.cell(BR + 1, 3 + i, f"='Budget FY27'!{gl(2+i)}$5")
    bud.cell(BR + 1, 3 + i).number_format = 'mmm yy'
for c in range(1, 15):
    bud.cell(BR + 1, c).font = Font(name='Calibri', size=10, bold=True, color=WHITE)
    bud.cell(BR + 1, c).fill = PatternFill('solid', fgColor=INK)

EXPNAMES = [e[0] for e in EXP]
r = BR + 2
BUDROWS = {}
for dept, (tag, *_rest) in DEPT.items():
    if dept == 'General Admin':
        continue
    for sub in EXPNAMES:
        BUDROWS[(dept, sub)] = r
        bud.cell(r, 1, dept)
        bud.cell(r, 2, sub)
        for m in range(12):
            v = BUD.get(f'{tag}|Expenses|{sub}|{m}', 0.0)
            bud.cell(r, 3 + m, round(v, 2))
            bud.cell(r, 3 + m).number_format = COST
        r += 1
BEND = r - 1
print(f'budget expenses by department written on Budget FY27, rows {BR+2} to {BEND}')

# ---------------------------------------------------------- 2. workings ----
W = 40                                     # the workings start here
ws.cell(W, 1, 'THE WORKINGS').font = Font(name='Calibri', size=11, bold=True, color=WHITE)
for c in range(1, 10):
    ws.cell(W, c).fill = PatternFill('solid', fgColor=TEAL)

ws.cell(W + 1, 1, 'Department index')
for i, h in enumerate(['Department', 'Tag', 'Income total row', 'COS total row',
                       'COS first row', 'First row on the Engine',
                       'Row on the overhead split'], start=1):
    ws.cell(W + 2, i, h).font = Font(name='Calibri', size=9, bold=True, color=INK)
for i, d in enumerate(ORDER):
    t, inc, cost, cst, pc, oc = DEPT[d]
    for j, v in enumerate([d, t, inc, cost, cst, pc, oc], start=1):
        ws.cell(W + 3 + i, j, v)
DI = f'$A${W+3}:$A${W+3+len(ORDER)-1}'          # department names
IDX = f'MATCH($B$5,{DI},0)'

C0 = W + 11
ws.cell(C0, 1, 'Cost of sales, the sixteen subcategories and which group each belongs to')
for i, h in enumerate(['#', 'Subcategory', 'Labour', 'Equipment & Services',
                       'Transportation', 'Other Direct Expenses'], start=1):
    ws.cell(C0 + 1, i, h).font = Font(name='Calibri', size=9, bold=True, color=INK)
for i, nm in enumerate(COS):
    ws.cell(C0 + 2 + i, 1, i + 1)
    ws.cell(C0 + 2 + i, 2, nm)
    for j, ty in enumerate(['Labour', 'Equipment & Services', 'Transportation',
                            'Other Direct Expenses'], start=3):
        ws.cell(C0 + 2 + i, j, 1 if (i + 1) in TYPE[ty] else 0)
MASK = {ty: f'${gl(3+j)}${C0+2}:${gl(3+j)}${C0+17}'
        for j, ty in enumerate(['Labour', 'Equipment & Services', 'Transportation',
                                'Other Direct Expenses'])}

# ------------------------------------------------------- 3. the card -------
thin = Side(style='thin', color=GREY)
ws['A1'] = 'P&L by Department, the month'
ws['A1'].font = Font(name='Calibri', size=16, bold=True, color=INK)
ws['A2'] = ('Pick the department below. This month comes off the P&L by Dept tab so it '
            'ties to the P&L. Last year comes off Actual FY26 and budget off Budget FY27.')
ws['A2'].font = Font(name='Calibri', size=9, italic=True, color=INK)
ws['A3'] = '=TEXT(Setup!$B$5,"mmm yyyy")&"   |   last year "&TEXT(Setup!$B$10,"mmm yyyy")'
ws['A3'].font = Font(name='Calibri', size=10, bold=True, color=TEAL)

ws['A5'] = 'Department'
ws['A5'].font = Font(name='Calibri', size=11, bold=True, color=INK)
ws['B5'] = 'Support'
ws['B5'].font = Font(name='Calibri', size=11, bold=True, color=WHITE)
ws['B5'].fill = PatternFill('solid', fgColor=TEAL)
ws['B5'].alignment = Alignment(horizontal='center')
ws['B5'].border = Border(left=thin, right=thin, top=thin, bottom=thin)
dv = DataValidation(type='list', formula1='"' + ','.join(ORDER) + '"', allow_blank=False)
dv.prompt = 'Pick the department'; dv.promptTitle = 'Department'
ws.add_data_validation(dv); dv.add(ws['B5'])
ws['C5'] = 'type over this cell or use the dropdown'
ws['C5'].font = Font(name='Calibri', size=9, italic=True, color=INK)

H = 7
ws.cell(H, 1, '=$B$5')
ws.cell(H, 2, '=TEXT(Setup!$B$5,"mmm yyyy")')
ws.cell(H, 3, '=TEXT(Setup!$B$10,"mmm yyyy")')
ws.cell(H, 4, 'Var vs LY')
ws.cell(H, 6, '="Budget "&TEXT(Setup!$B$5,"mmm yyyy")')
ws.cell(H, 7, 'Var vs Budget')
for c in range(1, 8):
    ws.cell(H, c).font = Font(name='Calibri', size=11, bold=True, color=WHITE)
    ws.cell(H, c).fill = PatternFill('solid', fgColor=INK if c != 1 else TEAL)
    ws.cell(H, c).alignment = Alignment(horizontal='right' if c > 1 else 'left')
ws.cell(H, 5).fill = PatternFill('solid', fgColor=WHITE)

PBD, A26, BFY = "'P&L by Dept'", "'Actual FY26'", "'Budget FY27'"
EN = "Engine"
EM = f'MATCH(Setup!$B$5,{EN}!$E$4:$P$4,0)'          # which month column
EB = f"INDEX(${gl(6)}${W+3}:${gl(6)}${W+3+len(ORDER)-1},{IDX})"   # Engine base row
def eng(first, n):
    """A department's n rows on the Engine, for the month on the dropdown."""
    return f'OFFSET({EN}!$A$1,{EB}+{first},{EM}+3,{n},1)'
LYC = f"MATCH(Setup!$B$5,{A26}!$B$5:$M$5,0)"      # same month column, last year's book
BUC = f"MATCH(Setup!$B$5,{BFY}!$B$5:$M$5,0)"
PC  = f"INDEX(${gl(6)}${W+3}:${gl(6)}${W+3+len(ORDER)-1},{IDX})"    # P&L by Dept column
OC  = f"INDEX(${gl(7)}${W+3}:${gl(7)}${W+3+len(ORDER)-1},{IDX})"
RW  = lambda col: f"INDEX(${gl(col)}${W+3}:${gl(col)}${W+3+len(ORDER)-1},{IDX})"

def blk(sheet, first, n, colexpr):
    """A department's n-row block on a 233-row P&L sheet, for one month column."""
    return f"OFFSET({sheet}!$A$1,{first}-1,{colexpr},{n},1)"

ROWS = []
def line(r, label, b, c, f, fmt, bold=False, fill=None, group=None):
    ws.cell(r, 1, label)
    ws.cell(r, 2, b); ws.cell(r, 3, c); ws.cell(r, 6, f)
    ws.cell(r, 4, f'=IF(OR($C{r}="",NOT(ISNUMBER($C{r})),$C{r}=0),"",($B{r}-$C{r})/ABS($C{r}))')
    ws.cell(r, 7, f'=IF(OR($F{r}="",NOT(ISNUMBER($F{r})),$F{r}=0),"",($B{r}-$F{r})/ABS($F{r}))')
    for cc in (2, 3, 6):
        ws.cell(r, cc).number_format = fmt
    for cc in (4, 7):
        ws.cell(r, cc).number_format = VAR
    for cc in range(1, 8):
        cell = ws.cell(r, cc)
        cell.font = Font(name='Calibri', size=11, bold=bold, color=INK)
        cell.border = Border(bottom=thin)
        if fill:
            cell.fill = PatternFill('solid', fgColor=fill)
        if cc > 1:
            cell.alignment = Alignment(horizontal='right')
    ROWS.append(r)

NA = '"no LY split"'
r = 8
# revenue -------------------------------------------------------------------
line(r, 'Total Sales Revenue',
     f'=SUM({eng(-1, 7)})',
     f'=INDEX({A26}!$B$1:$M$233,{RW(3)},{LYC})',
     f'=INDEX({BFY}!$B$1:$M$233,{RW(3)},{BUC})',
     MONEY, bold=True, fill=GREYL)
REV = r; r += 1

# cost of sales by her Type grouping ---------------------------------------
for ty in ['Labour', 'Equipment & Services', 'Transportation', 'Other Direct Expenses']:
    m = MASK[ty]
    line(r, ty,
         f'=SUMPRODUCT({eng(6, 16)},{m})',
         f'=SUMPRODUCT({blk(A26,"$E$"+str(W+3)+"",16,LYC)},{m})'.replace(
             '$E$' + str(W + 3), RW(5)),
         f'=SUMPRODUCT({blk(BFY,"X",16,BUC)},{m})'.replace('X', RW(5)),
         COST)
    r += 1
line(r, 'Total COGS',
     f'=SUM({eng(6, 16)})',
     f'=INDEX({A26}!$B$1:$M$233,{RW(4)},{LYC})',
     f'=INDEX({BFY}!$B$1:$M$233,{RW(4)},{BUC})',
     COST, bold=True, fill=GREYL)
COS_T = r; r += 1
line(r, 'Gross Profit', f'=$B{REV}+$B{COS_T}', f'=$C{REV}+$C{COS_T}',
     f'=$F{REV}+$F{COS_T}', MONEY, bold=True, fill=GREYL)
GP = r; r += 1
for cc, col in ((2, 'B'), (3, 'C'), (6, 'F')):
    pass
ws.cell(r, 1, 'GPM')
for col in ('B', 'C', 'F'):
    ws.cell(r, {'B': 2, 'C': 3, 'F': 6}[col],
            f'=IF(OR(NOT(ISNUMBER({col}{REV})),{col}{REV}=0),"",{col}{GP}/{col}{REV})')
ws.cell(r, 4, f'=IF(OR($C{r}="",NOT(ISNUMBER($C{r})),$C{r}=0),"",($B{r}-$C{r})/ABS($C{r}))')
ws.cell(r, 7, f'=IF(OR($F{r}="",NOT(ISNUMBER($F{r})),$F{r}=0),"",($B{r}-$F{r})/ABS($F{r}))')
for cc in (2, 3, 6):
    ws.cell(r, cc).number_format = PCT
for cc in (4, 7):
    ws.cell(r, cc).number_format = VAR
for cc in range(1, 8):
    ws.cell(r, cc).font = Font(name='Calibri', size=11, bold=True, color=INK)
    ws.cell(r, cc).fill = PatternFill('solid', fgColor=GREYL)
    ws.cell(r, cc).border = Border(bottom=thin)
    if cc > 1:
        ws.cell(r, cc).alignment = Alignment(horizontal='right')
GPM = r; r += 1

# expenses ------------------------------------------------------------------
OHS = "'Overhead Split'"
OHM = f"MATCH(Setup!$B$5,{OHS}!$C$22:$N$22,0)"
E0 = r
for nm, pbdrow in EXP:
    own = f'INDEX({eng(22, 6)},{EXPOFF[nm]})'
    alloc = (f'IF({OC}=0,-SUMPRODUCT(({OHS}!$U$23:$U$86=$A{r})*'
             f'INDEX({OHS}!$C$23:$N$86,0,{OHM})),'
             f'SUMPRODUCT(({OHS}!$U$23:$U$86=$A{r})*INDEX({OHS}!$P$23:$T$86,0,{OC})*'
             f'INDEX({OHS}!$C$23:$N$86,0,{OHM})))')
    bq = (f"=IF($B$5=\"General Admin\",{own},IFERROR(SUMIF({BFY}!$A${BR+2}:$A${BEND},$B$5,"
          f"INDEX({BFY}!$C${BR+2}:$N${BEND},0,{BUC}))*0,0))")   # placeholder, set below
    line(r, nm,
         f'={own}+{alloc}',
         f'={NA}',
         f'=IF($B$5="General Admin","",SUMIFS(INDEX({BFY}!$C${BR+2}:$N${BEND},0,{BUC}),'
         f'{BFY}!$A${BR+2}:$A${BEND},$B$5,{BFY}!$B${BR+2}:$B${BEND},$A{r}))',
         COST)
    r += 1
# Whatever the six headings above do not pick up. Taking it as the department's
# whole expense (its own, plus its share of admin, both already on P&L by Dept)
# less the six means Total Expenses always ties back to that tab, so nothing
# can quietly leak out of the bottom of the card.
SPREAD = (f'IF({OC}=0,-INDEX({OHS}!$C$95:$N$95,1,{OHM}),'
          f'INDEX({OHS}!$C$90:$N$94,{OC},{OHM}))')
line(r, 'Other admin overhead',
     f'=SUM({eng(22, 6)})+{SPREAD}-SUM($B{E0}:$B{r-1})',
     f'={NA}', '=""', COST)
r += 1
E1 = r - 1
line(r, 'Total Expenses', f'=SUM($B{E0}:$B{E1})', f'={NA}', f'=SUM($F{E0}:$F{E1})',
     COST, bold=True, fill=GREYL)
TE = r; r += 1
line(r, 'Net Profit/Loss', f'=$B{GP}+$B{TE}', f'={NA}', f'=$F{GP}+$F{TE}',
     MONEY, bold=True, fill=GREYL)
NP = r; r += 1
ws.cell(r, 1, 'NPM')
for col, cc in (('B', 2), ('C', 3), ('F', 6)):
    ws.cell(r, cc, f'=IF(OR(NOT(ISNUMBER({col}{NP})),NOT(ISNUMBER({col}{REV})),'
                   f'{col}{REV}=0),"",{col}{NP}/{col}{REV})')
ws.cell(r, 4, f'=IF(OR($C{r}="",NOT(ISNUMBER($C{r})),$C{r}=0),"",($B{r}-$C{r})/ABS($C{r}))')
ws.cell(r, 7, f'=IF(OR($F{r}="",NOT(ISNUMBER($F{r})),$F{r}=0),"",($B{r}-$F{r})/ABS($F{r}))')
for cc in (2, 3, 6):
    ws.cell(r, cc).number_format = PCT
for cc in (4, 7):
    ws.cell(r, cc).number_format = VAR
for cc in range(1, 8):
    ws.cell(r, cc).font = Font(name='Calibri', size=11, bold=True, color=INK)
    ws.cell(r, cc).fill = PatternFill('solid', fgColor=GREYL)
    ws.cell(r, cc).border = Border(bottom=thin)
    if cc > 1:
        ws.cell(r, cc).alignment = Alignment(horizontal='right')
NPM = r

ws.cell(NPM + 2, 1, 'Last year’s overheads are not split by department anywhere in '
                    'the pack, so the expense lines above have no last year column. '
                    'Revenue, COGS and gross profit do.')
ws.cell(NPM + 2, 1).font = Font(name='Calibri', size=9, italic=True, color=RED)

ws.column_dimensions['A'].width = 34
for c in 'BCDFG':
    ws.column_dimensions[c].width = 15
ws.column_dimensions['E'].width = 3
ws.freeze_panes = 'A8'

# the subcategory column the allocation needs -------------------------------
oh = wb['Overhead Split']
oh.cell(22, 21, 'Subcategory').font = Font(name='Calibri', size=10, bold=True, color=WHITE)
oh.cell(22, 21).fill = PatternFill('solid', fgColor=INK)
for rr in range(23, 87):
    if oh.cell(rr, 1).value:
        oh.cell(rr, 21,
                f'=IFERROR(INDEX(Lists!$D$6:$D$214,MATCH($A{rr},Lists!$B$6:$B$214,0)),"Other")')

wb.save(OUT)
print(f'Dept P&L built: card rows {REV} to {NPM}, workings from row {W}')
