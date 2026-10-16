"""P&L by Dept: a full P&L per department, laid out like P&L FY27.

The same department tree the P&L has, and it collapses the same way:

    Production            the two production departments added together
        Events
        Video
    Consulting            integration and consulting added together
        Integration
        Consulting
    Support
    General Admin

Four outline levels. The 1 button leaves the department headings, 2 opens a
department's P&L, 3 opens its subcategory detail.

Actuals come off the Engine, which already carries a row per department and
subcategory a month, so this is a lookup and not another pass over the GL.
Forecast months come off Budget FY27 where the budget is set by department,
which is income, cost of sales and overheads. Other income, other expenses
and depreciation are not budgeted by department anywhere, so those lines show
actuals and then stop.
"""
import os, json, copy
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.properties import Outline
from openpyxl.utils import get_column_letter as gl

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = f'{HERE}/CTS Financial Controller Pack - v20.xlsx'
B = json.load(open(f'{HERE}/data/brand.json'))
INK, WHITE = B['charcoal'], B['white']

# tag -> first row of its block on the Engine, first addback row
ENG = {'ONSITE': (5, 192), 'PRODUCTION': (36, 195), 'VIDEO': (67, 198),
       'INTEGRATION': (98, 201), 'CONSULTING': (129, 204), 'ADMIN': (160, 207)}
# tag -> first budget income row, first budget cost of sales row, label
BUDR = {'PRODUCTION': (9, 62, 'Events'), 'VIDEO': (17, 79, 'Video'),
        'INTEGRATION': (26, 97, 'Integration'), 'CONSULTING': (34, 114, 'Consulting'),
        'ONSITE': (42, 131, 'Support'), 'ADMIN': (50, 148, 'General Admin')}
# heading, outline level of the heading, the tags it adds up
BLOCKS = [('Production', 0, ['PRODUCTION', 'VIDEO']),
          ('Events', 1, ['PRODUCTION']),
          ('Video', 1, ['VIDEO']),
          ('Consulting', 0, ['INTEGRATION', 'CONSULTING']),
          ('Integration', 1, ['INTEGRATION']),
          ('Consulting', 1, ['CONSULTING']),
          ('Support', 0, ['ONSITE']),
          ('General Admin', 0, ['ADMIN'])]
LEAF = ['Events', 'Video', 'Integration', 'Consulting', 'Support', 'General Admin']
INC, COS, EXP = range(0, 7), range(7, 23), range(23, 29)
OTH_INC, OTH_EXP = 29, 30

wb = openpyxl.load_workbook(OUT)
PL, EN, BUD = wb['P&L FY27'], wb['Engine'], wb['Budget FY27']
names = lambda off: [EN.cell(ENG['ONSITE'][0] + i, 4).value for i in off]
INC_N, COS_N, EXP_N = names(INC), names(COS), names(EXP)
BUDEXP = {}
for r in range(197, 227):
    d, s = BUD.cell(r, 1).value, BUD.cell(r, 2).value
    if d:
        BUDEXP[(d, s)] = r

if 'P&L by Dept' in wb.sheetnames:
    idx = wb.sheetnames.index('P&L by Dept')
    del wb['P&L by Dept']
else:
    idx = wb.sheetnames.index('Dept P&L')
ws = wb.create_sheet('P&L by Dept', idx)
ws.sheet_properties.tabColor = INK
ws.sheet_properties.outlinePr = Outline(applyStyles=False, summaryBelow=False,
                                        summaryRight=False)
ws['A1'] = 'P&L by Department'
ws['A1'].font = Font(name='Calibri', size=16, bold=True, color=INK)
ws['A2'] = '=TEXT(Setup!$B$5,"mmm yyyy")'
ws['A2'].font = Font(name='Calibri', size=10, bold=True, color=INK)

st = lambda dr, dc, sr, sc: ws.cell(dr, dc)._style.__setattr__ if False else None
def style_from(dr, dc, sr, sc):
    ws.cell(dr, dc)._style = copy.copy(PL.cell(sr, sc)._style)

MON = range(2, 14)
ECOL = lambda c: gl(c + 3)
ROW = [4]
BANNER = {}

for head, lvl, tags in BLOCKS:
    r = ROW[0]
    BANNER.setdefault(head, []).append(r)
    ws.cell(r, 1, ('    ' if lvl else '') + head)
    for c in range(1, 16):
        ws.cell(r, c).fill = PatternFill('solid', fgColor=INK)
        ws.cell(r, c).font = Font(name='Calibri', size=12 - lvl, bold=True, color=WHITE)
    ws.row_dimensions[r].outlineLevel = lvl
    IN = lvl + 1                       # everything inside the block
    AF, HDR = r + 1, r + 2
    for rr in (AF, HDR):
        ws.row_dimensions[rr].outlineLevel = IN
    for c in MON:
        ws.cell(AF, c, f'=IF({gl(c)}${HDR}<=Setup!$B$5,"Actual","Forecast")')
        style_from(AF, c, 4, c)
        ws.cell(HDR, c, f"='P&L FY27'!{gl(c)}$5")
        style_from(HDR, c, 5, c)
        ws.cell(HDR, c).number_format = PL.cell(5, c).number_format
    ws.cell(HDR, 1, 'Line'); style_from(HDR, 1, 5, 1)
    ws.cell(HDR, 14, 'FY total'); style_from(HDR, 14, 5, 14)
    ws.cell(HDR, 15, 'YTD'); style_from(HDR, 15, 5, 15)
    ROW[0] = HDR + 1

    def line(label, src, eoff=None, boff=None, bexp=None, formula=None, deep=False):
        """One P&L line. eoff is the offset inside a department's Engine block,
        boff the offset inside its Budget FY27 block."""
        r = ROW[0]
        ws.cell(r, 1, label)
        style_from(r, 1, src, 1)
        for c in MON:
            if formula:
                ws.cell(r, c, formula(gl(c), c))
            else:
                act = '+'.join(f'Engine!${ECOL(c)}{ENG[t][0] + eoff}' for t in tags)
                if bexp is not None:
                    rows = [BUDEXP.get((BUDR[t][2], bexp)) for t in tags]
                    bud = '+'.join(f"'Budget FY27'!{gl(c)}{x}" for x in rows if x) or '0'
                elif boff is not None:
                    k = 0 if boff[0] == 'i' else 1
                    bud = '+'.join(f"'Budget FY27'!{gl(c)}{BUDR[t][k] + boff[1]}"
                                   for t in tags)
                else:
                    bud = '0'
                ws.cell(r, c, f'=IF({gl(c)}${HDR}<=Setup!$B$5,{act},{bud})')
            style_from(r, c, src, 2)
        ws.cell(r, 14, f'=SUM($B{r}:$M{r})')
        ws.cell(r, 15, f'=SUMIF($B${HDR}:$M${HDR},"<="&Setup!$B$5,$B{r}:$M{r})')
        for c in (14, 15):
            style_from(r, c, src, 2)
        ws.row_dimensions[r].outlineLevel = IN + 1 if deep else IN
        ws.row_dimensions[r].hidden = deep
        ROW[0] = r + 1
        return r

    TI = line('Total Income', 6, formula=lambda c, cc: f'=SUM({c}{ROW[0]+1}:{c}{ROW[0]+7})')
    for i, nm in zip(INC, INC_N):
        line('    ' + nm, 42, eoff=i, boff=('i', i), deep=True)
    OI = line('Other income', 57, eoff=OTH_INC)
    CS = line('Cost of Sales', 59, formula=lambda c, cc: f'=SUM({c}{ROW[0]+1}:{c}{ROW[0]+16})')
    for i, nm in zip(COS, COS_N):
        line('    ' + nm, 131, eoff=i, boff=('c', i - 7), deep=True)
    EX = line('Expenses Overheds', 164, formula=lambda c, cc: f'=SUM({c}{ROW[0]+1}:{c}{ROW[0]+6})')
    for i, nm in zip(EXP, EXP_N):
        line('    ' + nm, 165, eoff=i, bexp=nm, deep=True)
    OE = line('Other Expenses', 171, eoff=OTH_EXP)
    GP = line('Gross profit', 173, formula=lambda c, cc: f'={c}{TI}+{c}{CS}')
    line('Gross profit %', 174, formula=lambda c, cc: f'=IF({c}{TI}=0,"",{c}{GP}/{c}{TI})')
    da = '+'.join(f'Engine!${{E}}{ENG[t][1]}+IF(Setup!$B$28="Yes",Engine!${{E}}{ENG[t][1]+1},0)'
                  for t in tags)
    DA = line('Depreciation & amortisation', 175, formula=lambda c, cc:
              f'=IF({c}${HDR}<=Setup!$B$5,' + da.replace('{E}', ECOL(cc)) + ',0)')
    EB = line('EBITDA', 176, formula=lambda c, cc: f'={c}{TI}+{c}{CS}+{c}{EX}-{c}{DA}')
    line('EBITDA %', 177, formula=lambda c, cc: f'=IF({c}{TI}=0,"",{c}{EB}/{c}{TI})')
    EI = line('EBIT', 178, formula=lambda c, cc: f'={c}{EB}+{c}{DA}')
    line('EBIT %', 179, formula=lambda c, cc: f'=IF({c}{TI}=0,"",{c}{EI}/{c}{TI})')
    NP = line('Net profit', 180, formula=lambda c, cc: f'={c}{TI}+{c}{OI}+{c}{CS}+{c}{EX}+{c}{OE}')
    line('Net profit %', 181, formula=lambda c, cc: f'=IF({c}{TI}=0,"",{c}{NP}/{c}{TI})')
    for pct, money in (('Gross profit %', GP), ('EBITDA %', EB),
                       ('EBIT %', EI), ('Net profit %', NP)):
        pr = next(x for x in range(HDR, ROW[0]) if ws.cell(x, 1).value == pct)
        ws.cell(pr, 14, f'=IF(N{TI}=0,"",N{money}/N{TI})')
        ws.cell(pr, 15, f'=IF(O{TI}=0,"",O{money}/O{TI})')
    ROW[0] += 2

# ------------------------------------------- the six side by side ---------
LEAFROW = {}
for head, lvl, tags in BLOCKS:
    if head in LEAF and lvl == (0 if head in ('Support', 'General Admin') else 1):
        LEAFROW[head] = BANNER[head][-1 if head == 'Consulting' else 0]
S = ROW[0]
ws.cell(S, 1, 'ALL SIX SIDE BY SIDE, YEAR TO DATE')
for c in range(1, 8):
    ws.cell(S, c).fill = PatternFill('solid', fgColor=INK)
    ws.cell(S, c).font = Font(name='Calibri', size=12, bold=True, color=WHITE)
for i, h in enumerate(['Department', 'Total Income', 'Cost of Sales', 'Gross profit',
                       'EBITDA', 'Net profit', 'Net profit %'], start=1):
    ws.cell(S + 1, i, h)
    style_from(S + 1, i, 5, 1 if i == 1 else 2)
    ws.cell(S + 1, i).alignment = Alignment(horizontal='right' if i > 1 else 'left')
OFF = {'Total Income': 3, 'Cost of Sales': 12, 'Gross profit': 37,
       'EBITDA': 40, 'Net profit': 44}
for j, nm in enumerate(LEAF):
    rr, b = S + 2 + j, LEAFROW[nm]
    ws.cell(rr, 1, nm); style_from(rr, 1, 8, 1)
    for i, key in enumerate(OFF, start=2):
        ws.cell(rr, i, f'=$O${b + OFF[key]}')
        style_from(rr, i, 59 if key == 'Cost of Sales' else 173, 2)
    ws.cell(rr, 7, f'=IF($B{rr}=0,"",$F{rr}/$B{rr})')
    style_from(rr, 7, 181, 2)
T = S + 8
ws.cell(T, 1, 'CTS')
for i in range(2, 7):
    ws.cell(T, i, f'=SUM({gl(i)}{S+2}:{gl(i)}{S+7})')
ws.cell(T, 7, f'=IF($B{T}=0,"",$F{T}/$B{T})')
for i in range(1, 8):
    style_from(T, i, 180, 1 if i == 1 else 2)
ws.cell(T, 7).number_format = PL.cell(181, 2).number_format

for c in range(1, 16):
    ws.column_dimensions[gl(c)].width = PL.column_dimensions[gl(c)].width or 13
ws.column_dimensions['A'].width = 36
ws.freeze_panes = 'B5'
json.dump({'leaf_rows': LEAFROW, 'offsets': OFF, 'side_by_side': S},
          open(f'{HERE}/data/pldept.json', 'w'), indent=1)
print('block headings:', {k: v for k, v in BANNER.items()})
print('leaf blocks used by the table:', LEAFROW)
print(f'side by side at {S}, total {T}')
wb.save(OUT)
print('saved')
