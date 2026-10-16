"""P&L by Dept rebuilt: one full P&L per department, laid out like P&L FY27.

Same columns as the P&L, the twelve months then FY total then YTD. Same row
order, same styling, same + and - grouping, same Actual and Forecast heading.
Six blocks down the sheet, one a department.

Actuals come off the Engine, which already carries one row per department and
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

# label, Engine tag, first Budget FY27 income row, first budget COS row
DEPTS = [('Events', 'PRODUCTION', 9, 62), ('Video', 'VIDEO', 17, 79),
         ('Integration', 'INTEGRATION', 26, 97), ('Consulting', 'CONSULTING', 34, 114),
         ('Support', 'ONSITE', 42, 131), ('General Admin', 'ADMIN', 50, 148)]
ENG_BASE = {'ONSITE': 5, 'PRODUCTION': 36, 'VIDEO': 67,
            'INTEGRATION': 98, 'CONSULTING': 129, 'ADMIN': 160}
ADDBACK = {'ONSITE': 192, 'PRODUCTION': 195, 'VIDEO': 198,
           'INTEGRATION': 201, 'CONSULTING': 204, 'ADMIN': 207}
INC, COS, EXP = range(0, 7), range(7, 23), range(23, 29)
OTH_INC, OTH_EXP = 29, 30

wb = openpyxl.load_workbook(OUT)
PL, ENG, BUD = wb['P&L FY27'], wb['Engine'], wb['Budget FY27']

# the subcategory names, in the order the Engine holds them
names = lambda off: [ENG.cell(ENG_BASE['ONSITE'] + i, 4).value for i in off]
INC_N, COS_N, EXP_N = names(INC), names(COS), names(EXP)

# where the departmental overhead budget sits, by department and subcategory
BUDEXP = {}
for r in range(197, 227):
    d, s = BUD.cell(r, 1).value, BUD.cell(r, 2).value
    if d:
        BUDEXP[(d, s)] = r
print('departmental overhead budget rows found:', len(BUDEXP))

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

def style_from(dst_r, dst_c, src_r, src_c):
    ws.cell(dst_r, dst_c)._style = copy.copy(PL.cell(src_r, src_c)._style)

MON = range(2, 14)                 # B..M
ECOL = lambda c: gl(c + 3)         # B -> E on the Engine
ROW = [4]                          # a list, so the helper below can move it
for label, tag, binc, bcos in DEPTS:
    base, ab = ENG_BASE[tag], ADDBACK[tag]

    r = ROW[0]
    ws.cell(r, 1, label)
    for c in range(1, 16):
        ws.cell(r, c).fill = PatternFill('solid', fgColor=INK)
        ws.cell(r, c).font = Font(name='Calibri', size=12, bold=True, color=WHITE)
    AF = r + 1
    for c in MON:
        ws.cell(AF, c, f'=IF({gl(c)}${r+2}<=Setup!$B$5,"Actual","Forecast")')
        style_from(AF, c, 4, c)
    HDR = r + 2
    ws.cell(HDR, 1, 'Line'); style_from(HDR, 1, 5, 1)
    for c in MON:
        ws.cell(HDR, c, f"='P&L FY27'!{gl(c)}$5")
        style_from(HDR, c, 5, c)
        ws.cell(HDR, c).number_format = PL.cell(5, c).number_format
    ws.cell(HDR, 14, 'FY total'); style_from(HDR, 14, 5, 14)
    ws.cell(HDR, 15, 'YTD'); style_from(HDR, 15, 5, 15)
    ROW[0] = HDR + 1

    def line(label_, src, eng=None, bud=None, formula=None, level=0, hide=False):
        """One P&L line. eng is an Engine row, bud a Budget FY27 row."""
        r = ROW[0]
        ws.cell(r, 1, label_)
        style_from(r, 1, src, 1)
        for c in MON:
            if formula:
                ws.cell(r, c, formula(gl(c), c))
            else:
                tail = f"'Budget FY27'!{gl(c)}{bud}" if bud else '0'
                ws.cell(r, c, f'=IF({gl(c)}${HDR}<=Setup!$B$5,Engine!${ECOL(c)}{eng},{tail})')
            style_from(r, c, src, 2)
        ws.cell(r, 14, f'=SUM($B{r}:$M{r})')
        ws.cell(r, 15, f'=SUMIF($B${HDR}:$M${HDR},"<="&Setup!$B$5,$B{r}:$M{r})')
        for c in (14, 15):
            style_from(r, c, src, 2)
        ws.row_dimensions[r].outlineLevel = level
        ws.row_dimensions[r].hidden = hide
        ROW[0] = r + 1
        return r

    TI = line('Total Income', 6, formula=lambda c, cc:
              f'=SUM({c}{ROW[0]+1}:{c}{ROW[0]+7})')
    for i, nm in zip(INC, INC_N):
        line('    ' + nm, 42, eng=base + i, bud=binc + i, level=1, hide=True)
    OI = line('Other income', 57, eng=base + OTH_INC)
    CS = line('Cost of Sales', 59, formula=lambda c, cc:
              f'=SUM({c}{ROW[0]+1}:{c}{ROW[0]+16})')
    for i, nm in zip(COS, COS_N):
        line('    ' + nm, 131, eng=base + i, bud=bcos + (i - 7), level=1, hide=True)
    EX = line('Expenses Overheds', 164, formula=lambda c, cc:
              f'=SUM({c}{ROW[0]+1}:{c}{ROW[0]+6})')
    for i, nm in zip(EXP, EXP_N):
        line('    ' + nm, 165, eng=base + i, bud=BUDEXP.get((label, nm)),
             level=1, hide=True)
    OE = line('Other Expenses', 171, eng=base + OTH_EXP)
    GP = line('Gross profit', 173, formula=lambda c, cc: f'={c}{TI}+{c}{CS}')
    line('Gross profit %', 174, formula=lambda c, cc: f'=IF({c}{TI}=0,"",{c}{GP}/{c}{TI})')
    DA = line('Depreciation & amortisation', 175, formula=lambda c, cc:
              f'=IF({c}${HDR}<=Setup!$B$5,Engine!${ECOL(cc)}{ab}'
              f'+IF(Setup!$B$28="Yes",Engine!${ECOL(cc)}{ab+1},0),0)')
    EB = line('EBITDA', 176, formula=lambda c, cc: f'={c}{TI}+{c}{CS}+{c}{EX}-{c}{DA}')
    line('EBITDA %', 177, formula=lambda c, cc: f'=IF({c}{TI}=0,"",{c}{EB}/{c}{TI})')
    EI = line('EBIT', 178, formula=lambda c, cc: f'={c}{EB}+{c}{DA}')
    line('EBIT %', 179, formula=lambda c, cc: f'=IF({c}{TI}=0,"",{c}{EI}/{c}{TI})')
    NP = line('Net profit', 180,
              formula=lambda c, cc: f'={c}{TI}+{c}{OI}+{c}{CS}+{c}{EX}+{c}{OE}')
    line('Net profit %', 181, formula=lambda c, cc: f'=IF({c}{TI}=0,"",{c}{NP}/{c}{TI})')

    for pct, money in (('Gross profit %', GP), ('EBITDA %', EB),
                       ('EBIT %', EI), ('Net profit %', NP)):
        pr = next(x for x in range(HDR, ROW[0]) if ws.cell(x, 1).value == pct)
        ws.cell(pr, 14, f'=IF(N{TI}=0,"",N{money}/N{TI})')
        ws.cell(pr, 15, f'=IF(O{TI}=0,"",O{money}/O{TI})')
    ROW[0] += 2

for c in range(1, 16):
    ws.column_dimensions[gl(c)].width = PL.column_dimensions[gl(c)].width or 13
ws.column_dimensions['A'].width = 36
ws.freeze_panes = 'B5'
print(f'P&L by Dept rebuilt: 6 blocks, rows 4 to {ROW[0]-3}')
wb.save(OUT)
print('saved')

# --------------------------------------------- the six side by side -------
# Each block is 48 rows, so the lines sit at a fixed offset from its banner.
import json as _json
wb2 = openpyxl.load_workbook(OUT)
ws2 = wb2['P&L by Dept']
BANNERS = [4 + 48 * i for i in range(6)]
OFF = {'Total Income': 3, 'Cost of Sales': 12, 'Gross profit': 37,
       'EBITDA': 40, 'Net profit': 44}
S = BANNERS[-1] + 48
ws2.cell(S, 1, 'ALL SIX SIDE BY SIDE, YEAR TO DATE')
for c in range(1, 8):
    ws2.cell(S, c).fill = PatternFill('solid', fgColor=INK)
    ws2.cell(S, c).font = Font(name='Calibri', size=12, bold=True, color=WHITE)
for i, h in enumerate(['Department', 'Total Income', 'Cost of Sales', 'Gross profit',
                       'EBITDA', 'Net profit', 'Net profit %'], start=1):
    ws2.cell(S + 1, i, h)
    ws2.cell(S + 1, i)._style = copy.copy(wb2['P&L FY27'].cell(5, 1 if i == 1 else 2)._style)
    ws2.cell(S + 1, i).alignment = Alignment(horizontal='right' if i > 1 else 'left')
for j, (b, (label, *_)) in enumerate(zip(BANNERS, DEPTS)):
    rr = S + 2 + j
    ws2.cell(rr, 1, label)
    ws2.cell(rr, 1)._style = copy.copy(wb2['P&L FY27'].cell(8, 1)._style)
    for i, key in enumerate(['Total Income', 'Cost of Sales', 'Gross profit',
                             'EBITDA', 'Net profit'], start=2):
        ws2.cell(rr, i, f'=$O${b + OFF[key]}')
        ws2.cell(rr, i)._style = copy.copy(
            wb2['P&L FY27'].cell(173 if key != 'Cost of Sales' else 59, 2)._style)
    ws2.cell(rr, 7, f'=IF($B{rr}=0,"",$F{rr}/$B{rr})')
    ws2.cell(rr, 7)._style = copy.copy(wb2['P&L FY27'].cell(181, 2)._style)
T = S + 8
ws2.cell(T, 1, 'CTS')
for i in range(2, 7):
    ws2.cell(T, i, f'=SUM({gl(i)}{S+2}:{gl(i)}{S+7})')
ws2.cell(T, 7, f'=IF($B{T}=0,"",$F{T}/$B{T})')
for i in range(1, 8):
    ws2.cell(T, i)._style = copy.copy(wb2['P&L FY27'].cell(180, 1 if i == 1 else 2)._style)
ws2.cell(T, 7).number_format = wb2['P&L FY27'].cell(181, 2).number_format
print(f'side by side block: banner {S}, departments {S+2} to {S+7}, total {T}')
_json.dump({'banners': BANNERS, 'offsets': OFF, 'side_by_side': S},
           open(f'{HERE}/data/pldept.json', 'w'), indent=1)
wb2.save(OUT)
print('saved')
