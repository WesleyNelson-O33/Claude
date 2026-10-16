"""P&L by Dept: one P&L a department, with the sub-departments inside it.

Four P&Ls, because there are four departments:

    Production        with Events and Video inside it
    Consulting        with Integration and Consulting inside it
    Support
    General Admin

Each one is laid out like P&L FY27, the twelve months then FY total then YTD,
and the income and cost of sales sections carry the sub-department tree the
P&L carries. Overheads are not split by sub-department on the P&L, so they are
not split here either.

Outline levels, summary above detail as everywhere else in the pack:
    1  the four department headings
    2  a department's P&L
    3  its sub-departments
    4  the subcategory detail

Actuals come off the Engine, which already carries a row per department and
subcategory a month, so this is a lookup and not another pass over the GL.
Forecast months come off Budget FY27 where the budget is set by department,
which is income, cost of sales and overheads. Other income, other expenses and
depreciation are not budgeted by department anywhere, so those lines show
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

ENG = {'ONSITE': (5, 192), 'PRODUCTION': (36, 195), 'VIDEO': (67, 198),
       'INTEGRATION': (98, 201), 'CONSULTING': (129, 204), 'ADMIN': (160, 207)}
BUDR = {'PRODUCTION': (9, 62, 'Events'), 'VIDEO': (17, 79, 'Video'),
        'INTEGRATION': (26, 97, 'Integration'), 'CONSULTING': (34, 114, 'Consulting'),
        'ONSITE': (42, 131, 'Support'), 'ADMIN': (50, 148, 'General Admin')}
# heading, the tags it covers, the sub-departments shown inside it
BLOCKS = [('Production', ['PRODUCTION', 'VIDEO'],
           [('Events', 'PRODUCTION'), ('Video', 'VIDEO')]),
          ('Consulting', ['INTEGRATION', 'CONSULTING'],
           [('Integration', 'INTEGRATION'), ('Consulting', 'CONSULTING')]),
          ('Support', ['ONSITE'], []),
          ('General Admin', ['ADMIN'], [])]
INC, COS, EXP = range(0, 7), range(7, 23), range(23, 29)
OTH_INC, OTH_EXP = 29, 30

wb = openpyxl.load_workbook(OUT)
PL, EN, BUD = wb['P&L FY27'], wb['Engine'], wb['Budget FY27']
names = lambda off: [EN.cell(ENG['ONSITE'][0] + i, 4).value for i in off]
INC_N, COS_N, EXP_N = names(INC), names(COS), names(EXP)
BUDEXP = {}
for r in range(197, 227):
    d, s_ = BUD.cell(r, 1).value, BUD.cell(r, 2).value
    if d:
        BUDEXP[(d, s_)] = r
# That block's months sit in C to N, not B to M like the rest of the sheet, so
# its own heading row is the criteria range. Reading B to M puts every month
# one out and quietly picks up the wrong month's budget.
EH = min(BUDEXP.values()) - 1
EXPRNG = f"'Budget FY27'!$C${EH}:$N${EH}"
assert BUD.cell(EH, 3).value, 'the overhead budget heading row moved'
expbud = lambda r: f"SUMIF({EXPRNG},Setup!$B$5,'Budget FY27'!$C{r}:$N{r})"
print('overhead budget months read from row', EH, 'columns C to N')

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

def style_from(dr, dc, sr, sc):
    ws.cell(dr, dc)._style = copy.copy(PL.cell(sr, sc)._style)

MON = range(2, 14)
ECOL = lambda c: gl(c + 3)
ROW = [4]
BLOCKROW = {}

for head, tags, subs in BLOCKS:
    r = ROW[0]
    BLOCKROW[head] = r
    ws.cell(r, 1, head)
    for c in range(1, 16):
        ws.cell(r, c).fill = PatternFill('solid', fgColor=INK)
        ws.cell(r, c).font = Font(name='Calibri', size=12, bold=True, color=WHITE)
    ws.row_dimensions[r].outlineLevel = 0
    AF, HDR = r + 1, r + 2
    for rr in (AF, HDR):
        ws.row_dimensions[rr].outlineLevel = 1
    for c in MON:
        ws.cell(AF, c, f'=IF({gl(c)}${HDR}<=Setup!$B$5,"Actual","Forecast")')
        style_from(AF, c, 4, c)
        ws.cell(HDR, c, f"='P&L FY27'!{gl(c)}$5")
        style_from(HDR, c, 5, c)
        ws.cell(HDR, c).number_format = PL.cell(5, c).number_format
    ws.cell(HDR, 1, 'Line'); style_from(HDR, 1, 5, 1)
    ws.cell(HDR, 14, 'FY total'); style_from(HDR, 14, 5, 14)
    ws.cell(HDR, 15, 'YTD'); style_from(HDR, 15, 5, 15)
    # the month block, same as the one on P&L FY27, after a blank column
    for col, lab, src in ((17, 'Month', 15), (18, 'Month budget', 17),
                          (19, 'Var $', 18), (20, 'Var %', 19)):
        ws.cell(HDR, col, lab); style_from(HDR, col, 5, src)
        ws.cell(HDR, col).alignment = Alignment(horizontal='right')
    ROW[0] = HDR + 1

    def put(label, src, level, hide, body, pct=False, generic=False, budget=None):
        """One P&L line.

        pct      the body is a ratio, so the month column is the ratio too
                 rather than a sum of percentages
        generic  the body is written purely in column letters, so the budget
                 column can reuse it
        budget   an explicit month budget formula, for the leaf rows
        """
        r = ROW[0]
        ws.cell(r, 1, label)
        style_from(r, 1, src, 1)
        for c in MON:
            ws.cell(r, c, body(gl(c), c))
            style_from(r, c, src, 2)
        ws.cell(r, 14, f'=SUM($B{r}:$M{r})')
        ws.cell(r, 15, f'=SUMIF($B${HDR}:$M${HDR},"<="&Setup!$B$5,$B{r}:$M{r})')
        ws.cell(r, 17, body('Q', None) if pct else
                f'=SUMIF($B${HDR}:$M${HDR},Setup!$B$5,$B{r}:$M{r})')
        ws.cell(r, 18, budget(r) if budget else
                (body('R', None) if (pct or generic) else 0))
        ws.cell(r, 19, f'=IFERROR(Q{r}-R{r},"")')
        ws.cell(r, 20, f'=IFERROR(S{r}/ABS(R{r}),0)')
        for c in (14, 15, 17, 18, 19):
            style_from(r, c, src, 2)
        ws.cell(r, 20)._style = copy.copy(PL.cell(r if src != 165 else 165, 2)._style)
        ws.cell(r, 20).number_format = PL.cell(179, 2).number_format
        ws.row_dimensions[r].outlineLevel = level
        ws.row_dimensions[r].hidden = hide
        ROW[0] = r + 1
        return r

    def leaf(tag, off, nm, level, budkind):
        """One subcategory, for one tag."""
        def body(c, cc):
            act = f'Engine!${ECOL(cc)}{ENG[tag][0] + off}'
            if budkind == 'e':
                br = BUDEXP.get((BUDR[tag][2], nm))
                bud = f"'Budget FY27'!{gl(cc)}{br}" if br else '0'
            else:
                k = 0 if budkind == 'i' else 1
                base = BUDR[tag][k] + (off if budkind == 'i' else off - 7)
                bud = f"'Budget FY27'!{gl(cc)}{base}"
            return f'=IF({c}${HDR}<=Setup!$B$5,{act},{bud})'
        def mbud(rr):
            if budkind == 'e':
                br = BUDEXP.get((BUDR[tag][2], nm))
                return f'={expbud(br)}' if br else 0
            k = 0 if budkind == 'i' else 1
            base = BUDR[tag][k] + (off if budkind == 'i' else off - 7)
            return (f"=SUMIF('Budget FY27'!$B$5:$M$5,Setup!$B$5,"
                    f"'Budget FY27'!$B{base}:$M{base})")
        return put('        ' + nm if level > 2 else '    ' + nm,
                   42 if budkind == 'i' else (131 if budkind == 'c' else 165),
                   level, True, body, budget=mbud)

    def section(label, src, offs, nms, budkind, sub_src):
        """A category, with its sub-departments inside it when it has any."""
        head_row = ROW[0]
        ROW[0] += 1                       # the headline is written last
        parts = []
        if subs:
            for sublab, tag in subs:
                sub_row = ROW[0]
                ROW[0] += 1
                kids = [leaf(tag, o, n, 3, budkind) for o, n in zip(offs, nms)]
                cur = ROW[0]
                ROW[0] = sub_row
                put('    ' + sublab, sub_src, 2, False,
                    lambda c, cc, k=kids: f'=SUM({c}{k[0]}:{c}{k[-1]})', generic=True)
                ROW[0] = cur
                parts.append(sub_row)
        else:
            kids = [leaf(tags[0], o, n, 2, budkind) for o, n in zip(offs, nms)]
            parts = None
        cur = ROW[0]
        ROW[0] = head_row
        if parts:
            put(label, src, 1, False,
                lambda c, cc, p=parts: '=' + '+'.join(f'{c}{x}' for x in p),
                generic=True)
        else:
            put(label, src, 1, False,
                lambda c, cc, k=kids: f'=SUM({c}{k[0]}:{c}{k[-1]})', generic=True)
        ROW[0] = cur
        return head_row

    def plain(label, src, off, level=1):
        def body(c, cc):
            act = '+'.join(f'Engine!${ECOL(cc)}{ENG[t][0] + off}' for t in tags)
            return f'=IF({c}${HDR}<=Setup!$B$5,{act},0)'
        return put(label, src, level, False, body)

    TI = section('Total Income', 6, INC, INC_N, 'i', 7)
    OI = plain('Other income', 57, OTH_INC)
    CS = section('Cost of Sales', 59, COS, COS_N, 'c', 60)
    EXr = ROW[0]
    ROW[0] += 1
    ekids = []
    for o, n in zip(EXP, EXP_N):
        rr = ROW[0]
        def ebody(c, cc, o=o, n=n):
            act = '+'.join(f'Engine!${ECOL(cc)}{ENG[t][0] + o}' for t in tags)
            br = [BUDEXP.get((BUDR[t][2], n)) for t in tags]
            bud = '+'.join(f"'Budget FY27'!{gl(cc)}{x}" for x in br if x) or '0'
            return f'=IF({c}${HDR}<=Setup!$B$5,{act},{bud})'
        def embud(rr, n=n):
            br = [BUDEXP.get((BUDR[t][2], n)) for t in tags]
            parts = [expbud(x) for x in br if x]
            return '=' + '+'.join(parts) if parts else 0
        ekids.append(put('    ' + n, 165, 2, True, ebody, budget=embud))
    cur = ROW[0]; ROW[0] = EXr
    EX = put('Expenses Overheds', 164, 1, False,
             lambda c, cc: f'=SUM({c}{ekids[0]}:{c}{ekids[-1]})', generic=True)
    ROW[0] = cur
    OE = plain('Other Expenses', 171, OTH_EXP)
    GP = put('Gross profit', 173, 1, False, lambda c, cc: f'={c}{TI}+{c}{CS}',
             generic=True)
    put('Gross profit %', 174, 1, False,
        lambda c, cc: f'=IF({c}{TI}=0,"",{c}{GP}/{c}{TI})', pct=True)
    da = '+'.join(f'Engine!${{E}}{ENG[t][1]}+IF(Setup!$B$28="Yes",Engine!${{E}}{ENG[t][1]+1},0)'
                  for t in tags)
    DA = put('Depreciation & amortisation', 175, 1, False, lambda c, cc:
             f'=IF({c}${HDR}<=Setup!$B$5,' + da.replace('{E}', ECOL(cc)) + ',0)')
    EB = put('EBITDA', 176, 1, False,
             lambda c, cc: f'={c}{TI}+{c}{CS}+{c}{EX}-{c}{DA}', generic=True)
    put('EBITDA %', 177, 1, False, lambda c, cc: f'=IF({c}{TI}=0,"",{c}{EB}/{c}{TI})',
        pct=True)
    EI = put('EBIT', 178, 1, False, lambda c, cc: f'={c}{EB}+{c}{DA}', generic=True)
    put('EBIT %', 179, 1, False, lambda c, cc: f'=IF({c}{TI}=0,"",{c}{EI}/{c}{TI})',
        pct=True)
    NP = put('Net profit', 180, 1, False,
             lambda c, cc: f'={c}{TI}+{c}{OI}+{c}{CS}+{c}{EX}+{c}{OE}', generic=True)
    put('Net profit %', 181, 1, False, lambda c, cc: f'=IF({c}{TI}=0,"",{c}{NP}/{c}{TI})',
        pct=True)
    for pct, money in (('Gross profit %', GP), ('EBITDA %', EB),
                       ('EBIT %', EI), ('Net profit %', NP)):
        pr = next(x for x in range(HDR, ROW[0]) if ws.cell(x, 1).value == pct)
        ws.cell(pr, 14, f'=IF(N{TI}=0,"",N{money}/N{TI})')
        ws.cell(pr, 15, f'=IF(O{TI}=0,"",O{money}/O{TI})')
    BLOCKROW[head] = (r, TI, CS, GP, EB, NP)
    ROW[0] += 2

# ------------------------------------------ the four side by side ---------
S = ROW[0]
ws.cell(S, 1, 'ALL FOUR SIDE BY SIDE, YEAR TO DATE')
for c in range(1, 8):
    ws.cell(S, c).fill = PatternFill('solid', fgColor=INK)
    ws.cell(S, c).font = Font(name='Calibri', size=12, bold=True, color=WHITE)
for i, h in enumerate(['Department', 'Total Income', 'Cost of Sales', 'Gross profit',
                       'EBITDA', 'Net profit', 'Net profit %'], start=1):
    ws.cell(S + 1, i, h)
    style_from(S + 1, i, 5, 1 if i == 1 else 2)
    ws.cell(S + 1, i).alignment = Alignment(horizontal='right' if i > 1 else 'left')
for j, (head, *_rest) in enumerate(BLOCKS):
    rr = S + 2 + j
    _, TI, CS, GP, EB, NP = BLOCKROW[head]
    ws.cell(rr, 1, head); style_from(rr, 1, 8, 1)
    for i, src in enumerate([TI, CS, GP, EB, NP], start=2):
        ws.cell(rr, i, f'=$O${src}')
        style_from(rr, i, 59 if i == 3 else 173, 2)
    ws.cell(rr, 7, f'=IF($B{rr}=0,"",$F{rr}/$B{rr})')
    style_from(rr, 7, 181, 2)
T = S + 6
ws.cell(T, 1, 'CTS')
for i in range(2, 7):
    ws.cell(T, i, f'=SUM({gl(i)}{S+2}:{gl(i)}{S+5})')
ws.cell(T, 7, f'=IF($B{T}=0,"",$F{T}/$B{T})')
for i in range(1, 8):
    style_from(T, i, 180, 1 if i == 1 else 2)
ws.cell(T, 7).number_format = PL.cell(181, 2).number_format

for c in range(1, 21):
    ws.column_dimensions[gl(c)].width = PL.column_dimensions[gl(c)].width or 13
ws.column_dimensions['A'].width = 36
ws.freeze_panes = 'B5'
json.dump({'blocks': {k: v[0] for k, v in BLOCKROW.items()}, 'side_by_side': S},
          open(f'{HERE}/data/pldept.json', 'w'), indent=1)
for k, v in BLOCKROW.items():
    print(f'  {k:<14} banner {v[0]:>4}  income {v[1]:>4}  COS {v[2]:>4}  net {v[5]:>4}')
print(f'side by side {S}, total {T}')
wb.save(OUT)
print('saved')
