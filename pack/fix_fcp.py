"""
Repair Danica's CTS Financial Controller Pack and add what she asked for.

    python3.13 fix_fcp.py
    python3.13 /mnt/skills/public/xlsx/scripts/recalc.py "CTS Financial Controller Pack - fixed.xlsx" 1200
    python3.13 ../tools/fix_outline.py "CTS Financial Controller Pack - fixed.xlsx" "Summary,P&L FY27,Budget FY27,Actual FY26,Budget FY26"

This works ON HER FILE. Same sheets, same column layout, same headers, same
labels, same styles, same formula patterns. Nothing of hers is removed.

REPAIRS
  Cleanup          every formula read GL_Paste!#REF!. Rewritten against the
                   real paste rows, and now tolerant of the export's title rows.
  Setup B20/24/25  the same #REF!.
  Setup B6         was 8 (August) against a July year end, so September came
                   out as period 2. Now 7.

ADDED
  Eight rows at the bottom of each department block: gross profit, gross
  profit %, depreciation & amortisation, EBITDA, EBITDA %, EBIT, EBIT %,
  net profit %. On the Summary, the P&L and all three budget / prior year
  sheets, so they stay the same shape.
  A TOTAL COMPANY block under the six departments, same lines.
  BS_Paste and a Balance Sheet sheet.
  PL_Paste, feeding two new columns on PL_Check so the five category totals
  no longer have to be typed. Her typed column stays.
  Lists: the 14 accounts live in Xero that the master did not carry, and an
  Addback column driving the EBITDA bridge.
"""
import json, copy, datetime as dt
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.properties import Outline
from openpyxl.utils import get_column_letter as CL
from openpyxl.workbook.defined_name import DefinedName

HERE = __file__.rsplit('/', 1)[0]
SRC = '/tmp/claude-0/-home-user-Claude/a00ffe59-433c-507b-9006-917d8c60083a/scratchpad/fcp/orig.xlsx'
OUT = f'{HERE}/CTS Financial Controller Pack - fixed.xlsx'
MASTER = json.load(open(f'{HERE}/data/lists_master.json'))
BSACC = json.load(open(f'{HERE}/data/bs_accounts.json'))

AR = 'Arial'
NAVY = PatternFill('solid', fgColor='1F3864')
YELL = PatternFill('solid', fgColor='FFFF00')
NEWF = PatternFill('solid', fgColor='FFF2CC')
THIN = Side(style='thin', color='BFBFBF')
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
MONEY = '$#,##0;($#,##0);"-"'
MON2 = '$#,##0.00;($#,##0.00);"-"'
PCT = '0.0%;(0.0%);"-"'

def f(sz=10, b=False, col=None): return Font(name=AR, size=sz, bold=b, color=col)

# ------------------------------------------- her layout, exactly as it is ---
OLD_STARTS = [6, 44, 82, 120, 158, 196]
OLD_LEN = 37
DEPTS = ['ONSITE', 'PRODUCTION', 'VIDEO', 'INTEGRATION', 'CONSULTING', 'ADMIN']
O_DEPT, O_INC, O_COS, O_EXP, O_OI, O_OIS, O_OE, O_OES = 0, 1, 9, 26, 33, 34, 35, 36
INC_OFFS = list(range(2, 9))
COS_OFFS = list(range(10, 26))
EXP_OFFS = list(range(27, 33))
SUBCAT_OFFS = INC_OFFS + COS_OFFS + EXP_OFFS + [O_OIS, O_OES]
CAT_OFFS = {O_INC: INC_OFFS, O_COS: COS_OFFS, O_EXP: EXP_OFFS,
            O_OI: [O_OIS], O_OE: [O_OES]}

# ------------------------------------- the eight rows she asked for, added --
# (offset, label, kind)
NEW_ROWS = [
    (37, '    Gross profit', 'gp'),
    (38, '    Gross profit %', 'gp_pct'),
    (39, '    Depreciation & amortisation', 'dep'),
    (40, '    EBITDA', 'ebitda'),
    (41, '    EBITDA %', 'ebitda_pct'),
    (42, '    EBIT', 'ebit'),
    (43, '    EBIT %', 'ebit_pct'),
    (44, '    Net profit %', 'net_pct'),
]
NEW_LEN = OLD_LEN + len(NEW_ROWS)                 # 45
NEW_STARTS = [6 + i * (NEW_LEN + 1) for i in range(6)]      # 6,52,98,144,190,236
TOTAL_START = NEW_STARTS[-1] + NEW_LEN + 1                  # 282
ALL_STARTS = NEW_STARTS + [TOTAL_START]
NEW_LAST = TOTAL_START + NEW_LEN - 1                        # 326
PCT_ROWS = {'gp_pct': 'gp', 'ebitda_pct': 'ebitda', 'ebit_pct': 'ebit', 'net_pct': O_DEPT}
NEW_KIND = {off: kind for off, lab, kind in NEW_ROWS}
NEW_LABEL = {off: lab for off, lab, kind in NEW_ROWS}

def margin_formula(kind, c, s):
    """Her sign convention: income positive, costs negative."""
    if kind == 'gp':      return f'={c}{s+O_INC}+{c}{s+O_COS}'
    if kind == 'dep':     return None                       # engine, set by caller
    if kind == 'ebitda':  return f'={c}{s+O_INC}+{c}{s+O_COS}+{c}{s+O_EXP}-{c}{s+39}'
    if kind == 'ebit':    return f'={c}{s+O_INC}+{c}{s+O_COS}+{c}{s+O_EXP}'
    src = s + (O_DEPT if kind == 'net_pct' else
               {'gp_pct': 37, 'ebitda_pct': 40, 'ebit_pct': 42}[kind])
    return f'=IF({c}{s+O_INC}=0,"",{c}{src}/{c}{s+O_INC})'

# -------------------------------------------- her engine, and its new rows --
SUBS_ORDER = []          # the 31 subcategories in her engine order
INC = ['Contract Income', 'Non Contract Income', 'Work in Progess',
       'Freight Collected & Travel', 'Credit Card Fees Collected',
       'Meal & Travel Allowances', 'Miscellaneous Income']
COS = ['Equipment - Purchase', 'Equipment - Hires', 'Sub-Contract Labour',
       'Service - hires', 'Subscriptions & Licences', 'Equipment -Maintenance Support',
       'Stamp Duty', 'Damage Waiver', 'Supplier Per Diems', 'Direct Wages & Salaries',
       'Direct Salaries', 'Direct Wages', 'Direct Superannuation',
       'Direct Workers Comp', 'Other Direct Expenses', 'Credit Card fees paid - COGS']
EXP = ['General & Administrative Exp', 'Business Insurances',
       'Advertising & Promotion Exp', 'Operating Expenses',
       'Employment Expenses', 'Credit Card fees paid']
SUBS_ORDER = ([('Income', s) for s in INC] + [('Cost of Sales', s) for s in COS]
              + [('Expenses', s) for s in EXP]
              + [('Other Income', 'Other Income'), ('Other Expenses', 'Other Expenses')])
ENG_SUB_ROW = {}
r = 5
for d in DEPTS:
    for cat, sub in SUBS_ORDER:
        ENG_SUB_ROW[(d, cat, sub)] = r; r += 1
ENG_SUB_LAST = r - 1                             # 190, unchanged
ADDBACKS = ['D&A', 'IAWO', 'TAX']
ADDLABEL = {'D&A': 'Depreciation & amortisation', 'IAWO': 'Instant asset write off',
            'TAX': 'Income tax'}
ENG_ADD_ROW = {}
r = ENG_SUB_LAST + 2                             # 192, one blank row between
for d in DEPTS:
    for a in ADDBACKS:
        ENG_ADD_ROW[(d, a)] = r; r += 1
ENG_ADD_LAST = r - 1

# offset -> (category, subcategory) for her subcategory rows
OFF_SUB = {}
for i, off in enumerate(INC_OFFS): OFF_SUB[off] = ('Income', INC[i])
for i, off in enumerate(COS_OFFS): OFF_SUB[off] = ('Cost of Sales', COS[i])
for i, off in enumerate(EXP_OFFS): OFF_SUB[off] = ('Expenses', EXP[i])
OFF_SUB[O_OIS] = ('Other Income', 'Other Income')
OFF_SUB[O_OES] = ('Other Expenses', 'Other Expenses')

CLN_R0, CLN_R1 = 5, 20004
GL_R0 = 2                                        # her paste starts at A2
print('new block starts:', ALL_STARTS, 'last row', NEW_LAST)

wb = openpyxl.load_workbook(SRC)

# ===================================================================== 1 ===
# Capture each shaped sheet: label, style and typed value per (offset, column),
# taken from the ONSITE block because her six blocks are styled identically.
SHAPED = ['Summary', 'P&L FY27', 'Budget FY27', 'Actual FY26', 'Budget FY26']
MAXC = 30
style_of = {}       # (sheet, offset, col) -> cell style
rowmeta = {}        # (sheet, offset) -> (outlineLevel, hidden, height)
labels = {}         # (sheet, offset) -> label text from the ONSITE block
typed = {}          # (sheet, block, offset, col) -> a typed number to carry over
for nm in SHAPED:
    ws = wb[nm]
    for off in range(OLD_LEN):
        src_r = OLD_STARTS[0] + off
        labels[(nm, off)] = ws.cell(src_r, 1).value
        rd = ws.row_dimensions[src_r]
        rowmeta[(nm, off)] = (rd.outlineLevel, rd.hidden, rd.height)
        for c in range(1, MAXC + 1):
            style_of[(nm, off, c)] = copy.copy(ws.cell(src_r, c)._style)
    for b, s0 in enumerate(OLD_STARTS):
        for off in range(OLD_LEN):
            for c in range(2, 14):
                v = ws.cell(s0 + off, c).value
                if isinstance(v, (int, float)):
                    typed[(nm, b, off, c)] = v
print('budget cells captured:', sum(1 for k in typed if k[0] == 'Budget FY27'))
# the blank separator row style
sep_style = {nm: {c: copy.copy(wb[nm].cell(OLD_STARTS[0] + OLD_LEN, c)._style)
                  for c in range(1, MAXC + 1)} for nm in SHAPED}

# ===================================================================== 2 ===
# Clear the old block area, then write every block back at its new position.
QSTART = 'DATE(YEAR(Setup!$B$5),FLOOR(MONTH(Setup!$B$5)-1,3)+1,1)'

def pl_month(off, kind, c, ec, s, eng_row):
    """c is the P&L month column B..M.  ec is the engine's, E..P."""
    if off == O_DEPT:
        return f'={c}{s+O_INC}+{c}{s+O_COS}+{c}{s+O_EXP}+{c}{s+O_OI}+{c}{s+O_OE}'
    if off in CAT_OFFS:
        ms = CAT_OFFS[off]
        return f'=SUM({c}{s+ms[0]}:{c}{s+ms[-1]})'
    if off in SUBCAT_OFFS:
        return f"=IF({c}$5<=Setup!$B$5,Engine!${ec}{eng_row},'Budget FY27'!{c}{s+off})"
    if kind == 'dep':
        er = eng_row
        return (f"=IF({c}$5<=Setup!$B$5,Engine!${ec}{er}+IF(Setup!$B$28=\"Yes\","
                f"Engine!${ec}{er+1},0),'Budget FY27'!{c}{s+off})")
    return margin_formula(kind, c, s)

def budget_month(off, kind, c, s):
    if off == O_DEPT:
        return f'={c}{s+O_INC}+{c}{s+O_COS}+{c}{s+O_EXP}+{c}{s+O_OI}+{c}{s+O_OE}'
    if off in CAT_OFFS:
        ms = CAT_OFFS[off]
        return f'=SUM({c}{s+ms[0]}:{c}{s+ms[-1]})'
    if off in SUBCAT_OFFS:
        return None                       # a typed input, left to the caller
    return margin_formula(kind, c, s)

def total_month(off, kind, c, s):
    if kind in PCT_ROWS: return margin_formula(kind, c, s)
    return '=' + '+'.join(f'{c}{st + off}' for st in NEW_STARTS)

for nm in SHAPED:
    ws = wb[nm]
    for r in range(OLD_STARTS[0], 233):
        for c in range(1, MAXC + 1):
            ws.cell(r, c).value = None
        ws.row_dimensions[r].outlineLevel = 0
        ws.row_dimensions[r].hidden = False
    for b, s in enumerate(ALL_STARTS):
        is_total = b == 6
        dept = None if is_total else DEPTS[b]
        for off in range(NEW_LEN):
            r = s + off
            kind = NEW_KIND.get(off)
            pct = kind in PCT_ROWS if kind else False
            # label
            if off == O_DEPT:
                lab = 'TOTAL COMPANY' if is_total else dept
            elif off in NEW_LABEL:
                lab = NEW_LABEL[off]
            else:
                lab = labels[(nm, off)]
            ws.cell(r, 1, lab)
            # styles and row shape
            base = off if off < OLD_LEN else O_INC          # margins look like a category row
            for c in range(1, MAXC + 1):
                ws.cell(r, c)._style = copy.copy(style_of[(nm, base, c)])
            lvl, hid, ht = rowmeta[(nm, base)]
            ws.row_dimensions[r].outlineLevel = 0 if off == O_DEPT else max(lvl, 1)
            ws.row_dimensions[r].hidden = hid if off < OLD_LEN else False
            if ht: ws.row_dimensions[r].height = ht
            if off == O_DEPT:
                ws.cell(r, 1).font = f(11, True, '1F3864')
            elif off in NEW_LABEL:
                ws.cell(r, 1).font = f(10, True)
            # ---------------- the month columns -------------------------
            src_off = (O_DEPT if kind == 'net_pct' else
                       {'gp_pct': 37, 'ebitda_pct': 40, 'ebit_pct': 42}.get(kind))
            eng_row = None
            if not is_total:
                if off in SUBCAT_OFFS:
                    eng_row = ENG_SUB_ROW[(dept, *OFF_SUB[off])]
                elif kind == 'dep':
                    eng_row = ENG_ADD_ROW[(dept, 'D&A')]
            if nm != 'Summary':
                for i in range(12):
                    c, ec = CL(2 + i), CL(5 + i)
                    if is_total:
                        v = total_month(off, kind, c, s)
                    elif nm == 'P&L FY27':
                        v = pl_month(off, kind, c, ec, s, eng_row)
                    else:
                        v = budget_month(off, kind, c, s)
                        if v is None or kind == 'dep':
                            v = typed.get((nm, b, off, 2 + i))
                    ws.cell(r, 2 + i, v)
                    if pct: ws.cell(r, 2 + i).number_format = PCT
            # ---------------- the analysis columns ----------------------
            if pct and nm == 'P&L FY27':
                for col in (14, 15, 17, 21, 22):
                    cc = CL(col)
                    ws.cell(r, col, f'=IF({cc}{s+O_INC}=0,"",{cc}{s+src_off}/{cc}{s+O_INC})')
            if nm == 'P&L FY27':
                if pct:
                    ws.cell(r, 18, f'=IF(OR($O{r}="",$Q{r}=""),"",$O{r}-$Q{r})')
                    ws.cell(r, 23, f'=IF(OR($U{r}="",$V{r}=""),"",$U{r}-$V{r})')
                    ws.cell(r, 26, f'=IF(OR($O{r}="",$U{r}=""),"",$O{r}-$U{r})')
                    ws.cell(r, 27, f'=IF(OR($Q{r}="",$V{r}=""),"",$Q{r}-$V{r})')
                    for col in (14, 15, 17, 18, 21, 22, 23, 26, 27):
                        ws.cell(r, col).number_format = PCT
                else:
                    ws.cell(r, 14, f'=SUM(B{r}:M{r})')
                    ws.cell(r, 15, f"=SUMIF('P&L FY27'!$B$5:$M$5,\"<=\"&Setup!$B$5,"
                                   f"'P&L FY27'!$B{r}:$M{r})")
                    ws.cell(r, 17, f"=SUMIF('Budget FY27'!$B$5:$M$5,\"<=\"&Setup!$B$5,"
                                   f"'Budget FY27'!$B{r}:$M{r})")
                    ws.cell(r, 18, f'=O{r}-Q{r}')
                    ws.cell(r, 19, f'=IFERROR(R{r}/ABS(Q{r}),0)')
                    ws.cell(r, 21, f"=SUMIF('Actual FY26'!$B$5:$M$5,\"<=\"&Setup!$B$5,"
                                   f"'Actual FY26'!$B{r}:$M{r})")
                    ws.cell(r, 22, f"=SUMIF('Budget FY26'!$B$5:$M$5,\"<=\"&Setup!$B$5,"
                                   f"'Budget FY26'!$B{r}:$M{r})")
                    ws.cell(r, 23, f'=U{r}-V{r}')
                    ws.cell(r, 24, f'=IFERROR(W{r}/ABS(V{r}),0)')
                    ws.cell(r, 26, f'=O{r}-U{r}')
                    ws.cell(r, 27, f'=Q{r}-V{r}')
                    ws.cell(r, 28, f'=Z{r}-AA{r}')
                    ws.cell(r, 29, f'=IFERROR(AB{r}/ABS(AA{r}),0)')
                    ws.cell(r, 30,
                            f'=IF(AND(O{r}=0,Q{r}=0),"No Activity",'
                            f'IF(ABS(R{r})<Setup!$B$14,"Low",'
                            f'IF(OR(ABS(R{r})>=Setup!$B$17,ABS(S{r})>=Setup!$B$16),"High",'
                            f'IF(ABS(S{r})>=Setup!$B$15,"Medium","Low"))))')
            elif nm == 'Summary':
                for i, col in enumerate((2, 3, 4)):
                    cc = CL(col)
                    if pct:
                        ws.cell(r, col, f'=IF({cc}{s+O_INC}=0,"",{cc}{s+src_off}/{cc}{s+O_INC})')
                        ws.cell(r, col).number_format = PCT
                    else:
                        ws.cell(r, col, f"=SUMIF('P&L FY27'!$B$5:$M$5,{cc}$5,"
                                        f"'P&L FY27'!$B{r}:$M{r})")
                if pct:
                    ws.cell(r, 5, f'=IF(E{s+O_INC}=0,"",E{s+src_off}/E{s+O_INC})')
                    ws.cell(r, 7, f'=IF(G{s+O_INC}=0,"",G{s+src_off}/G{s+O_INC})')
                    ws.cell(r, 8, f'=IF(OR($E{r}="",$G{r}=""),"",$E{r}-$G{r})')
                    ws.cell(r, 11, f'=IF(K{s+O_INC}=0,"",K{s+src_off}/K{s+O_INC})')
                    ws.cell(r, 12, f'=IF(L{s+O_INC}=0,"",L{s+src_off}/L{s+O_INC})')
                    ws.cell(r, 13, f'=IF(OR($K{r}="",$L{r}=""),"",$K{r}-$L{r})')
                    for col in (5, 7, 8, 11, 12, 13): ws.cell(r, col).number_format = PCT
                else:
                    ws.cell(r, 5, f"=SUMIF('P&L FY27'!$B$5:$M$5,\"<=\"&Setup!$B$5,"
                                  f"'P&L FY27'!$B{r}:$M{r})")
                    ws.cell(r, 7, f"=SUMIF('Budget FY27'!$B$5:$M$5,\"<=\"&Setup!$B$5,"
                                  f"'Budget FY27'!$B{r}:$M{r})")
                    ws.cell(r, 8, f'=E{r}-G{r}')
                    ws.cell(r, 9, f'=IFERROR(H{r}/ABS(G{r}),0)')
                    ws.cell(r, 11, f"=SUMIFS('P&L FY27'!$B{r}:$M{r},"
                                   f"'P&L FY27'!$B$5:$M$5,\"<=\"&Setup!$B$5,"
                                   f"'P&L FY27'!$B$5:$M$5,\">=\"&{QSTART})")
                    ws.cell(r, 12, f"=SUMIFS('Budget FY27'!$B{r}:$M{r},"
                                   f"'Budget FY27'!$B$5:$M$5,\"<=\"&Setup!$B$5,"
                                   f"'Budget FY27'!$B$5:$M$5,\">=\"&{QSTART})")
                    ws.cell(r, 13, f'=K{r}-L{r}')
                    ws.cell(r, 14, f'=IFERROR(M{r}/ABS(L{r}),0)')
                    ws.cell(r, 16, f"=E{r}-'Actual FY26'!O{r}")
                    ws.cell(r, 27, f"=G{r}-'Budget FY26'!O{r}")
                    ws.cell(r, 28, f'=R{r}-AA{r}')
                    ws.cell(r, 29, f'=IFERROR(AB{r}/ABS(AA{r}),0)')
                    ws.cell(r, 30,
                            f'=IF(AND(E{r}=0,G{r}=0),"No Activity",'
                            f'IF(ABS(H{r})<Setup!$B$14,"Low",'
                            f'IF(OR(ABS(H{r})>=Setup!$B$17,ABS(I{r})>=Setup!$B$16),"High",'
                            f'IF(ABS(I{r})>=Setup!$B$15,"Medium","Low"))))')
            else:
                if pct:
                    ws.cell(r, 14, f'=IF(N{s+O_INC}=0,"",N{s+src_off}/N{s+O_INC})')
                    ws.cell(r, 14).number_format = PCT
                else:
                    ws.cell(r, 14, f'=SUM(B{r}:M{r})')
        # the blank separator row between blocks
        if s + NEW_LEN <= NEW_LAST + 1:
            for c in range(1, MAXC + 1):
                ws.cell(s + NEW_LEN, c)._style = copy.copy(sep_style[nm][c])
    ws.sheet_properties.outlinePr = Outline(summaryBelow=False, summaryRight=False,
                                            applyStyles=False)
wb['P&L FY27'].freeze_panes = 'B6'          # was I108, which showed nothing useful
print('shaped sheets rewritten')

# ===================================================================== 3 ===
# Lists: the 14 accounts live in Xero that the master did not carry, plus the
# Addback column that drives the EBITDA bridge.
ls = wb['Lists']
ls.cell(5, 6, 'Addback').font = f(10, True, 'FFFFFF')
ls.cell(5, 6).fill = NAVY
ls.cell(5, 6).alignment = Alignment(horizontal='center', wrap_text=True)
ls.column_dimensions['F'].width = 14
have = {str(ls.cell(r, 1).value).strip() for r in range(6, 198) if ls.cell(r, 1).value}
ADD_BY_CODE = {m['code']: m['addback'] for m in MASTER if m['addback']}
for r in range(6, 198):
    code = ls.cell(r, 1).value
    if code is None: continue
    cell = ls.cell(r, 6, ADD_BY_CODE.get(str(code).strip(), ''))
    cell.font = f(9); cell.border = BOX
newrows = [m for m in MASTER if m['new']]
LS_LAST = 197 + len(newrows)
for i, m in enumerate(newrows):
    r = 198 + i
    for c, v in enumerate([m['code'], m['name'], m['cat'], m['sub'], '', m['addback']], 1):
        cell = ls.cell(r, c, v); cell.font = f(9); cell.border = BOX; cell.fill = NEWF
ls.cell(LS_LAST + 2, 1,
        'The rows shaded cream are accounts live in Xero that this master did not carry, so every '
        'GL line on them was being left out of the P&L. 41500 Contract Help Desk is revenue. '
        'Their category and subcategory are my read of the account name and its Xero type. '
        'Please check them.').font = f(9, False, 'C00000')
ls.cell(LS_LAST + 4, 1,
        'Addback: D&A and IAWO are added back to reach EBITDA. TAX is held out of net profit '
        'before tax. Setup B27 decides whether the instant asset write off counts as depreciation.'
        ).font = f(9, False, '595959')
for nm_, ref in [('AcctCode', f'Lists!$A$6:$A${LS_LAST}'),
                 ('AcctCat', f'Lists!$C$6:$C${LS_LAST}'),
                 ('AcctSub', f'Lists!$D$6:$D${LS_LAST}'),
                 ('AcctAdd', f'Lists!$F$6:$F${LS_LAST}')]:
    if nm_ in wb.defined_names: del wb.defined_names[nm_]
    wb.defined_names.add(DefinedName(nm_, attr_text=ref))

# balance sheet master, on spare columns of her Lists sheet
ls.cell(4, 20, 'Balance sheet master').font = f(10, True, '1F3864')
for i, h in enumerate(['Code', 'Account name', 'Group', 'Flag']):
    c = ls.cell(5, 20 + i, h); c.font = f(10, True, 'FFFFFF'); c.fill = NAVY
    c.alignment = Alignment(horizontal='center')
for i, a in enumerate(BSACC):
    flag = ('Cash' if a['cash'] else 'Borrowings' if a['debt'] else
            'Trade debtors' if a['debtors'] else 'Trade creditors' if a['creditors'] else
            'WIP' if a['wip'] else '')
    for c, v in enumerate([a['code'], a['name'], a['group'], flag], 20):
        cell = ls.cell(6 + i, c, v); cell.font = f(9); cell.border = BOX
for col, w in zip('TUVW', (10, 34, 24, 15)): ls.column_dimensions[col].width = w

# ===================================================================== 4 ===
# Cleanup: every formula read GL_Paste!#REF!. Rewritten, and now it ignores any
# row with no real date so the export's title rows do not matter.
cu = wb['Cleanup']
cu.cell(2, 1, 'Amount is Credit less Debit, so income is positive and costs are negative. '
              'Department comes from the Cost Centres column, falling back to the bracket tag on '
              'the job number. Any row with no real date in the Date column is ignored, so the '
              "export's title rows do no harm.")
cu.cell(4, 7, 'Addback').font = f(10, True, 'FFFFFF')
cu.cell(4, 7).fill = NAVY
cu.cell(4, 7).alignment = Alignment(horizontal='center')
cu.column_dimensions['G'].width = 14
for i in range(CLN_R1 - CLN_R0 + 1):
    r = CLN_R0 + i; g = GL_R0 + i
    gate = f'N(GL_Paste!$E{g})=0'
    cu.cell(r, 1, f'=IF({gate},"",N(GL_Paste!$H{g})-N(GL_Paste!$G{g}))')
    cu.cell(r, 2, f'=IF({gate},"",DATE(YEAR(GL_Paste!$E{g}),MONTH(GL_Paste!$E{g}),1))')
    cu.cell(r, 3, f'=IF({gate},"",IF(TRIM(GL_Paste!$M{g})<>"",TRIM(GL_Paste!$M{g}),'
                  f'IFERROR(MID(GL_Paste!$I{g},FIND("[",GL_Paste!$I{g})+1,'
                  f'FIND("]",GL_Paste!$I{g})-FIND("[",GL_Paste!$I{g})-1),"")))')
    cu.cell(r, 4, f'=IF($C{r}="","",IFERROR(INDEX(DeptName,MATCH($C{r},DeptTag,0)),"UNALLOCATED"))')
    cu.cell(r, 5, f'=IF({gate},"",IFERROR(INDEX(AcctCat,MATCH(GL_Paste!$B{g}&"",AcctCode,0)),'
                  f'"UNKNOWN"))')
    cu.cell(r, 6, f'=IF({gate},"",IFERROR(INDEX(AcctSub,MATCH(GL_Paste!$B{g}&"",AcctCode,0)),'
                  f'"UNKNOWN"))')
    cu.cell(r, 7, f'=IF({gate},"",IFERROR(INDEX(AcctAdd,MATCH(GL_Paste!$B{g}&"",AcctCode,0)),""))')
    cu.cell(r, 7).font = f(9)
cu.auto_filter.ref = f'A4:G{CLN_R1}'
CL_A = f'Cleanup!$A${CLN_R0}:$A${CLN_R1}'
CL_B = f'Cleanup!$B${CLN_R0}:$B${CLN_R1}'
CL_D = f'Cleanup!$D${CLN_R0}:$D${CLN_R1}'
CL_G = f'Cleanup!$G${CLN_R0}:$G${CLN_R1}'

# ===================================================================== 5 ===
# Engine: her 186 rows are untouched. Eighteen add-back rows go underneath.
en = wb['Engine']
for d in DEPTS:
    for a in ADDBACKS:
        r = ENG_ADD_ROW[(d, a)]
        en.cell(r, 1, f'{d}|{a}').font = f(9, False, '808080')
        for c, v in zip((2, 3, 4), (d, 'Addback', ADDLABEL[a])): en.cell(r, c, v).font = f(9)
        for i in range(12):
            c = CL(5 + i)
            cell = en.cell(r, 5 + i,
                           f'=SUMIFS({CL_A},{CL_D},$B{r},{CL_G},"{a}",{CL_B},{c}$4)')
            cell.number_format = MON2; cell.font = f(9)
        en.cell(r, 17, f'=SUM(E{r}:P{r})').number_format = MON2
        en.cell(r, 17).font = f(9)
print('engine add-back rows', ENG_ADD_ROW[(DEPTS[0], "D&A")], 'to', ENG_ADD_LAST)

# ===================================================================== 6 ===
# Setup: the three #REF! cells, the wrong financial year start, and the few
# new controls. Nothing existing moves, because the risk thresholds on B14 to
# B17 are referenced by every risk formula in the pack.
st = wb['Setup']
st['B6'] = 7                       # was 8, which made September period 2
st['B20'] = '=COUNT(GL_Paste!$E$2:$E$20001)'
st['B24'] = '=IFERROR(MIN(GL_Paste!$E$2:$E$20001),0)'
st['B25'] = '=IFERROR(MAX(GL_Paste!$E$2:$E$20001),0)'
st['C6'] = 'July'
def srow(r, label, val, note='', nf='General', is_inp=False):
    st.cell(r, 1, label).font = f(10)
    c = st.cell(r, 2, val)
    if is_inp:
        c.font = f(10, False, '0000FF'); c.fill = YELL; c.border = BOX
    else:
        c.font = f(10)
    c.number_format = nf
    if note:
        n = st.cell(r, 3, note); n.font = f(9, False, '595959')
        n.alignment = Alignment(wrap_text=True, vertical='top')
st.cell(27, 1, '4.  The pasted reports, and the EBITDA bridge').font = f(10, True, '1F3864')
srow(28, 'Count Instant Asset Write Off as depreciation', 'Yes',
     'Yes or No. It is capital spending put through the P&L, so it changes EBITDA either way. '
     'Yes adds it back.', 'General', True)
srow(29, 'P&L paste: column holding the reporting month', 2,
     'Column B is 2, C is 3, and so on. PL_Check shows the heading it found in that column so you '
     'can see you picked the right one. Set it once.', '0', True)
srow(30, 'BS paste: column holding this month', 2, '', '0', True)
srow(31, 'BS paste: column holding last month', 3,
     'Leave at 0 if your export has no comparison columns.', '0', True)
srow(32, 'BS paste: column holding the start of the year', 4,
     'Leave at 0 if your export has no comparison columns.', '0', True)
srow(34, 'P&L lines pasted', '=COUNTA(PL_Paste!$A$1:$A$300)', '', '#,##0')
srow(35, 'BS lines pasted', '=COUNTA(BS_Paste!$A$1:$A$400)', '', '#,##0')
srow(36, 'GL lines dated outside this financial year',
     '=$B$20-COUNTIFS(Cleanup!$B$5:$B$20004,">="&DATE($B$7-1,$B$6,1),'
     'Cleanup!$B$5:$B$20004,"<="&DATE($B$7,$B$6-1,1))',
     'These are left out. If it is not nil, check what you pasted.', '#,##0')

# ===================================================================== 7 ===
# PL_Paste, and two new columns on her PL_Check so the five category totals no
# longer have to be typed. Her typed column C and its check stay exactly as is.
pp = wb.create_sheet('PL_Paste')
for col, w in [('A', 44)] + [(CL(c), 15) for c in range(2, 17)]:
    pp.column_dimensions[col].width = w
for r in range(1, 301):
    for c in range(1, 17):
        cell = pp.cell(r, c); cell.font = f(9); cell.fill = YELL; cell.border = BOX
        if c > 1: cell.number_format = MON2
pp.freeze_panes = 'B2'
pp.sheet_view.showGridLines = False

pc = wb['PL_Check']
pc['A2'] = ('Either paste the Xero profit and loss on PL_Paste, or type the five totals into the '
            'yellow cells in column C. Every status should read OK before you report.')
for i, h in enumerate(['Label in your Xero P&L', 'Per the pasted P&L', 'Found?',
                       'Status vs the paste']):
    c = pc.cell(4, 7 + i, h)
    c.font = f(10, True, 'FFFFFF'); c.fill = NAVY; c.border = BOX
    c.alignment = Alignment(horizontal='center', wrap_text=True)
for col, w in zip('GHIJ', (30, 16, 12, 16)): pc.column_dimensions[col].width = w
PLP, PLN = 'PL_Paste!$A$1:$P$300', 'PL_Paste!$A$1:$A$300'
DEFLBL = ['Total Income', 'Total Cost of Sales', 'Total Operating Expenses',
          'Total Other Income', 'Total Other Expenses']
for i, lbl in enumerate(DEFLBL):
    r = 5 + i
    c = pc.cell(r, 7, lbl)
    c.font = f(10, False, '0000FF'); c.fill = YELL; c.border = BOX
    pc.cell(r, 8, f'=IF($G{r}="",0,IFERROR(INDEX({PLP},MATCH($G{r},{PLN},0),'
                  f'Setup!$B$29),0))').number_format = MON2
    pc.cell(r, 9, f'=IF($G{r}="","no label",IF(ISNA(MATCH($G{r},{PLN},0)),"not found","found"))')
    pc.cell(r, 10, f'=IF(Setup!$B$34=0,"NO PASTE",IF($I{r}<>"found","CHECK LABEL",'
                   f'IF(ABS($B{r}-$H{r})<1,"OK","CHECK")))')
    for cc in (8, 9, 10):
        pc.cell(r, cc).font = f(10, cc == 10)
        pc.cell(r, cc).alignment = Alignment(horizontal='center' if cc != 8 else 'right')
pc.cell(10, 8, '=SUM(H5:H9)').number_format = MON2
pc.cell(10, 8).font = f(10, True)
pc.cell(10, 10, '=IF(Setup!$B$34=0,"NO PASTE",IF(ABS($B10-$H10)<1,"OK","CHECK"))')
pc.cell(10, 10).font = f(10, True)
pc.cell(10, 10).alignment = Alignment(horizontal='center')
pc.cell(12, 7, 'If a label reads "not found", change the text in column G to match your export '
               'word for word.').font = f(9, False, '595959')

# ===================================================================== 8 ===
# BS_Paste and a balance sheet, in the same style as the rest of the pack.
bp = wb.create_sheet('BS_Paste')
for col, w in [('A', 44)] + [(CL(c), 15) for c in range(2, 17)]:
    bp.column_dimensions[col].width = w
for r in range(1, 401):
    for c in range(1, 17):
        cell = bp.cell(r, c); cell.font = f(9); cell.fill = YELL; cell.border = BOX
        if c > 1: cell.number_format = MON2
bp.freeze_panes = 'B2'
bp.sheet_view.showGridLines = False

bs = wb.create_sheet('Balance_Sheet')
bs.sheet_view.showGridLines = False
bs['A1'] = 'Balance sheet'
bs['A1'].font = f(14, True, '1F3864')
bs['A2'] = ('Paste the Xero balance sheet on BS_Paste, whole export, at A1. This reads it by '
            'matching the account name in column A. Tell Setup rows 30 to 32 which columns hold '
            'the balances.')
bs['A2'].font = f(9, False, '595959')
bs['A3'] = '=TEXT(Setup!$B$5,"mmm yyyy")'
bs['A3'].font = f(10, True, '1F3864')
for i, h in enumerate(['Line', 'This month', 'Last month', 'Start of year',
                       'Movement in month', 'Movement in year', '', 'Code and flag']):
    c = bs.cell(5, 1 + i, h)
    c.font = f(10, True, 'FFFFFF'); c.fill = NAVY; c.border = BOX
    c.alignment = Alignment(horizontal='center', wrap_text=True)
for col, w in zip('ABCDEFGH', (44, 15, 15, 15, 15, 15, 2, 26)):
    bs.column_dimensions[col].width = w
bs.freeze_panes = 'B6'
BSP, BSN = 'BS_Paste!$A$1:$P$400', 'BS_Paste!$A$1:$A$400'
GROUPS = ['Current Assets', 'Non-Current Assets', 'Current Liabilities',
          'Non-Current Liabilities', 'Equity']
grp_row = {}
r = 6
for g in GROUPS:
    members = [a for a in BSACC if a['group'] == g]
    grp_row[g] = r
    bs.cell(r, 1, g).font = f(11, True, '1F3864')
    first, last = r + 1, r + len(members)
    for c in (2, 3, 4):
        bs.cell(r, c, f'=SUM({CL(c)}{first}:{CL(c)}{last})')
    r += 1
    for a in members:
        bs.cell(r, 1, '    ' + a['name']).font = f(9)
        for c, sc in ((2, 30), (3, 31), (4, 32)):
            bs.cell(r, c, f'=IF(Setup!$B${sc}=0,0,IFERROR(INDEX({BSP},'
                          f'MATCH(TRIM($A{r}),{BSN},0),Setup!$B${sc}),0))')
        flag = ('Cash' if a['cash'] else 'Borrowings' if a['debt'] else
                'Trade debtors' if a['debtors'] else 'Trade creditors' if a['creditors'] else
                'WIP' if a['wip'] else '')
        bs.cell(r, 8, f"{a['code']}  {flag}").font = f(8, False, '808080')
        bs.row_dimensions[r].outlineLevel = 1
        bs.row_dimensions[r].hidden = True
        r += 1
    r += 1
BS_DET_LAST = r - 2
TA, TL, NA, TE, CHK = r, r + 1, r + 2, r + 3, r + 5
for rr, lab in ((TA, 'Total assets'), (TL, 'Total liabilities'), (NA, 'Net assets'),
                (TE, 'Total equity'), (CHK, 'Check  -  net assets less total equity')):
    bs.cell(rr, 1, lab).font = f(11, True)
for c in (2, 3, 4):
    cc = CL(c)
    bs.cell(TA, c, f'={cc}{grp_row["Current Assets"]}+{cc}{grp_row["Non-Current Assets"]}')
    bs.cell(TL, c, f'={cc}{grp_row["Current Liabilities"]}+{cc}{grp_row["Non-Current Liabilities"]}')
    bs.cell(NA, c, f'={cc}{TA}-{cc}{TL}')
    bs.cell(TE, c, f'={cc}{grp_row["Equity"]}')
    bs.cell(CHK, c, f'=ROUND({cc}{NA}-{cc}{TE},2)')
bs.cell(CHK, 6, f'=IF(Setup!$B$35=0,"NO PASTE",IF(AND(ABS(B{CHK})<0.5,ABS(C{CHK})<0.5,'
                f'ABS(D{CHK})<0.5),"OK","CHECK"))')
bs.cell(CHK, 6).font = f(11, True)
for rr in range(6, CHK + 1):
    if bs.cell(rr, 1).value is None: continue
    top = not str(bs.cell(rr, 1).value).startswith('    ')
    for c in (2, 3, 4, 5, 6):
        cell = bs.cell(rr, c); cell.font = f(10, top); cell.number_format = MONEY
    if rr <= BS_DET_LAST or rr in (TA, TL, NA, TE):
        bs.cell(rr, 5, f'=$B{rr}-$C{rr}')
        bs.cell(rr, 6, f'=$B{rr}-$D{rr}')
bs.cell(CHK, 6).number_format = 'General'
KF = CHK + 2
bs.cell(KF, 1, 'Key figures').font = f(11, True, '1F3864')
kf_row = {}
for i, (lab, flag) in enumerate([('Cash and bank', 'Cash'), ('Trade debtors', 'Trade debtors'),
                                 ('Trade creditors', 'Trade creditors'),
                                 ('Work in progress', 'WIP'),
                                 ('Borrowings and lease liabilities', 'Borrowings')]):
    rr = KF + 1 + i; kf_row[flag] = rr
    bs.cell(rr, 1, lab).font = f(10)
    for c in (2, 3, 4):
        cc = CL(c)
        bs.cell(rr, c, f'=SUMIF($H$6:$H${BS_DET_LAST},"*"&"{flag}",{cc}$6:{cc}${BS_DET_LAST})')
        bs.cell(rr, c).number_format = MONEY; bs.cell(rr, c).font = f(10)
WK = kf_row['Borrowings'] + 2
bs.cell(WK, 1, 'Working capital and gearing').font = f(11, True, '1F3864')
CA, CLB = grp_row['Current Assets'], grp_row['Current Liabilities']
for i, (lab, fx, nf) in enumerate([
        ('Working capital', f'=$B${CA}-$B${CLB}', MONEY),
        ('Current ratio', f'=IF($B${CLB}=0,"",$B${CA}/$B${CLB})', '0.00x;(0.00x);"-"'),
        ('Quick ratio', f'=IF($B${CLB}=0,"",($B${CA}-$B${kf_row["WIP"]})/$B${CLB})',
         '0.00x;(0.00x);"-"'),
        ('Net debt', f'=$B${kf_row["Borrowings"]}-$B${kf_row["Cash"]}', MONEY)]):
    rr = WK + 1 + i
    bs.cell(rr, 1, lab).font = f(10)
    bs.cell(rr, 2, fx); bs.cell(rr, 2).number_format = nf; bs.cell(rr, 2).font = f(10, True)
bs.sheet_properties.outlinePr = Outline(summaryBelow=False, summaryRight=False, applyStyles=False)

# ===================================================================== 9 ===
rd = wb['README']
base = max(r for r in range(1, 60) if rd.cell(r, 2).value) + 2
rd.cell(base, 2, 'WHAT CHANGED IN THIS VERSION').font = f(10, True, 'C00000')
for i, (a, b) in enumerate([
    ('Cleanup was dead', 'Every formula on Cleanup pointed at deleted rows, so no sheet in '
     'the pack could calculate. Deleting the rows on GL_Paste is what broke them. Rewritten.'),
    ('Setup B6', 'The financial year start month was 8, August, against a July year end. '
     'September was coming out as period 2. It now reads 7.'),
    ('Setup B20, B24, B25', 'Broken the same way. Rewritten.'),
    ('Margins', 'Eight rows at the bottom of every department block: gross profit, gross profit %, '
     'depreciation & amortisation, EBITDA, EBITDA %, EBIT, EBIT % and net profit %. On the '
     'Summary, the P&L and all three budget and prior year sheets, so they stay the same shape.'),
    ('TOTAL COMPANY', 'A seventh block under the six departments, with the same lines.'),
    ('EBIT and EBITDA', 'EBITDA is gross profit plus overheads with depreciation added back. '
     'EBIT is gross profit plus overheads. Other income and other expenses sit below EBIT, which '
     'is where interest and tax already were in this pack.'),
    ('PL_Paste', 'Paste the Xero profit and loss and the five category totals read themselves. '
     'The typed column on PL_Check still works if you would rather type them.'),
    ('BS_Paste and Balance_Sheet', 'A balance sheet, which the pack did not have. All 79 live '
     'Xero accounts, grouped, with the movement in the month and in the year, working capital, '
     'current and quick ratio and net debt.'),
    ('Lists', '14 accounts live in Xero were missing from the master, so every GL line on them '
     'was being dropped. One of them, 41500 Contract Help Desk, is revenue. Added and shaded '
     'cream. Please check how I have categorised them.'),
    ('GL_Paste', 'Still paste at A2, as before. Rows with no real date are now ignored, so the '
     "export's title rows do no harm."),
]):
    r = base + 1 + i
    rd.cell(r, 2, a).font = f(10, True)
    c = rd.cell(r, 3, b); c.font = f(10)
    c.alignment = Alignment(wrap_text=True, vertical='top')
    rd.row_dimensions[r].height = 13 * (1 + len(b) // 105)

# ==================================================================== 10 ===
# The risk column had no colour on it. Red for High, amber for Medium, on the
# Risk cell only, so the rest of the sheet looks the same as it always did.
from openpyxl.formatting.rule import FormulaRule
RED = PatternFill('solid', fgColor='FFC7CE')
AMB = PatternFill('solid', fgColor='FFEB9C')
for nm_ in ('Summary', 'P&L FY27'):
    w = wb[nm_]
    rng = f'$AD$6:$AD${NEW_LAST}'
    w.conditional_formatting.add(rng, FormulaRule(formula=['$AD6="High"'], fill=RED,
                                                  font=Font(name=AR, size=10, bold=True,
                                                            color='9C0006')))
    w.conditional_formatting.add(rng, FormulaRule(formula=['$AD6="Medium"'], fill=AMB,
                                                  font=Font(name=AR, size=10, color='9C6500')))

# the three new sheets go where she would look for them, and nothing of hers moves
ORDER = ['README', 'requirements ', 'Setup', 'Summary', 'Engine', 'PL_Check', 'P&L FY27',
         'Balance_Sheet', 'Actual FY26', 'Budget FY27', 'Budget FY26', 'GL_Paste', 'PL_Paste',
         'BS_Paste', 'Utilisation', 'Lists', 'Cleanup']
assert sorted(ORDER) == sorted(wb.sheetnames), set(ORDER) ^ set(wb.sheetnames)
wb._sheets = [wb[n] for n in ORDER]
wb.active = ORDER.index('Summary')
wb['Summary'].sheet_view.tabSelected = True
for n in ORDER:
    wb[n].sheet_view.tabSelected = (n == 'Summary')

wb.save(OUT)
print('saved', OUT)
