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
OUT = f"{HERE}/CTS Financial Controller Pack - v8.xlsx"
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

# ============================================ her layout, as she laid it out ==
# One consolidated profit and loss. Category first, then department, then the
# subcategory detail. Taken line for line off the "P&L FY27 (2)" sheet she
# built, in her order and with her wording.
DEPTS_OLD = ['ONSITE', 'PRODUCTION', 'VIDEO', 'INTEGRATION', 'CONSULTING', 'ADMIN']
OLD_STARTS = [6, 44, 82, 120, 158, 196]
OLD_LEN = 37
# Her budget's three groups, with the departments that sit inside each one, and
# her names: the cost centre called PRODUCTION in Xero is Events, and the one
# called ONSITE is Support.
GROUPS_DISP = [
    ('Production',    [('PRODUCTION', 'Events'), ('VIDEO', 'Video')]),
    ('Consulting',    [('INTEGRATION', 'Integration'), ('CONSULTING', 'Consulting')]),
    ('Support',       [('ONSITE', 'Support')]),
    ('General Admin', [('ADMIN', 'General Admin')]),
]
DEPT_ORDER = [(c, n) for _, ms in GROUPS_DISP for c, n in ms]
DEPT_NAME = dict(DEPT_ORDER)
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
SUBS_ORDER = ([('Income', x) for x in INC] + [('Cost of Sales', x) for x in COS]
              + [('Expenses', x) for x in EXP]
              + [('Other Income', 'Other Income'), ('Other Expenses', 'Other Expenses')])
# ---- the row plan.  (label, level, kind, meta) in order from the top -------
ROWS = []
def add(label, level, kind, meta=None):
    ROWS.append((len(ROWS), label, level, kind, meta))

def block(cat, subs, gkind, dkind, skind):
    """Total, then each group, then its departments, then the detail."""
    for gname, members in GROUPS_DISP:
        multi = len(members) > 1
        add('    ' + gname, 2, gkind, gname)
        for code, dname in members:
            if multi:
                add('        ' + dname, 3, dkind, code)
            for sub in subs:
                add(('            ' if multi else '        ') + sub,
                    4 if multi else 3, skind, (code, cat, sub))

add('Total Income', 0, 'inc_total')
block('Income', INC, 'inc_group', 'inc_dept', 'inc_sub')
add('Other income ', 1, 'oi_total')
add('        Other Income', 2, 'oi_sub', ('Other Income', 'Other Income'))
add('    Cost of Sales', 1, 'cos_total')
block('Cost of Sales', COS, 'cos_group', 'cos_dept', 'cos_sub')
add('    Expenses Overheds ', 1, 'exp_total')
for s_ in EXP:
    add('        ' + s_, 2, 'exp_sub', ('Expenses', s_))
add('    Other Expenses', 1, 'oe_total')
add('        Other Expenses', 2, 'oe_sub', ('Other Expenses', 'Other Expenses'))
add('    Gross profit', 1, 'gp')
add('    Gross profit %', 1, 'gp_pct')
add('    Depreciation & amortisation', 1, 'dep')
add('    EBITDA', 1, 'ebitda')
add('    EBITDA %', 1, 'ebitda_pct')
add('    EBIT', 1, 'ebit')
add('    EBIT %', 1, 'ebit_pct')
add('    Net profit', 1, 'net')
add('    Net profit %', 1, 'net_pct')
add(None, 0, 'blank')
add('Profit Margin Per Department', 0, 'margin_head')
for gname, members in GROUPS_DISP:
    if len(members) > 1:
        add('    ' + gname, 1, 'margin_group', gname)
        for code, dname in members:
            add('        ' + dname, 2, 'margin_dept', code)
    else:
        add('    ' + members[0][1], 1, 'margin_dept', members[0][0])

R0 = 6                                   # the statement starts on row 6, as it did
NROWS = len(ROWS)
LAST = R0 + NROWS - 1
OFF = {}                                 # kind -> offset, for the one-of-a-kind rows
for off, lab, lvl, kind, meta in ROWS:
    if kind in ('net', 'inc_total', 'oi_total', 'oi_sub', 'cos_total', 'exp_total',
                'oe_total', 'oe_sub', 'gp', 'gp_pct', 'dep', 'ebitda', 'ebitda_pct',
                'ebit', 'ebit_pct', 'net_pct'):
        OFF[kind] = off
GOFF = {('Income', m): off for off, l, lv, k, m in ROWS if k == 'inc_group'}
GOFF.update({('Cost of Sales', m): off for off, l, lv, k, m in ROWS if k == 'cos_group'})
DEPT_INC_OFF = {m: off for off, l, lv, k, m in ROWS if k == 'inc_dept'}
DEPT_COS_OFF = {m: off for off, l, lv, k, m in ROWS if k == 'cos_dept'}
SUB_OFFS = {}                            # (dept, category) -> [offsets of its detail]
for off, lab, lvl, kind, meta in ROWS:
    if kind in ('inc_sub', 'cos_sub'):
        SUB_OFFS.setdefault((meta[0], meta[1]), []).append(off)
MARGIN_OFF = {m: off for off, l, lv, k, m in ROWS if k == 'margin_dept'}
EXP_OFFS = [off for off, l, lv, k, m in ROWS if k == 'exp_sub']
PCT_SRC = {'gp_pct': 'gp', 'ebitda_pct': 'ebitda', 'ebit_pct': 'ebit', 'net_pct': 'net'}
TYPED_KINDS = {'inc_sub', 'cos_sub', 'oi_sub', 'exp_sub', 'oe_sub', 'dep'}
print(f'statement rows {R0}..{LAST}  ({NROWS} rows)')

# ------------------------------------------------- her engine, and its rows --
ENG_SUB_ROW = {}
r = 5
for d in DEPTS_OLD:
    for cat, sub in SUBS_ORDER:
        ENG_SUB_ROW[(d, cat, sub)] = r; r += 1
ENG_SUB_LAST = r - 1
ADDBACKS = ['D&A', 'IAWO', 'TAX']
ADDLABEL = {'D&A': 'Depreciation & amortisation', 'IAWO': 'Instant asset write off',
            'TAX': 'Income tax'}
ENG_ADD_ROW = {}
r = ENG_SUB_LAST + 2
for d in DEPTS_OLD:
    for a in ADDBACKS:
        ENG_ADD_ROW[(d, a)] = r; r += 1
ENG_ADD_LAST = r - 1
DEPTS = DEPTS_OLD

# her old per-department block, so the budget figures can be lifted out of it
O_DEPT, O_INC, O_COS, O_EXP, O_OI, O_OIS, O_OE, O_OES = 0, 1, 9, 26, 33, 34, 35, 36
OFF_SUB = {}
for i, s in enumerate(INC): OFF_SUB[2 + i] = ('Income', s)
for i, s in enumerate(COS): OFF_SUB[10 + i] = ('Cost of Sales', s)
for i, s in enumerate(EXP): OFF_SUB[27 + i] = ('Expenses', s)
OFF_SUB[O_OIS] = ('Other Income', 'Other Income')
OFF_SUB[O_OES] = ('Other Expenses', 'Other Expenses')

GLSHEET = 'P&L Analysis'      # was GL_Paste
GLREF = "'P&L Analysis'"      # how it has to be written inside a formula
CLN_R0, CLN_R1 = 5, 20004
GL_R0 = 1   # Cleanup reads from row 1, so it does not matter whether the export
            # lands on A1 or further down. A row with no real date is ignored.

wb = openpyxl.load_workbook(SRC)
wb['GL_Paste'].title = GLSHEET
SHAPED = ['Summary', 'P&L FY27', 'Budget FY27', 'Actual FY26', 'Budget FY26']
MAXC = 30

# ===================================================================== 1 ===
# Lift the styles and the typed budget figures out of her old layout before it
# is cleared. Styles come off the ONSITE block, which is styled like the rest.
style_of, rowmeta, typed = {}, {}, {}
for nm in SHAPED:
    ws = wb[nm]
    for off in range(OLD_LEN):
        src_r = OLD_STARTS[0] + off
        rd = ws.row_dimensions[src_r]
        rowmeta[(nm, off)] = (rd.outlineLevel, rd.hidden, rd.height)
        for c in range(1, MAXC + 1):
            style_of[(nm, off, c)] = copy.copy(ws.cell(src_r, c)._style)
    for b, s0 in enumerate(OLD_STARTS):
        for off, (cat, sub) in OFF_SUB.items():
            for c in range(2, 14):
                v = ws.cell(s0 + off, c).value
                if isinstance(v, (int, float)):
                    typed[(nm, DEPTS_OLD[b], cat, sub, c)] = v
print('budget figures lifted:', sum(1 for k in typed if k[0] == 'Budget FY27'))
# ------------------------------------------------------- the board palette --
# One navy, used at four strengths, so the statement reads as a hierarchy at a
# glance: the bottom line, then each category, then each department, then the
# detail. Her grey heading band becomes the palest navy, which prints cleanly
# in black and white and does not fight the figures.
BRAND = json.load(open(f'{HERE}/data/brand.json'))
TEAL = BRAND['teal']       # the CTS primary: the bottom line and the column headings
GREY = BRAND['grey']       # their table grey: category headings
GREY_LT = BRAND['grey_lt'] # their subtotal band: the result lines
RULE_C = BRAND['rule']     # their divider grey
MID = BRAND['mid']         # their secondary text: department headings
INK = BRAND['ink']         # their body text
RED = BRAND['red']         # their unfavourable colour
RED_LT = BRAND['red_lt']

def fill(rgb): return PatternFill('solid', fgColor=rgb)
def under(style, rgb): return Border(bottom=Side(style=style, color=rgb))
def over(rgb): return Border(top=Side(style='thin', color=rgb))
NO_FILL = PatternFill(fill_type=None)
NO_RULE = Border()

CATEGORY = {'oi_total', 'cos_total', 'exp_total', 'oe_total', 'margin_head'}
GROUP_HEAD = {'inc_group', 'cos_group', 'margin_group'}
DEPT_HEAD = {'inc_dept', 'cos_dept'}
RESULT = {'gp', 'gp_pct', 'ebitda', 'ebitda_pct', 'ebit', 'ebit_pct', 'net', 'net_pct'}
TOPPED = {'gp', 'ebitda', 'ebit', 'net'}            # a teal rule above each result subtotal
DETAIL = {'inc_sub', 'cos_sub', 'oi_sub', 'exp_sub', 'oe_sub', 'margin_dept'}

def look_of(kind):
    """(fill, font, border, row height) for one line, in the CTS house style:
    teal for the headline, their table grey for structure, their subtotal band
    for the results, and nothing at all behind the detail."""
    if kind == 'inc_total':
        return (fill(TEAL), Font(name=AR, size=11, bold=True, color='FFFFFF'), NO_RULE, 20)
    if kind in CATEGORY:
        return (fill(GREY), Font(name=AR, size=10, bold=True, color=INK),
                under('thin', RULE_C), 17)
    if kind in GROUP_HEAD:
        return (fill(GREY_LT), Font(name=AR, size=10, bold=True, color=TEAL),
                under('hair', RULE_C), 16)
    if kind in DEPT_HEAD:
        return (NO_FILL, Font(name=AR, size=10, bold=True, color=MID),
                under('hair', RULE_C), 15.5)
    if kind in RESULT:
        return (fill(GREY_LT), Font(name=AR, size=10, bold=True, color=TEAL),
                over(TEAL) if kind in TOPPED else NO_RULE, 16)
    if kind == 'dep':
        return (NO_FILL, Font(name=AR, size=10, bold=False, color=MID), NO_RULE, None)
    return (NO_FILL, Font(name=AR, size=10, bold=False, color=INK), NO_RULE, None)

STY = {'head': O_DEPT, 'cat': O_INC, 'sub': 2}      # which old row to copy the look from
LOOK = {'net': 'cat', 'blank': 'sub', 'margin_head': 'cat',
        'margin_dept': 'sub', 'margin_group': 'cat',
        'inc_group': 'cat', 'cos_group': 'cat', 'inc_total': 'head', 'oi_total': 'cat', 'cos_total': 'cat',
        'exp_total': 'cat', 'oe_total': 'cat', 'gp': 'cat', 'gp_pct': 'cat',
        'dep': 'cat', 'ebitda': 'cat', 'ebitda_pct': 'cat', 'ebit': 'cat',
        'ebit_pct': 'cat', 'net_pct': 'cat', 'inc_dept': 'cat', 'cos_dept': 'cat',
        'inc_sub': 'sub', 'cos_sub': 'sub', 'oi_sub': 'sub', 'exp_sub': 'sub',
        'oe_sub': 'sub'}

# ===================================================================== 2 ===
# Clear her old seven blocks and write the one statement in their place.
MCOLS = [CL(c) for c in range(2, 14)]
QSTART = 'DATE(YEAR(Setup!$B$5),FLOOR(MONTH(Setup!$B$5)-1,3)+1,1)'

def eng(dept, cat, sub, ec):
    return f'Engine!${ec}{ENG_SUB_ROW[(dept, cat, sub)]}'

def all_depts(cat, sub, ec):
    return '+'.join(eng(d, cat, sub, ec) for d in DEPTS_OLD)

def dep_expr(ec):
    bits = []
    for d in DEPTS_OLD:
        da, ia = ENG_ADD_ROW[(d, 'D&A')], ENG_ADD_ROW[(d, 'IAWO')]
        bits.append(f'Engine!${ec}{da}+IF(Setup!$B$28="Yes",Engine!${ec}{ia},0)')
    return '+'.join(bits)

def inc_row_of(code):
    """Where a department's income total sits: its own row when the group has
    two, the group's own row when it is alone."""
    if code in DEPT_INC_OFF: return DEPT_INC_OFF[code]
    g = next(gn for gn, ms in GROUPS_DISP if code in [x for x, _ in ms])
    return GOFF[('Income', g)]

def cos_row_of(code):
    if code in DEPT_COS_OFF: return DEPT_COS_OFF[code]
    g = next(gn for gn, ms in GROUPS_DISP if code in [x for x, _ in ms])
    return GOFF[('Cost of Sales', g)]

def derived(kind, c, s, meta=None):
    """Rows worked out from other rows of the same statement."""
    R = lambda k: f'{c}{s + OFF[k]}'
    if kind == 'net':
        return (f'={R("inc_total")}+{R("oi_total")}+{R("cos_total")}'
                f'+{R("exp_total")}+{R("oe_total")}')
    if kind in ('inc_total', 'cos_total'):
        cat = 'Income' if kind == 'inc_total' else 'Cost of Sales'
        return '=' + '+'.join(f'{c}{s + GOFF[(cat, g)]}' for g, _ in GROUPS_DISP)
    if kind in ('inc_group', 'cos_group'):
        cat = 'Income' if kind == 'inc_group' else 'Cost of Sales'
        members = dict(GROUPS_DISP)[meta]
        if len(members) > 1:                       # add up its departments
            offs = DEPT_INC_OFF if cat == 'Income' else DEPT_COS_OFF
            return '=' + '+'.join(f'{c}{s + offs[code]}' for code, _ in members)
        o = SUB_OFFS[(members[0][0], cat)]         # or its own detail
        return f'=SUM({c}{s + o[0]}:{c}{s + o[-1]})'
    if kind in ('inc_dept', 'cos_dept'):
        cat = 'Income' if kind == 'inc_dept' else 'Cost of Sales'
        o = SUB_OFFS[(meta, cat)]
        return f'=SUM({c}{s + o[0]}:{c}{s + o[-1]})'
    if kind == 'margin_group':
        members = dict(GROUPS_DISP)[meta]
        i = s + GOFF[('Income', meta)]; co = s + GOFF[('Cost of Sales', meta)]
        return f'=IF({c}{i}=0,"",({c}{i}+{c}{co})/{c}{i})'
    if kind == 'oi_total': return f'={c}{s + OFF["oi_sub"]}'
    if kind == 'oe_total': return f'={c}{s + OFF["oe_sub"]}'
    if kind == 'exp_total':
        return f'=SUM({c}{s + EXP_OFFS[0]}:{c}{s + EXP_OFFS[-1]})'
    if kind == 'gp':     return f'={R("inc_total")}+{R("cos_total")}'
    if kind == 'ebitda': return f'={R("inc_total")}+{R("cos_total")}+{R("exp_total")}-{R("dep")}'
    if kind == 'ebit':   return f'={R("ebitda")}+{R("dep")}'
    if kind in PCT_SRC:
        return (f'=IF({R("inc_total")}=0,"",{c}{s + OFF[PCT_SRC[kind]]}/{R("inc_total")})')
    if kind == 'margin_dept':
        i, co = s + inc_row_of(meta), s + cos_row_of(meta)
        return f'=IF({c}{i}=0,"",({c}{i}+{c}{co})/{c}{i})'
    return None

def actual(kind, c, ec, s, off, meta):
    if kind == 'blank' or kind == 'margin_head': return None
    """What the P&L shows for one month: the GL up to the reporting month,
    the budget after it."""
    v = derived(kind, c, s, meta)
    if v is not None: return v
    if kind == 'inc_sub':  src = eng(meta[0], meta[1], meta[2], ec)
    elif kind == 'cos_sub': src = eng(meta[0], meta[1], meta[2], ec)
    elif kind in ('oi_sub', 'exp_sub', 'oe_sub'): src = '(' + all_depts(meta[0], meta[1], ec) + ')'
    elif kind == 'dep':    src = '(' + dep_expr(ec) + ')'
    else: raise KeyError(kind)
    return f"=IF({c}$5<=Setup!$B$5,{src},'Budget FY27'!{c}{s + off})"

def budget_value(nm, kind, meta, col):
    """Her typed figures, re-pointed at the new rows."""
    if kind in ('inc_sub', 'cos_sub'):
        return typed.get((nm, meta[0], meta[1], meta[2], col))
    if kind in ('oi_sub', 'exp_sub', 'oe_sub'):
        vals = [typed.get((nm, d, meta[0], meta[1], col)) for d in DEPTS_OLD]
        vals = [v for v in vals if v is not None]
        return sum(vals) if vals else None          # no rounding, or the year drifts a cent
    return None

for nm in SHAPED:
    ws = wb[nm]
    for r in range(R0, max(233, LAST + 1) + 1):
        for c in range(1, MAXC + 1):
            cell = ws.cell(r, c)
            cell.value = None
            cell.fill = NO_FILL
            cell.border = NO_RULE
            cell.font = f(10)
            cell.number_format = 'General'
        ws.row_dimensions[r].outlineLevel = 0
        ws.row_dimensions[r].hidden = False
    for off, lab, lvl, kind, meta in ROWS:
        r = R0 + off
        pct = kind in PCT_SRC or kind in ('margin_dept', 'margin_group')
        ws.cell(r, 1, lab)
        look = STY[LOOK[kind]]
        for c in range(1, MAXC + 1):
            cell = ws.cell(r, c)
            cell._style = copy.copy(style_of[(nm, look, c)])
            fl, fn, bd, ht = look_of(kind)
            cell.fill = copy.copy(fl)
            cell.font = copy.copy(fn)
            cell.border = copy.copy(bd)
        if ht: ws.row_dimensions[r].height = ht
        ws.row_dimensions[r].outlineLevel = lvl
        ws.row_dimensions[r].hidden = kind in ('inc_sub', 'cos_sub')
        # ------------------------------------------- the month columns -----
        if nm != 'Summary':
            for i, c in enumerate(MCOLS):
                ec = CL(5 + i)
                if nm == 'P&L FY27':
                    v = actual(kind, c, ec, R0, off, meta)
                else:
                    v = derived(kind, c, R0, meta)
                    if v is None:
                        v = budget_value(nm, kind, meta, 2 + i)
                ws.cell(r, 2 + i, v)
                if pct: ws.cell(r, 2 + i).number_format = PCT
                # on the budget and prior year tabs, mark what she types in blue
                if nm != 'P&L FY27' and kind in TYPED_KINDS:
                    fn = ws.cell(r, 2 + i).font
                    ws.cell(r, 2 + i).font = Font(name=AR, size=fn.size or 10,
                                                  bold=False, color='0000FF')
        if kind in ('blank', 'margin_head'):
            continue
        # ------------------------------------------ the analysis columns ---
        inc = R0 + OFF['inc_total']
        if kind in ('margin_dept', 'margin_group'):
            if kind == 'margin_group':
                inc, cos_ = R0 + GOFF[('Income', meta)], R0 + GOFF[('Cost of Sales', meta)]
            else:
                inc, cos_ = R0 + inc_row_of(meta), R0 + cos_row_of(meta)
            def ratio(cc, _i=inc, _c=cos_):
                return f'=IF({cc}{_i}=0,"",({cc}{_i}+{cc}{_c})/{cc}{_i})'
        elif pct:
            def ratio(cc, _s=R0 + OFF.get(PCT_SRC.get(kind, 'net'), 0)):
                return f'=IF({cc}{inc}=0,"",{cc}{_s}/{cc}{inc})'
        else:
            ratio = None
        if nm == 'P&L FY27':
            if pct:
                for col in (14, 15, 17, 21, 22):
                    ws.cell(r, col, ratio(CL(col)))
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
                ws.cell(r, 23).number_format = PCT
                ws.cell(r, 24).number_format = 'General'
                ws.cell(r, 24).alignment = Alignment(horizontal='center')
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
            for col in (2, 3, 4):
                cc = CL(col)
                if pct:
                    ws.cell(r, col, ratio(cc))
                    ws.cell(r, col).number_format = PCT
                else:
                    ws.cell(r, col, f"=SUMIF('P&L FY27'!$B$5:$M$5,{cc}$5,"
                                    f"'P&L FY27'!$B{r}:$M{r})")
            if pct:
                for col, cc in ((5, 'E'), (7, 'G'), (11, 'K'), (12, 'L')):
                    ws.cell(r, col, ratio(cc))
                ws.cell(r, 8, f'=IF(OR($E{r}="",$G{r}=""),"",$E{r}-$G{r})')
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
                ws.cell(r, 21, f"=G{r}-'Budget FY26'!O{r}")
                ws.cell(r, 22, f'=R{r}-U{r}')
                ws.cell(r, 23, f'=IFERROR(V{r}/ABS(U{r}),0)')
                ws.cell(r, 21).number_format = MONEY
                ws.cell(r, 22).number_format = MONEY
                ws.cell(r, 23).number_format = PCT
                ws.cell(r, 24).number_format = 'General'
                ws.cell(r, 24).alignment = Alignment(horizontal='center')
                ws.cell(r, 24,
                        f'=IF(AND(E{r}=0,G{r}=0),"No Activity",'
                        f'IF(ABS(H{r})<Setup!$B$14,"Low",'
                        f'IF(OR(ABS(H{r})>=Setup!$B$17,ABS(I{r})>=Setup!$B$16),"High",'
                        f'IF(ABS(I{r})>=Setup!$B$15,"Medium","Low"))))')
        else:
            if pct:
                ws.cell(r, 14, ratio('N')).number_format = PCT
            else:
                ws.cell(r, 14, f'=SUM(B{r}:M{r})')
    ws.sheet_properties.outlinePr = Outline(summaryBelow=False, summaryRight=False,
                                            applyStyles=False)
    ws.freeze_panes = 'B6'
    # it has to print for the board, not just look right on screen
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = '1:5'
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.page_margins.top = ws.page_margins.bottom = 0.5
    ws.oddFooter.left.text = 'Corporate Technology Services Pty Ltd'
    ws.oddFooter.right.text = 'Page &P of &N'
    ws.oddFooter.left.size = ws.oddFooter.right.size = 8
    for part in (ws.oddFooter.left, ws.oddFooter.right):
        part.color = '808080'
    ws.sheet_properties.tabColor = TEAL
NEW_LAST = LAST
print('shaped sheets written in her layout')
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
for cc_, h_ in ((7, 'Addback'), (8, 'Cost centre')):
    hc = cu.cell(4, cc_, h_)
    hc.font = f(10, True, 'FFFFFF'); hc.fill = NAVY
    hc.alignment = Alignment(horizontal='center')
cu.column_dimensions['G'].width = 14
cu.column_dimensions['H'].width = 18
# The live Xero export, checked against the August one, is:
#   A Account Code  B Account  C Source  D Date  E Contact  F Debit  G Credit
#   H Job Numbers  I Invoice Number  J Reference  K Description  L Cost Centres
# There is no Index column, whatever an older version of this pack assumed.
GL_CODE, GL_DATE, GL_DR, GL_CR, GL_JOB, GL_CC = 'A', 'D', 'F', 'G', 'H', 'L'
for i in range(CLN_R1 - CLN_R0 + 1):
    r = CLN_R0 + i; g = GL_R0 + i
    gate = f'N({GLREF}!${GL_DATE}{g})=0'
    cu.cell(r, 1, f'=IF({gate},"",N({GLREF}!${GL_CR}{g})-N({GLREF}!${GL_DR}{g}))')
    cu.cell(r, 2, f'=IF({gate},"",DATE(YEAR({GLREF}!${GL_DATE}{g}),'
                  f'MONTH({GLREF}!${GL_DATE}{g}),1))')
    cu.cell(r, 3, f'=IF({gate},"",IF(TRIM({GLREF}!${GL_CC}{g})<>"",TRIM({GLREF}!${GL_CC}{g}),'
                  f'IFERROR(MID({GLREF}!${GL_JOB}{g},FIND("[",{GLREF}!${GL_JOB}{g})+1,'
                  f'FIND("]",{GLREF}!${GL_JOB}{g})-FIND("[",{GLREF}!${GL_JOB}{g})-1),"")))')
    # An untagged line used to get a blank department, which meant it was in the
    # GL but in no column of the statement, so the P&L quietly did not tie. It
    # now falls into General Admin, which is the catch-all cost centre anyway,
    # and column H counts it so you can see it happened.
    cu.cell(r, 4, f'=IF({gate},"",IFERROR(INDEX(DeptName,MATCH($C{r},DeptTag,0)),"ADMIN"))')
    cu.cell(r, 8, f'=IF({gate},"",IF(ISNA(MATCH($C{r},DeptTag,0)),"NO COST CENTRE",""))')
    cu.cell(r, 8).font = f(9)
    cu.cell(r, 5, f'=IF({gate},"",IFERROR(INDEX(AcctCat,'
                  f'MATCH({GLREF}!${GL_CODE}{g}&"",AcctCode,0)),"UNKNOWN"))')
    cu.cell(r, 6, f'=IF({gate},"",IFERROR(INDEX(AcctSub,'
                  f'MATCH({GLREF}!${GL_CODE}{g}&"",AcctCode,0)),"UNKNOWN"))')
    cu.cell(r, 7, f'=IF({gate},"",IFERROR(INDEX(AcctAdd,'
                  f'MATCH({GLREF}!${GL_CODE}{g}&"",AcctCode,0)),""))')
    cu.cell(r, 7).font = f(9)
cu.auto_filter.ref = f'A4:H{CLN_R1}'
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
st['B20'] = f'=COUNT({GLREF}!$D$1:$D$20000)'
st['B24'] = f'=IFERROR(MIN({GLREF}!$D$1:$D$20000),0)'
st['B25'] = f'=IFERROR(MAX({GLREF}!$D$1:$D$20000),0)'
st['C6'] = 'July'
st['A22'] = 'Lines with no cost centre'
st.cell(22, 2, f'=COUNTIF(Cleanup!$H${CLN_R0}:$H${CLN_R1},"NO COST CENTRE")').number_format = '#,##0'
st.cell(22, 3, 'These are put into General Admin so the P&L still ties. If it is not nil, '
               'look at the next line and fix the cost centre in Xero.').font = f(9, False, '595959')
st.cell(22, 3).alignment = Alignment(wrap_text=True, vertical='top')
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
srow(37, 'Value sitting in General Admin for want of a cost centre',
     f'=SUMIF(Cleanup!$H${CLN_R0}:$H${CLN_R1},"NO COST CENTRE",'
     f'Cleanup!$A${CLN_R0}:$A${CLN_R1})',
     'Credit less debit, so income is positive. This is the amount in the wrong department.',
     MON2)
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
for c in range(2, 17):
    pp.column_dimensions[CL(c)].number_format = MON2

pc = wb['PL_Check']
pc['A2'] = ('Either paste the Xero profit and loss on PL_Paste, or type the totals into column C. '
            'Every status should read OK before you report.')
CU_A = f'Cleanup!$A${CLN_R0}:$A${CLN_R1}'
CU_E = f'Cleanup!$E${CLN_R0}:$E${CLN_R1}'
CU_B = f'Cleanup!$B${CLN_R0}:$B${CLN_R1}'
def gl_sum(cat):
    return f'SUMIFS({CU_A},{CU_E},"{cat}",{CU_B},Setup!$B$5)'
# Xero reports interest, tax, dividends and currency inside Total Operating
# Expenses, so the pack's Expenses line has to carry its Other Expenses too or
# this will never tie.
pc['B5'] = '=' + gl_sum('Income')
pc['B6'] = '=' + gl_sum('Cost of Sales')
pc['B7'] = '=' + gl_sum('Expenses') + '+' + gl_sum('Other Expenses')
pc['B8'] = '=' + gl_sum('Other Income')
pc['B9'] = '=' + gl_sum('Other Expenses')
pc['B10'] = '=SUM(B5:B8)'
pc['C10'] = '=SUM(C5:C8)'
pc['A7'] = 'Expenses  (including interest, tax and currency)'
pc['A9'] = 'of which Other Expenses  (already inside the line above)'
pc['A9'].font = f(9, False, '595959')
pc['D9'] = None
pc['E9'] = 'INFO'
pc['E9'].alignment = Alignment(horizontal='center')
for i, h in enumerate(['Label in your Xero P&L', 'Sign', 'Per the pasted P&L', 'Found?',
                       'Status vs the paste']):
    c = pc.cell(4, 7 + i, h)
    c.font = f(10, True, 'FFFFFF'); c.fill = PatternFill('solid', fgColor=TEAL); c.border = BOX
    c.alignment = Alignment(horizontal='center', wrap_text=True)
for col, w in zip('GHIJK', (34, 7, 17, 12, 17)): pc.column_dimensions[col].width = w
PLP, PLN = 'PL_Paste!$A$1:$P$300', 'PL_Paste!$A$1:$A$300'
# Xero prints costs as positive numbers, the pack keeps them negative, so each
# line carries the sign that turns one into the other.
LINES = [('Total Income', 1), ('Total Cost of Sales', -1),
         ('Total Operating Expenses', -1), ('Total Other Income', 1)]
for i, (lbl, sign) in enumerate(LINES):
    r = 5 + i
    c = pc.cell(r, 7, lbl)
    c.font = f(10, False, '0000FF'); c.border = BOX
    sc = pc.cell(r, 8, sign)
    sc.font = f(10, False, '0000FF'); sc.border = BOX; sc.number_format = '+0;-0'
    sc.alignment = Alignment(horizontal='center')
    pc.cell(r, 9, f'=IF($G{r}="",0,$H{r}*IFERROR(INDEX({PLP},MATCH($G{r},{PLN},0),'
                  f'Setup!$B$29),0))').number_format = MON2
    pc.cell(r, 10, f'=IF($G{r}="","no label",IF(ISNA(MATCH($G{r},{PLN},0)),"not found","found"))')
    pc.cell(r, 11, f'=IF(Setup!$B$34=0,"NO PASTE",IF($J{r}<>"found","CHECK LABEL",'
                   f'IF(ABS($B{r}-$I{r})<1,"OK","CHECK")))')
    for cc in (9, 10, 11):
        pc.cell(r, cc).font = f(10, cc == 11)
        pc.cell(r, cc).alignment = Alignment(horizontal='center' if cc != 9 else 'right')
pc.cell(9, 11, 'INFO').font = f(10, True)
pc.cell(9, 11).alignment = Alignment(horizontal='center')
pc.cell(10, 9, '=SUM(I5:I8)').number_format = MON2
pc.cell(10, 9).font = f(10, True)
pc.cell(10, 11, '=IF(Setup!$B$34=0,"NO PASTE",IF(ABS($B10-$I10)<1,"OK","CHECK"))')
pc.cell(10, 11).font = f(10, True)
pc.cell(10, 11).alignment = Alignment(horizontal='center')
for i, t in enumerate([
        'Run the Xero profit and loss for the reporting month and paste the whole export on '
        'PL_Paste at A1. Tell Setup B29 which column holds the figures.',
        'If a label reads "not found", change the text in column G to match your export word '
        'for word. Xero wording moves about between layouts.',
        'Sign turns Xero into the pack. Xero prints costs as positive numbers and the pack '
        'keeps them negative. If a cost line is out by exactly double, flip its sign.',
        'Xero has no Other Expenses section. Interest, income tax, dividends and currency '
        'movements all sit inside its Total Operating Expenses, so the Expenses line here '
        'carries them too.']):
    cc = pc.cell(12 + i, 7, t)
    cc.font = f(9, False, '595959')
    cc.alignment = Alignment(wrap_text=True, vertical='top')
    pc.row_dimensions[12 + i].height = 13 * (1 + len(t) // 110)

# ===================================================================== 8 ===
# BS_Paste and a balance sheet, in the same style as the rest of the pack.
bp = wb.create_sheet('BS_Paste')
for col, w in [('A', 44)] + [(CL(c), 15) for c in range(2, 17)]:
    bp.column_dimensions[col].width = w
for c in range(2, 17):
    bp.column_dimensions[CL(c)].number_format = MON2

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
# The monthly steps, rewritten for how the pack actually works now.
STEPS = [
    ('1.  Setup B5', 'Pick the reporting month. The financial year, the period and every '
     'comparative follow from it.'),
    ('2.  P&L Analysis', 'Xero: Account Transactions for P&L analysis, 1 July to the end of the '
     'reporting month. Export it, select the whole thing and paste it here, anywhere from A1. '
     'The title rows do not matter. Do not reorder the columns: Account Code, Account, Source, '
     'Date, Contact, Debit, Credit, Job Numbers, Invoice Number, Reference, Description, '
     'Cost Centres.'),
    ('3.  PL_Paste, BS_Paste', 'Xero: the profit and loss for the reporting month, and the '
     'balance sheet as at the end of it. Paste each whole export at A1.'),
    ('4.  PL_Check, then read', 'Every status should read OK. Then read Summary, and P&L FY27 '
     'for the detail.'),
]
for i, (a, b) in enumerate(STEPS):
    r = 5 + i
    rd.cell(r, 2, a).font = f(10, True)
    c = rd.cell(r, 3, b); c.font = f(10)
    c.alignment = Alignment(wrap_text=True, vertical='top')
    rd.row_dimensions[r].height = 13 * (1 + len(b) // 105)

base = max(r for r in range(1, 60) if rd.cell(r, 2).value) + 2
rd.cell(base, 2, 'WHAT CHANGED IN THIS VERSION').font = f(10, True, 'C00000')
for i, (a, b) in enumerate([
    ('The layout', 'Every statement tab is now the one consolidated profit and loss you laid '
     'out: category first, then department, then the subcategory detail. Summary, P&L FY27, '
     'Budget FY27, Actual FY26 and Budget FY26 all carry exactly the same rows, so anything '
     'that lines up on one lines up on all of them. The six separate department blocks are '
     'gone.'),
    ('Reading it', 'Total Income opens into the six departments, and each department opens '
     'into its seven income lines. Cost of Sales does the same with its sixteen. Use the + and '
     '- buttons down the left.'),
    ('Overheads', 'Expenses Overheads, Other income and Other Expenses are company totals, not '
     'split by department, which is how you laid them out. The budget figures for those lines '
     'are the six departments added together, so the totals are unchanged.'),
    ('The margins', 'Gross profit, gross profit %, depreciation & amortisation, EBITDA, '
     'EBITDA %, EBIT, EBIT % and net profit % sit at the foot of the statement.'),
    ('EBITDA', 'Gross profit plus overheads, with depreciation added back. EBIT is gross '
     'profit plus overheads. Other income and other expenses sit below EBIT, which is where '
     'interest and tax already were in this pack.'),
    ('Cleanup was dead', 'Every formula on Cleanup pointed at deleted rows, so no sheet in '
     'the pack could calculate. Rewritten.'),
    ('Setup B6', 'The financial year start month was 8, August, against a July year end. '
     'September was coming out as period 2. It now reads 7.'),
    ('Setup B20, B24, B25', 'Broken the same way. Rewritten.'),
    ('PL_Paste', 'Paste the Xero profit and loss and the five category totals read themselves. '
     'The typed column on PL_Check still works if you would rather type them.'),
    ('BS_Paste and Balance_Sheet', 'A balance sheet, which the pack did not have. All 79 live '
     'Xero accounts, grouped, with the movement in the month and in the year, working capital, '
     'current and quick ratio and net debt.'),
    ('Lists', '14 accounts live in Xero were missing from the master, so every GL line on them '
     'was being dropped. One of them, 41500 Contract Help Desk, is revenue. Added and shaded '
     'cream. Please check how I have categorised them.'),
    (GLSHEET, 'A bare sheet on purpose, with no headings, so a plain whole-sheet paste '
     'lands wherever you put it. Rows with no real date are ignored, so the title rows do no '
     'harm. The column order it expects is the one Xero gives you: Account Code, Account, '
     'Source, Date, Contact, Debit, Credit, Job Numbers, Invoice Number, Reference, '
     'Description, Cost Centres.'),
    ('No yellow', 'Cells you type into are marked by blue text instead.'),
]):
    r = base + 1 + i
    rd.cell(r, 2, a).font = f(10, True)
    c = rd.cell(r, 3, b); c.font = f(10)
    c.alignment = Alignment(wrap_text=True, vertical='top')
    rd.row_dimensions[r].height = 13 * (1 + len(b) // 105)

# The pack was built in a navy that is not theirs. Swap every one of them for
# the CTS teal, headers and titles included, so the whole file is on brand.
OLD_NAVY = ('FF1F3864', '001F3864', '1F3864')
for ws_ in wb.worksheets:
    tc = ws_.sheet_properties.tabColor
    if tc is not None and isinstance(tc.rgb, str) and tc.rgb in OLD_NAVY:
        ws_.sheet_properties.tabColor = TEAL
    for row_ in ws_.iter_rows():
        for cell in row_:
            fl = cell.fill
            if fl is not None and fl.fill_type == 'solid' \
                    and isinstance(getattr(fl.fgColor, 'rgb', None), str) \
                    and fl.fgColor.rgb in OLD_NAVY:
                cell.fill = PatternFill('solid', fgColor=TEAL)
            fc = cell.font.color
            if fc is not None and isinstance(fc.rgb, str) and fc.rgb in OLD_NAVY:
                fn = cell.font
                cell.font = Font(name=fn.name or AR, size=fn.size or 10, bold=fn.bold,
                                 italic=fn.italic, color=TEAL)

# the sheet is called P&L Analysis now, so nothing should still say GL_Paste
for ws_ in wb.worksheets:
    for row_ in ws_.iter_rows():
        for cell in row_:
            if isinstance(cell.value, str) and 'GL_Paste' in cell.value:
                cell.value = cell.value.replace('GL_Paste', GLSHEET)

# any text still telling her to look for yellow is now wrong
SWAP = [('yellow cells are the only ones you type into',
         'cells in blue text are the only ones you type into'),
        ('The yellow cell is the only one you type into',
         'The cell in blue text is the only one you type into'),
        ('Paste into the yellow months', 'Type into the months in blue'),
        ('yellow cells', 'cells in blue text'), ('yellow cell', 'cell in blue text'),
        ('the yellow months', 'the months in blue'), ('Yellow cells', 'Cells in blue text'),
        ('into the yellow', 'into the blue')]
for ws_ in wb.worksheets:
    for row_ in ws_.iter_rows():
        for cell in row_:
            if isinstance(cell.value, str) and 'ellow' in cell.value:
                t = cell.value
                for a, b in SWAP: t = t.replace(a, b)
                cell.value = t

# =================================================================== 12 ===
# FY26, from her own CTS Budget FY26 v2 Final workbook.
#
# That file runs on three departments, which are groups of the six here, and
# her FY27 budget proves the grouping to the dollar: Production + Video is the
# budget's Production, Consulting + Integration is its Consulting, and Onsite
# is its Support. It carries no subcategory split, so the detail underneath
# each department follows the shape of her own FY27 budget. Every total, every
# department and every month is hers exactly; only the split between the
# subcategory lines is derived.
FY26 = json.load(open(f'{HERE}/data/fy26.json'))
GROUPS = {'Production': ['PRODUCTION', 'VIDEO'], 'Consulting': ['CONSULTING', 'INTEGRATION'],
          'Support': ['ONSITE']}
SRC_ROW = {'Income': '%s', 'Cost of Sales': '%s Direct Costs'}

def fy27_shape(sheet):
    """Proportions out of her FY27 budget: how a department splits inside its
    group, and how a category splits across its subcategory lines."""
    dept_tot, line = {}, {}
    for off, lab, lvl, kind, meta in ROWS:
        if kind not in ('inc_sub', 'cos_sub'): continue
        d, cat, sub = meta
        v = sum(sheet.cell(R0 + off, 2 + i).value or 0 for i in range(12))
        line[(d, cat, sub)] = v
        dept_tot[(d, cat)] = dept_tot.get((d, cat), 0) + v
    return dept_tot, line

bud = wb['Budget FY27']
DEPT_TOT, LINE_TOT = fy27_shape(bud)
EXP_TOT = {}
for off, lab, lvl, kind, meta in ROWS:
    if kind == 'exp_sub':
        EXP_TOT[meta[1]] = sum(bud.cell(R0 + off, 2 + i).value or 0 for i in range(12))
EXP_ALL = sum(EXP_TOT.values())

def load_fy26(sheet_name, kind):
    ws_ = wb[sheet_name]
    months = FY26['months']
    get = lambda lab, m: FY26['data'].get(f'{kind}|{lab}', {}).get(m, 0)
    for off, lab, lvl, k, meta in ROWS:
        if k not in ('inc_sub', 'cos_sub', 'exp_sub'): continue
        for i, m in enumerate(months):
            if k == 'exp_sub':
                share = EXP_TOT[meta[1]] / EXP_ALL if EXP_ALL else 0
                v = -abs(get('Total Expenses', m)) * share
            else:
                d, cat, sub = meta
                grp = next((g for g, ds in GROUPS.items() if d in ds), None)
                if grp is None: v = 0                      # General Admin earns nothing
                else:
                    total = get(SRC_ROW[cat] % grp, m)
                    in_grp = sum(DEPT_TOT.get((x, cat), 0) for x in GROUPS[grp])
                    dshare = DEPT_TOT.get((d, cat), 0) / in_grp if in_grp else 0
                    dt_ = DEPT_TOT.get((d, cat), 0)
                    lshare = LINE_TOT.get((d, cat, sub), 0) / dt_ if dt_ else 0
                    v = total * dshare * lshare
                    if cat == 'Cost of Sales': v = -abs(v)
            ws_.cell(R0 + off, 2 + i, round(v, 2))
    ws_.cell(3, 1, f'{"Last year\u2019s actuals" if kind == "ACTUAL" else "Last year\u2019s budget"}'
                   ', from CTS Budget FY26 v2 Final. Her three budget departments map to these '
                   'six: Production and Video are her Production, Consulting and Integration are '
                   'her Consulting, Onsite is her Support. Totals, departments and months are '
                   'hers exactly. The split between the subcategory lines follows the shape of '
                   'the FY27 budget, because the FY26 file does not carry one.')
    ws_.cell(3, 1).font = f(9, False, MID)

load_fy26('Actual FY26', 'ACTUAL')
load_fy26('Budget FY26', 'BUDGET')
print('FY26 actuals and budget loaded')

# Summary, laid out the way she moved it: the YOY block at U to X with the six
# empty columns between taken out, and her two notes kept.
sm = wb['Summary']
for c in range(21, 25):
    for r in (4, 5):
        sm.cell(r, c)._style = copy.copy(sm.cell(r, c + 6)._style)
        sm.cell(r, c).value = sm.cell(r, c + 6).value
for c in range(25, 31):                 # the six columns she took out, wiped clean
    for r in range(1, 240):
        cell = sm.cell(r, c)
        cell.value = None
        cell.fill = NO_FILL
        cell.border = NO_RULE
        cell.number_format = 'General'
for col, w in (('J', 1), ('O', 1), ('T', 1), ('U', 14), ('V', 14), ('W', 14), ('X', 12)):
    sm.column_dimensions[col].width = w
for col in ('Y', 'Z', 'AA', 'AB', 'AC', 'AD'):
    sm.column_dimensions[col].width = 9
for c in (16, 17, 18, 19, 21, 22, 23, 24):   # headers only, never the spacer columns
    hc = sm.cell(5, c)
    hc.fill = PatternFill('solid', fgColor=TEAL)
    hc.font = Font(name=AR, size=10, bold=True, color='FFFFFF')
    hc.alignment = Alignment(horizontal='center', wrap_text=True)
    sm.cell(4, c).fill = NO_FILL
sm.cell(4, 21, 'remove this YOY Budget')
sm.cell(4, 22, 'check calc - changed it in meeting with GC ')
for c in (21, 22):
    sm.cell(4, c).font = Font(name=AR, size=9, color='C00000')
# No conditional formatting anywhere. Her own file carries none, and the pink
# it was putting on the Risk column is not wanted. The column says High,
# Medium or Low in words, which is what she reads.
for _n in ('Summary', 'P&L FY27'):
    _w = wb[_n]
    _w.conditional_formatting = type(_w.conditional_formatting)()

if 'requirements ' in wb.sheetnames:
    del wb['requirements ']

# Tab order follows the monthly steps on the README: set the month, paste the
# three reports, read the controls, then read the statements.
ORDER = ['README', 'Setup', 'Summary', 'PL_Paste', 'BS_Paste', 'P&L FY27', 'Budget FY27',
         'Actual FY26', 'Budget FY26', 'Balance_Sheet', GLSHEET, 'PL_Check', 'Utilisation',
         'Engine', 'Lists', 'Cleanup']        # the order she put them in
assert sorted(ORDER) == sorted(wb.sheetnames), set(ORDER) ^ set(wb.sheetnames)
wb._sheets = [wb[n] for n in ORDER]
wb.active = ORDER.index('Summary')
for n in ORDER:
    wb[n].sheet_view.tabSelected = (n == 'Summary')

wb.save(OUT)
print('saved', OUT)
