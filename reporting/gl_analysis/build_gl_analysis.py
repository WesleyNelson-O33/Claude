"""
Build the GL Month-on-Month / Year-on-Year Analysis workbook.

Design rule: the only sheets anyone types into are Setup, GL_Data and
COA_Mapping. Data_Engine aggregates once (account x period) and every
analysis sheet reads the engine, so the whole pack refreshes from one paste.
"""
import json
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as CL
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.comments import Comment
from openpyxl.worksheet.properties import Outline

ROOT = Path("/home/user/Claude/reporting")
OUT = ROOT / "gl_analysis" / "GL Month-on-Month & Year-on-Year Analysis.xlsx"
ACCOUNTS = json.loads((ROOT / "data/accounts.json").read_text())

# Division and group corrections, evidenced from the live Xero P&L and invoice
# line descriptions. accounts.json stays as the raw Xero-derived mapping.
_OV = json.loads((ROOT / "data/account_overrides.json").read_text())
OVERRIDE_NOTE = {}
for _a in ACCOUNTS:
    _n = _a["account"]
    if _n in _OV["division"]:
        _a["division"] = _OV["division"][_n]["to"]
        OVERRIDE_NOTE[_n] = _OV["division"][_n]["why"]
    if _n in _OV["group"]:
        _a["group"] = _OV["group"][_n]["to"]
        OVERRIDE_NOTE[_n] = _OV["group"][_n]["why"]

# ---- capacity -------------------------------------------------------------
GL_R0, GL_R1 = 8, 5007          # GL_Data data rows
COA_R0, COA_R1 = 7, 256         # COA_Mapping data rows (250 accounts)
RAW_R0, RAW_R1 = 1, 12000       # Raw_Paste dump zone starts at A1 (12,000 rows, 20 columns)
RAW_NCOL = 20
CLN_R0, CLN_R1 = 6, 12005       # Cleanup rows, one per dump row
ENG_R0, ENG_R1 = 6, 255         # Data_Engine data rows (aligned 1:1 with COA)
NACC = COA_R1 - COA_R0 + 1

# ---- styling --------------------------------------------------------------
ARIAL = "Arial"
NAVY, MUTED, WHITE = "1F3864", "595959", "FFFFFF"
HEAD = PatternFill("solid", fgColor=NAVY)
SUBHEAD = PatternFill("solid", fgColor="2E5C8A")
BAND = PatternFill("solid", fgColor="D9E2F3")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
TOTAL_FILL = PatternFill("solid", fgColor="EDEDED")
ENGINE_FILL = PatternFill("solid", fgColor="F7F7F7")
GREY_FILL = PatternFill("solid", fgColor="F2F2F2")
EXAMPLE_FILL = PatternFill("solid", fgColor="EAF3EA")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
TOPLINE = Border(top=Side(style="thin", color="1F3864"))

BLUE_FONT = "0000FF"      # hardcoded input
BLACK_FONT = "000000"     # formula on this sheet
GREEN_FONT = "008000"     # link to another sheet

# risk palette
F_HIGH, T_HIGH = "FFC7CE", "9C0006"
F_MED,  T_MED  = "FFEB9C", "9C6500"
F_LOW,  T_LOW  = "C6EFCE", "006100"
F_NA,   T_NA   = "F2F2F2", "808080"
ROW_HIGH = PatternFill("solid", fgColor="FDECEA")
ROW_MED = PatternFill("solid", fgColor="FFF7E0")

MONEY = '$#,##0;($#,##0);"-"'
MONEY2 = '$#,##0.00;($#,##0.00);"-"'
PCT = '0.0%;(0.0%);"-"'
PP = '0.0"pp";(0.0"pp");"-"'
NUM = '#,##0;(#,##0);"-"'
DATEF = 'dd/mm/yyyy'

wb = Workbook()

def sheet(name, tab=None):
    ws = wb.create_sheet(name)
    ws.sheet_view.showGridLines = False
    if tab:
        ws.sheet_properties.tabColor = tab
    return ws

def title(ws, text, sub=None):
    ws["A1"] = text
    ws["A1"].font = Font(name=ARIAL, size=15, bold=True, color=NAVY)
    if sub:
        ws["A2"] = sub
        ws["A2"].font = Font(name=ARIAL, size=9, italic=True, color=MUTED)

def header(ws, row, labels, start_col=1, fill=HEAD, wrap=True, size=9):
    for i, lab in enumerate(labels):
        c = ws.cell(row=row, column=start_col + i, value=lab)
        c.font = Font(name=ARIAL, size=size, bold=True, color=WHITE)
        c.fill = fill
        c.border = BOX
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=wrap)
    ws.row_dimensions[row].height = 32

def band(ws, row, text, first_col, last_col, fill=SUBHEAD):
    c = ws.cell(row=row, column=first_col, value=text)
    c.font = Font(name=ARIAL, size=10, bold=True, color=WHITE)
    for col in range(first_col, last_col + 1):
        ws.cell(row=row, column=col).fill = fill
    ws.row_dimensions[row].height = 18

def widths(ws, spec):
    for col, w in spec.items():
        ws.column_dimensions[col].width = w

def inp(cell, note=None):
    cell.font = Font(name=ARIAL, size=10, color=BLUE_FONT, bold=True)
    cell.fill = INPUT_FILL
    cell.border = BOX
    if note:
        cell.comment = Comment(note, "Financial Controller Pack")

def lbl(ws, row, col, text, bold=False, size=10, color="000000", indent=0):
    c = ws.cell(row=row, column=col, value=text)
    c.font = Font(name=ARIAL, size=size, bold=bold, color=color)
    if indent:
        c.alignment = Alignment(indent=indent)
    return c

# ---- department reporting plan --------------------------------------------
# Overheads are the expense accounts with no division. They get their own
# department rather than being buried inside Admin, so a $2m cost base cannot
# hide behind an operating division. Change any account on COA_Mapping.
DEPTS = ["Onsite", "Production", "Video", "Consulting", "Integration",
         "Admin", "Unallocated"]
CORE_GROUPS = ["Income", "Other Income", "Cost of Sales", "Expenses"]
EXTRA_GROUPS = ["Depreciation & Amortisation", "Finance Costs", "Income Tax Expense"]
# Overheads report under Admin rather than as a department of their own.
# Flip this back to "Overheads" (and add it to DEPTS) to split them out again.
OVERHEAD_DEPT = "Admin"

def _is_overhead(a):
    return a["group"] == "Expenses" and a["division"] == "Unallocated"

def _overhead_flag(a):
    return "Yes" if _is_overhead(a) else "No"

def _rdept(a):
    return OVERHEAD_DEPT if _is_overhead(a) else a["division"]

_dist = {}
for _a in ACCOUNTS:
    _dist[(_rdept(_a), _a["group"])] = _dist.get((_rdept(_a), _a["group"]), 0) + 1

DEPT_PLAN = {}                      # dept -> [(group, slots), ...]
for _d in DEPTS:
    _blocks = []
    for _g in CORE_GROUPS + EXTRA_GROUPS:
        _n = _dist.get((_d, _g), 0)
        if _g in CORE_GROUPS or _n > 0:
            _blocks.append((_g, max(_n + 4, 6)))
    DEPT_PLAN[_d] = _blocks

MONTH_N = 84                        # rolling month list for the Setup dropdown
MONTH_START = (2023, 7)

# Setup cell addresses, named once so the layout can move without hunting refs
S_MONTH, S_MDATE, S_CCY = "Setup!$B$5", "Setup!$B$10", "Setup!$B$19"
# =========================================================================== LISTS  - group master, departments, month picker, block sizes
# ===========================================================================
GROUPS = [
    # order, group, statement, sign, net effect on profit, block
    (10,  "Income",                     "P&L", -1,  1, "Trading"),
    (20,  "Cost of Sales",              "P&L",  1, -1, "Trading"),
    (30,  "Other Income",               "P&L", -1,  1, "Trading"),
    (40,  "Expenses",                   "P&L",  1, -1, "Trading"),
    (50,  "Depreciation & Amortisation","P&L",  1, -1, "Below EBITDA"),
    (60,  "Finance Costs",              "P&L",  1, -1, "Below EBITDA"),
    (70,  "Income Tax Expense",         "P&L",  1, -1, "Tax"),
    (80,  "Current Assets",             "BS",   1,  0, "Assets"),
    (90,  "Non-Current Assets",         "BS",   1,  0, "Assets"),
    (100, "Current Liabilities",        "BS",  -1,  0, "Liabilities"),
    (110, "Non-Current Liabilities",    "BS",  -1,  0, "Liabilities"),
    (120, "Equity",                     "BS",  -1,  0, "Equity"),
]
DIVISIONS = ["Onsite", "Production", "Video", "Consulting", "Integration", "Admin", "Unallocated"]

ls = sheet("Lists", "808080")
title(ls, "Reference Lists",
      "Drives the dropdowns, the sign convention, the report order and the department layout. "
      "Sign flips credit balances so income reads positive. Net effect is what the account does to profit.")
header(ls, 4, ["Order", "Group", "Statement", "Sign", "Net", "Block"])
for i, g in enumerate(GROUPS):
    r = 5 + i
    for j, v in enumerate(g):
        c = ls.cell(row=r, column=1 + j, value=v)
        c.font = Font(name=ARIAL, size=10)
        c.border = BOX
        if j in (0, 3, 4):
            c.alignment = Alignment(horizontal="center")
GRP_R0, GRP_R1 = 5, 4 + len(GROUPS)

header(ls, 19, ["Department (report order)"])
for i, d in enumerate(DEPTS):
    c = ls.cell(row=20 + i, column=1, value=d)
    c.font = Font(name=ARIAL, size=10); c.border = BOX
DEP_R0, DEP_R1 = 20, 19 + len(DEPTS)

header(ls, 31, ["Division (for COA_Mapping)"])
for i, d in enumerate(DIVISIONS):
    c = ls.cell(row=32 + i, column=1, value=d)
    c.font = Font(name=ARIAL, size=10); c.border = BOX
DIV_R0, DIV_R1 = 32, 31 + len(DIVISIONS)

header(ls, 42, ["Risk Flag"])
for i, d in enumerate(["High", "Medium", "Low", "No Activity"]):
    c = ls.cell(row=43 + i, column=1, value=d)
    c.font = Font(name=ARIAL, size=10); c.border = BOX

# month picker - label and the real date side by side, so the dropdown is
# plain text and no locale can mis-parse it
band(ls, 3, "MONTH PICKER", 8, 9)
header(ls, 4, ["Month", "Date"], start_col=8)
_y, _m = MONTH_START
for i in range(MONTH_N):
    import datetime as _dt
    d = _dt.date(_y + (_m - 1 + i) // 12, (_m - 1 + i) % 12 + 1, 1)
    c = ls.cell(row=5 + i, column=8, value=d.strftime("%b %Y"))
    c.font = Font(name=ARIAL, size=10); c.border = BOX
    c2 = ls.cell(row=5 + i, column=9, value=d)
    c2.number_format = DATEF
    c2.font = Font(name=ARIAL, size=10); c2.border = BOX
MTH_R0, MTH_R1 = 5, 4 + MONTH_N

# how many display slots each department/group block has, so a control can
# tell you when an account has nowhere to appear
band(ls, 3, "DEPARTMENT BLOCK SIZES", 11, 14)
header(ls, 4, ["Department", "Group", "Slots"], start_col=11)
_r = 5
LIMIT_R0 = _r
for _d in DEPTS:
    for _g, _slots in DEPT_PLAN[_d]:
        for _j, _v in enumerate((_d, _g, _slots)):
            c = ls.cell(row=_r, column=11 + _j, value=_v)
            c.font = Font(name=ARIAL, size=9); c.border = BOX
        ls.cell(row=_r, column=14, value=f"{_d}|{_g}").font = Font(name=ARIAL, size=8, color="BFBFBF")
        _r += 1
LIMIT_R1 = _r - 1
widths(ls, {"A": 34, "B": 30, "C": 12, "D": 8, "E": 8, "F": 16, "H": 14, "I": 14,
            "K": 16, "L": 28, "M": 8, "N": 30})
# =========================================================================== SETUP
# ===========================================================================
st = sheet("Setup", "1F3864")
title(st, "Setup & Control Panel",
      "Yellow cells are the only ones you type into. Pick the month you are reporting on and everything else follows.")
widths(st, {"A": 46, "B": 22, "C": 62})

def setup_row(r, label, value, fmt=None, note=None, is_input=True):
    lbl(st, r, 1, label, size=10)
    c = st.cell(row=r, column=2, value=value)
    if is_input:
        inp(c, note)
    else:
        c.font = Font(name=ARIAL, size=10, bold=True, color=BLACK_FONT)
        c.border = BOX
        c.fill = TOTAL_FILL
    c.alignment = Alignment(horizontal="center")
    if fmt:
        c.number_format = fmt
    return c

band(st, 3, "1.  WHAT AM I REPORTING ON", 1, 3)
setup_row(4, "Entity name", "Corporate Technology Services Pty Ltd")
setup_row(5, "Reporting month", "Jun 2026",
          note="Pick the month from the dropdown. The financial year and the period number below work themselves "
               "out from it - you no longer set them by hand.")
lbl(st, 5, 3, "<-  THIS is the control you change each month. Everything else follows from it.",
    size=10, color="C00000", bold=True)
setup_row(6, "Financial year starts in month", 7,
          note="7 = July, the Australian standard. You set this once and leave it.")
setup_row(7, "Financial year being reported (year it ENDS)",
          '=IF($B$10="","",IF($B$6=1,YEAR($B$10),YEAR($B$10)+IF(MONTH($B$10)>=$B$6,1,0)))', NUM, is_input=False)
setup_row(8, "Period number within that financial year",
          '=IF($B$10="","",MOD(MONTH($B$10)-$B$6+12,12)+1)', NUM, is_input=False)
setup_row(9, "Prior financial year (the comparative)", '=IF($B$7="","",$B$7-1)', NUM, is_input=False)
setup_row(10, "Reporting month as a date",
          f'=IFERROR(INDEX(Lists!$I${MTH_R0}:$I${MTH_R1},MATCH($B$5,Lists!$H${MTH_R0}:$H${MTH_R1},0)),"")',
          DATEF, is_input=False)
setup_row(11, "Comparative month (same month last year)",
          '=IF($B$10="","",TEXT(DATE(YEAR($B$10)-1,MONTH($B$10),1),"MMM YYYY"))', is_input=False)
setup_row(12, "Data source  ->  which sheet the reports read", "Raw_Paste",
          note="Raw_Paste = dump your Xero export as-is and let the workbook clean it. "
               "GL_Data = you have already tidied the data yourself.")
lbl(st, 12, 3, "Raw_Paste = dirty Xero dump, cleaned automatically.   GL_Data = tidy data you paste yourself.",
    size=9, color=MUTED)
dv_mth = DataValidation(type="list", formula1=f"=Lists!$H${MTH_R0}:$H${MTH_R1}", allow_blank=False,
                        showErrorMessage=True, errorTitle="Pick a month",
                        error="Choose a month from the list.")
st.add_data_validation(dv_mth); dv_mth.add(st["B5"])
dv_src = DataValidation(type="list", formula1='"Raw_Paste,GL_Data"', allow_blank=False, showErrorMessage=True)
st.add_data_validation(dv_src); dv_src.add(st["B12"])
st.conditional_formatting.add("B12", FormulaRule(formula=['$B$12="Raw_Paste"'],
                              fill=PatternFill("solid", fgColor="D9E2F3"), font=Font(color="1F3864", bold=True)))
st.conditional_formatting.add("B10", FormulaRule(formula=['$B$10=""'],
                              fill=PatternFill("solid", fgColor=F_HIGH), font=Font(color=T_HIGH, bold=True)))

band(st, 13, "2.  RISK FLAG THRESHOLDS", 1, 3)
setup_row(14, "Materiality floor  ($) - below this a variance is always Low", 5000, MONEY,
          note="Stops a 400% swing on an $80 account being flagged High. Roughly 0.5% of revenue is a fair start.")
setup_row(15, "Medium risk threshold  (% variance)", 0.10, PCT)
setup_row(16, "High risk threshold  (% variance)", 0.25, PCT)
setup_row(17, "High risk threshold  ($ variance) - overrides the % test", 50000, MONEY)
setup_row(18, "Look-back months for the rolling average", 3, NUM, is_input=False)
lbl(st, 18, 3, "Fixed at 3 in the formulas - shown for reference.", size=9, color=MUTED)
setup_row(19, "Reporting currency", "AUD")

band(st, 20, "3.  DATA STATUS  (calculated - do not type here)", 1, 3)
status = [
    ("GL lines loaded (GL_Data sheet)", f'=COUNT(GL_Data!$A${GL_R0}:$A${GL_R1})', NUM),
    ("GL capacity (rows available)", f'={GL_R1-GL_R0+1}', NUM),
    ("Earliest transaction date", f'=IF($B$21=0,"",MIN(GL_Data!$A${GL_R0}:$A${GL_R1}))', DATEF),
    ("Latest transaction date", f'=IF($B$21=0,"",MAX(GL_Data!$A${GL_R0}:$A${GL_R1}))', DATEF),
    ("Accounts in COA_Mapping", f'=COUNTA(COA_Mapping!$A${COA_R0}:$A${COA_R1})', NUM),
    ("GL lines not mapped to an account", f'=COUNTIF(GL_Data!$T${GL_R0}:$T${GL_R1},"UNMAPPED")', NUM),
    ("GL lines dated outside FY range", f'=COUNTIF(GL_Data!$T${GL_R0}:$T${GL_R1},"OUTSIDE FY RANGE")', NUM),
    ("Total debits", f'=SUM(GL_Data!$G${GL_R0}:$G${GL_R1})', MONEY2),
    ("Total credits", f'=SUM(GL_Data!$H${GL_R0}:$H${GL_R1})', MONEY2),
    ("Debits less credits (must be nil)", '=ROUND($B$28-$B$29,2)', MONEY2),
]
for i, (label, f, fmt) in enumerate(status):
    setup_row(21 + i, label, f, fmt, is_input=False)
st.conditional_formatting.add("B30", FormulaRule(formula=['ROUND($B$30,2)<>0'],
                              fill=PatternFill("solid", fgColor=F_HIGH), font=Font(color=T_HIGH, bold=True)))
st.conditional_formatting.add("B30", FormulaRule(formula=['ROUND($B$30,2)=0'],
                              fill=PatternFill("solid", fgColor=F_LOW), font=Font(color=T_LOW, bold=True)))
st.conditional_formatting.add("B26", FormulaRule(formula=['$B$26>0'],
                              fill=PatternFill("solid", fgColor=F_HIGH), font=Font(color=T_HIGH, bold=True)))

band(st, 32, "3b.  RAW_PASTE CLEANING RESULTS  (only relevant when the source is Raw_Paste)", 1, 3)
raw_status = [
    ("Rows found in the dump zone", f'=SUMPRODUCT(--(Cleanup!$E${CLN_R0}:$E${CLN_R1}<>"Blank"))', NUM),
    ("Transaction lines kept", f'=SUM(Cleanup!$K${CLN_R0}:$K${CLN_R1})', NUM),
    ("Dropped - account heading rows", f'=COUNTIF(Cleanup!$E${CLN_R0}:$E${CLN_R1},"Heading")', NUM),
    ("Dropped - subtotal / balance rows", f'=COUNTIF(Cleanup!$E${CLN_R0}:$E${CLN_R1},"Subtotal")', NUM),
    ("Dropped - lines with no value", f'=COUNTIF(Cleanup!$E${CLN_R0}:$E${CLN_R1},"Nil value")', NUM),
    ("Dropped - text with no date or amount", f'=COUNTIF(Cleanup!$E${CLN_R0}:$E${CLN_R1},"Text only")', NUM),
    ("Kept debits", f'=SUMPRODUCT(Cleanup!$K${CLN_R0}:$K${CLN_R1},Cleanup!$C${CLN_R0}:$C${CLN_R1})', MONEY2),
    ("Kept credits", f'=SUMPRODUCT(Cleanup!$K${CLN_R0}:$K${CLN_R1},Cleanup!$D${CLN_R0}:$D${CLN_R1})', MONEY2),
    ("Kept debits less credits (must be nil)", '=ROUND($B$39-$B$40,2)', MONEY2),
    ("Kept lines not mapped to an account", f'=COUNTIF(Cleanup!$Q${CLN_R0}:$Q${CLN_R1},"UNMAPPED")', NUM),
]
for i, (label, f, fmt) in enumerate(raw_status):
    setup_row(33 + i, label, f, fmt, is_input=False)
for cell, bad in (("B41", 'ROUND($B$41,2)<>0'), ("B42", '$B$42>0')):
    st.conditional_formatting.add(cell, FormulaRule(formula=[bad],
                                  fill=PatternFill("solid", fgColor=F_HIGH), font=Font(color=T_HIGH, bold=True)))
st.conditional_formatting.add("B41", FormulaRule(formula=['ROUND($B$41,2)=0'],
                              fill=PatternFill("solid", fgColor=F_LOW), font=Font(color=T_LOW, bold=True)))

band(st, 45, "4.  RISK FLAG LOGIC  (read once, then trust the colours)", 1, 3)
logic = [
    ("No Activity", "Nil in both the current and comparative period.", F_NA, T_NA),
    ("Low", "Variance is under the materiality floor, or under the Medium % threshold.", F_LOW, T_LOW),
    ("Medium", "Variance is at or over the Medium % threshold and over the materiality floor.", F_MED, T_MED),
    ("High", "Variance is at or over the High % threshold, OR over the High $ threshold, OR the comparative was nil "
             "and the movement is material (a new or ceased account).", F_HIGH, T_HIGH),
]
for i, (flag, desc, fill, font) in enumerate(logic):
    r = 46 + i
    c = st.cell(row=r, column=1, value=flag)
    c.font = Font(name=ARIAL, size=10, bold=True, color=font)
    c.fill = PatternFill("solid", fgColor=fill)
    c.alignment = Alignment(horizontal="center")
    c.border = BOX
    d = st.cell(row=r, column=2, value=desc)
    d.font = Font(name=ARIAL, size=9)
    d.alignment = Alignment(wrap_text=True, vertical="top")
    st.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
    st.row_dimensions[r].height = 28

dv_month = DataValidation(type="whole", operator="between", formula1=1, formula2=12, showErrorMessage=True,
                          errorTitle="Out of range", error="Enter a whole number from 1 to 12.")
st.add_data_validation(dv_month); dv_month.add(st["B6"])
# =========================================================================== COA_MAPPING  - the bridge between GL account names and the report
# ===========================================================================
coa = sheet("COA_Mapping", "C00000")
title(coa, "Chart of Accounts Mapping",
      "Column A must match the account name in your GL export EXACTLY. Group, Division and Overhead drive every "
      "report. Overwrite rows in place - do NOT insert or delete rows, it breaks the link to Data_Engine (Controls "
      "check C5 catches this).")
widths(coa, {"A": 42, "B": 11, "C": 24, "D": 14, "E": 10, "F": 6, "G": 7, "H": 9, "I": 15, "J": 11, "K": 16})
header(coa, 6, ["Account (must match GL)", "Code (optional)", "Group", "Division",
                "Statement", "Sign", "Order", "GL Lines", "Status", "Overhead?", "Reports under"])
coa.freeze_panes = "A7"
lbl(coa, 5, 10, "you set this", size=8, color=MUTED)
lbl(coa, 5, 11, "calculated", size=8, color=MUTED)

GRP_RNG = f"Lists!$B${GRP_R0}:$B${GRP_R1}"
for i in range(NACC):
    r = COA_R0 + i
    src_a = ACCOUNTS[i] if i < len(ACCOUNTS) else None
    a = coa.cell(row=r, column=1, value=(src_a["account"] if src_a else None))
    b = coa.cell(row=r, column=2)
    c = coa.cell(row=r, column=3, value=(src_a["group"] if src_a else None))
    d = coa.cell(row=r, column=4, value=(src_a["division"] if src_a else None))
    ov = coa.cell(row=r, column=10, value=("Yes" if (src_a and _is_overhead(src_a)) else ("No" if src_a else None)))
    for cell in (a, b, c, d, ov):
        cell.font = Font(name=ARIAL, size=9, color=BLUE_FONT)
        cell.fill = INPUT_FILL
        cell.border = BOX
    if src_a and src_a["account"] in OVERRIDE_NOTE:
        a.comment = Comment(f"Reallocated from the Xero default.\n\n{OVERRIDE_NOTE[src_a['account']]}",
                            "Financial Controller Pack")
        for cell in (c, d):
            cell.fill = PatternFill("solid", fgColor="E2EFDA")
    ov.alignment = Alignment(horizontal="center")
    e = coa.cell(row=r, column=5, value=f'=IF($A{r}="","",IFERROR(INDEX(Lists!$C${GRP_R0}:$C${GRP_R1},MATCH($C{r},{GRP_RNG},0)),"?"))')
    f = coa.cell(row=r, column=6, value=f'=IF($A{r}="","",IFERROR(INDEX(Lists!$D${GRP_R0}:$D${GRP_R1},MATCH($C{r},{GRP_RNG},0)),1))')
    g = coa.cell(row=r, column=7, value=f'=IF($A{r}="","",IFERROR(INDEX(Lists!$A${GRP_R0}:$A${GRP_R1},MATCH($C{r},{GRP_RNG},0)),999))')
    h = coa.cell(row=r, column=8, value=f'=IF($A{r}="","",COUNTIF(GL_Data!$B${GL_R0}:$B${GL_R1},$A{r}))')
    istat = coa.cell(row=r, column=9, value=(
        f'=IF($A{r}="","",IF($C{r}="","NO GROUP",IF($E{r}="?","BAD GROUP",'
        f'IF($H{r}=0,"No GL activity","Mapped"))))'))
    k = coa.cell(row=r, column=11, value=f'=IF($A{r}="","",IF($J{r}="Yes","{OVERHEAD_DEPT}",$D{r}))')
    for cell in (e, f, g, h, istat, k):
        cell.font = Font(name=ARIAL, size=9, color=GREEN_FONT)
        cell.border = BOX
        cell.alignment = Alignment(horizontal="center")
    h.number_format = NUM
    istat.alignment = Alignment(horizontal="left")
    k.alignment = Alignment(horizontal="left")

for f1, rng in ((f"={GRP_RNG}", f"C{COA_R0}:C{COA_R1}"),
                (f"=Lists!$A${DIV_R0}:$A${DIV_R1}", f"D{COA_R0}:D{COA_R1}"),
                ('"Yes,No"', f"J{COA_R0}:J{COA_R1}")):
    dv = DataValidation(type="list", formula1=f1, allow_blank=True, showErrorMessage=False)
    coa.add_data_validation(dv); dv.add(rng)

coa.conditional_formatting.add(f"A{COA_R0}:K{COA_R1}", FormulaRule(
    formula=[f'AND($A{COA_R0}<>"",OR($I{COA_R0}="NO GROUP",$I{COA_R0}="BAD GROUP"))'],
    fill=PatternFill("solid", fgColor=F_HIGH), font=Font(color=T_HIGH)))
coa.conditional_formatting.add(f"A{COA_R0}:K{COA_R1}", FormulaRule(
    formula=[f'AND($A{COA_R0}<>"",$I{COA_R0}="No GL activity")'], fill=GREY_FILL, font=Font(color=T_NA)))
coa.conditional_formatting.add(f"J{COA_R0}:K{COA_R1}", FormulaRule(
    formula=[f'$J{COA_R0}="Yes"'], fill=PatternFill("solid", fgColor="E2EFDA"), font=Font(color="375623", bold=True)))
# ===========================================================================
# GL_DATA  - the paste target
# ===========================================================================
gl = sheet("GL_Data", "C00000")
title(gl, "GL Data  -  paste your general ledger export here",
      "Paste into columns A to H starting at row 8. Columns I to V are formulas - never type in them. "
      "Works with transaction-level detail or a month-by-account summary export.")
widths(gl, {"A": 11, "B": 38, "C": 34, "D": 14, "E": 22, "F": 14, "G": 13, "H": 13,
            "I": 13, "J": 7, "K": 7, "L": 10, "M": 20, "N": 9, "O": 6, "P": 13,
            "Q": 7, "R": 7, "S": 11, "T": 17, "U": 30, "V": 7})

for col, txt in ((1, "INPUT  -  type or paste here"), (9, "CALCULATED  -  do not type")):
    last = 8 if col == 1 else 22
    band(gl, 5, txt, col, last, fill=PatternFill("solid", fgColor=("7F6000" if col == 1 else "404040")))

lbl(gl, 6, 1, "Required", size=8, color=MUTED)
lbl(gl, 6, 2, "Required - must match COA_Mapping column A", size=8, color=MUTED)
lbl(gl, 6, 7, "Enter Dr and Cr, or put a signed amount in Debit only", size=8, color=MUTED)

GL_HEAD = ["Date", "Account", "Description", "Journal / Ref", "Contact", "Division (info)",
           "Debit", "Credit",
           "Amount", "FY", "Per", "Period", "Group", "Stmt", "Sign", "Reported Amt",
           "Acct Idx", "FY Off", "Match Key", "Line Status", "Duplicate Key", "Dup #"]
header(gl, 7, GL_HEAD)
gl.freeze_panes = "C8"

B6, B7, B8, B9 = "Setup!$B$6", "Setup!$B$7", "Setup!$B$8", "Setup!$B$9"
COA_A = f"COA_Mapping!$A${COA_R0}:$A${COA_R1}"

for r in range(GL_R0, GL_R1 + 1):
    for col in range(1, 9):
        c = gl.cell(row=r, column=col)
        c.font = Font(name=ARIAL, size=9, color=BLUE_FONT)
        c.border = BOX
        if col == 1:
            c.number_format = DATEF
        if col in (7, 8):
            c.number_format = MONEY2
    fml = {
        9:  f'=IF($A{r}="","",$G{r}-$H{r})',
        10: f'=IF($A{r}="","",IF({B6}=1,YEAR($A{r}),YEAR($A{r})+IF(MONTH($A{r})>={B6},1,0)))',
        11: f'=IF($A{r}="","",MOD(MONTH($A{r})-{B6}+12,12)+1)',
        12: f'=IF($A{r}="","",TEXT($A{r},"MMM-YY"))',
        17: f'=IF($B{r}="",0,IFERROR(MATCH($B{r},{COA_A},0),0))',
        13: f'=IF($A{r}="","",IF($Q{r}=0,"UNMAPPED",INDEX(COA_Mapping!$C${COA_R0}:$C${COA_R1},$Q{r})))',
        14: f'=IF($A{r}="","",IF($Q{r}=0,"",INDEX(COA_Mapping!$E${COA_R0}:$E${COA_R1},$Q{r})))',
        15: f'=IF($A{r}="",1,IF($Q{r}=0,1,INDEX(COA_Mapping!$F${COA_R0}:$F${COA_R1},$Q{r})))',
        16: f'=IF($A{r}="","",$I{r}*$O{r})',
        18: f'=IF($A{r}="",-1,IF($J{r}={B7},1,IF($J{r}={B9},0,-1)))',
        19: f'=IF(OR($A{r}="",$Q{r}=0,$R{r}=-1),0,$Q{r}*1000+$R{r}*100+$K{r})',
        20: (f'=IF($A{r}="","",IF($B{r}="","NO ACCOUNT",IF($Q{r}=0,"UNMAPPED",'
             f'IF($R{r}=-1,"OUTSIDE FY RANGE","OK"))))'),
        21: f'=IF($A{r}="","",$B{r}&"|"&TEXT($A{r},"yyyymmdd")&"|"&TEXT($I{r},"0.00"))',
        22: f'=IF($A{r}="","",COUNTIF($U${GL_R0}:$U${GL_R1},$U{r}))',
    }
    for col in sorted(fml):
        c = gl.cell(row=r, column=col, value=fml[col])
        c.font = Font(name=ARIAL, size=9, color=(BLACK_FONT if col in (9, 10, 11, 12, 16, 19, 21, 22) else GREEN_FONT))
        c.border = BOX
        c.fill = ENGINE_FILL
        if col in (9, 16):
            c.number_format = MONEY2
        elif col in (10, 11, 17, 18, 19, 22):
            c.number_format = NUM
            c.alignment = Alignment(horizontal="center")
        elif col in (12, 14, 15):
            c.alignment = Alignment(horizontal="center")

gl.conditional_formatting.add(f"A{GL_R0}:V{GL_R1}", FormulaRule(
    formula=[f'AND($A{GL_R0}<>"",OR($T{GL_R0}="UNMAPPED",$T{GL_R0}="NO ACCOUNT"))'],
    fill=PatternFill("solid", fgColor=F_HIGH), font=Font(color=T_HIGH)))
gl.conditional_formatting.add(f"A{GL_R0}:V{GL_R1}", FormulaRule(
    formula=[f'AND($A{GL_R0}<>"",$T{GL_R0}="OUTSIDE FY RANGE")'],
    fill=PatternFill("solid", fgColor=F_MED), font=Font(color=T_MED)))
gl.conditional_formatting.add(f"V{GL_R0}:V{GL_R1}", FormulaRule(
    formula=[f'AND($A{GL_R0}<>"",$V{GL_R0}>1)'],
    fill=PatternFill("solid", fgColor=F_MED), font=Font(color=T_MED, bold=True)))

# ===========================================================================
# RAW_PASTE  - dump the Xero export exactly as it comes out
# ===========================================================================
rp = sheet("Raw_Paste", "E36C0A")
# Deliberately bare. Excel refuses to paste a whole-sheet or whole-column selection anywhere
# except A1, so there is nothing above the data - no title, no headings, no panel. Select all,
# copy, paste into A1. The column mapping lives on Setup instead.
rp.sheet_view.showGridLines = True
for i in range(1, RAW_NCOL + 1):
    rp.column_dimensions[CL(i)].width = 18

band(st, 51, "5.  RAW_PASTE COLUMN MAPPING   (count columns from the left of what you pasted: A=1, B=2, C=3 ...)", 1, 3)
lbl(st, 52, 1, "Paste your Xero export into Raw_Paste cell A1, then set these.", bold=True, size=10, color=NAVY)
RP_OPTS = [
    (53, "Layout of your export", "Section headings",
     "Section headings = the account name sits on its own row above its transactions (Xero's General Ledger Detail "
     "and Account Transactions reports). Account in a column = every row carries its own account name."),
    (54, "Column holding the account name", 1, "For section-heading layouts, the column the heading sits in."),
    (55, "Column holding the date", 1, "Rows without a readable date are treated as headings or subtotals."),
    (56, "Column holding the description", 3, "0 if you do not have one."),
    (57, "Column holding the reference / journal no.", 4, "0 if you do not have one."),
    (58, "Column holding the contact / payee", 2, "0 if you do not have one."),
    (59, "Amount layout", "Debit and Credit",
     "Debit and Credit = two separate columns. Single amount = one signed column, positive is a debit."),
    (60, "Column holding Debit (or the single amount)", 5, None),
    (61, "Column holding Credit", 6, "Ignored if you chose Single amount."),
    (62, "Drop lines that have no value", "Yes",
     "Yes strips the empty lines Xero pads the report with. This is the setting that saves you deleting rows."),
    (63, "Strip a leading account code", "Yes",
     "Turns '200 - Sales' into 'Sales'. Only strips a code that is numeric and at the very start, so an account "
     "genuinely named like 'Meal & Travel - CONS' is left alone."),
]
for r, label, val, note in RP_OPTS:
    lbl(st, r, 1, label, size=10)
    c = st.cell(row=r, column=2, value=val)
    inp(c, note)
    c.alignment = Alignment(horizontal="center")
    if note:
        n = st.cell(row=r, column=3, value=note)
        n.font = Font(name=ARIAL, size=8, color=MUTED)
        n.alignment = Alignment(wrap_text=True, vertical="top")
        st.row_dimensions[r].height = 26
for formula1, cells in (('"Section headings,Account in a column"', ["B53"]),
                        ('"Debit and Credit,Single amount"', ["B59"]),
                        ('"Yes,No"', ["B62", "B63"])):
    dv = DataValidation(type="list", formula1=formula1, allow_blank=False, showErrorMessage=True)
    st.add_data_validation(dv)
    for cc in cells:
        dv.add(st[cc])
dv_col = DataValidation(type="whole", operator="between", formula1=0, formula2=RAW_NCOL,
                        showErrorMessage=True, errorTitle="Column number",
                        error=f"Enter a column number from 1 to {RAW_NCOL}, or 0 for none.")
st.add_data_validation(dv_col)
for cc in ("B54", "B55", "B56", "B57", "B58", "B60", "B61"):
    dv_col.add(st[cc])

LAY, ACOL, DCOL = "Setup!$B$53", "Setup!$B$54", "Setup!$B$55"
DESCC, REFC, CONC = "Setup!$B$56", "Setup!$B$57", "Setup!$B$58"
AMTL, DRC, CRC = "Setup!$B$59", "Setup!$B$60", "Setup!$B$61"
DROPNIL, STRIPC = "Setup!$B$62", "Setup!$B$63"

# ===========================================================================
# CLEANUP  - classifies every pasted row and keeps only the real transactions
# ===========================================================================
cu = sheet("Cleanup", "404040")
title(cu, "Cleanup  -  what the workbook made of your dump",
      "Calculated. One row here for every row in the dump zone. Column E says what each row was judged to be and "
      "column K says whether it was kept. Nothing here should ever be typed into.")
CU_HEAD = ["Account text\nas pasted", "Date", "Debit", "Credit", "Row type", "Account\n(code stripped)",
           "Account applied", "Description", "Reference", "Contact", "Keep", "Acct Idx", "FY", "Per",
           "Reported Amt", "Match Key", "Status", "Gather key"]
header(cu, 5, CU_HEAD)
cu.freeze_panes = "A6"
widths(cu, {"A": 30, "B": 11, "C": 12, "D": 12, "E": 13, "F": 28, "G": 28, "H": 30,
            "I": 16, "J": 18, "K": 6, "L": 8, "M": 7, "N": 6, "O": 14, "P": 11, "Q": 16, "R": 10})

for i in range(CLN_R1 - CLN_R0 + 1):
    r = CLN_R0 + i
    rr = RAW_R0 + i
    rng = f"Raw_Paste!$A{rr}:${CL(RAW_NCOL)}{rr}"
    prev = f"$G{r - 1}" if r > CLN_R0 else '""'
    has_amt = f'OR(ROUND($C{r},2)<>0,ROUND($D{r},2)<>0)'
    excl = (f'OR(ISNUMBER(SEARCH("opening balance",$A{r}&" "&$H{r})),'
            f'ISNUMBER(SEARCH("closing balance",$A{r}&" "&$H{r})),'
            f'LEFT(LOWER($A{r}),5)="total")')
    fml = {
        1:  f'=IF({ACOL}=0,"",IFERROR(TRIM(INDEX({rng},{ACOL})&""),""))',
        2:  (f'=IF(ISNUMBER(INDEX({rng},{DCOL})),INDEX({rng},{DCOL}),'
             f'IFERROR(DATEVALUE(INDEX({rng},{DCOL})&""),""))'),
        3:  (f'=IF({AMTL}="Single amount",MAX(0,IFERROR(N(INDEX({rng},{DRC})),0)),'
             f'IFERROR(N(INDEX({rng},{DRC})),0))'),
        4:  (f'=IF({AMTL}="Single amount",MAX(0,-IFERROR(N(INDEX({rng},{DRC})),0)),'
             f'IF({CRC}=0,0,IFERROR(N(INDEX({rng},{CRC})),0)))'),
        5:  (f'=IF({excl},"Subtotal",'
             f'IF(ISNUMBER($B{r}),'
             f'IF({has_amt},"Transaction",IF({DROPNIL}="Yes","Nil value","Transaction")),'
             f'IF({has_amt},"Subtotal",'
             f'IF($A{r}<>"",IF({LAY}="Section headings","Heading","Text only"),"Blank"))))'),
        6:  (f'=IF($A{r}="","",IF({STRIPC}="Yes",'
             f'IF(AND(ISNUMBER(IFERROR(VALUE(LEFT($A{r},1)),"x")),ISNUMBER(SEARCH(" - ",$A{r}))),'
             f'TRIM(MID($A{r},SEARCH(" - ",$A{r})+3,300)),TRIM($A{r})),TRIM($A{r})))'),
        7:  (f'=IF({LAY}="Section headings",IF($E{r}="Heading",$F{r},{prev}),'
             f'IF($E{r}="Transaction",$F{r},{prev}))'),
        8:  f'=IF({DESCC}=0,"",IFERROR(INDEX({rng},{DESCC})&"",""))',
        9:  f'=IF({REFC}=0,"",IFERROR(INDEX({rng},{REFC})&"",""))',
        10: f'=IF({CONC}=0,"",IFERROR(INDEX({rng},{CONC})&"",""))',
        11: f'=IF(AND($E{r}="Transaction",$G{r}<>""),1,0)',
        12: f'=IF($K{r}=0,0,IFERROR(MATCH($G{r},{COA_A},0),0))',
        13: f'=IF($K{r}=0,0,IF({B6}=1,YEAR($B{r}),YEAR($B{r})+IF(MONTH($B{r})>={B6},1,0)))',
        14: f'=IF($K{r}=0,0,MOD(MONTH($B{r})-{B6}+12,12)+1)',
        15: (f'=IF(OR($K{r}=0,$L{r}=0),0,'
             f'($C{r}-$D{r})*INDEX(COA_Mapping!$F${COA_R0}:$F${COA_R1},$L{r}))'),
        16: (f'=IF(OR($K{r}=0,$L{r}=0),0,IF($M{r}={B7},$L{r}*1000+100+$N{r},'
             f'IF($M{r}={B9},$L{r}*1000+$N{r},0)))'),
        17: (f'=IF($K{r}=0,"",IF($L{r}=0,"UNMAPPED",'
             f'IF(AND($M{r}<>{B7},$M{r}<>{B9}),"OUTSIDE FY RANGE","OK")))'),
        18: (f'={f"INT($R{r - 1})" if r > CLN_R0 else "0"}+$K{r}+IF($K{r}=1,0,0.5)'),
        19: f'=IF($K{r}=1,$B{r},0)',
        20: (f'=IF($K{r}=0,"",$G{r}&"|"&TEXT($B{r},"yyyymmdd")&"|"&'
             f'TEXT($C{r}-$D{r},"0.00"))'),
        21: f'=IF($K{r}=0,"",COUNTIF($T${CLN_R0}:$T${CLN_R1},$T{r}))',
    }
    for col, f in fml.items():
        c = cu.cell(row=r, column=col, value=f)
        c.font = Font(name=ARIAL, size=9)
        if col == 2:
            c.number_format = DATEF
        elif col in (3, 4, 15):
            c.number_format = MONEY2
        elif col == 19:
            c.number_format = DATEF
        elif col in (11, 12, 13, 14, 16, 18, 21):
            c.number_format = NUM

CU_KEEP = f"Cleanup!$K${CLN_R0}:$K${CLN_R1}"
cu.conditional_formatting.add(f"A{CLN_R0}:R{CLN_R1}", FormulaRule(
    formula=[f'$E{CLN_R0}="Transaction"'], fill=PatternFill("solid", fgColor="F2F9F2")))
cu.conditional_formatting.add(f"A{CLN_R0}:R{CLN_R1}", FormulaRule(
    formula=[f'AND($E{CLN_R0}="Transaction",$Q{CLN_R0}="UNMAPPED")'],
    fill=PatternFill("solid", fgColor=F_HIGH), font=Font(color=T_HIGH)))
cu.conditional_formatting.add(f"E{CLN_R0}:E{CLN_R1}", FormulaRule(
    formula=[f'OR($E{CLN_R0}="Heading",$E{CLN_R0}="Subtotal")'],
    fill=PatternFill("solid", fgColor=F_MED), font=Font(color=T_MED)))

# preview of the first kept lines, so the cleaning can be eyeballed before it is trusted
PV_N, PV_C0 = 150, 23
band(cu, 3, "PREVIEW  -  the first kept transaction lines, in order", PV_C0, PV_C0 + 6)
header(cu, 5, ["#", "Date", "Account applied", "Description", "Debit", "Credit", "Status"], start_col=PV_C0)
for j in range(PV_N):
    r = CLN_R0 + j
    k = j + 1
    src = cu.cell(row=r, column=PV_C0 + 7,
                  value=f'=IF({k}>Setup!$B$34,"",MATCH({k},Cleanup!$R${CLN_R0}:$R${CLN_R1},1)+{CLN_R0 - 1})')
    src.font = Font(name=ARIAL, size=8, color="BFBFBF")
    for m, (col_letter, fmt) in enumerate([("B", DATEF), ("G", None), ("H", None),
                                           ("C", MONEY2), ("D", MONEY2), ("Q", None)]):
        c = cu.cell(row=r, column=PV_C0 + 1 + m,
                    value=f'=IF(${CL(PV_C0 + 7)}{r}="","",INDEX(Cleanup!${col_letter}:${col_letter},'
                          f'${CL(PV_C0 + 7)}{r}))')
        c.font = Font(name=ARIAL, size=9)
        c.border = BOX
        if fmt:
            c.number_format = fmt
    n = cu.cell(row=r, column=PV_C0, value=f'=IF(${CL(PV_C0 + 7)}{r}="","",{k})')
    n.font = Font(name=ARIAL, size=9)
    n.border = BOX
    n.alignment = Alignment(horizontal="center")
for m, w in enumerate((5, 11, 28, 30, 12, 12, 14, 8)):
    cu.column_dimensions[CL(PV_C0 + m)].width = w

# ===========================================================================
# shared flag formula builders
# ===========================================================================
MAT, MEDP, HIP, HID = "Setup!$B$14", "Setup!$B$15", "Setup!$B$16", "Setup!$B$17"

def flag_formula(r, cur, pri, var):
    """High/Medium/Low/No Activity from a current value, a comparative and the variance."""
    return (f'=IF($A{r}="","",'
            f'IF(AND(ROUND(${cur}{r},2)=0,ROUND(${pri}{r},2)=0),"No Activity",'
            f'IF(ABS(${var}{r})<{MAT},"Low",'
            f'IF(ROUND(${pri}{r},2)=0,"High",'
            f'IF(OR(ABS(${var}{r}/${pri}{r})>={HIP},ABS(${var}{r})>={HID}),"High",'
            f'IF(ABS(${var}{r}/${pri}{r})>={MEDP},"Medium","Low"))))))')

def pct_formula(r, var, base):
    return f'=IF($A{r}="","",IF(ROUND(${base}{r},2)=0,"n/a",${var}{r}/ABS(${base}{r})))'

def risk_cf(ws, rng, flag_col, first_row):
    """Whole-row highlight plus a coloured flag cell."""
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'${flag_col}{first_row}="High"'], fill=ROW_HIGH, font=Font(color=T_HIGH)))
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'${flag_col}{first_row}="Medium"'], fill=ROW_MED, font=Font(color=T_MED)))
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'${flag_col}{first_row}="No Activity"'], fill=GREY_FILL, font=Font(color=T_NA)))

def flag_cell_cf(ws, rng, first_row, col):
    for txt, fill, font in (("High", F_HIGH, T_HIGH), ("Medium", F_MED, T_MED),
                            ("Low", F_LOW, T_LOW), ("No Activity", F_NA, T_NA)):
        ws.conditional_formatting.add(rng, FormulaRule(
            formula=[f'${col}{first_row}="{txt}"'],
            fill=PatternFill("solid", fgColor=fill), font=Font(color=font, bold=True)))

# =========================================================================== DATA_ENGINE  - account x period aggregation, calculated once
# ===========================================================================
en = sheet("Data_Engine", "404040")
title(en, "Data Engine  -  account by period aggregation",
      "Calculated sheet. Every analysis sheet reads from here, so the GL is scanned once instead of "
      "once per report. Nothing on this sheet should ever be typed into.")
PY_C0, CY_C0 = 6, 18
DER = {"CY_MTH": 30, "PY_MTH": 31, "CY_PRI": 32, "AVG3": 33,
       "CY_YTD": 34, "PY_YTD": 35, "PY_FY": 36, "CY_TOT": 37, "ACTIVE": 38,
       "OVH": 39, "RDEPT": 40, "NETSIGN": 41, "SEQ": 42, "SLOTS": 43, "SHOWN": 44,
       "MOM_VAR": 45, "MOM_FLAG": 46, "YOY_VAR": 47, "YOY_FLAG": 48,
       "MOM_KEY": 49, "YOY_KEY": 50}

def month_hdr(fy_ref, p):
    return (f'=TEXT(DATE({fy_ref}-IF(AND({B6}>1,MOD({B6}+{p}-2,12)+1>={B6}),1,0),'
            f'MOD({B6}+{p}-2,12)+1,1),"MMM-YY")')

for p in range(1, 13):
    for c0, fy_ref, key in ((PY_C0, B9, p), (CY_C0, B7, 100 + p)):
        col = c0 + p - 1
        en.cell(row=3, column=col, value=p).font = Font(name=ARIAL, size=8, color=MUTED)
        en.cell(row=4, column=col, value=key).font = Font(name=ARIAL, size=8, color=MUTED)
        for rr in (3, 4):
            en.cell(row=rr, column=col).alignment = Alignment(horizontal="center")
        h = en.cell(row=5, column=col, value=month_hdr(fy_ref, p))
        h.font = Font(name=ARIAL, size=9, bold=True, color=WHITE)
        h.fill = HEAD if c0 == CY_C0 else SUBHEAD
        h.border = BOX
        h.alignment = Alignment(horizontal="center")
lbl(en, 3, 1, "period no", size=8, color=MUTED)
lbl(en, 4, 1, "match key", size=8, color=MUTED)
header(en, 5, ["Account", "Group", "Statement", "Division", "Sign"], size=9)
header(en, 5, ["Current\nMonth", "PY Same\nMonth", "CY Prior\nMonth", "Avg Prior\n3 Months",
               "CY YTD", "PY YTD", "PY Full\nYear", "CY Total", "Active",
               "Overhead", "Reports\nunder", "Net\nsign", "Seq in\nblock", "Slots", "Shown",
               "MoM Var $", "MoM Flag", "YoY YTD\nVar $", "YoY Flag", "MoM key", "YoY key"],
       start_col=DER["CY_MTH"], size=9)
en.freeze_panes = "F6"
widths(en, {"A": 38, "B": 24, "C": 10, "D": 14, "E": 6})
for p in range(24):
    en.column_dimensions[CL(PY_C0 + p)].width = 12
for k in DER.values():
    en.column_dimensions[CL(k)].width = 13

SRC = "Setup!$B$12"
CU_KEY = f"Cleanup!$P${CLN_R0}:$P${CLN_R1}"
CU_AMT = f"Cleanup!$O${CLN_R0}:$O${CLN_R1}"
GL_KEY = f"GL_Data!$S${GL_R0}:$S${GL_R1}"
GL_AMT = f"GL_Data!$P${GL_R0}:$P${GL_R1}"
LIM_KEY = f"Lists!$N${LIMIT_R0}:$N${LIMIT_R1}"
LIM_SLOT = f"Lists!$M${LIMIT_R0}:$M${LIMIT_R1}"
NET_RNG = f"Lists!$E${GRP_R0}:$E${GRP_R1}"
GRP_LIST = f"Lists!$B${GRP_R0}:$B${GRP_R1}"

for i in range(NACC):
    r = ENG_R0 + i
    cr = COA_R0 + i
    idx = i + 1
    meta = {
        1: f'=IF(COA_Mapping!$A{cr}="","",COA_Mapping!$A{cr})',
        2: f'=IF($A{r}="","",COA_Mapping!$C{cr})',
        3: f'=IF($A{r}="","",COA_Mapping!$E{cr})',
        4: f'=IF($A{r}="","",COA_Mapping!$D{cr})',
        5: f'=IF($A{r}="",1,COA_Mapping!$F{cr})',
    }
    for col, f in meta.items():
        c = en.cell(row=r, column=col, value=f)
        c.font = Font(name=ARIAL, size=9, color=GREEN_FONT)
        c.border = BOX
        if col == 5:
            c.alignment = Alignment(horizontal="center")
    for p in range(1, 13):
        for c0, key in ((PY_C0, p), (CY_C0, 100 + p)):
            col = c0 + p - 1
            c = en.cell(row=r, column=col,
                        value=f'=IF($A{r}="",0,IF({SRC}="Raw_Paste",'
                              f'SUMIF({CU_KEY},{idx * 1000 + key},{CU_AMT}),'
                              f'SUMIF({GL_KEY},{idx * 1000 + key},{GL_AMT})))')
            c.font = Font(name=ARIAL, size=9)
            c.number_format = MONEY
            c.border = BOX
            if c0 == PY_C0:
                c.fill = ENGINE_FILL
    ser = f"$F{r}:$AC{r}"
    D = {CL(v): v for v in DER.values()}
    derived = {
        DER["CY_MTH"]: f'=IF($A{r}="",0,INDEX({ser},12+{B8}))',
        DER["PY_MTH"]: f'=IF($A{r}="",0,INDEX({ser},{B8}))',
        DER["CY_PRI"]: f'=IF($A{r}="",0,INDEX({ser},11+{B8}))',
        DER["AVG3"]:   f'=IF($A{r}="",0,AVERAGE(INDEX({ser},9+{B8}):INDEX({ser},11+{B8})))',
        DER["CY_YTD"]: f'=IF($A{r}="",0,SUMPRODUCT(($R$3:$AC$3<={B8})*$R{r}:$AC{r}))',
        DER["PY_YTD"]: f'=IF($A{r}="",0,SUMPRODUCT(($F$3:$Q$3<={B8})*$F{r}:$Q{r}))',
        DER["PY_FY"]:  f'=IF($A{r}="",0,SUM($F{r}:$Q{r}))',
        DER["CY_TOT"]: f'=IF($A{r}="",0,SUM($R{r}:$AC{r}))',
        DER["ACTIVE"]: f'=IF($A{r}="",0,IF(SUMPRODUCT(ABS($F{r}:$AC{r}))>0,1,0))',
        DER["OVH"]:    f'=IF($A{r}="","",COA_Mapping!$J{cr})',
        DER["RDEPT"]:  f'=IF($A{r}="","",COA_Mapping!$K{cr})',
        DER["NETSIGN"]: f'=IF($A{r}="",0,IFERROR(INDEX({NET_RNG},MATCH($B{r},{GRP_LIST},0)),0))',
        DER["SEQ"]:    (f'=IF($A{r}="",0,COUNTIFS(${CL(DER["RDEPT"])}${ENG_R0}:${CL(DER["RDEPT"])}{r},'
                        f'${CL(DER["RDEPT"])}{r},$B${ENG_R0}:$B{r},$B{r}))'),
        DER["SLOTS"]:  (f'=IF($A{r}="",0,IFERROR(INDEX({LIM_SLOT},'
                        f'MATCH(${CL(DER["RDEPT"])}{r}&"|"&$B{r},{LIM_KEY},0)),0))'),
        DER["SHOWN"]:  (f'=IF($A{r}="",0,IF(AND(${CL(DER["SEQ"])}{r}>0,'
                        f'${CL(DER["SEQ"])}{r}<=${CL(DER["SLOTS"])}{r}),1,0))'),
        DER["MOM_VAR"]: f'=IF($A{r}="",0,${CL(DER["CY_MTH"])}{r}-${CL(DER["CY_PRI"])}{r})',
        DER["MOM_FLAG"]: flag_formula(r, CL(DER["CY_MTH"]), CL(DER["CY_PRI"]), CL(DER["MOM_VAR"])),
        DER["YOY_VAR"]: f'=IF($A{r}="",0,${CL(DER["CY_YTD"])}{r}-${CL(DER["PY_YTD"])}{r})',
        DER["YOY_FLAG"]: flag_formula(r, CL(DER["CY_YTD"]), CL(DER["PY_YTD"]), CL(DER["YOY_VAR"])),
        DER["MOM_KEY"]: (f'=IF($A{r}="",-1E+15,IF(ROUND(ABS(${CL(DER["MOM_VAR"])}{r}),0)=0,-1E+15,'
                         f'ABS(${CL(DER["MOM_VAR"])}{r})+ROW()/1000000))'),
        DER["YOY_KEY"]: (f'=IF($A{r}="",-1E+15,IF(ROUND(ABS(${CL(DER["YOY_VAR"])}{r}),0)=0,-1E+15,'
                         f'ABS(${CL(DER["YOY_VAR"])}{r})+ROW()/1000000))'),
    }
    for col, f in derived.items():
        c = en.cell(row=r, column=col, value=f)
        c.font = Font(name=ARIAL, size=9)
        c.border = BOX
        c.fill = TOTAL_FILL
        if col in (DER["ACTIVE"], DER["NETSIGN"], DER["SEQ"], DER["SLOTS"], DER["SHOWN"]):
            c.number_format = NUM
            c.alignment = Alignment(horizontal="center")
        elif col in (DER["MOM_KEY"], DER["YOY_KEY"]):
            c.number_format = NUM
        elif col in (DER["OVH"], DER["RDEPT"], DER["MOM_FLAG"], DER["YOY_FLAG"]):
            c.alignment = Alignment(horizontal="center")
        else:
            c.number_format = MONEY

ENG_RD = f"Data_Engine!${CL(DER['RDEPT'])}${ENG_R0}:${CL(DER['RDEPT'])}${ENG_R1}"
ENG_GRP = f"Data_Engine!$B${ENG_R0}:$B${ENG_R1}"
ENG_NET = f"Data_Engine!${CL(DER['NETSIGN'])}${ENG_R0}:${CL(DER['NETSIGN'])}${ENG_R1}"
ENG_SEQ = f"Data_Engine!${CL(DER['SEQ'])}${ENG_R0}:${CL(DER['SEQ'])}${ENG_R1}"
def ecol(k):
    return f"Data_Engine!${CL(DER[k])}${ENG_R0}:${CL(DER[k])}${ENG_R1}"
def ecol_letter(k):
    return CL(DER[k])
# =========================================================================== MoM_ANALYSIS  - by department, collapsible
# ===========================================================================
# --- shared department-block builder ---------------------------------------
# Level 0 = department (net contribution)   Level 1 = group subtotal
# Level 2 = the accounts.  Collapsed to level 1 on open; click + to drill in.
def dept_blocks(ws, first_row, src_col, name_col=1, grp_col=2, dep_col=3):
    """Lay out the department / group / account skeleton. Returns the row map."""
    rows = []
    r = first_row
    ws.sheet_properties.outlinePr.summaryBelow = False
    ws.sheet_properties.outlinePr.applyStyles = False
    for dept in DEPTS:
        rows.append(("dept", dept, None, None, r))
        c = ws.cell(row=r, column=name_col, value=f"{dept.upper()}   (net contribution)")
        c.font = Font(name=ARIAL, size=11, bold=True, color=WHITE)
        for cc in range(1, src_col + 1):
            ws.cell(row=r, column=cc).fill = SUBHEAD
        ws.cell(row=r, column=dep_col, value=dept).font = Font(name=ARIAL, size=9, color=WHITE)
        ws.row_dimensions[r].height = 17
        r += 1
        for group, slots in DEPT_PLAN[dept]:
            rows.append(("group", dept, group, None, r))
            g = ws.cell(row=r, column=name_col, value=f"    {group}")
            g.font = Font(name=ARIAL, size=10, bold=True)
            for cc in range(1, src_col + 1):
                ws.cell(row=r, column=cc).fill = BAND
            ws.cell(row=r, column=grp_col, value=group).font = Font(name=ARIAL, size=9)
            ws.cell(row=r, column=dep_col, value=dept).font = Font(name=ARIAL, size=9)
            ws.row_dimensions[r].outlineLevel = 1
            ws.row_dimensions[r].collapsed = True
            r += 1
            for k in range(1, slots + 1):
                rows.append(("acct", dept, group, k, r))
                ws.row_dimensions[r].outlineLevel = 2
                ws.row_dimensions[r].hidden = True
                src = ws.cell(row=r, column=src_col,
                              value=f'=SUMPRODUCT(({ENG_RD}="{dept}")*({ENG_GRP}="{group}")'
                                    f'*({ENG_SEQ}={k})*ROW({ENG_RD}))')
                src.font = Font(name=ARIAL, size=8, color="BFBFBF")
                n = ws.cell(row=r, column=name_col,
                            value=f'=IF(${CL(src_col)}{r}=0,"","        "&INDEX(Data_Engine!$A:$A,${CL(src_col)}{r}))')
                n.font = Font(name=ARIAL, size=9, color=GREEN_FONT)
                ws.cell(row=r, column=grp_col,
                        value=f'=IF(${CL(src_col)}{r}=0,"",INDEX(Data_Engine!$B:$B,${CL(src_col)}{r}))'
                        ).font = Font(name=ARIAL, size=9, color=GREEN_FONT)
                ws.cell(row=r, column=dep_col,
                        value=f'=IF(${CL(src_col)}{r}=0,"",INDEX(Data_Engine!${CL(DER["RDEPT"])}:'
                              f'${CL(DER["RDEPT"])},${CL(src_col)}{r}))'
                        ).font = Font(name=ARIAL, size=9, color=GREEN_FONT)
                r += 1
        ws.row_dimensions[r].height = 5
        r += 1
    return rows, r

def eng_range(col_letter):
    return f"Data_Engine!${col_letter}${ENG_R0}:${col_letter}${ENG_R1}"

def block_value(kind, dept, group, slot, eng_letter, r, src_col):
    """The same figure, worked out three ways depending on the row type."""
    if kind == "acct":
        return (f'=IF(${CL(src_col)}{r}=0,"",INDEX(Data_Engine!${eng_letter}:${eng_letter},'
                f'${CL(src_col)}{r}))')
    if kind == "group":
        return f'=SUMIFS({eng_range(eng_letter)},{ENG_RD},"{dept}",{ENG_GRP},"{group}")'
    return f'=SUMPRODUCT(({ENG_RD}="{dept}")*{ENG_NET}*{eng_range(eng_letter)})'

mm = sheet("MoM_Analysis", "2E75B6")
title(mm, "Month-on-Month Analysis  -  by department")
mm["A2"] = ('=IF(Setup!$B$4="","","Month on month for "&Setup!$B$5&".   |   Click the + and - buttons on the left to '
            'open a department up.   |   Department rows are net contribution: income less costs. Group and account '
            'rows are gross, income and costs both positive.")')
mm["A2"].font = Font(name=ARIAL, size=9, italic=True, color=MUTED)
MM_SRC = 26
band(mm, 4, "LINE", 1, 3)
band(mm, 4, "MONTHLY ACTUALS  -  CURRENT FINANCIAL YEAR", 4, 15, fill=PatternFill("solid", fgColor="2E75B6"))
band(mm, 4, "MONTH-ON-MONTH TEST", 16, 23, fill=PatternFill("solid", fgColor="7F6000"))
band(mm, 4, "REVIEW", 24, 25, fill=PatternFill("solid", fgColor="404040"))
header(mm, 5, ["Line item", "Group", "Reports under"] + [f"M{p}" for p in range(1, 13)] +
       ["YTD", "Current\nMonth", "Prior\nMonth", "Movement\n$", "Movement\n%",
        "Avg Prior\n3 Months", "Var vs\nAvg $", "Var vs\nAvg %", "Risk\nFlag", "Controller Comment"])
for p in range(1, 13):
    mm.cell(row=5, column=3 + p, value=month_hdr(B7, p))
mm.freeze_panes = "D6"
widths(mm, {"A": 44, "B": 22, "C": 14, "P": 13, "Q": 13, "R": 13, "S": 13, "T": 11,
            "U": 13, "V": 13, "W": 11, "X": 11, "Y": 34, "Z": 9})
for p in range(4, 16):
    mm.column_dimensions[CL(p)].width = 11

mm_rows, MM_END = dept_blocks(mm, 6, MM_SRC)
MM_R0, MM_R1 = 6, MM_END - 1
for kind, dept, group, slot, r in mm_rows:
    bold = kind in ("dept", "group")
    white = kind == "dept"
    vals = {}
    for p in range(1, 13):
        vals[3 + p] = block_value(kind, dept, group, slot, CL(CY_C0 + p - 1), r, MM_SRC)
    vals[16] = block_value(kind, dept, group, slot, CL(DER["CY_YTD"]), r, MM_SRC)
    vals[17] = block_value(kind, dept, group, slot, CL(DER["CY_MTH"]), r, MM_SRC)
    vals[18] = block_value(kind, dept, group, slot, CL(DER["CY_PRI"]), r, MM_SRC)
    vals[21] = block_value(kind, dept, group, slot, CL(DER["AVG3"]), r, MM_SRC)
    blank = f'$A{r}=""'
    vals[19] = f'=IF({blank},"",$Q{r}-$R{r})'
    vals[20] = f'=IF({blank},"",IF(ROUND($R{r},2)=0,"n/a",$S{r}/ABS($R{r})))'
    vals[22] = f'=IF({blank},"",$Q{r}-$U{r})'
    vals[23] = f'=IF({blank},"",IF(ROUND($U{r},2)=0,"n/a",$V{r}/ABS($U{r})))'
    vals[24] = flag_formula(r, "Q", "R", "S")
    for col, f in vals.items():
        c = mm.cell(row=r, column=col, value=f)
        c.font = Font(name=ARIAL, size=9, bold=bold, color=(WHITE if white else BLACK_FONT))
        c.number_format = PCT if col in (20, 23) else MONEY
        c.border = BOX
        if col == 24:
            c.number_format = "General"
            c.alignment = Alignment(horizontal="center")
    cm = mm.cell(row=r, column=25)
    cm.font = Font(name=ARIAL, size=9, color=BLUE_FONT)
    cm.fill = INPUT_FILL
    cm.border = BOX
    cm.alignment = Alignment(wrap_text=True, vertical="top")

risk_cf(mm, f"A{MM_R0}:Y{MM_R1}", "X", MM_R0)
flag_cell_cf(mm, f"X{MM_R0}:X{MM_R1}", MM_R0, "X")
lbl(mm, 4, MM_SRC, "helper", size=8, color="BFBFBF")
# =========================================================================== YoY_ANALYSIS  - by department, collapsible
# ===========================================================================
yy = sheet("YoY_Analysis", "548235")
title(yy, "Year-on-Year Analysis  -  by department")
yy["A2"] = ('=IF(Setup!$B$4="","",Setup!$B$5&" vs "&Setup!$B$11&", and year to date against last year to date.   |   '
            'Click the + and - buttons on the left to open a department up.")')
yy["A2"].font = Font(name=ARIAL, size=9, italic=True, color=MUTED)
YY_SRC = 19
band(yy, 4, "LINE", 1, 3)
band(yy, 4, "CURRENT MONTH vs PRIOR YEAR SAME MONTH", 4, 7, fill=PatternFill("solid", fgColor="548235"))
band(yy, 4, "YEAR TO DATE vs PRIOR YEAR TO DATE", 8, 11, fill=PatternFill("solid", fgColor="2E75B6"))
band(yy, 4, "FULL YEAR RUN-RATE", 12, 14, fill=PatternFill("solid", fgColor="7F6000"))
band(yy, 4, "REVIEW", 15, 17, fill=PatternFill("solid", fgColor="404040"))
header(yy, 5, ["Line item", "Group", "Reports under",
               "CY Month", "PY Same\nMonth", "Var $", "Var %",
               "CY YTD", "PY YTD", "YTD Var $", "YTD Var %",
               "PY Full\nYear", "CY Run-Rate\n(annualised)", "Run-Rate vs\nPY FY %",
               "Movement\nType", "Risk\nFlag", "Controller Comment"])
yy.freeze_panes = "D6"
widths(yy, {"A": 44, "B": 22, "C": 14, "D": 13, "E": 13, "F": 13, "G": 11,
            "H": 14, "I": 14, "J": 14, "K": 11, "L": 14, "M": 14, "N": 12,
            "O": 13, "P": 11, "Q": 34, "R": 9})

yy_rows, YY_END = dept_blocks(yy, 6, YY_SRC)
YY_R0, YY_R1 = 6, YY_END - 1
for kind, dept, group, slot, r in yy_rows:
    bold = kind in ("dept", "group")
    white = kind == "dept"
    blank = f'$A{r}=""'
    vals = {
        4: block_value(kind, dept, group, slot, CL(DER["CY_MTH"]), r, YY_SRC),
        5: block_value(kind, dept, group, slot, CL(DER["PY_MTH"]), r, YY_SRC),
        8: block_value(kind, dept, group, slot, CL(DER["CY_YTD"]), r, YY_SRC),
        9: block_value(kind, dept, group, slot, CL(DER["PY_YTD"]), r, YY_SRC),
        12: block_value(kind, dept, group, slot, CL(DER["PY_FY"]), r, YY_SRC),
        6: f'=IF({blank},"",$D{r}-$E{r})',
        7: f'=IF({blank},"",IF(ROUND($E{r},2)=0,"n/a",$F{r}/ABS($E{r})))',
        10: f'=IF({blank},"",$H{r}-$I{r})',
        11: f'=IF({blank},"",IF(ROUND($I{r},2)=0,"n/a",$J{r}/ABS($I{r})))',
        13: f'=IF({blank},"",IF({B8}=0,0,$H{r}/{B8}*12))',
        14: f'=IF({blank},"",IF(ROUND($L{r},2)=0,"n/a",($M{r}-$L{r})/ABS($L{r})))',
        15: (f'=IF({blank},"",IF(AND(ROUND($H{r},2)=0,ROUND($I{r},2)=0),"Dormant",'
             f'IF(ROUND($I{r},2)=0,"New this year",IF(ROUND($H{r},2)=0,"Ceased",'
             f'IF($J{r}>0,"Increase",IF($J{r}<0,"Decrease","Flat"))))))'),
        16: flag_formula(r, "H", "I", "J"),
    }
    for col, f in vals.items():
        c = yy.cell(row=r, column=col, value=f)
        c.font = Font(name=ARIAL, size=9, bold=bold, color=(WHITE if white else BLACK_FONT))
        c.number_format = PCT if col in (7, 11, 14) else MONEY
        c.border = BOX
        if col in (15, 16):
            c.number_format = "General"
            c.alignment = Alignment(horizontal="center")
    cm = yy.cell(row=r, column=17)
    cm.font = Font(name=ARIAL, size=9, color=BLUE_FONT)
    cm.fill = INPUT_FILL
    cm.border = BOX
    cm.alignment = Alignment(wrap_text=True, vertical="top")

risk_cf(yy, f"A{YY_R0}:Q{YY_R1}", "P", YY_R0)
flag_cell_cf(yy, f"P{YY_R0}:P{YY_R1}", YY_R0, "P")
lbl(yy, 4, YY_SRC, "helper", size=8, color="BFBFBF")
# =========================================================================== SUMMARY  - company first, then each department, collapsible
# ===========================================================================
sm = sheet("Summary", "C55A11")
title(sm, "Financial Controller Summary")
sm["A2"] = ('=IF(Setup!$B$4="","",Setup!$B$4&"   |   "&Setup!$B$5&"   |   FY"&Setup!$B$7&" vs FY"&Setup!$B$9'
            '&"   |   all figures "&Setup!$B$19)')
sm["A2"].font = Font(name=ARIAL, size=10, bold=True, color=MUTED)
sm.sheet_properties.outlinePr.summaryBelow = False
sm.sheet_properties.outlinePr.applyStyles = False

SM_R0 = 6
BASE = {"B": DER["CY_MTH"], "C": DER["PY_MTH"], "F": DER["CY_PRI"],
        "I": DER["CY_YTD"], "J": DER["PY_YTD"], "M": DER["PY_FY"]}
band(sm, 4, "", 1, 17, fill=PatternFill("solid", fgColor="FFFFFF"))
band(sm, 4, "CURRENT MONTH vs PRIOR YEAR", 2, 5, fill=PatternFill("solid", fgColor="548235"))
band(sm, 4, "CURRENT MONTH vs PRIOR MONTH", 6, 8, fill=PatternFill("solid", fgColor="7F6000"))
band(sm, 4, "YEAR TO DATE vs PRIOR YEAR TO DATE", 9, 12, fill=PatternFill("solid", fgColor="2E75B6"))
band(sm, 4, "PY FY", 13, 13, fill=PatternFill("solid", fgColor="808080"))
band(sm, 4, "RISK FLAGS & ACTION", 14, 17, fill=PatternFill("solid", fgColor="404040"))
header(sm, 5, ["Line Item", "CY Month", "PY Same\nMonth", "YoY\nVar $", "YoY\nVar %",
               "CY Prior\nMonth", "MoM\nVar $", "MoM\nVar %",
               "CY YTD", "PY YTD", "YTD\nVar $", "YTD\nVar %", "PY Full\nYear",
               "YoY\nFlag", "MoM\nFlag", "Overall\nRisk", "Review Action"])
sm.freeze_panes = "B6"
widths(sm, {"A": 42, "N": 10, "O": 10, "P": 10, "Q": 30})
for c in "BCDEFGHIJKLM":
    sm.column_dimensions[c].width = 14

def sm_derived(r, pct_row=False):
    if pct_row:
        d = {"D": f'=IF(OR(ISTEXT($B{r}),ISTEXT($C{r})),"n/a",($B{r}-$C{r})*100)', "E": '=""',
             "G": f'=IF(OR(ISTEXT($B{r}),ISTEXT($F{r})),"n/a",($B{r}-$F{r})*100)', "H": '=""',
             "K": f'=IF(OR(ISTEXT($I{r}),ISTEXT($J{r})),"n/a",($I{r}-$J{r})*100)', "L": '=""',
             "N": f'=IF(ISTEXT($K{r}),"No Activity",IF(ABS($K{r})>=5,"High",IF(ABS($K{r})>=2,"Medium","Low")))',
             "O": f'=IF(ISTEXT($G{r}),"No Activity",IF(ABS($G{r})>=5,"High",IF(ABS($G{r})>=2,"Medium","Low")))'}
        fmts = {"D": PP, "G": PP, "K": PP}
    else:
        d = {"D": f'=$B{r}-$C{r}', "E": f'=IF(ROUND($C{r},2)=0,"n/a",$D{r}/ABS($C{r}))',
             "G": f'=$B{r}-$F{r}', "H": f'=IF(ROUND($F{r},2)=0,"n/a",$G{r}/ABS($F{r}))',
             "K": f'=$I{r}-$J{r}', "L": f'=IF(ROUND($J{r},2)=0,"n/a",$K{r}/ABS($J{r}))',
             "N": (f'=IF(AND(ROUND($I{r},2)=0,ROUND($J{r},2)=0),"No Activity",IF(ABS($K{r})<{MAT},"Low",'
                   f'IF(ROUND($J{r},2)=0,"High",IF(OR(ABS($K{r}/$J{r})>={HIP},ABS($K{r})>={HID}),"High",'
                   f'IF(ABS($K{r}/$J{r})>={MEDP},"Medium","Low")))))'),
             "O": (f'=IF(AND(ROUND($B{r},2)=0,ROUND($F{r},2)=0),"No Activity",IF(ABS($G{r})<{MAT},"Low",'
                   f'IF(ROUND($F{r},2)=0,"High",IF(OR(ABS($G{r}/$F{r})>={HIP},ABS($G{r})>={HID}),"High",'
                   f'IF(ABS($G{r}/$F{r})>={MEDP},"Medium","Low")))))')}
        fmts = {"D": MONEY, "G": MONEY, "K": MONEY}
    d["P"] = (f'=IF(OR($N{r}="High",$O{r}="High"),"High",IF(OR($N{r}="Medium",$O{r}="Medium"),"Medium",'
              f'IF(AND($N{r}="No Activity",$O{r}="No Activity"),"No Activity","Low")))')
    return d, fmts

def sm_write(r, label, base_f, bold=False, pct_row=False, level=None, white=False, indent=0):
    c = lbl(sm, r, 1, label, bold=bold, size=(11 if white else 10), indent=indent)
    if white:
        c.font = Font(name=ARIAL, size=11, bold=True, color=WHITE)
    for col, f in base_f.items():
        cc = sm.cell(row=r, column=ord(col) - 64, value=f)
        cc.font = Font(name=ARIAL, size=10, bold=bold, color=(WHITE if white else BLACK_FONT))
        cc.number_format = PCT if pct_row else MONEY
        cc.border = BOX
    d, fmts = sm_derived(r, pct_row)
    for col, f in d.items():
        cc = sm.cell(row=r, column=ord(col) - 64, value=f)
        cc.font = Font(name=ARIAL, size=10, bold=(bold or col == "P"), color=(WHITE if white else BLACK_FONT))
        nf = fmts.get(col, PCT if col in ("E", "H", "L") else None)
        if nf:
            cc.number_format = nf
        cc.border = BOX
        if col in ("N", "O", "P"):
            cc.alignment = Alignment(horizontal="center")
    act = sm.cell(row=r, column=17)
    act.font = Font(name=ARIAL, size=9, color=BLUE_FONT)
    act.fill = INPUT_FILL
    act.border = BOX
    act.alignment = Alignment(wrap_text=True, vertical="top")
    if white:
        for col in range(1, 17):
            sm.cell(row=r, column=col).fill = SUBHEAD
    elif bold:
        for col in range(1, 17):
            sm.cell(row=r, column=col).fill = TOTAL_FILL
    if level:
        sm.row_dimensions[r].outlineLevel = level
    return r + 1

def grp_base(group):
    return {col: f'=SUMIF({ENG_GRP},"{group}",{eng_range(CL(ec))})' for col, ec in BASE.items()}

def deptgrp_base(dept, group):
    return {col: f'=SUMIFS({eng_range(CL(ec))},{ENG_RD},"{dept}",{ENG_GRP},"{group}")'
            for col, ec in BASE.items()}

def deptnet_base(dept):
    return {col: f'=SUMPRODUCT(({ENG_RD}="{dept}")*{ENG_NET}*{eng_range(CL(ec))})'
            for col, ec in BASE.items()}

def calc_base(expr, rowof):
    out = {}
    for col in BASE:
        e = expr
        for name, rr in rowof.items():
            e = e.replace("{" + name + "}", f"{col}{rr}")
        out[col] = "=" + e
    return out

def pct_base(num_r, den_r):
    return {col: f'=IF(ROUND({col}{den_r},2)=0,"n/a",{col}{num_r}/{col}{den_r})' for col in BASE}

r = SM_R0
band(sm, r, "COMPANY  -  PROFIT & LOSS", 1, 17); r += 1
co = {}
for label, group in (("Income", "Income"), ("Cost of Sales", "Cost of Sales")):
    co[label] = r; r = sm_write(r, label, grp_base(group), indent=1)
co["Gross Profit"] = r; r = sm_write(r, "Gross Profit", calc_base("{Income}-{Cost of Sales}", co), bold=True)
r = sm_write(r, "Gross Margin %", pct_base(co["Gross Profit"], co["Income"]), pct_row=True, indent=1)
for label, group in (("Other Income", "Other Income"), ("Operating Expenses", "Expenses")):
    co[label] = r; r = sm_write(r, label, grp_base(group), indent=1)
co["EBITDA"] = r
r = sm_write(r, "EBITDA", calc_base("{Gross Profit}+{Other Income}-{Operating Expenses}", co), bold=True)
r = sm_write(r, "EBITDA Margin %", pct_base(co["EBITDA"], co["Income"]), pct_row=True, indent=1)
for label, group in (("Depreciation & Amortisation", "Depreciation & Amortisation"),
                     ("Finance Costs", "Finance Costs")):
    co[label] = r; r = sm_write(r, label, grp_base(group), indent=1)
co["Net Profit Before Tax"] = r
r = sm_write(r, "Net Profit Before Tax",
             calc_base("{EBITDA}-{Depreciation & Amortisation}-{Finance Costs}", co), bold=True)
co["Income Tax Expense"] = r; r = sm_write(r, "Income Tax Expense", grp_base("Income Tax Expense"), indent=1)
co["Net Profit After Tax"] = r
r = sm_write(r, "Net Profit After Tax", calc_base("{Net Profit Before Tax}-{Income Tax Expense}", co), bold=True)
r = sm_write(r, "Net Margin %", pct_base(co["Net Profit After Tax"], co["Income"]), pct_row=True, indent=1)

sm.row_dimensions[r].height = 6; r += 1
band(sm, r, "BY DEPARTMENT  -  click the + and - buttons on the left to open one up", 1, 17); r += 1
DEPT_ROWS = {}
for dept in DEPTS:
    dr = r
    DEPT_ROWS[dept] = dr
    r = sm_write(r, f"{dept.upper()}   (net contribution)", deptnet_base(dept), bold=True, white=True)
    d = {}
    for label, group in (("Income", "Income"), ("Cost of Sales", "Cost of Sales")):
        d[label] = r; r = sm_write(r, label, deptgrp_base(dept, group), level=1, indent=2)
    d["Gross Profit"] = r
    r = sm_write(r, "Gross Profit", calc_base("{Income}-{Cost of Sales}", d), bold=True, level=1, indent=1)
    r = sm_write(r, "Gross Margin %", pct_base(d["Gross Profit"], d["Income"]), pct_row=True, level=1, indent=2)
    for label, group in (("Other Income", "Other Income"), ("Direct Expenses", "Expenses")):
        d[label] = r; r = sm_write(r, label, deptgrp_base(dept, group), level=1, indent=2)
    d["dept"] = dr
    r = sm_write(r, "Other (depreciation, finance, tax)",
                 calc_base("{dept}-({Gross Profit}+{Other Income}-{Direct Expenses})", d), level=1, indent=2)
    sm.row_dimensions[r].height = 5
    sm.row_dimensions[r].outlineLevel = 1
    r += 1

sm.row_dimensions[r].height = 6; r += 1
band(sm, r, "BALANCE SHEET  -  MOVEMENT IN THE PERIOD", 1, 17); r += 1
bs = {}
for label in ("Current Assets", "Non-Current Assets"):
    bs[label] = r; r = sm_write(r, label, grp_base(label), indent=1)
bs["Total Assets"] = r
r = sm_write(r, "Total Assets", calc_base("{Current Assets}+{Non-Current Assets}", bs), bold=True)
for label in ("Current Liabilities", "Non-Current Liabilities"):
    bs[label] = r; r = sm_write(r, label, grp_base(label), indent=1)
bs["Total Liabilities"] = r
r = sm_write(r, "Total Liabilities", calc_base("{Current Liabilities}+{Non-Current Liabilities}", bs), bold=True)
r = sm_write(r, "Net Assets", calc_base("{Total Assets}-{Total Liabilities}", bs), bold=True)
r = sm_write(r, "Equity", grp_base("Equity"), indent=1)

SM_LAST = r - 1
risk_cf(sm, f"A{SM_R0}:Q{SM_LAST}", "P", SM_R0)
for col in ("N", "O", "P"):
    flag_cell_cf(sm, f"{col}{SM_R0}:{col}{SM_LAST}", SM_R0, col)

r += 1
band(sm, r, "CONTROLLER REVIEW STATUS", 1, 17); r += 1
REVIEW = [
    ("High risk accounts - month on month", f'=COUNTIF({ecol("MOM_FLAG")},"High")', "review"),
    ("Medium risk accounts - month on month", f'=COUNTIF({ecol("MOM_FLAG")},"Medium")', "info"),
    ("High risk accounts - year on year", f'=COUNTIF({ecol("YOY_FLAG")},"High")', "review"),
    ("Medium risk accounts - year on year", f'=COUNTIF({ecol("YOY_FLAG")},"Medium")', "info"),
    ("High risk lines still without a comment",
     f'=SUMPRODUCT((MoM_Analysis!$X${MM_R0}:$X${MM_R1}="High")*(MoM_Analysis!$Y${MM_R0}:$Y${MM_R1}=""))'
     f'+SUMPRODUCT((YoY_Analysis!$P${YY_R0}:$P${YY_R1}="High")*(YoY_Analysis!$Q${YY_R0}:$Q${YY_R1}=""))', "must"),
    ("Accounts with nowhere to display", f'=SUMPRODUCT(({ecol("ACTIVE")}=1)*({ecol("SHOWN")}=0))', "must"),
    ("GL lines not mapped to an account", '=IF(Setup!$B$12="Raw_Paste",Setup!$B$42,Setup!$B$26)', "must"),
    ("Source debits less credits (must be nil)", '=IF(Setup!$B$12="Raw_Paste",Setup!$B$41,Setup!$B$30)', "must"),
    ("Controls failing on the Controls sheet", '=COUNTIF(Controls!$E$6:$E$40,"FAIL")', "must"),
]
REV_R0 = r
for label, f, sev in REVIEW:
    lbl(sm, r, 1, label, size=10)
    c = sm.cell(row=r, column=2, value=f)
    c.font = Font(name=ARIAL, size=10, bold=True)
    c.number_format = MONEY2 if "debits" in label else NUM
    c.alignment = Alignment(horizontal="center")
    c.border = BOX
    sev_label = {"must": "Must be cleared", "review": "Review required"}.get(sev, "For information")
    s2 = sm.cell(row=r, column=3, value=f'=IF(ROUND($B{r},2)=0,"Clear","{sev_label}")')
    s2.font = Font(name=ARIAL, size=9, bold=True)
    s2.alignment = Alignment(horizontal="center")
    s2.border = BOX
    sm.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
    r += 1
REV_R1 = r - 1
for txt, fill, font in (("Must be cleared", F_HIGH, T_HIGH), ("Review required", F_MED, T_MED),
                        ("Clear", F_LOW, T_LOW)):
    sm.conditional_formatting.add(f"B{REV_R0}:E{REV_R1}", FormulaRule(
        formula=[f'$C{REV_R0}="{txt}"'], fill=PatternFill("solid", fgColor=fill),
        font=Font(color=font, bold=True)))

r += 1
for txt in ("Department rows are NET CONTRIBUTION - income less costs. The lines underneath are gross, with income "
            "and costs both positive. 'Other' picks up anything mapped to depreciation, finance or tax so the "
            "department block always ties back to its own total.",
            "Overheads - the expense accounts with no division - report under ADMIN. Admin therefore carries the "
            "whole indirect cost base, so its contribution is deliberately a large negative: it is the cost of "
            "running the company, not a trading result. Column K on COA_Mapping shows where each account lands.",
            "The balance sheet block shows the MOVEMENT posted in each period, not the closing position, and stays "
            "nil until you map balance sheet accounts on COA_Mapping."):
    n = sm.cell(row=r, column=1, value=txt)
    n.font = Font(name=ARIAL, size=9, italic=True, color=MUTED)
    n.alignment = Alignment(wrap_text=True, vertical="top")
    sm.merge_cells(start_row=r, start_column=1, end_row=r + 1, end_column=13)
    r += 3
# =========================================================================== EXCEPTIONS  - ranked biggest movers, straight off the engine
# ===========================================================================
ex = sheet("Exceptions", "C00000")
title(ex, "Exception Report  -  biggest movers, ranked automatically",
      "Ranked from Data_Engine, so it is account level regardless of how the other sheets are grouped. "
      "Blank rows mean there were fewer movers than slots. Columns K and L are working cells.")
TOPN = 25
widths(ex, {"A": 7, "B": 38, "C": 22, "D": 15, "E": 14, "F": 14, "G": 14, "H": 11,
            "I": 11, "J": 3, "K": 14, "L": 9})

def exception_block(start_row, heading, key, cols, flag_key):
    band(ex, start_row, heading, 1, 9)
    header(ex, start_row + 1, ["Rank", "Account", "Group", "Reports under"] + cols + ["Risk Flag"])
    r0 = start_row + 2
    srng = ecol(key)
    for k in range(TOPN):
        r = r0 + k
        ex.cell(row=r, column=11, value=f'=LARGE({srng},{k + 1})').font = Font(name=ARIAL, size=8, color="BFBFBF")
        ex.cell(row=r, column=12,
                value=f'=IF($K{r}<0,"",MATCH($K{r},{srng},0))').font = Font(name=ARIAL, size=8, color="BFBFBF")
        rank = ex.cell(row=r, column=1, value=f'=IF($K{r}<0,"",{k + 1})')
        rank.alignment = Alignment(horizontal="center")
        rank.font = Font(name=ARIAL, size=9)
        rank.border = BOX
        src_cols = [("A", None), ("B", None), (CL(DER["RDEPT"]), None)] + COLMAP[key] + [(CL(DER[flag_key]), None)]
        for j, (sc, fmt) in enumerate(src_cols):
            c = ex.cell(row=r, column=2 + j,
                        value=f'=IF($K{r}<0,"",INDEX(Data_Engine!${sc}${ENG_R0}:${sc}${ENG_R1},$L{r}))')
            c.font = Font(name=ARIAL, size=9, color=GREEN_FONT)
            c.border = BOX
            if fmt:
                c.number_format = fmt
            if j == len(src_cols) - 1:
                c.alignment = Alignment(horizontal="center")
                c.font = Font(name=ARIAL, size=9, bold=True)
        pct = ex.cell(row=r, column=9, value=f'=IF($A{r}="","",IF(ROUND($F{r},2)=0,"n/a",$G{r}/ABS($F{r})))')
        pct.number_format = PCT
        pct.font = Font(name=ARIAL, size=9)
        pct.border = BOX
    risk_cf(ex, f"A{r0}:I{r0 + TOPN - 1}", "I", r0)
    flag_cell_cf(ex, f"I{r0}:I{r0 + TOPN - 1}", r0, "I")
    return r0 + TOPN

COLMAP = {
    "MOM_KEY": [(CL(DER["CY_MTH"]), MONEY), (CL(DER["CY_PRI"]), MONEY), (CL(DER["MOM_VAR"]), MONEY)],
    "YOY_KEY": [(CL(DER["CY_YTD"]), MONEY), (CL(DER["PY_YTD"]), MONEY), (CL(DER["YOY_VAR"]), MONEY)],
}
nxt = exception_block(4, f"TOP {TOPN} MONTH-ON-MONTH MOVEMENTS  (by absolute dollar movement)",
                      "MOM_KEY", ["Current\nMonth", "Prior\nMonth", "Movement $", "Movement %"], "MOM_FLAG")
nxt = exception_block(nxt + 2, f"TOP {TOPN} YEAR-ON-YEAR VARIANCES  (year to date, by absolute dollar variance)",
                      "YOY_KEY", ["CY YTD", "PY YTD", "YTD Var $", "YTD Var %"], "YOY_FLAG")
lbl(ex, 3, 11, "working cells", size=8, color="BFBFBF")
# ===========================================================================
# CONTROLS  - the data-integrity checks a controller signs off
# ===========================================================================
ct = sheet("Controls", "7030A0")
title(ct, "Controls & Data Integrity Checks",
      "Run these before you trust a single number on the other sheets. Anything marked FAIL means the analysis is wrong, "
      "not just untidy. REVIEW means look at it and satisfy yourself.")
widths(ct, {"A": 7, "B": 54, "C": 15, "D": 14, "E": 12, "F": 10, "G": 62})
header(ct, 5, ["Ref", "Control", "Result", "Target", "Status", "Severity", "What to do if it is not clean"])
ct.freeze_panes = "A6"

GA, GT = f"GL_Data!$A${GL_R0}:$A${GL_R1}", f"GL_Data!$T${GL_R0}:$T${GL_R1}"
GI = f"GL_Data!$I${GL_R0}:$I${GL_R1}"
# raw Dr-Cr: blank cells coerce to 0, so ABS/ROUND/MOD stay safe inside SUMPRODUCT
GAMT = f"(GL_Data!$G${GL_R0}:$G${GL_R1}-GL_Data!$H${GL_R0}:$H${GL_R1})"
# cleaned-source equivalents, so a control never reports PASS off an empty sheet
CD = f"Cleanup!$S${CLN_R0}:$S${CLN_R1}"     # date, numeric, kept rows only (0 elsewhere)
CC = f"Cleanup!$C${CLN_R0}:$C${CLN_R1}"
CR2 = f"Cleanup!$D${CLN_R0}:$D${CLN_R1}"
CK = f"Cleanup!$K${CLN_R0}:$K${CLN_R1}"
CH = f"Cleanup!$H${CLN_R0}:$H${CLN_R1}"
CHECKS = [
    ("C1", "Source data is in balance (total debits less total credits)",
     f'=IF({SRC}="Raw_Paste",Setup!$B$41,'
     f'ROUND(SUM(GL_Data!$G${GL_R0}:$G${GL_R1})-SUM(GL_Data!$H${GL_R0}:$H${GL_R1}),2))', "= 0", "eq0", "FAIL", MONEY2,
     "Tests whichever source Setup B12 is pointing at. If you dumped a single account rather than the whole ledger "
     "this will never be nil - that is expected, and the analysis still works, but you lose this control."),
    ("C2", "All GL lines net to nil (double entry intact)",
     f'=IF({SRC}="Raw_Paste",Setup!$B$41,ROUND(SUM({GI}),2))', "= 0", "eq0", "FAIL", MONEY2,
     "Same cause as C1. If you pasted signed amounts into one column, this still needs to be nil overall."),
    ("C3", "Every GL line maps to an account in COA_Mapping",
     f'=IF({SRC}="Raw_Paste",Setup!$B$42,COUNTIF({GT},"UNMAPPED"))', "= 0", "eq0", "FAIL", NUM,
     "Filter the live source for UNMAPPED (Cleanup column Q, or GL_Data column T), copy those account names and add "
     "them to COA_Mapping. Unmapped lines are silently excluded from every report, so never leave this red."),
    ("C4", "No GL lines with a date but no account name",
     f'=IF({SRC}="Raw_Paste",'
     f'SUMPRODUCT((Cleanup!$E${CLN_R0}:$E${CLN_R1}="Transaction")*(Cleanup!$G${CLN_R0}:$G${CLN_R1}="")),'
     f'COUNTIF({GT},"NO ACCOUNT"))', "= 0", "eq0", "FAIL", NUM,
     "Transaction rows the workbook could not attach to an account. On a Raw_Paste dump that means transactions "
     "appeared before the first account heading - check the account column number on Setup B54."),
    ("C5", "Data_Engine is still aligned to COA_Mapping",
     f'=SUMPRODUCT(--(Data_Engine!$A${ENG_R0}:$A${ENG_R1}&""<>COA_Mapping!$A${COA_R0}:$A${COA_R1}&""))',
     "= 0", "eq0", "FAIL", NUM,
     "Someone inserted or deleted rows on COA_Mapping. Undo it, or rebuild the workbook. Overwrite rows in place instead."),
    ("C6", "Current reporting period actually has data",
     f'=IF({SRC}="Raw_Paste",'
     f'COUNTIFS(Cleanup!$M${CLN_R0}:$M${CLN_R1},{B7},Cleanup!$N${CLN_R0}:$N${CLN_R1},{B8}),'
     f'COUNTIFS(GL_Data!$J${GL_R0}:$J${GL_R1},{B7},GL_Data!$K${GL_R0}:$K${GL_R1},{B8}))', "> 0", "gt0", "FAIL", NUM,
     "Either the period on Setup is wrong, or that month has not been posted yet. Check Setup cell B8."),
    ("C7", "Prior year comparative is loaded",
     f'=IF({SRC}="Raw_Paste",COUNTIF(Cleanup!$M${CLN_R0}:$M${CLN_R1},{B9}),'
     f'COUNTIF(GL_Data!$J${GL_R0}:$J${GL_R1},{B9}))', "> 0", "gt0", "REVIEW", NUM,
     "Without prior year data every year-on-year variance reads as a new account. Load both years into GL_Data."),
    ("C8", "No GL lines dated outside the two financial years on Setup",
     f'=IF({SRC}="Raw_Paste",COUNTIF(Cleanup!$Q${CLN_R0}:$Q${CLN_R1},"OUTSIDE FY RANGE"),'
     f'COUNTIF({GT},"OUTSIDE FY RANGE"))', "= 0", "eq0", "REVIEW", NUM,
     "Those lines are ignored by the analysis. Fine if you deliberately loaded three years; a problem if the dates are wrong."),
    ("C9", "Row capacity not exceeded on the live source",
     f'=IF({SRC}="Raw_Paste",IF(Setup!$B$33>={CLN_R1 - CLN_R0 + 1},1,0),'
     f'IF(Setup!$B$21>={GL_R1 - GL_R0 + 1},1,0))', "= 0", "eq0", "FAIL", NUM,
     f"Raw_Paste holds {RAW_R1 - RAW_R0 + 1:,} rows and GL_Data holds {GL_R1 - GL_R0 + 1:,}. If you have filled the "
     "sheet, data is being cut off. Export in two halves, or run a month-by-account summary report instead of "
     "transaction detail - month-on-month analysis does not need every line."),
    ("C10", "No possible duplicate journal lines (same account, date and amount)",
     f'=IF({SRC}="Raw_Paste",COUNTIF(Cleanup!$U${CLN_R0}:$U${CLN_R1},">1"),'
     f'COUNTIF(GL_Data!$V${GL_R0}:$V${GL_R1},">1"))', "= 0", "eq0", "REVIEW", NUM,
     "Sort GL_Data by column V. Recurring journals of identical value are legitimate; a double-posted invoice is not."),
    ("C11", "No transactions dated in the future",
     f'=IF({SRC}="Raw_Paste",SUMPRODUCT(({CD}>0)*({CD}>TODAY())),'
     f'SUMPRODUCT(({GA}<>"")*({GA}>TODAY())))', "= 0", "eq0", "REVIEW", NUM,
     "Usually a typed date error (wrong year). Check them before reporting."),
    ("C12", "No weekend postings",
     f'=IF({SRC}="Raw_Paste",SUMPRODUCT(({CD}>0)*(WEEKDAY({CD}+0,2)>5)),'
     f'SUMPRODUCT(({GA}>0)*(WEEKDAY({GA}+0,2)>5)))', "= 0", "info", "REVIEW", NUM,
     "Not wrong in itself, but weekend manual journals are a standard fraud-risk indicator. Check who posted them."),
    ("C13", "No round-dollar postings of $10,000 or more",
     f'=IF({SRC}="Raw_Paste",SUMPRODUCT(({CD}>0)*(ABS({CC}-{CR2})>=10000)*(MOD(ABS({CC}-{CR2}),1000)=0)),'
     f'SUMPRODUCT(({GA}<>"")*(ABS{GAMT}>=10000)*(MOD(ABS{GAMT},1000)=0)))', "= 0", "info", "REVIEW", NUM,
     "Large round numbers are usually accruals or estimates. Confirm each one is supported."),
    ("C14", "Every GL line has a description",
     f'=IF({SRC}="Raw_Paste",SUMPRODUCT({CK}*({CH}="")),'
     f'SUMPRODUCT(({GA}<>"")*(GL_Data!$C${GL_R0}:$C${GL_R1}="")))', "= 0", "info", "REVIEW", NUM,
     "Unexplained journals are the hardest thing to defend at audit. Get narrations added at source."),
    ("C15", "No nil-value GL lines",
     f'=IF({SRC}="Raw_Paste",SUMPRODUCT({CK}*(ROUND({CC}-{CR2},2)=0)),'
     f'SUMPRODUCT(({GA}<>"")*(ROUND({GAMT[1:-1]},2)=0)))', "= 0", "info", "REVIEW", NUM,
     "Harmless but they bloat the file and hide reversals. Consider stripping them from the export."),
    ("C16", "Every mapped account has a valid group",
     f'=COUNTIF(COA_Mapping!$I${COA_R0}:$I${COA_R1},"NO GROUP")+COUNTIF(COA_Mapping!$I${COA_R0}:$I${COA_R1},"BAD GROUP")',
     "= 0", "eq0", "FAIL", NUM,
     "Those accounts fall out of the summary entirely. Set the Group on COA_Mapping using the dropdown."),
    ("C17", "Accounts in COA_Mapping with no GL activity",
     f'=COUNTIFS(COA_Mapping!$A${COA_R0}:$A${COA_R1},"<>",COA_Mapping!$H${COA_R0}:$H${COA_R1},0)', "= 0", "info", "REVIEW", NUM,
     "Dormant accounts. Harmless, but if an account you expected to see is here the name probably does not match the GL exactly."),
    ("C18", "High risk movements all have a controller comment",
     f'=SUMPRODUCT((MoM_Analysis!$X${MM_R0}:$X${MM_R1}="High")*(MoM_Analysis!$Y${MM_R0}:$Y${MM_R1}=""))'
     f'+SUMPRODUCT((YoY_Analysis!$P${YY_R0}:$P${YY_R1}="High")*(YoY_Analysis!$Q${YY_R0}:$Q${YY_R1}=""))',
     "= 0", "eq0", "REVIEW", NUM,
     "This is your sign-off. Every High flag needs an explanation before the pack goes out."),
    ("C20", "Raw_Paste: the cleaning kept a sensible number of lines",
     f'=IF({SRC}<>"Raw_Paste",0,IF(Setup!$B$33=0,0,IF(Setup!$B$34/Setup!$B$33<0.3,1,0)))', "= 0", "info", "REVIEW", NUM,
     "Fewer than 30% of your pasted rows were judged to be transactions. Usually the column numbers on Raw_Paste "
     "rows 4 to 14 are pointing at the wrong columns. Check the preview on Cleanup before you go further."),
    ("C21", "Raw_Paste: no pasted rows were left unclassified",
     f'=IF({SRC}<>"Raw_Paste",0,COUNTIF(Cleanup!$E${CLN_R0}:$E${CLN_R1},"Text only"))', "= 0", "info", "REVIEW", NUM,
     "Rows with text but no date and no amount that were not treated as account headings. Harmless if they are "
     "report titles; a problem if they are transactions whose date column was misidentified."),
    ("C22", "Every active account has somewhere to display",
     f'=SUMPRODUCT(({ecol("ACTIVE")}=1)*({ecol("SHOWN")}=0)*(Data_Engine!$C${ENG_R0}:$C${ENG_R1}="P&L"))',
     "= 0", "eq0", "FAIL", NUM,
     "A P&L account has more entries in its department and group than the block on MoM_Analysis has room for, so it is "
     "invisible on those sheets even though its money is in the department total. Move some accounts to another "
     "department on COA_Mapping, or ask for the block to be made bigger."),
    ("C23", "The reporting month resolved to a real date",
     '=IF(Setup!$B$10="",1,0)', "= 0", "eq0", "FAIL", NUM,
     "Setup B5 does not match any month in the picker list. Choose one from the dropdown rather than typing it."),
    ("C24", "Every account carries an Overhead flag",
     f'=SUMPRODUCT((COA_Mapping!$A${COA_R0}:$A${COA_R1}<>"")*(COA_Mapping!$J${COA_R0}:$J${COA_R1}=""))',
     "= 0", "info", "REVIEW", NUM,
     "Accounts with a blank Overhead? flag report under their own division. Set Yes for anything that is a "
     "company-wide overhead so it lands in the Overheads department instead."),
    ("C19", "Thresholds on Setup are sensible (High % above Medium %)",
     f'=IF({HIP}>{MEDP},0,1)', "= 0", "eq0", "FAIL", NUM,
     "Setup B16 must be greater than B15, otherwise the Medium band never triggers."),
]
CT_R0 = 6
for i, (ref, desc, f, target, mode, sev, fmt, guide) in enumerate(CHECKS):
    r = CT_R0 + i
    a = ct.cell(row=r, column=1, value=ref)
    a.font = Font(name=ARIAL, size=9, bold=True)
    a.alignment = Alignment(horizontal="center")
    b = ct.cell(row=r, column=2, value=desc)
    b.font = Font(name=ARIAL, size=10)
    b.alignment = Alignment(wrap_text=True, vertical="center")
    c = ct.cell(row=r, column=3, value=f)
    c.font = Font(name=ARIAL, size=10, bold=True)
    c.number_format = fmt
    c.alignment = Alignment(horizontal="center")
    d = ct.cell(row=r, column=4, value=target)
    d.font = Font(name=ARIAL, size=9, color=MUTED)
    d.alignment = Alignment(horizontal="center")
    cond = f'ROUND($C{r},2)=0' if mode in ("eq0", "info") else f'$C{r}>0'
    e = ct.cell(row=r, column=5, value=f'=IF({cond},"PASS","{sev}")')
    e.font = Font(name=ARIAL, size=10, bold=True)
    e.alignment = Alignment(horizontal="center")
    fsev = ct.cell(row=r, column=6, value=sev)
    fsev.font = Font(name=ARIAL, size=9, color=MUTED)
    fsev.alignment = Alignment(horizontal="center")
    g = ct.cell(row=r, column=7, value=guide)
    g.font = Font(name=ARIAL, size=9, color=MUTED)
    g.alignment = Alignment(wrap_text=True, vertical="center")
    for col in range(1, 8):
        ct.cell(row=r, column=col).border = BOX
    ct.row_dimensions[r].height = 30
CT_R1 = CT_R0 + len(CHECKS) - 1
ct.conditional_formatting.add(f"A{CT_R0}:G{CT_R1}", FormulaRule(
    formula=[f'$E{CT_R0}="FAIL"'], fill=ROW_HIGH, font=Font(color=T_HIGH)))
ct.conditional_formatting.add(f"A{CT_R0}:G{CT_R1}", FormulaRule(
    formula=[f'$E{CT_R0}="REVIEW"'], fill=ROW_MED, font=Font(color=T_MED)))
for txt, fill, font in (("PASS", F_LOW, T_LOW), ("REVIEW", F_MED, T_MED), ("FAIL", F_HIGH, T_HIGH)):
    ct.conditional_formatting.add(f"E{CT_R0}:E{CT_R1}", FormulaRule(
        formula=[f'$E{CT_R0}="{txt}"'], fill=PatternFill("solid", fgColor=fill),
        font=Font(color=font, bold=True)))

r = CT_R1 + 2
band(ct, r, "OVERALL", 1, 7)
for j, (label, f) in enumerate((
        ("Controls failing (must be fixed before the pack is used)", f'=COUNTIF($E${CT_R0}:$E${CT_R1},"FAIL")'),
        ("Controls needing review", f'=COUNTIF($E${CT_R0}:$E${CT_R1},"REVIEW")'),
        ("Controls passed", f'=COUNTIF($E${CT_R0}:$E${CT_R1},"PASS")'))):
    rr = r + 1 + j
    lbl(ct, rr, 2, label, bold=True, size=10)
    c = ct.cell(row=rr, column=3, value=f)
    c.font = Font(name=ARIAL, size=11, bold=True)
    c.alignment = Alignment(horizontal="center")
    c.border = BOX
    c.number_format = NUM
ct.conditional_formatting.add(f"C{r + 1}", FormulaRule(formula=[f'$C${r + 1}>0'],
    fill=PatternFill("solid", fgColor=F_HIGH), font=Font(color=T_HIGH, bold=True)))
ct.conditional_formatting.add(f"C{r + 1}", FormulaRule(formula=[f'$C${r + 1}=0'],
    fill=PatternFill("solid", fgColor=F_LOW), font=Font(color=T_LOW, bold=True)))
ct.conditional_formatting.add(f"C{r + 2}", FormulaRule(formula=[f'$C${r + 2}>0'],
    fill=PatternFill("solid", fgColor=F_MED), font=Font(color=T_MED, bold=True)))

# ===========================================================================
# README
# ===========================================================================
rd = wb.create_sheet("README", 0)
rd.sheet_view.showGridLines = False
title(rd, "GL Month-on-Month & Year-on-Year Analysis",
      "Financial controller pack. Paste the GL, set two cells, read the flags.")
widths(rd, {"A": 4, "B": 26, "C": 96})

def rd_band(r, text):
    band(rd, r, text, 1, 3)
    return r + 1

def rd_line(r, left, right, bold=False):
    lbl(rd, r, 2, left, bold=True, size=10)
    c = rd.cell(row=r, column=3, value=right)
    c.font = Font(name=ARIAL, size=10, bold=bold)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    rd.row_dimensions[r].height = max(14, 13 * (1 + len(right) // 105))
    return r + 1

r = 4
r = rd_band(r, "HOW TO USE IT  -  five steps, in this order")
for i, (a, b) in enumerate([
    ("Step 1", "Open Setup and pick the month you are reporting on from the dropdown in cell B5. That is the only "
               "date control - the financial year and the period number work themselves out from it. Check the risk "
               "thresholds in B14:B17 suit your business."),
    ("Step 2", "Open Raw_Paste and dump your Xero export into cell A1, exactly as it comes out. The sheet is "
               "deliberately empty so you can select the whole export, copy, and paste straight into A1 - Excel "
               "will not let you paste a whole-sheet selection anywhere else. Leave the title rows, the account "
               "headings, the subtotals and the blank lines in. Then set the column numbers on Setup rows 53 to 63."),
    ("Step 3", "Open Cleanup and look at the preview block on the right. If those lines look like your transactions, "
               "the column numbers are right. If they look wrong, go back and fix the numbers on Raw_Paste - that is "
               "almost always the problem."),
    ("Step 4", "Open COA_Mapping. Every account name the cleaning found must appear in column A, spelled exactly "
               "the same. It is pre-loaded with your 190 accounts from Xero. Set the Group, the Division, and the "
               "Overhead? flag - anything marked Yes reports under Overheads instead of its division. Column K "
               "shows where each account will actually appear. Overwrite rows in place - never insert or delete."),
    ("Step 5", "Open Controls. Clear every FAIL before you read anything else. Then Summary, then Exceptions, "
               "then the two detail sheets."),
]):
    r = rd_line(r, a, b)

r += 1
r = rd_band(r, "THE DIRTY DUMP  -  what it throws away, and what it keeps")
for a, b in [
    ("The idea", "You should never have to delete a row by hand. Paste the export whole and the Cleanup sheet decides "
                 "what each row is. Nothing is hidden: Setup rows 33 to 42 count every row it dropped and why."),
    ("Kept", "A row with a readable date AND a debit or credit AND an account name it could work out. That is it."),
    ("Dropped - headings", "Rows with text but no date and no amount. In a Xero Account Transactions or General "
                           "Ledger export that is the account name sitting above its transactions, so the account "
                           "is remembered and applied to every line underneath until the next heading."),
    ("Dropped - subtotals", "Rows with an amount but no date - that is how every Total line behaves - plus anything "
                            "containing Opening Balance or Closing Balance, and anything starting with Total. "
                            "These are the rows that would double-count if you left them in."),
    ("Dropped - empty lines", "Rows with a date but no value, and completely blank rows. This is the setting on "
                              "Setup B62 and it is the one that saves you the deleting."),
    ("Two layouts", "Section headings (General Ledger Detail, Account Transactions - the account is a heading row) "
                    "or Account in a column (Journal report - every row names its own account). Setup B53."),
    ("Switching source", "Setup B12 chooses which sheet the reports read. Raw_Paste uses the cleaned dump. GL_Data "
                         "uses the tidy sheet you fill in yourself. Only one is live at a time, so there is no "
                         "double counting - but check B12 says what you think it says."),
    ("Check it balances", "Setup B41 is kept debits less kept credits. If you dumped the whole ledger it must be nil. "
                          "If it is not, the cleaning has taken something it should not have, or missed a column. "
                          "If you only dumped one account it will not be nil, and that is fine."),
]:
    r = rd_line(r, a, b)

r += 1
r = rd_band(r, "WHAT EACH SHEET IS FOR")
for a, b in [
    ("Setup", "Control panel. Financial year, reporting period, risk thresholds, and which data source is live."),
    ("Raw_Paste", "The dirty dump zone, starting at A1. Paste the Xero export as-is. Nothing is formatted, "
                  "nothing needs deleting, and nothing sits above the data to get in the way of a plain paste."),
    ("Cleanup", "Calculated. Judges every pasted row, carries account headings down, and keeps only real "
                "transactions. Has a preview so you can check the cleaning before trusting it."),
    ("GL_Data", "The alternative source, for data you have already tidied. Columns A-H are yours, I-V are calculated."),
    ("COA_Mapping", "Maps each GL account to a reporting group and a division. This is what makes the summary work."),
    ("Lists", "The group master. Sets the sign convention (income shows positive, expenses show positive) and report order."),
    ("Data_Engine", "Calculated. Aggregates the GL once by account and period so the reports are fast. Never type here."),
    ("MoM_Analysis", "Account level. Twelve months across, then this month vs last month and vs the rolling 3-month average."),
    ("YoY_Analysis", "Account level. This month vs the same month last year, year to date vs last year to date, and a run-rate."),
    ("Summary", "The one page. Group P&L, revenue and gross profit by division, and the review status block."),
    ("Exceptions", "The 25 biggest month-on-month movers and the 25 biggest year-on-year variances, ranked automatically."),
    ("Controls", "Nineteen data integrity checks. FAIL means the numbers are wrong, not just untidy."),
]:
    r = rd_line(r, a, b)

r += 1
r = rd_band(r, "READING IT BY DEPARTMENT")
for a, b in [
    ("The three levels", "Summary, MoM_Analysis and YoY_Analysis all group the same way. A department row, then its "
                         "group subtotals, then the accounts. Use the + and - buttons in the left margin, or the "
                         "small 1 2 3 buttons above them to open every department to the same depth at once."),
    ("It opens collapsed", "You see departments and their group subtotals. The accounts are hidden until you open "
                           "them, which is the point - you drill into the one department that moved, not all eight."),
    ("Department rows", "NET CONTRIBUTION: income less costs, for that department. The lines underneath are gross, "
                        "with income and costs both shown positive. The 'Other' line on Summary catches anything "
                        "mapped to depreciation, finance or tax so each department block ties to its own total."),
    ("Overheads", "The expense accounts with no division report under ADMIN, flagged Overhead? = Yes on "
                  "COA_Mapping column J. Column K shows where each account actually lands. Set an account's "
                  "Overhead? flag to No and give it a Division to move it out to an operating department."),
    ("Block sizes", "Each department and group block has a fixed number of account slots with room to spare. "
                    "Control C22 tells you if an account has nowhere to appear, which happens only if you re-map a "
                    "lot of accounts into one department."),
    ("Exceptions", "Ranked straight off Data_Engine at account level, so it ignores the grouping entirely. That is "
                   "the sheet to read first when you do not yet know which department moved."),
]:
    r = rd_line(r, a, b)

r += 1
r = rd_band(r, "THE COLOUR CODE")
for a, b, fill, font in [
    ("High risk", "Variance at or above the High % threshold, or above the High $ threshold, or a new/ceased account "
                  "with a material movement. Explain every one of these before the pack goes out.", F_HIGH, T_HIGH),
    ("Medium risk", "Variance at or above the Medium % threshold and above the materiality floor. Worth a look.", F_MED, T_MED),
    ("Low risk", "Either immaterial in dollars or within the Medium threshold. No action.", F_LOW, T_LOW),
    ("No Activity", "Nil in both periods. Greyed out so it stops competing for your attention.", F_NA, T_NA),
]:
    c = rd.cell(row=r, column=2, value=a)
    c.font = Font(name=ARIAL, size=10, bold=True, color=font)
    c.fill = PatternFill("solid", fgColor=fill)
    c.alignment = Alignment(horizontal="center")
    c.border = BOX
    d = rd.cell(row=r, column=3, value=b)
    d.font = Font(name=ARIAL, size=10)
    d.alignment = Alignment(wrap_text=True, vertical="top")
    rd.row_dimensions[r].height = max(16, 13 * (1 + len(b) // 105))
    r += 1
r += 1
for a, b, fill, font in [
    ("Blue text on yellow", "You type here.", "FFF2CC", BLUE_FONT),
    ("Black text", "A formula that works on that sheet.", "FFFFFF", BLACK_FONT),
    ("Green text", "A formula pulling from another sheet.", "FFFFFF", GREEN_FONT),
]:
    c = rd.cell(row=r, column=2, value=a)
    c.font = Font(name=ARIAL, size=10, bold=True, color=font)
    c.fill = PatternFill("solid", fgColor=fill)
    c.alignment = Alignment(horizontal="center")
    c.border = BOX
    rd.cell(row=r, column=3, value=b).font = Font(name=ARIAL, size=10)
    r += 1

r += 1
r = rd_band(r, "HOW THE RISK FLAG IS CALCULATED  (so you can defend it)")
for a, b in [
    ("Test order", "1. Both periods nil  ->  No Activity.   2. Dollar variance below the materiality floor  ->  Low.   "
                   "3. Comparative was nil and the movement is material  ->  High (new or ceased account).   "
                   "4. Percentage at or above the High threshold, OR dollar variance above the High dollar threshold  ->  High.   "
                   "5. Percentage at or above the Medium threshold  ->  Medium.   6. Otherwise  ->  Low."),
    ("Why dollars first", "A 400% swing on an $80 account is noise. The materiality floor stops the report crying wolf. "
                          "Set it on Setup B14 - roughly 0.5% of revenue is a sensible starting point."),
    ("Percentage base", "Variance divided by the absolute value of the comparative, so the sign shows direction, not the "
                        "quirk of a credit balance. Where the comparative is nil the percentage shows n/a and the "
                        "dollar test decides the flag."),
]:
    r = rd_line(r, a, b)

r += 1
r = rd_band(r, "THINGS THAT WILL CATCH YOU OUT")
for a, b in [
    ("Account names", "The match between GL_Data column B and COA_Mapping column A is exact. A trailing space or a "
                      "changed account name silently drops the account out of every report. Control C3 catches it."),
    ("Row inserts", "Do not insert or delete rows on COA_Mapping - Data_Engine is aligned to it row for row. "
                    "Control C5 catches it if you do."),
    ("Sign convention", "Income and liabilities are stored as credits in the GL and flipped on the reports, so income "
                        "reads positive and expenses read positive. Gross Profit = Income less Cost of Sales."),
    ("EBITDA", "Your accounts currently carry depreciation and interest inside the Expenses group, so the EBITDA line "
               "equals Net Profit Before Tax. Re-map those accounts on COA_Mapping to split them out."),
    ("Balance sheet", "The balance sheet block shows the movement posted in the period, not the closing position. "
                      "Include your opening balance journal in the export if you want positions."),
    ("Capacity", f"Raw_Paste holds {RAW_R1 - RAW_R0 + 1:,} rows, GL_Data holds {GL_R1 - GL_R0 + 1:,}, and there is room "
                 f"for {NACC} accounts. Control C9 warns you when the rows run out. If a year of transaction detail "
                 "will not fit, run the Xero report summarised by month instead - month-on-month analysis does not "
                 "need every individual line, and the workbook will be far quicker."),
    ("If the cleaning looks wrong", "It is nearly always the column numbers on Setup rows 53 to 63. Check the "
                                    "preview on Cleanup first, then Setup rows 33 to 42 to see what was dropped. "
                                    "Controls C20 and C21 flag the two usual symptoms."),
]:
    r = rd_line(r, a, b)

r += 1
lbl(rd, r, 2, "Built", bold=True, size=10)
src = rd.cell(row=r, column=3, value=(
    "Chart of accounts pre-loaded from reporting/data/accounts.json (190 accounts, sourced from Xero P&L, accrual "
    "basis, pulled 2026-09-03 per reporting/data/validation.json). Group and division per account come from that "
    "same file. No transaction data is included - GL_Data ships empty."))
src.font = Font(name=ARIAL, size=9, italic=True, color=MUTED)
src.alignment = Alignment(wrap_text=True, vertical="top")
rd.row_dimensions[r].height = 40

# ===========================================================================
# finish
# ===========================================================================
del wb["Sheet"]
wb.move_sheet("README", offset=-wb.sheetnames.index("README"))
order = ["README", "Setup", "Raw_Paste", "GL_Data", "COA_Mapping", "Summary", "MoM_Analysis",
         "YoY_Analysis", "Exceptions", "Controls", "Cleanup", "Data_Engine", "Lists"]
wb._sheets = [wb[n] for n in order]

for ws in wb.worksheets:                      # Arial everywhere, size preserved
    for row in ws.iter_rows():
        for c in row:
            if c.value is not None or c.has_style:
                f = c.font
                if f.name != ARIAL:
                    c.font = Font(name=ARIAL, size=f.size or 10, bold=f.bold, italic=f.italic,
                                  color=f.color, underline=f.underline)
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True

for name in ("Summary", "MoM_Analysis", "YoY_Analysis", "Exceptions", "Controls"):
    wb[name].print_title_rows = "1:5"

# Set last: the summary row sits ABOVE its detail, so Excel must be told
# summaryBelow=0 or the +/- buttons land on the wrong rows.
for name in ("Summary", "MoM_Analysis", "YoY_Analysis"):
    wb[name].sheet_properties.outlinePr = Outline(summaryBelow=False, summaryRight=False, applyStyles=False)

wb.active = wb.sheetnames.index("README")
OUT.parent.mkdir(parents=True, exist_ok=True)
wb.save(OUT)
print("saved:", OUT)
print("sheets:", wb.sheetnames)
