"""
CTS Financial Controller Pack - one workbook, three paste points.

Replaces the manual chain of PL_Analysis.xlsm -> 2026-08_Graph.xlsx ->
CTS_Budget_FY27.xlsx -> Utilisation_Report.xlsx. Paste the Xero GL export,
the budget and the hours; every other sheet is formulas off those three.

Sign convention follows the existing pack: Amount = Credit - Debit, so income
is positive, costs are negative, and any total is a contribution.
"""
import json
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as CL
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.properties import Outline
from openpyxl.formatting.rule import FormulaRule
from openpyxl.comments import Comment
from openpyxl.chart import LineChart, BarChart, Reference

ROOT = Path("/home/user/Claude/reporting")
OUT = ROOT / "cts_pack" / "CTS Financial Controller Pack.xlsx"
ACCTS = json.loads((ROOT / "data/cts_accounts.json").read_text())
JOBS = json.loads((ROOT / "data/cts_jobs.json").read_text())
BUDGET = json.loads((ROOT / "data/cts_budget_fy27.json").read_text())

# ---- capacity -------------------------------------------------------------
GL_R0, GL_R1 = 2, 30001            # GL_Paste: export pasted at A1, 30,000 rows
CL_R0, CL_R1 = 6, 30005            # Cleanup, one row per pasted row
NACC = len(ACCTS)
DEPTS = ["Onsite", "Production", "Video", "Integration", "Consulting", "CTS"]
CATS = ["Income", "Cost of Sales", "Expenses", "Other Income", "Other Expenses"]
TYPES = ["Labour", "Equipment & Misc", "Transport / Freight & Accommodation",
         "Admin, Recruitment & Training"]

# subcategories in the order they appear in the account master
SUBS, _seen = [], set()
for a in ACCTS:
    k = (a["category"], a["subcategory"])
    if k not in _seen and a["subcategory"]:
        _seen.add(k); SUBS.append(k)
SUB_NAMES = [s for _c, s in SUBS]

# ---- styling --------------------------------------------------------------
ARIAL = "Arial"
NAVY, MUTED, WHITE = "1F3864", "595959", "FFFFFF"
HEAD = PatternFill("solid", fgColor=NAVY)
SUBHEAD = PatternFill("solid", fgColor="2E5C8A")
BAND = PatternFill("solid", fgColor="D9E2F3")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
TOTAL_FILL = PatternFill("solid", fgColor="EDEDED")
GREY_FILL = PatternFill("solid", fgColor="F2F2F2")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
BLUE_FONT, BLACK_FONT, GREEN_FONT = "0000FF", "000000", "008000"
F_HIGH, T_HIGH = "FFC7CE", "9C0006"
F_MED, T_MED = "FFEB9C", "9C6500"
F_LOW, T_LOW = "C6EFCE", "006100"
F_NA, T_NA = "F2F2F2", "808080"
ROW_HIGH = PatternFill("solid", fgColor="FDECEA")
ROW_MED = PatternFill("solid", fgColor="FFF7E0")
MONEY = '$#,##0;($#,##0);"-"'
MONEY2 = '$#,##0.00;($#,##0.00);"-"'
PCT = '0.0%;(0.0%);"-"'
PP = '0.0"pp";(0.0"pp");"-"'
NUM = '#,##0;(#,##0);"-"'
HRS = '#,##0.0;(#,##0.0);"-"'
DATEF = 'dd/mm/yyyy'

wb = Workbook()

def sheet(name, tab=None):
    ws = wb.create_sheet(name)
    ws.sheet_view.showGridLines = False
    if tab: ws.sheet_properties.tabColor = tab
    return ws

def title(ws, text, sub=None):
    ws["A1"] = text
    ws["A1"].font = Font(name=ARIAL, size=15, bold=True, color=NAVY)
    if sub:
        ws["A2"] = sub
        ws["A2"].font = Font(name=ARIAL, size=9, italic=True, color=MUTED)

def header(ws, row, labels, start_col=1, fill=HEAD, size=9, h=30):
    for i, lab in enumerate(labels):
        c = ws.cell(row=row, column=start_col + i, value=lab)
        c.font = Font(name=ARIAL, size=size, bold=True, color=WHITE)
        c.fill = fill; c.border = BOX
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = h

def band(ws, row, text, c0, c1, fill=SUBHEAD):
    c = ws.cell(row=row, column=c0, value=text)
    c.font = Font(name=ARIAL, size=10, bold=True, color=WHITE)
    for col in range(c0, c1 + 1):
        ws.cell(row=row, column=col).fill = fill
    ws.row_dimensions[row].height = 18

def widths(ws, spec):
    for col, w in spec.items(): ws.column_dimensions[col].width = w

def lbl(ws, r, c, text, bold=False, size=10, color="000000", indent=0):
    cell = ws.cell(row=r, column=c, value=text)
    cell.font = Font(name=ARIAL, size=size, bold=bold, color=color)
    if indent: cell.alignment = Alignment(indent=indent)
    return cell

def inp(cell, note=None):
    cell.font = Font(name=ARIAL, size=10, color=BLUE_FONT, bold=True)
    cell.fill = INPUT_FILL; cell.border = BOX
    if note: cell.comment = Comment(note, "CTS Pack")

print(f"accounts {NACC} | subcategories {len(SUBS)} | departments {len(DEPTS)} | jobs {len(JOBS)}")

# ===========================================================================
# LISTS  - the reference data every lookup keys off
# ===========================================================================
ls = sheet("Lists", "808080")
title(ls, "Reference Lists",
      "Pulled from the GL code database in your PL Analysis workbook. Accounts, subcategories, the "
      "Labour / Equipment type split, departments and the job master.")
header(ls, 4, ["#", "Account Code", "Account Name", "Category", "Subcategory", "Type",
               "Cat #", "Sub #", "Type #"])
for i, a in enumerate(ACCTS):
    r = 5 + i
    vals = [i + 1, a["code"], a["name"], a["category"], a["subcategory"], a["type"],
            (CATS.index(a["category"]) + 1) if a["category"] in CATS else 0,
            (SUB_NAMES.index(a["subcategory"]) + 1) if a["subcategory"] in SUB_NAMES else 0,
            (TYPES.index(a["type"]) + 1) if a["type"] in TYPES else 0]
    for j, v in enumerate(vals):
        c = ls.cell(row=r, column=1 + j, value=v)
        c.font = Font(name=ARIAL, size=9); c.border = BOX
ACC_R0, ACC_R1 = 5, 4 + NACC

ROLLUP = {"Onsite": "Support", "Production": "Production", "Video": "Production",
          "Integration": "Consulting", "Consulting": "Consulting", "CTS": "CTS"}
ROLL_NAMES = ["Support", "Production", "Consulting", "CTS"]
header(ls, 4, ["#", "Department", "Cost Centre (as in GL)", "Reports as"], start_col=11)
for i, d in enumerate(DEPTS):
    for j, v in enumerate([i + 1, d, d.upper(), ROLLUP[d]]):
        c = ls.cell(row=5 + i, column=11 + j, value=v)
        c.font = Font(name=ARIAL, size=9); c.border = BOX
DEP_R0, DEP_R1 = 5, 4 + len(DEPTS)

header(ls, 4, ["#", "Subcategory", "Category"], start_col=15)
for i, (cat, sub) in enumerate(SUBS):
    for j, v in enumerate([i + 1, sub, cat]):
        c = ls.cell(row=5 + i, column=15 + j, value=v)
        c.font = Font(name=ARIAL, size=9); c.border = BOX
SUB_R0, SUB_R1 = 5, 4 + len(SUBS)

header(ls, 4, ["#", "Type"], start_col=19)
for i, t in enumerate(TYPES):
    for j, v in enumerate([i + 1, t]):
        c = ls.cell(row=5 + i, column=19 + j, value=v)
        c.font = Font(name=ARIAL, size=9); c.border = BOX
TYP_R0, TYP_R1 = 5, 4 + len(TYPES)

header(ls, 4, ["Month", "Date"], start_col=22)
import datetime as _dt
for i in range(84):
    d = _dt.date(2023 + (6 + i) // 12, (6 + i) % 12 + 1, 1)
    c = ls.cell(row=5 + i, column=22, value=d.strftime("%b %Y"))
    c.font = Font(name=ARIAL, size=9); c.border = BOX
    c2 = ls.cell(row=5 + i, column=23, value=d)
    c2.number_format = DATEF; c2.font = Font(name=ARIAL, size=9); c2.border = BOX
MTH_R0, MTH_R1 = 5, 88

header(ls, 4, ["Job No.", "Job Name", "Department"], start_col=25)
for i, j in enumerate(JOBS):
    for k, v in enumerate([j["job"], j["name"], j["department"]]):
        c = ls.cell(row=5 + i, column=25 + k, value=v)
        c.font = Font(name=ARIAL, size=8); c.border = BOX
JOB_R0, JOB_R1 = 5, 4 + len(JOBS)
widths(ls, {"A": 5, "B": 12, "C": 32, "D": 15, "E": 28, "F": 26, "G": 7, "H": 7, "I": 7,
            "K": 5, "L": 14, "M": 20, "N": 14, "O": 5, "P": 28, "Q": 15, "S": 5, "T": 30,
            "V": 12, "W": 12, "Y": 10, "Z": 34, "AA": 14})

L_ACC_CODE = f"Lists!$B${ACC_R0}:$B${ACC_R1}"
L_ACC_NAME = f"Lists!$C${ACC_R0}:$C${ACC_R1}"
L_ACC_SUB = f"Lists!$H${ACC_R0}:$H${ACC_R1}"
L_ACC_TYP = f"Lists!$I${ACC_R0}:$I${ACC_R1}"
L_DEP_CC = f"Lists!$M${DEP_R0}:$M${DEP_R1}"
L_JOB = f"Lists!$Y${JOB_R0}:$Y${JOB_R1}"
L_JOB_DEP = f"Lists!$AA${JOB_R0}:$AA${JOB_R1}"
L_DEP_NAME = f"Lists!$L${DEP_R0}:$L${DEP_R1}"

# ===========================================================================
# SETUP
# ===========================================================================
st = sheet("Setup", "1F3864")
title(st, "Setup & Control Panel",
      "Yellow cells are the only ones you type into. Pick the reporting month and everything follows.")
widths(st, {"A": 46, "B": 22, "C": 66})

def srow(r, label, value, fmt=None, note=None, is_input=True):
    lbl(st, r, 1, label, size=10)
    c = st.cell(row=r, column=2, value=value)
    if is_input: inp(c, note)
    else:
        c.font = Font(name=ARIAL, size=10, bold=True); c.border = BOX; c.fill = TOTAL_FILL
    c.alignment = Alignment(horizontal="center")
    if fmt: c.number_format = fmt
    return c

band(st, 3, "1.  WHAT AM I REPORTING ON", 1, 3)
srow(4, "Entity name", "Corporate Technology Services Pty Ltd")
srow(5, "Reporting month", "Aug 2026",
     note="Pick from the dropdown. The financial year and period number follow from it.")
lbl(st, 5, 3, "<-  the one control you change each month", size=10, color="C00000", bold=True)
srow(6, "Financial year starts in month", 7, NUM, note="7 = July. Set once.")
srow(7, "Financial year (the year it ENDS)",
     '=IF($B$10="","",IF($B$6=1,YEAR($B$10),YEAR($B$10)+IF(MONTH($B$10)>=$B$6,1,0)))', NUM, is_input=False)
srow(8, "Period number in the financial year",
     '=IF($B$10="","",MOD(MONTH($B$10)-$B$6+12,12)+1)', NUM, is_input=False)
srow(9, "Prior financial year", '=IF($B$7="","",$B$7-1)', NUM, is_input=False)
srow(10, "Reporting month as a date",
     f'=IFERROR(INDEX(Lists!$W${MTH_R0}:$W${MTH_R1},MATCH($B$5,Lists!$V${MTH_R0}:$V${MTH_R1},0)),"")',
     DATEF, is_input=False)
srow(11, "Comparative month",
     '=IF($B$10="","",TEXT(DATE(YEAR($B$10)-1,MONTH($B$10),1),"MMM YYYY"))', is_input=False)
srow(12, "Reporting currency", "AUD")
dv = DataValidation(type="list", formula1=f"=Lists!$V${MTH_R0}:$V${MTH_R1}", allow_blank=False,
                    showErrorMessage=True, errorTitle="Pick a month", error="Choose from the list.")
st.add_data_validation(dv); dv.add(st["B5"])

band(st, 14, "2.  RISK FLAG THRESHOLDS", 1, 3)
srow(15, "Materiality floor  ($)", 5000, MONEY,
     note="Below this a variance is always Low, so small accounts cannot cry wolf.")
srow(16, "Medium risk threshold  (% variance)", 0.10, PCT)
srow(17, "High risk threshold  (% variance)", 0.25, PCT)
srow(18, "High risk threshold  ($ variance)", 50000, MONEY)

band(st, 20, "3.  DATA STATUS  (calculated)", 1, 3)
stat = [
    ("GL lines pasted", f'=SUMPRODUCT(--(Cleanup!$A${CL_R0}:$A${CL_R1}<>""))', NUM),
    ("GL capacity", f'={GL_R1-GL_R0+1}', NUM),
    ("Lines kept for reporting", f'=SUM(Cleanup!$I${CL_R0}:$I${CL_R1})', NUM),
    ("Earliest date", f'=IF($B$21=0,"",MIN(Cleanup!$A${CL_R0}:$A${CL_R1}))', DATEF),
    ("Latest date", f'=IF($B$21=0,"",MAX(Cleanup!$A${CL_R0}:$A${CL_R1}))', DATEF),
    ("Lines with no account match", f'=COUNTIF(Cleanup!$Q${CL_R0}:$Q${CL_R1},"NO ACCOUNT")', NUM),
    ("Lines with no department", f'=COUNTIF(Cleanup!$Q${CL_R0}:$Q${CL_R1},"NO DEPARTMENT")', NUM),
    ("Lines outside the two financial years", f'=COUNTIF(Cleanup!$Q${CL_R0}:$Q${CL_R1},"OUTSIDE FY")', NUM),
    ("Total debits", f'=SUM(GL_Paste!$G${GL_R0}:$G${GL_R1})', MONEY2),
    ("Total credits", f'=SUM(GL_Paste!$H${GL_R0}:$H${GL_R1})', MONEY2),
    ("Debits less credits (must be nil)", '=ROUND($B$29-$B$30,2)', MONEY2),
]
for i, (t, f, fm) in enumerate(stat):
    srow(21 + i, t, f, fm, is_input=False)
for cell, bad in (("B31", 'ROUND($B$31,2)<>0'), ("B26", '$B$26>0'), ("B27", '$B$27>0')):
    st.conditional_formatting.add(cell, FormulaRule(formula=[bad],
        fill=PatternFill("solid", fgColor=F_HIGH), font=Font(color=T_HIGH, bold=True)))
    st.conditional_formatting.add(cell, FormulaRule(formula=[f"NOT({bad})"],
        fill=PatternFill("solid", fgColor=F_LOW), font=Font(color=T_LOW, bold=True)))

B5, B6, B7, B8, B9, B10 = ("Setup!$B$5", "Setup!$B$6", "Setup!$B$7", "Setup!$B$8",
                           "Setup!$B$9", "Setup!$B$10")
MAT, MEDP, HIP, HID = "Setup!$B$15", "Setup!$B$16", "Setup!$B$17", "Setup!$B$18"
print("lists + setup built")

# ===========================================================================
# PASTE SHEETS  - the only three places data goes in
# ===========================================================================
gl = sheet("GL_Paste", "C00000")
gl.sheet_view.showGridLines = True
for i in range(1, 21): gl.column_dimensions[CL(i)].width = 16
# Bare on purpose: Excel will not paste a whole-sheet selection anywhere but A1.
for j, h in enumerate(["Index", "Account Code", "Account Name", "Source", "Date", "Contact",
                       "Debit", "Credit", "Job Numbers", "Invoice Number", "Reference",
                       "Description", "Cost Centres", "Amount", "Fiscal Year", "Month",
                       "Quarter", "Account Type"]):
    c = gl.cell(row=1, column=1 + j, value=h)
    c.font = Font(name=ARIAL, size=9, bold=True, color=WHITE); c.fill = HEAD; c.border = BOX

bp = sheet("Budget_Paste", "C55A11")
title(bp, "Budget",
      "Loaded from CTS_Budget_FY27_Final.xlsx. One row per financial year, department and line, months "
      "July to June across D to O. Your budget names three departments; the GL has six, so they are "
      "mapped - see the note below. Overwrite any figure; nothing here is locked.")
lbl(bp, 3, 1, "Support = Onsite  |  Consulting = Integration  |  Production covers Production AND "
              "Video, so Video has no budget line of its own  |  CTS Overheads sits against CTS",
    size=9, color="C00000", bold=True)
bp.merge_cells(start_row=3, start_column=1, end_row=3, end_column=3)
header(bp, 4, ["FY (year it ends)", "Department", "Line"] +
       [_dt.date(2000, (6 + i) % 12 + 1, 1).strftime("%b") for i in range(12)])
BUD_R0, BUD_R1 = 5, 84
BUD_LINES = ["Income", "Cost of Sales", "Direct Expenses"]
seed = [(BUDGET["fy"], b["dept"], b["line"], b["months"]) for b in BUDGET["rows"]]
seeded = {(b["dept"], b["line"]) for b in BUDGET["rows"]}
seed += [(BUDGET["fy"], d, l, None) for d in DEPTS for l in BUD_LINES if (d, l) not in seeded]
for i in range(BUD_R1 - BUD_R0 + 1):
    r = BUD_R0 + i
    row = seed[i] if i < len(seed) else (None, None, None, None)
    for j in range(15):
        v = row[j] if j < 3 else (row[3][j - 3] if row[3] else None)
        c = bp.cell(row=r, column=1 + j, value=v)
        c.font = Font(name=ARIAL, size=9, color=BLUE_FONT); c.fill = INPUT_FILL; c.border = BOX
        if j >= 3: c.number_format = MONEY
widths(bp, {"A": 16, "B": 16, "C": 18})
for i in range(4, 16): bp.column_dimensions[CL(i)].width = 13
for f1, rng in ((f"=Lists!$L${DEP_R0}:$L${DEP_R1}", f"B{BUD_R0}:B{BUD_R1}"),
                ('"Income,Cost of Sales,Direct Expenses"', f"C{BUD_R0}:C{BUD_R1}")):
    d = DataValidation(type="list", formula1=f1, allow_blank=True, showErrorMessage=False)
    bp.add_data_validation(d); d.add(rng)

hp = sheet("Hours_Paste", "7030A0")
title(hp, "Utilisation  -  paste the hours here",
      "One row per financial year, department and hour type. Same month layout as the budget. "
      "Take it from the Job Group table on the Dashboard tab of your utilisation report.")
header(hp, 4, ["FY (year it ends)", "Department", "Hour Type"] +
       [_dt.date(2000, (6 + i) % 12 + 1, 1).strftime("%b") for i in range(12)])
HRS_R0, HRS_R1 = 5, 84
HRS_TYPES = ["Chargeable", "Non-chargeable", "Leave"]
seed_h = [(2027, d, t) for d in DEPTS for t in HRS_TYPES]
for i in range(HRS_R1 - HRS_R0 + 1):
    r = HRS_R0 + i
    vals = seed_h[i] if i < len(seed_h) else (None, None, None)
    for j in range(15):
        c = hp.cell(row=r, column=1 + j, value=(vals[j] if j < 3 else None))
        c.font = Font(name=ARIAL, size=9, color=BLUE_FONT); c.fill = INPUT_FILL; c.border = BOX
        if j >= 3: c.number_format = HRS
widths(hp, {"A": 16, "B": 16, "C": 18})
for i in range(4, 16): hp.column_dimensions[CL(i)].width = 13
for f1, rng in ((f"=Lists!$L${DEP_R0}:$L${DEP_R1}", f"B{HRS_R0}:B{HRS_R1}"),
                ('"Chargeable,Non-chargeable,Leave"', f"C{HRS_R0}:C{HRS_R1}")):
    d = DataValidation(type="list", formula1=f1, allow_blank=True, showErrorMessage=False)
    hp.add_data_validation(d); d.add(rng)

# ===========================================================================
# CLEANUP  - one row per pasted GL line, turned into numeric keys
# ===========================================================================
cu = sheet("Cleanup", "404040")
title(cu, "Cleanup  -  what the workbook made of the GL paste",
      "Calculated. Department comes from the Cost Centres column, falling back to the job number. "
      "Columns J to N are numeric match keys - that is what makes the reports fast.")
header(cu, 5, ["Date", "Amount", "Acct #", "Dept #", "Sub #", "Type #", "Period", "FY off",
               "Keep", "Key: acct x period", "Key: acct x dept (month)",
               "Key: acct x dept (YTD)", "Key: sub x dept", "Key: type x dept", "Status"])
cu.freeze_panes = "A6"
widths(cu, {"A": 11, "B": 13, "C": 8, "D": 8, "E": 8, "F": 8, "G": 8, "H": 8, "I": 7,
            "J": 15, "K": 17, "L": 17, "M": 15, "N": 15, "O": 18})

for i in range(CL_R1 - CL_R0 + 1):
    r = CL_R0 + i
    g = GL_R0 + i
    jobno = f'TRIM(LEFT(GL_Paste!$I{g}&" ",FIND(" ",GL_Paste!$I{g}&" ")-1))'
    fml = {
        1: f'=IF(GL_Paste!$E{g}="","",GL_Paste!$E{g})',
        2: f'=IF($A{r}="",0,IFERROR(N(GL_Paste!$H{g}),0)-IFERROR(N(GL_Paste!$G{g}),0))',
        3: (f'=IF($A{r}="",0,IFERROR(MATCH(TEXT(GL_Paste!$B{g},"0"),{L_ACC_CODE},0),'
            f'IFERROR(MATCH(GL_Paste!$C{g},{L_ACC_NAME},0),0)))'),
        4: (f'=IF($A{r}="",0,IFERROR(MATCH(UPPER(TRIM(GL_Paste!$M{g}&"")),{L_DEP_CC},0),'
            f'IFERROR(MATCH(INDEX({L_JOB_DEP},MATCH({jobno},{L_JOB},0)),{L_DEP_NAME},0),0)))'),
        5: f'=IF($C{r}=0,0,INDEX({L_ACC_SUB},$C{r}))',
        6: f'=IF($C{r}=0,0,INDEX({L_ACC_TYP},$C{r}))',
        7: f'=IF($A{r}="",0,MOD(MONTH($A{r})-{B6}+12,12)+1)',
        8: (f'=IF($A{r}="",-1,IF(IF({B6}=1,YEAR($A{r}),YEAR($A{r})+IF(MONTH($A{r})>={B6},1,0))={B7},1,'
            f'IF(IF({B6}=1,YEAR($A{r}),YEAR($A{r})+IF(MONTH($A{r})>={B6},1,0))={B9},0,-1)))'),
        9: f'=IF(AND($A{r}<>"",$C{r}>0,$D{r}>0,$H{r}>=0),1,0)',
        10: f'=IF($I{r}=0,0,$C{r}*1000+$H{r}*100+$G{r})',
        11: f'=IF(AND($I{r}=1,$H{r}=1,$G{r}={B8}),$C{r}*100+$D{r},0)',
        12: f'=IF(AND($I{r}=1,$G{r}<={B8}),$C{r}*1000+$D{r}*10+$H{r},0)',
        13: f'=IF($I{r}=0,0,$E{r}*100000+$D{r}*10000+$H{r}*1000+$G{r})',
        14: f'=IF($I{r}=0,0,$F{r}*100000+$D{r}*10000+$H{r}*1000+$G{r})',
        15: (f'=IF($A{r}="","",IF($C{r}=0,"NO ACCOUNT",IF($D{r}=0,"NO DEPARTMENT",'
             f'IF($H{r}=-1,"OUTSIDE FY","OK"))))'),
    }
    for col, f in fml.items():
        c = cu.cell(row=r, column=col, value=f)
        c.font = Font(name=ARIAL, size=9)
        if col == 1: c.number_format = DATEF
        elif col == 2: c.number_format = MONEY2
        elif col != 15: c.number_format = NUM
cu.conditional_formatting.add(f"A{CL_R0}:O{CL_R1}", FormulaRule(
    formula=[f'AND($A{CL_R0}<>"",$O{CL_R0}<>"OK")'],
    fill=PatternFill("solid", fgColor=F_HIGH), font=Font(color=T_HIGH)))

CU_KM = f"Cleanup!$K${CL_R0}:$K${CL_R1}"
CU_KY = f"Cleanup!$L${CL_R0}:$L${CL_R1}"
CU_K1 = f"Cleanup!$J${CL_R0}:$J${CL_R1}"
CU_K5 = f"Cleanup!$M${CL_R0}:$M${CL_R1}"
CU_K6 = f"Cleanup!$N${CL_R0}:$N${CL_R1}"
CU_AMT = f"Cleanup!$B${CL_R0}:$B${CL_R1}"
print("paste sheets + cleanup built")

# ===========================================================================
# ENGINE  - account level. Scanned once, every report reads from here.
# ===========================================================================
en = sheet("Engine", "404040")
title(en, "Engine  -  account by period and by department",
      "Calculated. Never type here. Income is positive and costs are negative, the same convention "
      "as your existing PL Analysis pack, so every total is a contribution.")
E_R0 = 6
E_R1 = E_R0 + NACC - 1
PY_C0, CY_C0 = 6, 18
D = {"CY_M": 30, "PY_M": 31, "CY_P": 32, "CY_YTD": 33, "PY_YTD": 34, "PY_FY": 35, "ACT": 36}
DEPT_C0 = 37                                  # 6 depts x 3 measures
FLAG_C0 = DEPT_C0 + len(DEPTS) * 3            # 55

def mhdr(fy_ref, p):
    return (f'=TEXT(DATE({fy_ref}-IF(AND({B6}>1,MOD({B6}+{p}-2,12)+1>={B6}),1,0),'
            f'MOD({B6}+{p}-2,12)+1,1),"MMM-YY")')

for p in range(1, 13):
    for c0, fy_ref in ((PY_C0, B9), (CY_C0, B7)):
        col = c0 + p - 1
        en.cell(row=3, column=col, value=p).font = Font(name=ARIAL, size=8, color=MUTED)
        h = en.cell(row=5, column=col, value=mhdr(fy_ref, p))
        h.font = Font(name=ARIAL, size=9, bold=True, color=WHITE)
        h.fill = HEAD if c0 == CY_C0 else SUBHEAD
        h.border = BOX; h.alignment = Alignment(horizontal="center")
header(en, 5, ["Code", "Account", "Category", "Subcategory", "Type"], size=9)
header(en, 5, ["CY Month", "PY Month", "CY Prior", "CY YTD", "PY YTD", "PY FY", "Active"],
       start_col=D["CY_M"], size=9)
for di, dep in enumerate(DEPTS):
    header(en, 5, [f"{dep}\nMonth", f"{dep}\nCY YTD", f"{dep}\nPY YTD"],
           start_col=DEPT_C0 + di * 3, size=8)
header(en, 5, ["MoM Var", "MoM Flag", "YoY Var", "YoY Flag", "MoM key", "YoY key"],
       start_col=FLAG_C0, size=9)
en.freeze_panes = "F6"
widths(en, {"A": 10, "B": 34, "C": 15, "D": 26, "E": 24})
for i in range(PY_C0, FLAG_C0 + 6): en.column_dimensions[CL(i)].width = 12

def flag(r, cur, pri, var):
    return (f'=IF($B{r}="","",IF(AND(ROUND(${cur}{r},2)=0,ROUND(${pri}{r},2)=0),"No Activity",'
            f'IF(ABS(${var}{r})<{MAT},"Low",IF(ROUND(${pri}{r},2)=0,"High",'
            f'IF(OR(ABS(${var}{r}/${pri}{r})>={HIP},ABS(${var}{r})>={HID}),"High",'
            f'IF(ABS(${var}{r}/${pri}{r})>={MEDP},"Medium","Low"))))))')

for i, a in enumerate(ACCTS):
    r = E_R0 + i
    idx = i + 1
    for col, v in ((1, a["code"]), (2, a["name"]), (3, a["category"]),
                   (4, a["subcategory"]), (5, a["type"])):
        c = en.cell(row=r, column=col, value=v)
        c.font = Font(name=ARIAL, size=9); c.border = BOX
    for p in range(1, 13):
        for c0, off in ((PY_C0, 0), (CY_C0, 1)):
            c = en.cell(row=r, column=c0 + p - 1,
                        value=f'=SUMIF({CU_K1},{idx*1000 + off*100 + p},{CU_AMT})')
            c.font = Font(name=ARIAL, size=9); c.number_format = MONEY; c.border = BOX
    ser = f"$F{r}:$AC{r}"
    der = {
        D["CY_M"]: f'=INDEX({ser},12+{B8})', D["PY_M"]: f'=INDEX({ser},{B8})',
        D["CY_P"]: f'=INDEX({ser},11+{B8})',
        D["CY_YTD"]: f'=SUMPRODUCT(($R$3:$AC$3<={B8})*$R{r}:$AC{r})',
        D["PY_YTD"]: f'=SUMPRODUCT(($F$3:$Q$3<={B8})*$F{r}:$Q{r})',
        D["PY_FY"]: f'=SUM($F{r}:$Q{r})',
        D["ACT"]: f'=IF(SUMPRODUCT(ABS($F{r}:$AC{r}))>0,1,0)',
    }
    for col, f in der.items():
        c = en.cell(row=r, column=col, value=f)
        c.font = Font(name=ARIAL, size=9, bold=(col in (D["CY_M"], D["CY_YTD"])))
        c.number_format = NUM if col == D["ACT"] else MONEY
        c.border = BOX; c.fill = TOTAL_FILL
    for di in range(len(DEPTS)):
        dn = di + 1
        for k, key, rng in ((0, idx * 100 + dn, CU_KM),
                            (1, idx * 1000 + dn * 10 + 1, CU_KY),
                            (2, idx * 1000 + dn * 10 + 0, CU_KY)):
            c = en.cell(row=r, column=DEPT_C0 + di * 3 + k,
                        value=f'=SUMIF({rng},{key},{CU_AMT})')
            c.font = Font(name=ARIAL, size=9); c.number_format = MONEY; c.border = BOX
    F = {k: CL(v) for k, v in D.items()}
    fl = {
        FLAG_C0: f'=${F["CY_M"]}{r}-${F["CY_P"]}{r}',
        FLAG_C0 + 1: flag(r, F["CY_M"], F["CY_P"], CL(FLAG_C0)),
        FLAG_C0 + 2: f'=${F["CY_YTD"]}{r}-${F["PY_YTD"]}{r}',
        FLAG_C0 + 3: flag(r, F["CY_YTD"], F["PY_YTD"], CL(FLAG_C0 + 2)),
        FLAG_C0 + 4: (f'=IF(ROUND(ABS(${CL(FLAG_C0)}{r}),0)=0,-1E+15,'
                      f'ABS(${CL(FLAG_C0)}{r})+ROW()/1000000)'),
        FLAG_C0 + 5: (f'=IF(ROUND(ABS(${CL(FLAG_C0+2)}{r}),0)=0,-1E+15,'
                      f'ABS(${CL(FLAG_C0+2)}{r})+ROW()/1000000)'),
    }
    for col, f in fl.items():
        c = en.cell(row=r, column=col, value=f)
        c.font = Font(name=ARIAL, size=9); c.border = BOX; c.fill = TOTAL_FILL
        c.number_format = MONEY if col in (FLAG_C0, FLAG_C0 + 2, FLAG_C0 + 4, FLAG_C0 + 5) else "General"
        if col in (FLAG_C0 + 1, FLAG_C0 + 3): c.alignment = Alignment(horizontal="center")

def ecol(key): return f"Engine!${CL(D[key])}${E_R0}:${CL(D[key])}${E_R1}"
E_CAT = f"Engine!$C${E_R0}:$C${E_R1}"
E_SUB = f"Engine!$D${E_R0}:$D${E_R1}"
E_NAME = f"Engine!$B${E_R0}:$B${E_R1}"
def edept(di, k): return f"Engine!${CL(DEPT_C0+di*3+k)}${E_R0}:${CL(DEPT_C0+di*3+k)}${E_R1}"

# ===========================================================================
# ENGINE_DEPT  - subcategory x department and type x department, by period
# ===========================================================================
ed = sheet("Engine_Dept", "404040")
title(ed, "Engine  -  subcategory and cost type by department, by period", "Calculated. Never type here.")
ED_R0 = 6
rows_spec = ([("Sub", s, c, di + 1, si + 1) for si, (c, s) in enumerate(SUBS) for di in range(len(DEPTS))] +
             [("Type", t, "", di + 1, ti + 1) for ti, t in enumerate(TYPES) for di in range(len(DEPTS))])
EDP_C0 = 6
for p in range(1, 13):
    for c0, fy_ref in ((EDP_C0, B9), (EDP_C0 + 12, B7)):
        col = c0 + p - 1
        ed.cell(row=3, column=col, value=p).font = Font(name=ARIAL, size=8, color=MUTED)
        h = ed.cell(row=5, column=col, value=mhdr(fy_ref, p))
        h.font = Font(name=ARIAL, size=9, bold=True, color=WHITE)
        h.fill = HEAD if c0 > EDP_C0 else SUBHEAD
        h.border = BOX; h.alignment = Alignment(horizontal="center")
header(ed, 5, ["Kind", "Name", "Category", "Dept", "Dept #"], size=9)
ED_D = {"CY_M": 30, "PY_M": 31, "CY_P": 32, "CY_YTD": 33, "PY_YTD": 34, "PY_FY": 35}
header(ed, 5, ["CY Month", "PY Month", "CY Prior", "CY YTD", "PY YTD", "PY FY"],
       start_col=ED_D["CY_M"], size=9)
ed.freeze_panes = "F6"
widths(ed, {"A": 8, "B": 30, "C": 15, "D": 14, "E": 7})
for i in range(EDP_C0, ED_D["PY_FY"] + 1): ed.column_dimensions[CL(i)].width = 12

for i, (kind, name, cat, dn, ix) in enumerate(rows_spec):
    r = ED_R0 + i
    rng = CU_K5 if kind == "Sub" else CU_K6
    for col, v in ((1, kind), (2, name), (3, cat), (4, DEPTS[dn - 1]), (5, dn)):
        c = ed.cell(row=r, column=col, value=v)
        c.font = Font(name=ARIAL, size=9); c.border = BOX
    for p in range(1, 13):
        for c0, off in ((EDP_C0, 0), (EDP_C0 + 12, 1)):
            key = ix * 100000 + dn * 10000 + off * 1000 + p
            c = ed.cell(row=r, column=c0 + p - 1, value=f'=SUMIF({rng},{key},{CU_AMT})')
            c.font = Font(name=ARIAL, size=9); c.number_format = MONEY; c.border = BOX
    ser = f"$F{r}:$AC{r}"
    for col, f in ((ED_D["CY_M"], f'=INDEX({ser},12+{B8})'), (ED_D["PY_M"], f'=INDEX({ser},{B8})'),
                   (ED_D["CY_P"], f'=INDEX({ser},11+{B8})'),
                   (ED_D["CY_YTD"], f'=SUMPRODUCT(($R$3:$AC$3<={B8})*$R{r}:$AC{r})'),
                   (ED_D["PY_YTD"], f'=SUMPRODUCT(($F$3:$Q$3<={B8})*$F{r}:$Q{r})'),
                   (ED_D["PY_FY"], f'=SUM($F{r}:$Q{r})')):
        c = ed.cell(row=r, column=col, value=f)
        c.font = Font(name=ARIAL, size=9, bold=True); c.number_format = MONEY
        c.border = BOX; c.fill = TOTAL_FILL
ED_R1 = ED_R0 + len(rows_spec) - 1
ED_KIND = f"Engine_Dept!$A${ED_R0}:$A${ED_R1}"
ED_NAME = f"Engine_Dept!$B${ED_R0}:$B${ED_R1}"
ED_CAT = f"Engine_Dept!$C${ED_R0}:$C${ED_R1}"
ED_DEPT = f"Engine_Dept!$D${ED_R0}:$D${ED_R1}"
def edcol(k): return f"Engine_Dept!${CL(ED_D[k])}${ED_R0}:${CL(ED_D[k])}${ED_R1}"
def edper(off, p): return f"Engine_Dept!${CL(EDP_C0+off*12+p-1)}${ED_R0}:${CL(EDP_C0+off*12+p-1)}${ED_R1}"
print(f"engines built: Engine {NACC} rows, Engine_Dept {len(rows_spec)} rows")

# period numbers above the budget / hours grids so YTD can be summed by formula
for sh in (bp, hp):
    for p in range(1, 13):
        c = sh.cell(row=3, column=3 + p, value=p)
        c.font = Font(name=ARIAL, size=8, color=MUTED)
        c.alignment = Alignment(horizontal="center")
BP_FY = f"Budget_Paste!$A${BUD_R0}:$A${BUD_R1}"
BP_DEP = f"Budget_Paste!$B${BUD_R0}:$B${BUD_R1}"
BP_LINE = f"Budget_Paste!$C${BUD_R0}:$C${BUD_R1}"
BP_GRID = f"Budget_Paste!$D${BUD_R0}:$O${BUD_R1}"
BP_PER = "Budget_Paste!$D$3:$O$3"
HP_FY = f"Hours_Paste!$A${HRS_R0}:$A${HRS_R1}"
HP_DEP = f"Hours_Paste!$B${HRS_R0}:$B${HRS_R1}"
HP_TYPE = f"Hours_Paste!$C${HRS_R0}:$C${HRS_R1}"
HP_GRID = f"Hours_Paste!$D${HRS_R0}:$O${HRS_R1}"
HP_PER = "Hours_Paste!$D$3:$O$3"

def budget(dept, line, ytd=True, fy=B7):
    per = f'({BP_PER}<={B8})' if ytd else f'({BP_PER}={B8})'
    return (f'=SUMPRODUCT(({BP_FY}={fy})*({BP_DEP}="{dept}")*({BP_LINE}="{line}")'
            f'*{per}*{BP_GRID})')

def hours(dept, htype, ytd=True, fy=B7):
    per = f'({HP_PER}<={B8})' if ytd else f'({HP_PER}={B8})'
    return (f'=SUMPRODUCT(({HP_FY}={fy})*({HP_DEP}="{dept}")*({HP_TYPE}="{htype}")'
            f'*{per}*{HP_GRID})')

def dsum(dept, cat, key):
    return f'=SUMIFS({edcol(key)},{ED_KIND},"Sub",{ED_DEPT},"{dept}",{ED_CAT},"{cat}")'

def csum(cat, key):
    return f'=SUMIF({E_CAT},"{cat}",{ecol(key)})'

# ===========================================================================
# FINANCIAL_SUMMARY  - the page you read first
# ===========================================================================
fs = sheet("Financial_Summary", "C55A11")
title(fs, "Financial Summary")
fs["A2"] = ('=IF(Setup!$B$4="","",Setup!$B$4&"   |   "&Setup!$B$5&"   |   FY"&Setup!$B$7&" vs FY"'
            '&Setup!$B$9&"   |   "&Setup!$B$12&"   |   income positive, costs negative")')
fs["A2"].font = Font(name=ARIAL, size=10, bold=True, color=MUTED)
widths(fs, {"A": 34})
for i in range(2, 14): fs.column_dimensions[CL(i)].width = 14

r = 4
band(fs, r, "COMPANY PROFIT & LOSS", 1, 12); r += 1
header(fs, r, ["", "CY Month", "PY Month", "Var $", "Var %", "CY YTD", "PY YTD", "Var $",
               "Var %", "Budget YTD", "vs Budget", "PY Full Year"]); r += 1

def money_row(row, label, base, bold=False, pct_of=None, budget_line=None, indent=0):
    c = lbl(fs, row, 1, label, bold=bold, size=10, indent=indent)
    vals = {2: base["CY_M"], 3: base["PY_M"], 6: base["CY_YTD"], 7: base["PY_YTD"], 12: base["PY_FY"]}
    for col, f in vals.items():
        cc = fs.cell(row=row, column=col, value=f)
        cc.font = Font(name=ARIAL, size=10, bold=bold); cc.number_format = MONEY; cc.border = BOX
    for col, f, fmt in ((4, f'=$B{row}-$C{row}', MONEY),
                        (5, f'=IF(ROUND($C{row},2)=0,"n/a",$D{row}/ABS($C{row}))', PCT),
                        (8, f'=$F{row}-$G{row}', MONEY),
                        (9, f'=IF(ROUND($G{row},2)=0,"n/a",$H{row}/ABS($G{row}))', PCT),
                        (10, budget_line or '=""', MONEY),
                        (11, f'=IF(N($J{row})=0,"n/a",$F{row}-N($J{row}))', MONEY)):
        cc = fs.cell(row=row, column=col, value=f)
        cc.font = Font(name=ARIAL, size=10, bold=bold); cc.number_format = fmt; cc.border = BOX
    if bold:
        for col in range(1, 13): fs.cell(row=row, column=col).fill = TOTAL_FILL
    return row + 1

def cat_base(cat): return {k: csum(cat, k) for k in ("CY_M", "PY_M", "CY_YTD", "PY_YTD", "PY_FY")}
def calc_base(expr, rowmap):
    out = {}
    for k, col in (("CY_M", "B"), ("PY_M", "C"), ("CY_YTD", "F"), ("PY_YTD", "G"), ("PY_FY", "L")):
        e = expr
        for name, rr in rowmap.items(): e = e.replace("{" + name + "}", f"{col}{rr}")
        out[k] = "=" + e
    return out

R = {}
R["Income"] = r
r = money_row(r, "Income", cat_base("Income"), indent=1,
              budget_line=f'=SUMPRODUCT(({BP_FY}={B7})*({BP_LINE}="Income")*({BP_PER}<={B8})*{BP_GRID})')
R["Cost of Sales"] = r
r = money_row(r, "Cost of Sales", cat_base("Cost of Sales"), indent=1,
              budget_line=f'=-SUMPRODUCT(({BP_FY}={B7})*({BP_LINE}="Cost of Sales")*({BP_PER}<={B8})*{BP_GRID})')
R["Gross Profit"] = r
r = money_row(r, "Gross Profit", calc_base("{Income}+{Cost of Sales}", R), bold=True)
gp_row = R["Gross Profit"]
gmr = r
lbl(fs, r, 1, "Gross Margin %", size=10, indent=1)
for col, num, den in ((2, "B", "B"), (3, "C", "C"), (6, "F", "F"), (7, "G", "G"), (12, "L", "L")):
    cc = fs.cell(row=r, column=col,
                 value=f'=IF(ROUND({den}{R["Income"]},2)=0,"n/a",{num}{gp_row}/{den}{R["Income"]})')
    cc.font = Font(name=ARIAL, size=10); cc.number_format = PCT; cc.border = BOX
for col, f in ((4, f'=IF(OR(ISTEXT($B{r}),ISTEXT($C{r})),"n/a",($B{r}-$C{r})*100)'),
               (8, f'=IF(OR(ISTEXT($F{r}),ISTEXT($G{r})),"n/a",($F{r}-$G{r})*100)')):
    cc = fs.cell(row=r, column=col, value=f); cc.number_format = PP
    cc.font = Font(name=ARIAL, size=10); cc.border = BOX
r += 1
for cat in ("Other Income", "Expenses", "Other Expenses"):
    R[cat] = r
    bl = (f'=-SUMPRODUCT(({BP_FY}={B7})*({BP_LINE}="Direct Expenses")*({BP_PER}<={B8})*{BP_GRID})'
          if cat == "Expenses" else None)
    r = money_row(r, cat, cat_base(cat), indent=1, budget_line=bl)
R["Net Profit"] = r
r = money_row(r, "Net Profit", calc_base(
    "{Gross Profit}+{Other Income}+{Expenses}+{Other Expenses}", R), bold=True)
npr = R["Net Profit"]
lbl(fs, r, 1, "Net Margin %", size=10, indent=1)
for col in ("B", "C", "F", "G", "L"):
    cc = fs.cell(row=r, column=(ord(col) - 64),
                 value=f'=IF(ROUND({col}{R["Income"]},2)=0,"n/a",{col}{npr}/{col}{R["Income"]})')
    cc.font = Font(name=ARIAL, size=10); cc.number_format = PCT; cc.border = BOX
r += 2

# ---- the bit they asked for: gross margin per department ------------------
band(fs, r, "GROSS MARGIN BY DEPARTMENT", 1, 12); r += 1
header(fs, r, ["Department", "Revenue\nMonth", "COGS\nMonth", "GP\nMonth", "GM %\nMonth",
               "Revenue\nYTD", "COGS\nYTD", "GP\nYTD", "GM %\nYTD", "GP\nPY YTD",
               "GP Var\nYTD", "GP Budget\nYTD"]); r += 1
DEPT_GM_R0 = r
for dep in DEPTS:
    lbl(fs, r, 1, dep, bold=True, size=10)
    cells = {
        2: dsum(dep, "Income", "CY_M"), 3: dsum(dep, "Cost of Sales", "CY_M"),
        4: f'=$B{r}+$C{r}', 5: f'=IF(ROUND($B{r},2)=0,"n/a",$D{r}/$B{r})',
        6: dsum(dep, "Income", "CY_YTD"), 7: dsum(dep, "Cost of Sales", "CY_YTD"),
        8: f'=$F{r}+$G{r}', 9: f'=IF(ROUND($F{r},2)=0,"n/a",$H{r}/$F{r})',
        10: f'={dsum(dep, "Income", "PY_YTD")[1:]}+{dsum(dep, "Cost of Sales", "PY_YTD")[1:]}',
        11: f'=$H{r}-$J{r}',
        12: f'={budget(dep, "Income")[1:]}-{budget(dep, "Cost of Sales")[1:]}',
    }
    for col, f in cells.items():
        cc = fs.cell(row=r, column=col, value=f)
        cc.font = Font(name=ARIAL, size=10, bold=(col in (4, 8)))
        cc.number_format = PCT if col in (5, 9) else MONEY
        cc.border = BOX
    r += 1
DEPT_GM_R1 = r - 1
lbl(fs, r, 1, "TOTAL", bold=True, size=10)
for col in (2, 3, 4, 6, 7, 8, 10, 11, 12):
    cc = fs.cell(row=r, column=col, value=f'=SUM({CL(col)}{DEPT_GM_R0}:{CL(col)}{DEPT_GM_R1})')
    cc.font = Font(name=ARIAL, size=10, bold=True); cc.number_format = MONEY
    cc.border = BOX; cc.fill = TOTAL_FILL
for col, num, den in ((5, 4, 2), (9, 8, 6)):
    cc = fs.cell(row=r, column=col,
                 value=f'=IF(ROUND({CL(den)}{r},2)=0,"n/a",{CL(num)}{r}/{CL(den)}{r})')
    cc.font = Font(name=ARIAL, size=10, bold=True); cc.number_format = PCT
    cc.border = BOX; cc.fill = TOTAL_FILL
lbl(fs, r, 1, "TOTAL", bold=True, size=10).fill = TOTAL_FILL
fs.cell(row=r, column=1).fill = TOTAL_FILL
GM_TOT_R = r
r += 2
for rr in range(DEPT_GM_R0, DEPT_GM_R1 + 1):
    fs.conditional_formatting.add(f"E{rr}", FormulaRule(formula=[f'AND(ISNUMBER($E{rr}),$E{rr}<0.15)'],
        fill=PatternFill("solid", fgColor=F_HIGH), font=Font(color=T_HIGH, bold=True)))
    fs.conditional_formatting.add(f"I{rr}", FormulaRule(formula=[f'AND(ISNUMBER($I{rr}),$I{rr}<0.15)'],
        fill=PatternFill("solid", fgColor=F_HIGH), font=Font(color=T_HIGH, bold=True)))
    fs.conditional_formatting.add(f"I{rr}", FormulaRule(formula=[f'AND(ISNUMBER($I{rr}),$I{rr}>=0.35)'],
        fill=PatternFill("solid", fgColor=F_LOW), font=Font(color=T_LOW, bold=True)))
print("financial summary built")

def tsum(dept, typ, key):
    return f'=SUMIFS({edcol(key)},{ED_KIND},"Type",{ED_DEPT},"{dept}",{ED_NAME},"{typ}")'

band(fs, r, "COST MIX  -  LABOUR vs EQUIPMENT  (year to date)", 1, 12); r += 1
header(fs, r, ["Department"] + TYPES + ["Total Cost", "Labour %"], start_col=1); r += 1
MIX_R0 = r
for dep in DEPTS:
    lbl(fs, r, 1, dep, bold=True, size=10)
    for k, t in enumerate(TYPES):
        cc = fs.cell(row=r, column=2 + k, value=tsum(dep, t, "CY_YTD"))
        cc.font = Font(name=ARIAL, size=10); cc.number_format = MONEY; cc.border = BOX
    cc = fs.cell(row=r, column=6, value=f'=SUM($B{r}:$E{r})')
    cc.font = Font(name=ARIAL, size=10, bold=True); cc.number_format = MONEY; cc.border = BOX
    cc = fs.cell(row=r, column=7, value=f'=IF(ROUND($F{r},2)=0,"n/a",$B{r}/$F{r})')
    cc.font = Font(name=ARIAL, size=10); cc.number_format = PCT; cc.border = BOX
    r += 1
MIX_R1 = r - 1
r += 1

band(fs, r, "UTILISATION  -  hours against revenue  (year to date)", 1, 12); r += 1
header(fs, r, ["Department", "Chargeable\nhours", "Non-chargeable", "Leave", "Total hours",
               "Utilisation %", "Revenue YTD", "Revenue per\nchargeable hour",
               "Gross profit\nYTD", "GP per\nchargeable hour"], start_col=1); r += 1
UTIL_R0 = r
for i, dep in enumerate(DEPTS):
    gm = DEPT_GM_R0 + i
    lbl(fs, r, 1, dep, bold=True, size=10)
    cells = {
        2: hours(dep, "Chargeable"), 3: hours(dep, "Non-chargeable"), 4: hours(dep, "Leave"),
        5: f'=SUM($B{r}:$D{r})',
        6: f'=IF(ROUND($B{r}+$C{r},2)=0,"n/a",$B{r}/($B{r}+$C{r}))',
        7: f'=Financial_Summary!$F{gm}',
        8: f'=IF(ROUND($B{r},2)=0,"n/a",$G{r}/$B{r})',
        9: f'=Financial_Summary!$H{gm}',
        10: f'=IF(ROUND($B{r},2)=0,"n/a",$I{r}/$B{r})',
    }
    for col, f in cells.items():
        cc = fs.cell(row=r, column=col, value=f)
        cc.font = Font(name=ARIAL, size=10, bold=(col in (8, 10)))
        cc.number_format = (HRS if col in (2, 3, 4, 5) else PCT if col == 6 else MONEY)
        cc.border = BOX
    r += 1
UTIL_R1 = r - 1
r += 1

band(fs, r, "CONTROLLER REVIEW STATUS", 1, 12); r += 1
REV = [
    ("GL lines pasted", "=Setup!$B$21", "info"),
    ("Lines the workbook could not use", "=Setup!$B$26+Setup!$B$27+Setup!$B$28", "must"),
    ("Debits less credits (must be nil)", "=Setup!$B$31", "must"),
    ("High risk accounts - month on month",
     f'=COUNTIF(Engine!${CL(FLAG_C0+1)}${E_R0}:${CL(FLAG_C0+1)}${E_R1},"High")', "review"),
    ("High risk accounts - year on year",
     f'=COUNTIF(Engine!${CL(FLAG_C0+3)}${E_R0}:${CL(FLAG_C0+3)}${E_R1},"High")', "review"),
    ("Budget rows entered", f'=SUMPRODUCT(--({BP_DEP}<>""))', "info"),
    ("Hours rows entered", f'=SUMPRODUCT(--({HP_DEP}<>""))', "info"),
    ("Controls failing", '=COUNTIF(Controls!$E$6:$E$40,"FAIL")', "must"),
]
REV_R0 = r
for label, f, sev in REV:
    lbl(fs, r, 1, label, size=10)
    c = fs.cell(row=r, column=2, value=f)
    c.font = Font(name=ARIAL, size=10, bold=True); c.border = BOX
    c.number_format = MONEY2 if "Debits" in label else NUM
    c.alignment = Alignment(horizontal="center")
    sl = {"must": "Must be cleared", "review": "Review required"}.get(sev, "For information")
    s2 = fs.cell(row=r, column=3, value=f'=IF(ROUND($B{r},2)=0,"Clear","{sl}")')
    s2.font = Font(name=ARIAL, size=9, bold=True); s2.border = BOX
    s2.alignment = Alignment(horizontal="center")
    fs.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
    r += 1
REV_R1 = r - 1
for txt, fill, font in (("Must be cleared", F_HIGH, T_HIGH), ("Review required", F_MED, T_MED),
                        ("Clear", F_LOW, T_LOW)):
    fs.conditional_formatting.add(f"B{REV_R0}:E{REV_R1}", FormulaRule(
        formula=[f'$C{REV_R0}="{txt}"'], fill=PatternFill("solid", fgColor=fill),
        font=Font(color=font, bold=True)))

# ===========================================================================
# DEPT_PL  - full P&L per department, collapsible
# ===========================================================================
dp = sheet("Dept_PL", "2E75B6")
title(dp, "Profit & Loss by Department")
dp["A2"] = ('=IF(Setup!$B$4="","",Setup!$B$5&"   |   click the + and - buttons on the left to open a '
            'department, then a category, down to subcategory detail")')
dp["A2"].font = Font(name=ARIAL, size=9, italic=True, color=MUTED)
dp.sheet_properties.outlinePr = Outline(summaryBelow=False, summaryRight=False, applyStyles=False)
header(dp, 5, ["Line", "CY Month", "PY Month", "Var $", "Var %", "CY YTD", "PY YTD",
               "Var $", "Var %", "Budget YTD", "vs Budget", "PY Full Year", "Risk"])
dp.freeze_panes = "B6"
widths(dp, {"A": 42, "M": 11})
for i in range(2, 13): dp.column_dimensions[CL(i)].width = 14
CATS_IN_ORDER = ["Income", "Cost of Sales", "Expenses", "Other Income", "Other Expenses"]
r = 6
for dep in DEPTS:
    lbl(dp, r, 1, f"{dep.upper()}   (net contribution)", bold=True, size=11,
        color=WHITE)
    for col in range(1, 14): dp.cell(row=r, column=col).fill = SUBHEAD
    base = {k: f'=SUMIFS({edcol(k)},{ED_KIND},"Sub",{ED_DEPT},"{dep}")'
            for k in ("CY_M", "PY_M", "CY_YTD", "PY_YTD", "PY_FY")}
    dept_row = r
    for col, key in ((2, "CY_M"), (3, "PY_M"), (6, "CY_YTD"), (7, "PY_YTD"), (12, "PY_FY")):
        c = dp.cell(row=r, column=col, value=base[key])
        c.font = Font(name=ARIAL, size=10, bold=True, color=WHITE); c.number_format = MONEY
    r += 1
    for cat in CATS_IN_ORDER:
        lbl(dp, r, 1, f"    {cat}", bold=True, size=10)
        for col in range(1, 14): dp.cell(row=r, column=col).fill = BAND
        for col, key in ((2, "CY_M"), (3, "PY_M"), (6, "CY_YTD"), (7, "PY_YTD"), (12, "PY_FY")):
            c = dp.cell(row=r, column=col, value=dsum(dep, cat, key))
            c.font = Font(name=ARIAL, size=10, bold=True); c.number_format = MONEY
        if cat in ("Income", "Cost of Sales", "Expenses"):
            line = {"Income": "Income", "Cost of Sales": "Cost of Sales",
                    "Expenses": "Direct Expenses"}[cat]
            sgn = "" if cat == "Income" else "-"
            dp.cell(row=r, column=10, value=f'={sgn}{budget(dep, line)[1:]}').number_format = MONEY
        dp.row_dimensions[r].outlineLevel = 1
        dp.row_dimensions[r].collapsed = True
        r += 1
        for c2, sub in SUBS:
            if c2 != cat: continue
            lbl(dp, r, 1, f"        {sub}", size=9)
            for col, key in ((2, "CY_M"), (3, "PY_M"), (6, "CY_YTD"), (7, "PY_YTD"), (12, "PY_FY")):
                c = dp.cell(row=r, column=col,
                            value=f'=SUMIFS({edcol(key)},{ED_KIND},"Sub",{ED_DEPT},"{dep}",'
                                  f'{ED_NAME},"{sub}")')
                c.font = Font(name=ARIAL, size=9); c.number_format = MONEY
            dp.row_dimensions[r].outlineLevel = 2
            dp.row_dimensions[r].hidden = True
            r += 1
    for rr in range(dept_row, r):
        for col, f, fmt in ((4, f'=$B{rr}-$C{rr}', MONEY),
                            (5, f'=IF(ROUND($C{rr},2)=0,"n/a",$D{rr}/ABS($C{rr}))', PCT),
                            (8, f'=$F{rr}-$G{rr}', MONEY),
                            (9, f'=IF(ROUND($G{rr},2)=0,"n/a",$H{rr}/ABS($G{rr}))', PCT),
                            (11, f'=IF(ROUND($J{rr},2)=0,"n/a",$F{rr}-$J{rr})', MONEY),
                            (13, (f'=IF(AND(ROUND($F{rr},2)=0,ROUND($G{rr},2)=0),"No Activity",'
                                  f'IF(ABS($H{rr})<{MAT},"Low",IF(ROUND($G{rr},2)=0,"High",'
                                  f'IF(OR(ABS($H{rr}/$G{rr})>={HIP},ABS($H{rr})>={HID}),"High",'
                                  f'IF(ABS($H{rr}/$G{rr})>={MEDP},"Medium","Low")))))'), "General")):
            c = dp.cell(row=rr, column=col, value=f)
            c.number_format = fmt
            c.font = Font(name=ARIAL, size=(10 if rr == dept_row else 9),
                          bold=(rr == dept_row), color=(WHITE if rr == dept_row else BLACK_FONT))
            if col == 13: c.alignment = Alignment(horizontal="center")
    dp.row_dimensions[r].height = 5
    r += 1
DP_R1 = r - 1
for txt, fill, font in (("High", F_HIGH, T_HIGH), ("Medium", F_MED, T_MED),
                        ("Low", F_LOW, T_LOW), ("No Activity", F_NA, T_NA)):
    dp.conditional_formatting.add(f"M6:M{DP_R1}", FormulaRule(
        formula=[f'$M6="{txt}"'], fill=PatternFill("solid", fgColor=fill),
        font=Font(color=font, bold=True)))
print("dept P&L built")

# ===========================================================================
# BUDGET_VARIANCE
# ===========================================================================
bv = sheet("Budget_Variance", "7F6000")
title(bv, "Budget Variance by Department",
      "Actual comes from the GL paste, budget from the Budget sheet. Nothing is retyped, so the two "
      "cannot drift apart. Note: your budget covers Production and Video as one line, so Video shows "
      "actual against no budget and Production carries the budget for both. Read those two together, "
      "or split the Production budget on Budget_Paste.")
header(bv, 5, ["Department / Line", "Actual\nMonth", "Budget\nMonth", "Var $", "Var %",
               "Actual\nYTD", "Budget\nYTD", "Var $", "Var %", "Budget\nFull Year",
               "Actual + remaining\nbudget (forecast)", "Flag"])
bv.freeze_panes = "B6"
widths(bv, {"A": 34, "L": 11})
for i in range(2, 12): bv.column_dimensions[CL(i)].width = 15
BV_LINES = [("Income", "Income", "Income", 1), ("Cost of Sales", "Cost of Sales", "Cost of Sales", -1),
            ("Direct Expenses", "Expenses", "Direct Expenses", -1)]
r = 6
for dep in DEPTS:
    lbl(bv, r, 1, dep.upper(), bold=True, size=11, color=WHITE)
    for col in range(1, 13): bv.cell(row=r, column=col).fill = SUBHEAD
    dept_hdr = r; r += 1
    first = r
    for label, cat, bline, sgn in BV_LINES:
        lbl(bv, r, 1, f"    {label}", size=10)
        s = "" if sgn == 1 else "-"
        cells = {
            2: dsum(dep, cat, "CY_M"), 3: f'={s}{budget(dep, bline, ytd=False)[1:]}',
            6: dsum(dep, cat, "CY_YTD"), 7: f'={s}{budget(dep, bline)[1:]}',
            10: (f'={s}SUMPRODUCT(({BP_FY}={B7})*({BP_DEP}="{dep}")*({BP_LINE}="{bline}")*{BP_GRID})'),
            11: f'=$F{r}+($J{r}-$G{r})',
        }
        for col, f in cells.items():
            c = bv.cell(row=r, column=col, value=f)
            c.font = Font(name=ARIAL, size=10); c.number_format = MONEY; c.border = BOX
        r += 1
    lbl(bv, r, 1, "    Contribution", bold=True, size=10)
    for col in (2, 3, 6, 7, 10, 11):
        c = bv.cell(row=r, column=col, value=f'=SUM({CL(col)}{first}:{CL(col)}{r-1})')
        c.font = Font(name=ARIAL, size=10, bold=True); c.number_format = MONEY
        c.border = BOX; c.fill = TOTAL_FILL
    for rr in list(range(first, r + 1)) + [dept_hdr]:
        if rr == dept_hdr:
            for col, key in ((2, "CY_M"), (6, "CY_YTD")):
                c = bv.cell(row=dept_hdr, column=col,
                            value=f'=SUMIFS({edcol(key)},{ED_KIND},"Sub",{ED_DEPT},"{dep}")')
                c.font = Font(name=ARIAL, size=10, bold=True, color=WHITE); c.number_format = MONEY
            for col in (3, 7, 10, 11):
                c = bv.cell(row=dept_hdr, column=col, value=f'={CL(col)}{r}')
                c.font = Font(name=ARIAL, size=10, bold=True, color=WHITE); c.number_format = MONEY
        for col, f, fmt in ((4, f'=$B{rr}-$C{rr}', MONEY),
                            (5, f'=IF(ROUND($C{rr},2)=0,"n/a",$D{rr}/ABS($C{rr}))', PCT),
                            (8, f'=$F{rr}-$G{rr}', MONEY),
                            (9, f'=IF(ROUND($G{rr},2)=0,"n/a",$H{rr}/ABS($G{rr}))', PCT),
                            (12, (f'=IF(ROUND($G{rr},2)=0,"No Budget",IF(ABS($H{rr})<{MAT},"Low",'
                                  f'IF(OR(ABS($H{rr}/$G{rr})>={HIP},ABS($H{rr})>={HID}),"High",'
                                  f'IF(ABS($H{rr}/$G{rr})>={MEDP},"Medium","Low"))))'), "General")):
            c = bv.cell(row=rr, column=col, value=f)
            c.number_format = fmt
            c.font = Font(name=ARIAL, size=10, bold=(rr in (dept_hdr, r)),
                          color=(WHITE if rr == dept_hdr else BLACK_FONT))
            if col == 12: c.alignment = Alignment(horizontal="center")
    r += 2
BV_R1 = r
for txt, fill, font in (("High", F_HIGH, T_HIGH), ("Medium", F_MED, T_MED),
                        ("Low", F_LOW, T_LOW), ("No Budget", F_NA, T_NA)):
    bv.conditional_formatting.add(f"L6:L{BV_R1}", FormulaRule(
        formula=[f'$L6="{txt}"'], fill=PatternFill("solid", fgColor=fill),
        font=Font(color=font, bold=True)))

# ===========================================================================
# CHARTS  - driven off the engine, nothing copied by hand
# ===========================================================================
ch = sheet("Charts", "548235")
title(ch, "Charts",
      "Every series below is a formula off the engine. Paste a new month and the charts move - "
      "there is no copying from another workbook.")
widths(ch, {"A": 16})
for i in range(2, 9): ch.column_dimensions[CL(i)].width = 13
CD_R = 4
lbl(ch, CD_R, 1, "Revenue by department, current financial year", bold=True, size=10)
header(ch, CD_R + 1, ["Month"] + DEPTS)
for p in range(1, 13):
    r = CD_R + 1 + p
    ch.cell(row=r, column=1, value=mhdr(B7, p)).font = Font(name=ARIAL, size=9)
    for di, dep in enumerate(DEPTS):
        c = ch.cell(row=r, column=2 + di,
                    value=f'=SUMIFS({edper(1,p)},{ED_KIND},"Sub",{ED_DEPT},"{dep}",{ED_CAT},"Income")')
        c.font = Font(name=ARIAL, size=9); c.number_format = MONEY
REV_R0, REV_R1 = CD_R + 2, CD_R + 13
c1 = LineChart(); c1.title = "Revenue by department"; c1.height, c1.width = 8, 20
c1.y_axis.title = "Revenue"; c1.x_axis.title = "Month"
c1.add_data(Reference(ch, min_col=2, max_col=1 + len(DEPTS), min_row=REV_R0 - 1, max_row=REV_R1),
            titles_from_data=True)
c1.set_categories(Reference(ch, min_col=1, min_row=REV_R0, max_row=REV_R1))
ch.add_chart(c1, "J4")

GP_R = REV_R1 + 3
lbl(ch, GP_R, 1, "Gross profit by department, current financial year", bold=True, size=10)
header(ch, GP_R + 1, ["Month"] + DEPTS)
for p in range(1, 13):
    r = GP_R + 1 + p
    ch.cell(row=r, column=1, value=mhdr(B7, p)).font = Font(name=ARIAL, size=9)
    for di, dep in enumerate(DEPTS):
        c = ch.cell(row=r, column=2 + di,
                    value=(f'=SUMIFS({edper(1,p)},{ED_KIND},"Sub",{ED_DEPT},"{dep}",{ED_CAT},"Income")'
                           f'+SUMIFS({edper(1,p)},{ED_KIND},"Sub",{ED_DEPT},"{dep}",{ED_CAT},"Cost of Sales")'))
        c.font = Font(name=ARIAL, size=9); c.number_format = MONEY
GP_R0, GP_R1 = GP_R + 2, GP_R + 13
c2 = BarChart(); c2.title = "Gross profit by department"; c2.height, c2.width = 8, 20
c2.type = "col"; c2.grouping = "clustered"
c2.add_data(Reference(ch, min_col=2, max_col=1 + len(DEPTS), min_row=GP_R0 - 1, max_row=GP_R1),
            titles_from_data=True)
c2.set_categories(Reference(ch, min_col=1, min_row=GP_R0, max_row=GP_R1))
ch.add_chart(c2, "J22")

MX_R = GP_R1 + 3
lbl(ch, MX_R, 1, "Cost mix year to date, labour against everything else", bold=True, size=10)
header(ch, MX_R + 1, ["Department"] + TYPES)
for di, dep in enumerate(DEPTS):
    r = MX_R + 2 + di
    ch.cell(row=r, column=1, value=dep).font = Font(name=ARIAL, size=9)
    for k, t in enumerate(TYPES):
        c = ch.cell(row=r, column=2 + k, value=tsum(dep, t, "CY_YTD"))
        c.font = Font(name=ARIAL, size=9); c.number_format = MONEY
MX_R0, MX_R1 = MX_R + 2, MX_R + 1 + len(DEPTS)
c3 = BarChart(); c3.title = "Cost mix by department (YTD)"; c3.height, c3.width = 8, 20
c3.type = "col"; c3.grouping = "stacked"; c3.overlap = 100
c3.add_data(Reference(ch, min_col=2, max_col=1 + len(TYPES), min_row=MX_R0 - 1, max_row=MX_R1),
            titles_from_data=True)
c3.set_categories(Reference(ch, min_col=1, min_row=MX_R0, max_row=MX_R1))
ch.add_chart(c3, "J40")

UT_R = MX_R1 + 3
lbl(ch, UT_R, 1, "Chargeable hours by department, current financial year", bold=True, size=10)
header(ch, UT_R + 1, ["Month"] + DEPTS)
for p in range(1, 13):
    r = UT_R + 1 + p
    ch.cell(row=r, column=1, value=mhdr(B7, p)).font = Font(name=ARIAL, size=9)
    for di, dep in enumerate(DEPTS):
        c = ch.cell(row=r, column=2 + di,
                    value=(f'=SUMPRODUCT(({HP_FY}={B7})*({HP_DEP}="{dep}")*({HP_TYPE}="Chargeable")'
                           f'*({HP_PER}={p})*{HP_GRID})'))
        c.font = Font(name=ARIAL, size=9); c.number_format = HRS
UT_R0, UT_R1 = UT_R + 2, UT_R + 13
c4 = LineChart(); c4.title = "Chargeable hours by department"; c4.height, c4.width = 8, 20
c4.add_data(Reference(ch, min_col=2, max_col=1 + len(DEPTS), min_row=UT_R0 - 1, max_row=UT_R1),
            titles_from_data=True)
c4.set_categories(Reference(ch, min_col=1, min_row=UT_R0, max_row=UT_R1))
ch.add_chart(c4, "J58")
print("budget variance + charts built")

# ===========================================================================
# REVENUE_COMPARISON  - actual, prior year, budget and a forecast that
# updates itself. Replaces the Sales Revenue Comparison sheet, whose numbers
# were typed in by hand on the Revenue tab.
# ===========================================================================
rc = sheet("Revenue_Comparison", "2E75B6")
title(rc, "Revenue Comparison  -  actual, prior year, budget, forecast")
rc["A2"] = ('=IF(Setup!$B$4="","","Every figure below is a formula off the GL paste and the budget. '
            'The forecast is actual for months already closed and budget for the rest, so it rolls '
            'forward on its own.")')
rc["A2"].font = Font(name=ARIAL, size=9, italic=True, color=MUTED)
widths(rc, {"A": 14})
for i in range(2, 10): rc.column_dimensions[CL(i)].width = 14

def rev_block(row, heading, mode, fill):
    """mode: cy | py | budget | forecast"""
    band(rc, row, heading, 1, 8, fill=fill)
    header(rc, row + 1, ["Month"] + DEPTS + ["Total"])
    r0 = row + 2
    for p in range(1, 13):
        r = r0 + p - 1
        c = rc.cell(row=r, column=1, value=mhdr(B7 if mode != "py" else B9, p))
        c.font = Font(name=ARIAL, size=9); c.border = BOX
        for di, dep in enumerate(DEPTS):
            if mode in ("cy", "py"):
                off = 1 if mode == "cy" else 0
                f = (f'=SUMIFS({edper(off,p)},{ED_KIND},"Sub",{ED_DEPT},"{dep}",'
                     f'{ED_CAT},"Income")')
            elif mode == "budget":
                f = (f'=SUMPRODUCT(({BP_FY}={B7})*({BP_DEP}="{dep}")*({BP_LINE}="Income")'
                     f'*({BP_PER}={p})*{BP_GRID})')
            else:
                act = f"{CL(2+di)}{CY_R0 + p - 1}"
                bud = f"{CL(2+di)}{BU_R0 + p - 1}"
                f = f'=IF({p}<={B8},{act},{bud})'
            cc = rc.cell(row=r, column=2 + di, value=f)
            cc.font = Font(name=ARIAL, size=9); cc.number_format = MONEY; cc.border = BOX
        t = rc.cell(row=r, column=8, value=f'=SUM($B{r}:$G{r})')
        t.font = Font(name=ARIAL, size=9, bold=True); t.number_format = MONEY
        t.border = BOX; t.fill = TOTAL_FILL
    tr = r0 + 12
    rc.cell(row=tr, column=1, value="Full year").font = Font(name=ARIAL, size=9, bold=True)
    for col in range(2, 9):
        c = rc.cell(row=tr, column=col, value=f'=SUM({CL(col)}{r0}:{CL(col)}{r0+11})')
        c.font = Font(name=ARIAL, size=9, bold=True); c.number_format = MONEY
        c.border = BOX; c.fill = TOTAL_FILL
    return r0, tr + 2

CY_R0, nxt = rev_block(4, "ACTUAL  -  current financial year", "cy", PatternFill("solid", fgColor="2E75B6"))
PY_R0, nxt = rev_block(nxt, "ACTUAL  -  prior financial year", "py", PatternFill("solid", fgColor="808080"))
BU_R0, nxt = rev_block(nxt, "BUDGET  -  current financial year", "budget", PatternFill("solid", fgColor="7F6000"))
FC_R0, nxt = rev_block(nxt, "ROLLING FORECAST  -  actual for closed months, budget for the rest",
                       "forecast", PatternFill("solid", fgColor="548235"))

band(rc, nxt, "ROLLED UP THE WAY THE BUDGET NAMES THEM", 1, 6); nxt += 1
lbl(rc, nxt, 1, "Support = Onsite.  Production = Production + Video.  Consulting = Integration + "
                "Consulting.  This is the view your old graph pack used.", size=9, color=MUTED)
rc.merge_cells(start_row=nxt, start_column=1, end_row=nxt, end_column=8); nxt += 1
header(rc, nxt, ["Month"] + ROLL_NAMES + ["Total"]); nxt += 1
ROLL_R0 = nxt
GRP_OF = {g: [d for d in DEPTS if ROLLUP[d] == g] for g in ROLL_NAMES}
for p in range(1, 13):
    r = ROLL_R0 + p - 1
    rc.cell(row=r, column=1, value=f'=$A{CY_R0 + p - 1}').font = Font(name=ARIAL, size=9)
    for gi, g in enumerate(ROLL_NAMES):
        cols = [CL(2 + DEPTS.index(d)) for d in GRP_OF[g]]
        f = "=" + "+".join(f"{c}{CY_R0 + p - 1}" for c in cols)
        cc = rc.cell(row=r, column=2 + gi, value=f)
        cc.font = Font(name=ARIAL, size=9); cc.number_format = MONEY; cc.border = BOX
    t = rc.cell(row=r, column=2 + len(ROLL_NAMES), value=f'=SUM($B{r}:${CL(1+len(ROLL_NAMES))}{r})')
    t.font = Font(name=ARIAL, size=9, bold=True); t.number_format = MONEY
    t.border = BOX; t.fill = TOTAL_FILL
ROLL_R1 = ROLL_R0 + 11

cmp_chart = BarChart()
cmp_chart.title = "Revenue by department - actual, prior year, budget"
cmp_chart.type = "col"; cmp_chart.grouping = "clustered"
cmp_chart.height, cmp_chart.width = 9, 24
CMP_R = ROLL_R1 + 3
lbl(rc, CMP_R, 1, "Reporting month comparison", bold=True, size=10)
header(rc, CMP_R + 1, ["Department", "Actual", "Prior year", "Budget"])
for di, dep in enumerate(DEPTS):
    r = CMP_R + 2 + di
    rc.cell(row=r, column=1, value=dep).font = Font(name=ARIAL, size=9)
    for k, r0 in enumerate((CY_R0, PY_R0, BU_R0)):
        c = rc.cell(row=r, column=2 + k, value=f'=INDEX({CL(2+di)}{r0}:{CL(2+di)}{r0+11},{B8})')
        c.font = Font(name=ARIAL, size=9); c.number_format = MONEY; c.border = BOX
CMP_R0, CMP_R1 = CMP_R + 2, CMP_R + 1 + len(DEPTS)
cmp_chart.add_data(Reference(rc, min_col=2, max_col=4, min_row=CMP_R0 - 1, max_row=CMP_R1),
                   titles_from_data=True)
cmp_chart.set_categories(Reference(rc, min_col=1, min_row=CMP_R0, max_row=CMP_R1))
rc.add_chart(cmp_chart, "J4")

trend = LineChart(); trend.title = "Total revenue - actual against budget and forecast"
trend.height, trend.width = 9, 24
TR_R = CMP_R1 + 3
lbl(rc, TR_R, 1, "Monthly trend", bold=True, size=10)
header(rc, TR_R + 1, ["Month", "Actual", "Prior year", "Budget", "Forecast"])
for p in range(1, 13):
    r = TR_R + 1 + p
    rc.cell(row=r, column=1, value=f'=$A{CY_R0 + p - 1}').font = Font(name=ARIAL, size=9)
    for k, r0 in enumerate((CY_R0, PY_R0, BU_R0, FC_R0)):
        c = rc.cell(row=r, column=2 + k, value=f'=$H{r0 + p - 1}')
        c.font = Font(name=ARIAL, size=9); c.number_format = MONEY; c.border = BOX
TR_R0, TR_R1 = TR_R + 2, TR_R + 13
trend.add_data(Reference(rc, min_col=2, max_col=5, min_row=TR_R0 - 1, max_row=TR_R1),
               titles_from_data=True)
trend.set_categories(Reference(rc, min_col=1, min_row=TR_R0, max_row=TR_R1))
rc.add_chart(trend, "J24")

# ===========================================================================
# CONTROLS
# ===========================================================================
ct = sheet("Controls", "7030A0")
title(ct, "Controls & Data Integrity",
      "Clear every FAIL before you read a number. FAIL means the pack is wrong, not untidy.")
widths(ct, {"A": 6, "B": 52, "C": 15, "D": 12, "E": 11, "F": 9, "G": 62})
header(ct, 5, ["Ref", "Control", "Result", "Target", "Status", "Severity", "What to do"])
ct.freeze_panes = "A6"
CU_ST = f"Cleanup!$O${CL_R0}:$O${CL_R1}"
E_YTD = ecol("CY_YTD")
CHECKS = [
    ("C1", "Every kept line reached the engine",
     f'=ROUND(SUMPRODUCT(Cleanup!$I${CL_R0}:$I${CL_R1},Cleanup!$B${CL_R0}:$B${CL_R1})'
     f'-SUM(Engine!$F${E_R0}:$AC${E_R1}),2)', "= 0", "eq0", "FAIL", MONEY2,
     "Nothing may be lost between the paste and the reports. If this is not nil a line was kept but "
     "its key did not land in the engine grid - check Setup B5 is a month the paste actually covers."),
    ("C1b", "Net movement of the paste (nil only if you pasted a full ledger)",
     f'=ROUND(SUMPRODUCT(Cleanup!$I${CL_R0}:$I${CL_R1},Cleanup!$B${CL_R0}:$B${CL_R1}),2)',
     "info", "info", "REVIEW", MONEY2,
     "A profit and loss only export nets to the profit for the period, which is expected. A full "
     "ledger export including balance sheet accounts nets to nil."),
    ("C2", "Every pasted line matched an account code",
     f'=COUNTIF({CU_ST},"NO ACCOUNT")', "= 0", "eq0", "FAIL", NUM,
     "Filter Cleanup column O for NO ACCOUNT. The account code is missing from the Lists sheet - "
     "add it there with its category, subcategory and type."),
    ("C3", "Every pasted line has a department",
     f'=COUNTIF({CU_ST},"NO DEPARTMENT")', "= 0", "eq0", "FAIL", NUM,
     "The Cost Centres column was blank and the job number did not match the job master. Either fix "
     "the cost centre in Xero or add the job to the Lists sheet."),
    ("C4", "No lines outside the two financial years on Setup",
     f'=COUNTIF({CU_ST},"OUTSIDE FY")', "= 0", "eq0", "REVIEW", NUM,
     "Those lines are excluded. Fine if you pasted three years; a problem if the reporting month is wrong."),
    ("C5", "Reporting month has data",
     f'=SUMPRODUCT((Cleanup!$I${CL_R0}:$I${CL_R1}=1)*(Cleanup!$H${CL_R0}:$H${CL_R1}=1)'
     f'*(Cleanup!$G${CL_R0}:$G${CL_R1}={B8}))', "> 0", "gt0", "FAIL", NUM,
     "Either the month on Setup B5 is wrong or that month is not in the paste."),
    ("C6", "Prior year comparative is loaded",
     f'=SUMPRODUCT(--(Cleanup!$H${CL_R0}:$H${CL_R1}=0))', "> 0", "gt0", "REVIEW", NUM,
     "Without the prior year every year-on-year variance reads as a new account. Paste both years."),
    ("C7", "Paste capacity not exceeded",
     f'=IF(Setup!$B$21>={GL_R1-GL_R0+1},1,0)', "= 0", "eq0", "FAIL", NUM,
     f"GL_Paste holds {GL_R1-GL_R0+1:,} rows. If it is full, data is being cut off - paste the prior "
     "year summarised by month instead of transaction by transaction."),
    ("C8", "Account engine and department engine agree",
     f'=ROUND(SUM({E_YTD})-SUMIFS({edcol("CY_YTD")},{ED_KIND},"Sub"),2)', "= 0", "eq0", "FAIL", MONEY2,
     "The two engines are built from different keys off the same data, so they must agree. If they do "
     "not, an account has no subcategory on the Lists sheet and is dropping out of the department view."),
    ("C9", "Departments add back to the company total",
     f'=ROUND(SUMIF({E_CAT},"Income",{E_YTD})'
     f'-SUMIFS({edcol("CY_YTD")},{ED_KIND},"Sub",{ED_CAT},"Income"),2)', "= 0", "eq0", "FAIL", MONEY2,
     "Same cause as C8, on the revenue lines specifically."),
    ("C10", "Budget loaded for the reporting year",
     f'=SUMPRODUCT(--({BP_FY}={B7}))', "> 0", "gt0", "REVIEW", NUM,
     "Budget_Variance will read nil until you enter the budget for this financial year."),
    ("C11", "Hours loaded for the reporting year",
     f'=SUMPRODUCT(--({HP_FY}={B7}))', "> 0", "gt0", "REVIEW", NUM,
     "The utilisation block on the summary needs the hours from your utilisation report."),
    ("C12", "Income is positive on the sign convention",
     f'=IF(SUMIF({E_CAT},"Income",{E_YTD})>=0,0,1)', "= 0", "eq0", "FAIL", NUM,
     "Amount is Credit less Debit, so income must come out positive. If this fails, the Debit and "
     "Credit columns are the wrong way round in the paste."),
    ("C13", "No future-dated lines",
     f'=SUMPRODUCT((Cleanup!$A${CL_R0}:$A${CL_R1}<>"")*(Cleanup!$A${CL_R0}:$A${CL_R1}>TODAY()))',
     "= 0", "eq0", "REVIEW", NUM, "Usually a typed date error. Check before reporting."),
    ("C15", "Cost type view agrees with the subcategory view",
     f'=ROUND(SUMIFS({edcol("CY_YTD")},{ED_KIND},"Sub")-SUMIFS({edcol("CY_YTD")},{ED_KIND},"Type"),2)',
     "= 0", "eq0", "FAIL", MONEY2,
     "The Labour / Equipment split and the subcategory split are built from different keys over the "
     "same lines, so they must agree. A gap means an account has no Type on the Lists sheet."),
    ("C14", "Every account on Lists carries a subcategory",
     f'=SUMPRODUCT(--({L_ACC_SUB}=0))', "= 0", "eq0", "REVIEW", NUM,
     "An account with no subcategory still appears in the company P&L but vanishes from the "
     "department view. This is what C8 would catch."),
]
for i, (ref, desc, f, target, mode, sev, fmt, guide) in enumerate(CHECKS):
    r = 6 + i
    for col, v, al in ((1, ref, "center"), (2, desc, "left"), (4, target, "center"),
                       (6, sev, "center"), (7, guide, "left")):
        c = ct.cell(row=r, column=col, value=v)
        c.font = Font(name=ARIAL, size=(9 if col in (1, 6, 7) else 10),
                      bold=(col == 1), color=(MUTED if col in (4, 6, 7) else BLACK_FONT))
        c.alignment = Alignment(horizontal=al, wrap_text=(col in (2, 7)), vertical="center")
    c = ct.cell(row=r, column=3, value=f)
    c.font = Font(name=ARIAL, size=10, bold=True); c.number_format = fmt
    c.alignment = Alignment(horizontal="center")
    if mode == "info":
        e = ct.cell(row=r, column=5, value='="INFO"')      # neutral: nothing to pass or fail
    else:
        cond = f'ROUND($C{r},2)=0' if mode == "eq0" else f'$C{r}>0'
        e = ct.cell(row=r, column=5, value=f'=IF({cond},"PASS","{sev}")')
    e.font = Font(name=ARIAL, size=10, bold=True); e.alignment = Alignment(horizontal="center")
    for col in range(1, 8): ct.cell(row=r, column=col).border = BOX
    ct.row_dimensions[r].height = 30
CT_R1 = 5 + len(CHECKS)
ct.conditional_formatting.add(f"A6:G{CT_R1}", FormulaRule(formula=['$E6="FAIL"'],
    fill=ROW_HIGH, font=Font(color=T_HIGH)))
ct.conditional_formatting.add(f"A6:G{CT_R1}", FormulaRule(formula=['$E6="REVIEW"'],
    fill=ROW_MED, font=Font(color=T_MED)))
for txt, fill, font in (("PASS", F_LOW, T_LOW), ("REVIEW", F_MED, T_MED),
                        ("FAIL", F_HIGH, T_HIGH), ("INFO", F_NA, T_NA)):
    ct.conditional_formatting.add(f"E6:E{CT_R1}", FormulaRule(formula=[f'$E6="{txt}"'],
        fill=PatternFill("solid", fgColor=fill), font=Font(color=font, bold=True)))

# ===========================================================================
# README
# ===========================================================================
rd = wb.create_sheet("README", 0)
rd.sheet_view.showGridLines = False
title(rd, "CTS Financial Controller Pack",
      "One workbook. Three paste points. Everything else is formulas.")
widths(rd, {"A": 4, "B": 24, "C": 96})
def rb(r, t):
    band(rd, r, t, 1, 3); return r + 1
def rl(r, a, b):
    lbl(rd, r, 2, a, bold=True, size=10)
    c = rd.cell(row=r, column=3, value=b)
    c.font = Font(name=ARIAL, size=10); c.alignment = Alignment(wrap_text=True, vertical="top")
    rd.row_dimensions[r].height = max(14, 13 * (1 + len(b) // 100))
    return r + 1

r = 4
r = rb(r, "EACH MONTH  -  three pastes, then read")
for a, b in [
    ("1. Setup B5", "Pick the reporting month from the dropdown. The financial year and period follow."),
    ("2. GL_Paste", "Run the same Xero GL export you already use and paste it at cell A1, exactly as it "
                    "comes. The sheet is bare so a whole-sheet paste works. Columns must stay in their "
                    "current order: Account Code, Account Name, Source, Date, Contact, Debit, Credit, "
                    "Job Numbers ... Cost Centres."),
    ("3. Budget_Paste", "One row per financial year, department and line, with the twelve months across. "
                        "Enter it once a year, not monthly."),
    ("4. Hours_Paste", "Chargeable, non-chargeable and leave hours per department per month, from the Job "
                       "Group table on your utilisation Dashboard. Once a month."),
    ("5. Controls", "Clear every FAIL. Then Financial_Summary, then Dept_PL, then Budget_Variance."),
]:
    r = rl(r, a, b)
r += 1
r = rb(r, "WHAT REPLACED WHAT")
for a, b in [
    ("PL Analysis.xlsm", "Tabs 1 to 8 become GL_Paste, Cleanup, Engine and Dept_PL. The SUMIFS design is "
                         "yours - it was already right. What changed is that it now feeds the budget, the "
                         "hours and the charts as well, instead of stopping at the P&L."),
    ("2026-08_Graph.xlsx", "The Charts sheet. Every series is a formula off the engine, so there is no "
                           "copying tab 5 and tab 7 across each month."),
    ("CTS Budget FY27", "Budget_Paste plus Budget_Variance. Actuals are never retyped - they come from "
                        "the same GL paste as everything else, which removes the step most likely to "
                        "introduce an error."),
    ("Utilisation Report", "Hours_Paste plus the utilisation block on Financial_Summary: chargeable hours, "
                           "utilisation rate, revenue per chargeable hour and gross profit per chargeable "
                           "hour, per department."),
]:
    r = rl(r, a, b)
r += 1
r = rb(r, "THINGS THAT WILL CATCH YOU OUT")
for a, b in [
    ("Sign convention", "Amount is Credit less Debit, matching your existing pack. Income is positive, "
                        "costs are negative, and every subtotal is therefore a contribution. Gross profit "
                        "is Income plus Cost of Sales, not minus."),
    ("Department", "Comes from the Cost Centres column on each GL line, falling back to the job number "
                   "against the job master on Lists. It is per transaction, not per account, so one "
                   "account can serve several departments - which is how your GL actually behaves."),
    ("Adding an account", "A new Xero account must be added to the Lists sheet with its category, "
                          "subcategory and type, or its lines are rejected by control C2."),
    ("Capacity", f"GL_Paste holds {GL_R1-GL_R0+1:,} rows, about sixteen months at your volume. For two "
                 "full years, paste the prior year summarised by month. Control C7 warns you."),
    ("Budget department names", "Your budget calls them Support, Consulting and Production. On the GL "
                                "they are ONSITE, INTEGRATION and PRODUCTION plus VIDEO and CONSULTING. "
                                "Enter the budget against the GL department names and the variance ties."),
]:
    r = rl(r, a, b)

# ===========================================================================
# finish
# ===========================================================================
del wb["Sheet"]
order = ["README", "Setup", "GL_Paste", "Budget_Paste", "Hours_Paste", "Financial_Summary",
         "Dept_PL", "Budget_Variance", "Revenue_Comparison", "Charts", "Controls", "Cleanup", "Engine",
         "Engine_Dept", "Lists"]
wb._sheets = [wb[n] for n in order]
for ws in wb.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if c.value is not None or c.has_style:
                f = c.font
                if f.name != ARIAL:
                    c.font = Font(name=ARIAL, size=f.size or 10, bold=f.bold, italic=f.italic,
                                  color=f.color, underline=f.underline)
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
dp.sheet_properties.outlinePr = Outline(summaryBelow=False, summaryRight=False, applyStyles=False)
wb.active = 0
OUT.parent.mkdir(parents=True, exist_ok=True)
wb.save(OUT)
print("saved:", OUT)
print("sheets:", wb.sheetnames)
