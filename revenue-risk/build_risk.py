"""Build the CTS Revenue Risk Profile workbook.

Only four sheets are typed or pasted into: Zoho Dump, P&L Dump, Adjustments
and Budget. Settings holds the levers. Everything else is formulas, so a new
Zoho export or a new month of P&L flows straight through.
"""
import datetime as dt
import json
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as CL
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule

ROOT = Path("/home/user/Claude/revenue-risk")
ZOHO_SRC = ROOT / "source/All Deals by Stage FY26-FY27 (12-10-2026).xlsx"
ACCOUNTS = json.loads(Path("/home/user/Claude/reporting/data/accounts.json").read_text())
OUT = ROOT / "CTS Revenue Risk Profile.xlsx"

NAVY, MUTED = "1F4E79", "595959"
HEAD = PatternFill("solid", fgColor=NAVY)
BAND = PatternFill("solid", fgColor="D9E2F3")
INPUT = PatternFill("solid", fgColor="FFF2CC")
TOTAL = PatternFill("solid", fgColor="F2F2F2")
HELP = PatternFill("solid", fgColor="E2EFDA")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
ARIAL = "Arial"
MONEY = '$#,##0;($#,##0);-'
PCT = '0.0%;(0.0%);-'
MON = 'mmm-yy'

DEPTS = ["Onsite", "Production", "Video", "Consulting", "Integration", "Admin", "Unallocated"]
ZOHO_ROWS = 3000            # helper formulas cover Zoho Dump rows 8 to 3007
ZF, ZL = 8, 8 + ZOHO_ROWS - 1
PL_ROWS = 1000              # P&L Dump paste rows 7 to 1006
PF, PL = 7, 7 + PL_ROWS - 1
ADJ_ROWS = 200
AF, AL = 8, 8 + ADJ_ROWS - 1
CUST_ROWS = 400


def font(ws, ref, **kw):
    ws[ref].font = Font(name=ARIAL, **kw)


def title(ws, text, sub=None):
    ws["A1"] = text
    ws["A1"].font = Font(name=ARIAL, size=16, bold=True, color=NAVY)
    if sub:
        ws["A2"] = sub
        ws["A2"].font = Font(name=ARIAL, size=9, color=MUTED)
    ws.sheet_view.showGridLines = False


def head(ws, row, col, labels, fill=HEAD, color="FFFFFF"):
    for i, label in enumerate(labels):
        c = ws.cell(row=row, column=col + i, value=label)
        c.font = Font(name=ARIAL, size=10, bold=True, color=color)
        c.fill = fill
        c.border = BOX
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")


def put(ws, row, col, value, fmt=None, bold=False, fill=None, color="000000", border=True, italic=False):
    c = ws.cell(row=row, column=col, value=value)
    c.font = Font(name=ARIAL, size=10, bold=bold, color=color, italic=italic)
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill
    if border:
        c.border = BOX
    return c


def widths(ws, spec):
    for col, w in spec.items():
        ws.column_dimensions[col].width = w


def note(ws, row, text, col=1, size=9, color=MUTED, bold=False):
    c = ws.cell(row=row, column=col, value=text)
    c.font = Font(name=ARIAL, size=size, color=color, bold=bold)
    return c


wb = Workbook()
wb.remove(wb.active)

# ---------------------------------------------------------------- Settings
st = wb.create_sheet("Settings")
title(st, "Settings", "Yellow cells are inputs. Change these when the month or the year rolls over.")
S = {}
rows = [
    ("Financial year starts", dt.date(2026, 7, 1), "dd/mm/yyyy", "Australian financial year, 1 July."),
    ("Last month of P&L actuals (month end)", dt.date(2026, 9, 30), "dd/mm/yyyy",
     "Update each month after the P&L is final. Zoho deals closing after this date count as future revenue."),
    ("Financial year ends", "=EDATE(B4,12)-1", "dd/mm/yyyy", "Formula."),
    ("First forecast month", "=EOMONTH(B5,0)+1", "dd/mm/yyyy", "Formula."),
    ("Count overdue Zoho deals in the forecast?", "No", None,
     "Overdue means an open deal whose closing date has already passed. Yes puts them in the first forecast month at their stage weighting."),
]
head(st, 3, 1, ["Setting", "Value", "Note"])
for i, (lab, val, fmt, nt) in enumerate(rows):
    r = 4 + i
    put(st, r, 1, lab)
    is_input = not (isinstance(val, str) and val.startswith("="))
    put(st, r, 2, val, fmt, fill=INPUT if is_input else None, color="0000FF" if is_input else "000000")
    put(st, r, 3, nt, italic=True, color=MUTED)
FYSTART, ASAT, FYEND, FIRSTF, INCOVER = ("Settings!$B$4", "Settings!$B$5", "Settings!$B$6",
                                          "Settings!$B$7", "Settings!$B$8")
dv = DataValidation(type="list", formula1='"Yes,No"', allow_blank=False)
st.add_data_validation(dv)
dv.add("B8")

# Departments and months
head(st, 11, 1, ["Departments (as in the P&L)"])
for i, d in enumerate(DEPTS):
    put(st, 12 + i, 1, d)
DEPT_RNG = "Settings!$A$12:$A$18"
head(st, 11, 4, ["Months in this FY"])
for i in range(12):
    put(st, 12 + i, 4, f"=EDATE($B$4,{i})", MON)
MONTH_RNG = "Settings!$D$12:$D$23"

# Stage table
head(st, 26, 1, ["Zoho stage", "Type", "Probability"])
note(st, 25, "Probability is the chance an open deal closes. These are placeholders. Management should set them.", bold=True, color="C00000")
stages = [("Closed Won", "Won", None), ("Closed Lost", "Lost", None),
          ("Verbal Approval", "Open", 0.9), ("Negotiation/Review", "Open", 0.7),
          ("Proposal Sent", "Open", 0.5), ("Building Value Proposition", "Open", 0.2)]
for i in range(10):
    r = 27 + i
    s = stages[i] if i < len(stages) else ("", "", None)
    put(st, r, 1, s[0] or None, fill=INPUT, color="0000FF")
    put(st, r, 2, s[1] or None, fill=INPUT, color="0000FF")
    put(st, r, 3, s[2], PCT, fill=INPUT, color="0000FF")
STAGE_N, STAGE_T, STAGE_P = "Settings!$A$27:$A$36", "Settings!$B$27:$B$36", "Settings!$C$27:$C$36"
dv = DataValidation(type="list", formula1='"Won,Open,Lost"', allow_blank=True)
st.add_data_validation(dv)
dv.add("B27:B36")
note(st, 37, "Type must be Won, Open or Lost. Add any new Zoho stage in a spare row.")

# Zoho department map
head(st, 40, 1, ["Zoho CTS department", "P&L department", "Status"])
zmap = [("Event Production", "Production", "From the name"),
        ("Video Production", "Video", "From the name"),
        ("Consulting", "Consulting", "Same name"),
        ("Integration", "Integration", "Same name"),
        ("Support", "Onsite", "ASSUMED. Please confirm"),
        ("Multi-Service", "Unallocated", "ASSUMED. Please confirm")]
for i in range(10):
    r = 41 + i
    z = zmap[i] if i < len(zmap) else ("", "", "")
    put(st, r, 1, z[0] or None, fill=INPUT, color="0000FF")
    put(st, r, 2, z[1] or None, fill=INPUT, color="0000FF")
    put(st, r, 3, z[2] or None, italic=True, color="C00000" if "ASSUMED" in z[2] else MUTED)
ZMAP_F, ZMAP_T = "Settings!$A$41:$A$50", "Settings!$B$41:$B$50"
dv = DataValidation(type="list", formula1=DEPT_RNG, allow_blank=True)
st.add_data_validation(dv)
dv.add("B41:B50")

# Risk rating thresholds
head(st, 53, 1, ["Risk rating test", "Medium from", "High from"])
note(st, 52, "Thresholds are placeholders. Management should set them.", bold=True, color="C00000")
thr = [("Forecast short of budget by", 0.0, 0.10),
       ("Largest customer share of forward revenue", 0.25, 0.40),
       ("Share of forward revenue that is pipeline, not contracted", 0.30, 0.50)]
for i, (lab, m, h) in enumerate(thr):
    r = 54 + i
    put(st, r, 1, lab)
    put(st, r, 2, m, PCT, fill=INPUT, color="0000FF")
    put(st, r, 3, h, PCT, fill=INPUT, color="0000FF")
note(st, 58, "Forecast short of budget: medium if the forecast is below budget at all, high if it is more than 10% below.")
widths(st, {"A": 44, "B": 18, "C": 70, "D": 14})

# ---------------------------------------------------------------- Zoho Dump
zd = wb.create_sheet("Zoho Dump")
src = load_workbook(ZOHO_SRC).active
for row in src.iter_rows():
    for c in row:
        if c.value is None:
            continue
        n = zd.cell(row=c.row, column=c.column, value=c.value)
        n.number_format = c.number_format
        n.font = Font(name=ARIAL, size=10, bold=(c.row == 7))
for col, w in zip("ABCDEFGHIJ", [28, 50, 34, 16, 18, 13, 16, 10, 10, 3]):
    zd.column_dimensions[col].width = w

# helper block, columns L onward
H = dict(stage="L", deal="M", acct="N", zdept="O", dept="P", amt="Q", close="R", stype="S",
         cat="T", prob="U", wtd="V", land="W", fwd="X", full="Y", stress="Z", wd="AA", cno="AB")
labels = dict(stage="Stage", deal="Deal", acct="Customer", zdept="Zoho dept", dept="P&L dept",
              amt="Amount", close="Closing date", stype="Stage type", cat="Category",
              prob="Probability", wtd="Weighted", land="Lands in month", fwd="Forward (weighted, this FY)",
              full="Forward (full value, this FY)", stress="Stress case value", wd="Withdrawn customer",
              cno="Customer #")
finders = dict(stage="Stage", deal="Deal Name", acct="Account Name", zdept="CTS DEPARTMENT",
               amt="Amount*", close="Closing Date")
zd["L1"] = "Helper columns. Formulas only. Do not type here."
zd["L1"].font = Font(name=ARIAL, size=11, bold=True, color="C00000")
zd["L3"] = "Column found at:"
zd["L3"].font = Font(name=ARIAL, size=9, color=MUTED)
for k, col in H.items():
    zd[f"{col}7"] = labels[k]
    zd[f"{col}7"].font = Font(name=ARIAL, size=9, bold=True, color="FFFFFF")
    zd[f"{col}7"].fill = HEAD
    zd[f"{col}7"].alignment = Alignment(wrap_text=True, vertical="center")
    zd.column_dimensions[col].width = 13
    if k in finders:
        zd[f"{col}5"] = finders[k].rstrip("*")
        zd[f"{col}5"].font = Font(name=ARIAL, size=8, color=MUTED)
        zd[f"{col}6"] = f'=IFERROR(MATCH("{finders[k]}",$A$7:$J$7,0),0)'
        zd[f"{col}6"].font = Font(name=ARIAL, size=8, color=MUTED)
zd.column_dimensions["M"].width = 30
zd.column_dimensions["N"].width = 26
zd.row_dimensions[7].height = 40


def raw(r, k):
    return f"INDEX($A{r}:$J{r},${H[k]}$6)"


for r in range(ZF, ZL + 1):
    f = {}
    f["stage"] = (f'=IF({raw(r,"stage")}<>"",TRIM(LEFT({raw(r,"stage")},FIND("(",{raw(r,"stage")}&"(")-1)),L{r-1})')
    f["deal"] = f'=IF({raw(r,"deal")}="","",{raw(r,"deal")}&"")'
    f["acct"] = f'=IF($M{r}="","",TRIM({raw(r,"acct")}&""))'
    f["zdept"] = f'=IF($M{r}="","",TRIM({raw(r,"zdept")}&""))'
    f["dept"] = (f'=IF($M{r}="","",IFERROR(IF(INDEX({ZMAP_T},MATCH(O{r},{ZMAP_F},0))&""="","UNMAPPED",'
                 f'INDEX({ZMAP_T},MATCH(O{r},{ZMAP_F},0))&""),"UNMAPPED"))')
    f["amt"] = f'=IF($M{r}="",0,N({raw(r,"amt")}))'
    f["close"] = f'=IF($M{r}="","",{raw(r,"close")})'
    f["stype"] = f'=IF($M{r}="","",IFERROR(INDEX({STAGE_T},MATCH(L{r},{STAGE_N},0))&"","UNMAPPED"))'
    f["cat"] = (f'=IF($M{r}="","",IF(NOT(ISNUMBER(R{r})),"NO DATE",IF(S{r}="Won",IF(R{r}>{ASAT},"Backlog","Delivered"),'
                f'IF(S{r}="Open",IF(R{r}>{ASAT},"Pipeline","Overdue"),IF(S{r}="Lost","Lost","UNMAPPED")))))')
    f["prob"] = (f'=IF(T{r}="Backlog",1,IF(OR(T{r}="Pipeline",T{r}="Overdue"),'
                 f'N(INDEX({STAGE_P},MATCH(L{r},{STAGE_N},0))),0))')
    f["wtd"] = f"=Q{r}*U{r}"
    f["land"] = (f'=IF($M{r}="","",IF(T{r}="Overdue",IF({INCOVER}="Yes",{FIRSTF},""),'
                 f'IF(ISNUMBER(R{r}),DATE(YEAR(R{r}),MONTH(R{r}),1),"")))')
    counted = f'AND(OR(T{r}="Backlog",T{r}="Pipeline",T{r}="Overdue"),ISNUMBER(W{r}))'
    f["fwd"] = f"=IF({counted},IF(W{r}<={FYEND},V{r},0),0)"
    f["full"] = f"=IF({counted},IF(W{r}<={FYEND},Q{r},0),0)"
    f["stress"] = f'=IF(Y{r}=0,0,IF(T{r}="Backlog",Q{r},Q{r}*Dashboard!$C$10))'
    f["wd"] = f'=IF($N{r}="",0,IF(COUNTIF(Dashboard!$C$13:$C$22,$N{r})>0,1,0))'
    f["cno"] = f'=IF($N{r}="","",IF(COUNTIF($N${ZF}:$N{r},$N{r})=1,MAX($AB${ZF-1}:$AB{r-1})+1,""))'
    for k, col in H.items():
        c = zd[f"{col}{r}"]
        c.value = f[k]
        c.font = Font(name=ARIAL, size=9)
        if k == "close":
            c.number_format = "dd/mm/yyyy"
        elif k == "land":
            c.number_format = MON
        elif k == "prob":
            c.number_format = PCT
        elif k in ("amt", "wtd", "fwd", "full", "stress"):
            c.number_format = MONEY
zd.freeze_panes = "A8"
Z = lambda k: f"'Zoho Dump'!${H[k]}${ZF}:${H[k]}${ZL}"

# ---------------------------------------------------------------- P&L Dump
pd_ = wb.create_sheet("P&L Dump")
title(pd_, "P&L Dump: paste the Xero Profit and Loss here",
      "Paste the Xero P&L with account names in column A and months across. Start at cell A7. "
      "Then pick the month for each column in row 5. Only Income accounts are used.")
note(pd_, 3, "Clear A7:Z1006 before each paste so old rows do not stay behind. Months can be in any order.", color="C00000", bold=True)
put(pd_, 5, 1, "Month in this column", bold=True)
for c in range(2, 27):
    put(pd_, 5, c, None, MON, fill=INPUT, color="0000FF")
dv = DataValidation(type="list", formula1=MONTH_RNG, allow_blank=True)
pd_.add_data_validation(dv)
dv.add("B5:Z5")
head(pd_, 6, 1, ["Paste starts below (A7)"])
head(pd_, 6, 28, ["Account", "Group", "Department", "Not in Account Map"])
for r in range(PF, PL + 1):
    pd_[f"AB{r}"] = f'=TRIM($A{r}&"")'
    pd_[f"AC{r}"] = f'=IF(AB{r}="","",IFERROR(INDEX(\'Account Map\'!$B:$B,MATCH(AB{r},\'Account Map\'!$A:$A,0)),""))'
    pd_[f"AD{r}"] = f'=IF(AB{r}="","",IFERROR(INDEX(\'Account Map\'!$C:$C,MATCH(AB{r},\'Account Map\'!$A:$A,0)),""))'
    pd_[f"AE{r}"] = (f'=IF(AND(AB{r}<>"",AC{r}="",COUNT($B{r}:$Z{r})>0,LEFT(AB{r},5)<>"Total",'
                     f'AB{r}<>"Gross Profit",AB{r}<>"Net Profit",AB{r}<>"Operating Profit"),1,0)')
    for col in ("AB", "AC", "AD", "AE"):
        pd_[f"{col}{r}"].font = Font(name=ARIAL, size=9)
    for c in range(2, 27):
        pd_.cell(row=r, column=c).number_format = MONEY
widths(pd_, {"A": 38, "AB": 34, "AC": 12, "AD": 13, "AE": 12})
for c in range(2, 27):
    pd_.column_dimensions[CL(c)].width = 12
pd_.freeze_panes = "B7"

# ---------------------------------------------------------------- Account Map
am = wb.create_sheet("Account Map")
title(am, "Account Map", "Copied from the account list in CTS P&L Reporting. Department is taken from the account suffix (ONS, PRD, VID, CONS, INTEGRATION, ADMIN).")
note(am, 3, "Most income accounts carry no suffix, so they sit in Unallocated. Change the department in column C if they belong to one team.", color="C00000", bold=True)
head(am, 4, 1, ["Account", "Group", "Department"])
for i, a in enumerate(ACCOUNTS):
    r = 5 + i
    put(am, r, 1, a["account"])
    put(am, r, 2, a["group"])
    put(am, r, 3, a["division"], fill=INPUT if a["group"] == "Income" else None,
        color="0000FF" if a["group"] == "Income" else "000000")
dv = DataValidation(type="list", formula1=DEPT_RNG, allow_blank=True)
am.add_data_validation(dv)
dv.add(f"C5:C{5 + len(ACCOUNTS) + 50}")
widths(am, {"A": 40, "B": 16, "C": 16})
am.freeze_panes = "A5"

# ---------------------------------------------------------------- Adjustments
ad = wb.create_sheet("Adjustments")
title(ad, "Adjustments: contracted or likely revenue that is not in Zoho",
      "For example recurring support contracts, licence renewals, or a deal sales has not logged yet. Yellow cells only.")
note(ad, 3, "Example of a line:  Onsite | Customer name exactly as in Zoho | Monthly help desk contract | Backlog | (leave blank) | 01/11/2026 | 12000")
note(ad, 4, "Type is Backlog (signed) or Pipeline (not signed). Pipeline needs a probability. One row per month.", color="C00000")
cols = ["P&L department", "Customer", "Description", "Type", "Probability", "Month", "Amount",
        "Category", "Weighted", "Lands in month", "Forward (weighted, this FY)", "Forward (full value, this FY)",
        "Stress case value", "Withdrawn customer"]
head(ad, 7, 1, cols)
for r in range(AF, AL + 1):
    for c in range(1, 8):
        put(ad, r, c, None, {5: PCT, 6: "dd/mm/yyyy", 7: MONEY}.get(c), fill=INPUT, color="0000FF")
    hf = {
        8: f'=IF(OR(G{r}="",A{r}=""),"",IF(NOT(ISNUMBER(F{r})),"NO DATE",IF(D{r}="Backlog",IF(F{r}>{ASAT},"Backlog","Delivered"),IF(F{r}>{ASAT},"Pipeline","Overdue"))))',
        9: f'=IF(H{r}="",0,N(G{r})*IF(D{r}="Backlog",1,N(E{r})))',
        10: f'=IF(OR(H{r}="",H{r}="NO DATE"),"",IF(H{r}="Overdue",IF({INCOVER}="Yes",{FIRSTF},""),DATE(YEAR(F{r}),MONTH(F{r}),1)))',
        11: f'=IF(AND(OR(H{r}="Backlog",H{r}="Pipeline",H{r}="Overdue"),ISNUMBER(J{r})),IF(J{r}<={FYEND},I{r},0),0)',
        12: f'=IF(AND(OR(H{r}="Backlog",H{r}="Pipeline",H{r}="Overdue"),ISNUMBER(J{r})),IF(J{r}<={FYEND},N(G{r}),0),0)',
        13: f'=IF(L{r}=0,0,IF(H{r}="Backlog",N(G{r}),N(G{r})*Dashboard!$C$10))',
        14: f'=IF(B{r}="",0,IF(COUNTIF(Dashboard!$C$13:$C$22,B{r})>0,1,0))',
    }
    for c, v in hf.items():
        put(ad, r, c, v, {9: MONEY, 10: MON, 11: MONEY, 12: MONEY, 13: MONEY}.get(c), color="000000")
for rng, f1 in ((f"A{AF}:A{AL}", DEPT_RNG), (f"D{AF}:D{AL}", '"Backlog,Pipeline"')):
    dv = DataValidation(type="list", formula1=f1, allow_blank=True)
    ad.add_data_validation(dv)
    dv.add(rng)
widths(ad, {"A": 15, "B": 30, "C": 34, "D": 11, "E": 11, "F": 12, "G": 13, "H": 11, "I": 12, "J": 11,
            "K": 13, "L": 13, "M": 12, "N": 11})
ad.row_dimensions[7].height = 42
ad.freeze_panes = "A8"
A = {"dept": "A", "cust": "B", "acct": "B", "amt": "G", "cat": "H", "land": "J", "fwd": "K", "full": "L", "stress": "M", "wd": "N", "month": "F"}
AR = lambda k: f"Adjustments!${A[k]}${AF}:${A[k]}${AL}"

# ---------------------------------------------------------------- Budget
bu = wb.create_sheet("Budget")
title(bu, "Budget: revenue by department", "Optional. Paste or type the FY budgeted revenue. It drives the gap to budget and the risk rating.")
head(bu, 4, 1, ["Department"] + [None] * 12 + ["FY total"])
for i in range(12):
    c = bu.cell(row=4, column=2 + i, value=f"={CL(4)}{12+i}".replace("D", "Settings!$D$"))
    c.number_format = MON
for i, d in enumerate(DEPTS):
    r = 5 + i
    put(bu, r, 1, f"=Settings!$A${12+i}")
    for c in range(2, 14):
        put(bu, r, c, None, MONEY, fill=INPUT, color="0000FF")
    put(bu, r, 14, f"=SUM(B{r}:M{r})", MONEY, bold=True)
put(bu, 12, 1, "Total", bold=True, fill=TOTAL)
for c in range(2, 15):
    put(bu, 12, c, f"=SUM({CL(c)}5:{CL(c)}11)", MONEY, bold=True, fill=TOTAL)
widths(bu, {"A": 16, **{CL(c): 12 for c in range(2, 15)}})

# ---------------------------------------------------------------- Monthly Outlook
mo = wb.create_sheet("Monthly Outlook")
title(mo, "Monthly Outlook by department",
      "Months up to the last P&L month show actuals. Later months show Zoho backlog plus weighted pipeline, plus Adjustments.")
put(mo, 4, 1, "Month", bold=True, border=False)
put(mo, 5, 1, "Actual or forecast", bold=True, border=False)
for i in range(12):
    c = 3 + i
    put(mo, 4, c, f"=Settings!$D${12+i}", MON, bold=True, fill=BAND)
    put(mo, 5, c, f'=IF({CL(c)}$4<={ASAT},"Actual","Forecast")', bold=False, fill=BAND)
put(mo, 4, 15, "FY total", bold=True, fill=BAND)
MCOLS = [CL(3 + i) for i in range(12)]
blocks = {}
row = 7


def block(name, sub, cell_formula):
    """Write one department block. cell_formula(dept_cell, month_cell, col) -> formula."""
    global row
    put(mo, row, 1, name, bold=True, color="FFFFFF", fill=HEAD)
    put(mo, row, 2, sub, color="FFFFFF", fill=HEAD, italic=True)
    for c in range(3, 16):
        put(mo, row, c, None, fill=HEAD)
    first = row + 1
    for i, d in enumerate(DEPTS):
        r = first + i
        put(mo, r, 1, f"=Settings!$A${12+i}")
        put(mo, r, 2, None)
        for col in MCOLS:
            put(mo, r, MCOLS.index(col) + 3, cell_formula(f"$A{r}", f"{col}$4", col, i), MONEY)
        put(mo, r, 15, f"=SUM(C{r}:N{r})", MONEY, bold=True)
    tr = first + len(DEPTS)
    put(mo, tr, 1, "Total", bold=True, fill=TOTAL)
    put(mo, tr, 2, None, fill=TOTAL)
    for c in range(3, 16):
        put(mo, tr, c, f"=SUM({CL(c)}{first}:{CL(c)}{tr-1})", MONEY, bold=True, fill=TOTAL)
    blocks[name] = first
    row = tr + 2


PLD = "'P&L Dump'"
block("Actual revenue", "From P&L Dump, Income accounts",
      lambda d, m, col, i: (f"=IFERROR(SUMIFS(INDEX({PLD}!$B${PF}:$Z${PL},0,MATCH({m},{PLD}!$B$5:$Z$5,0)),"
                            f"{PLD}!$AC${PF}:$AC${PL},\"Income\",{PLD}!$AD${PF}:$AD${PL},{d}),0)"))
block("Backlog", "Signed. Zoho Closed Won after the last P&L month, plus Adjustments",
      lambda d, m, col, i: (f"=SUMIFS({Z('fwd')},{Z('dept')},{d},{Z('land')},{m},{Z('cat')},\"Backlog\")"
                            f"+SUMIFS({AR('fwd')},{AR('dept')},{d},{AR('land')},{m},{AR('cat')},\"Backlog\")"))


def both(key, *crit):
    """SUMIFS over the Zoho Dump plus the same SUMIFS over Adjustments."""
    z = ",".join(f"{Z(k)},{v}" for k, v in crit)
    a = ",".join(f"{AR(k)},{v}" for k, v in crit)
    return f"SUMIFS({Z(key)},{z})+SUMIFS({AR(key)},{a})"


def br(name, i):
    return blocks[name] + i


block("Pipeline (weighted)", "Open Zoho deals times stage probability, plus Adjustments",
      lambda d, m, col, i: "=" + both("fwd", ("dept", d), ("land", m), ("cat", '"Pipeline"'))
      + "+" + both("fwd", ("dept", d), ("land", m), ("cat", '"Overdue"')))
block("Forecast", "Actuals to date, then backlog plus weighted pipeline",
      lambda d, m, col, i: (f"=IF({m}<={ASAT},{col}{br('Actual revenue', i)},"
                            f"{col}{br('Backlog', i)}+{col}{br('Pipeline (weighted)', i)})"))
block("If pipeline falls through", "Actuals, then backlog plus pipeline at the stress case rate on the Dashboard",
      lambda d, m, col, i: f"=IF({m}<={ASAT},{col}{br('Actual revenue', i)}," + both("stress", ("dept", d), ("land", m)) + ")")
block("Budget", "From the Budget sheet",
      lambda d, m, col, i: f"=Budget!{CL(MCOLS.index(col) + 2)}{5 + i}")
block("Forecast less budget", "Negative means short of budget",
      lambda d, m, col, i: f"={col}{br('Forecast', i)}-{col}{br('Budget', i)}")
widths(mo, {"A": 26, "B": 30, **{c: 12 for c in MCOLS}, "O": 14})
mo.freeze_panes = "C6"

# ---------------------------------------------------------------- Customer Risk
cr = wb.create_sheet("Customer Risk")
title(cr, "Customer Risk", "Every customer in the Zoho export. Forward revenue means backlog plus weighted pipeline for the rest of this FY.")
ccols = ["#", "Customer", "Won this FY to date (Zoho)", "Backlog this FY", "Pipeline full value",
         "Pipeline weighted", "Forward revenue", "Share of all forward revenue", "Lost this FY (Zoho)",
         "Withdrawn in scenario"] + [f"Forward: {d}" for d in DEPTS] + ["Sort key"]
head(cr, 4, 1, ccols)
cr.row_dimensions[4].height = 45
CF, CLR = 5, 5 + CUST_ROWS - 1
for n in range(1, CUST_ROWS + 1):
    r = CF + n - 1
    b = f"$B{r}"
    fm = {
        1: n,
        2: f'=IFERROR(INDEX({Z("acct")},MATCH(A{r},{Z("cno")},0)),"")',
        3: f'=IF({b}="",0,SUMIFS({Z("amt")},{Z("acct")},{b},{Z("cat")},"Delivered",{Z("close")},">="&{FYSTART}))',
        4: f'=IF({b}="",0,' + both("fwd", ("acct", b), ("cat", '"Backlog"')) + ")",
        5: f'=IF({b}="",0,' + (both("full", ("acct", b), ("cat", '"Pipeline"')) + "+" + both("full", ("acct", b), ("cat", '"Overdue"'))) + ")",
        6: f'=IF({b}="",0,' + (both("fwd", ("acct", b), ("cat", '"Pipeline"')) + "+" + both("fwd", ("acct", b), ("cat", '"Overdue"'))) + ")",
        7: f"=D{r}+F{r}",
        8: f"=IF(SUM($G${CF}:$G${CLR})=0,0,G{r}/SUM($G${CF}:$G${CLR}))",
        9: f'=IF({b}="",0,SUMIFS({Z("amt")},{Z("acct")},{b},{Z("cat")},"Lost",{Z("close")},">="&{FYSTART}))',
        10: f'=IF({b}="","",IF(COUNTIF(Dashboard!$C$13:$C$22,{b})>0,"Yes",""))',
    }
    for j, d in enumerate(DEPTS):
        fm[11 + j] = (f'=IF({b}="",0,' + both("fwd", ("acct", b), ("dept", f"Settings!$A${12+j}")) + ")")
    fm[18] = f'=IF(G{r}>0,G{r}+A{r}/10000000,"")'
    for c, v in fm.items():
        fmt = PCT if c == 8 else (None if c in (1, 2, 10) else MONEY)
        if c == 18:
            fmt = "0.0000000"
        put(cr, r, c, v, fmt, color="000000")
put(cr, CLR + 1, 2, "Total", bold=True, fill=TOTAL)
for c in list(range(3, 10)) + list(range(11, 18)):
    put(cr, CLR + 1, c, f"=SUM({CL(c)}{CF}:{CL(c)}{CLR})", PCT if c == 8 else MONEY, bold=True, fill=TOTAL)
widths(cr, {"A": 5, "B": 34, **{CL(c): 13 for c in range(3, 19)}})
cr.column_dimensions["R"].hidden = True
cr.freeze_panes = "C5"
CR = lambda col: f"'Customer Risk'!${col}${CF}:${col}${CLR}"

# ---------------------------------------------------------------- Dashboard
db = wb.create_sheet("Dashboard")
title(db, "Revenue Risk Profile")
db["A2"] = f'="Financial year "&TEXT({FYSTART},"dd/mm/yyyy")&" to "&TEXT({FYEND},"dd/mm/yyyy")&". P&L actuals to "&TEXT({ASAT},"dd/mm/yyyy")&". Zoho export: "&MID(\'Zoho Dump\'!$A$3,FIND(" on ",\'Zoho Dump\'!$A$3)+4,40)'
db["A2"].font = Font(name=ARIAL, size=9, color=MUTED)
T0, T1 = 26, 26 + len(DEPTS) - 1     # department rows
TT = T1 + 1                           # total row
db["A4"] = f'="Base forecast for the year is "&TEXT(F{TT},"$#,##0")&". Signed so far (actuals plus backlog) is "&TEXT(B{TT}+C{TT},"$#,##0")&"."'
db["A5"] = f'="If no more pipeline comes through than the stress case below, revenue falls by "&TEXT(L{TT},"$#,##0")&"."'
db["A6"] = f'="The selected customers withdrawing would cost "&TEXT(M{TT},"$#,##0")&" of forward revenue. Both together leave "&TEXT(O{TT},"$#,##0")&" for the year."'
db["A7"] = f'="Largest single customer: "&B37&", "&TEXT(F37,"0.0%")&" of all forward revenue. Top 10 customers: "&TEXT(G46,"0.0%")&"."'
for r in range(4, 8):
    db[f"A{r}"].font = Font(name=ARIAL, size=11, bold=(r == 4), color="000000")

put(db, 9, 1, "Scenario levers (yellow cells)", bold=True, color="FFFFFF", fill=HEAD)
for c in range(2, 4):
    put(db, 9, c, None, fill=HEAD)
put(db, 10, 1, "Stress case: share of pipeline that still comes through")
put(db, 10, 2, None)
put(db, 10, 3, 0, PCT, fill=INPUT, color="0000FF")
put(db, 11, 1, "0% means no open deal lands. Backlog is kept in full.", italic=True, color=MUTED, border=False)
put(db, 12, 1, "Customers who withdraw (pick from the list)", bold=True, border=False)
for k in range(10):
    put(db, 13 + k, 2, k + 1)
    put(db, 13 + k, 3, None, fill=INPUT, color="0000FF")
dv = DataValidation(type="list", formula1=CR("B"), allow_blank=True)
db.add_data_validation(dv)
dv.add("C13:C22")
note(db, 23, "A withdrawal removes that customer's backlog and pipeline for the rest of the year. It does not touch actuals already earned, or revenue that is not in Zoho or Adjustments.")

dcols = ["Department", "Actual to date (P&L)", "Backlog, signed", "Pipeline, full value", "Pipeline, weighted",
         "Base forecast for FY", "Budget for FY", "Forecast less budget", "Signed as % of budget",
         "Pipeline share of forward revenue", "If pipeline falls through", "Revenue at risk from pipeline",
         "Lost if selected customers withdraw", "Worst case: both", "Largest customer",
         "Largest customer share", "Overdue Zoho deals, full value", "Won to date in Zoho (reference)", "Risk rating"]
# columns: A..S  (O = worst case moved to N? keep order below)
dcols = ["Department", "Actual to date (P&L)", "Backlog, signed", "Pipeline, full value", "Pipeline, weighted",
         "Base forecast for FY", "Budget for FY", "Forecast less budget", "Signed as % of budget",
         "Pipeline share of forward revenue", "If pipeline falls through", "Revenue at risk from pipeline",
         "Lost if selected customers withdraw", "Forecast after withdrawals", "Worst case: both",
         "Largest customer", "Largest customer share", "Overdue Zoho deals, full value",
         "Won to date in Zoho (reference)", "Risk rating"]
head(db, T0 - 1, 1, dcols)
db.row_dimensions[T0 - 1].height = 58
MO = "'Monthly Outlook'"
for i, d in enumerate(DEPTS):
    r = T0 + i
    dref = f"$A{r}"
    act, bkl, pip, fc, flo, bud = (br(n, i) for n in ("Actual revenue", "Backlog", "Pipeline (weighted)",
                                                      "Forecast", "If pipeline falls through", "Budget"))
    ccol = CL(11 + i)
    fm = {
        1: f"=Settings!$A${12+i}",
        2: f'=SUMIFS({MO}!$C${act}:$N${act},{MO}!$C$5:$N$5,"Actual")',
        3: f"={MO}!$O${bkl}",
        4: "=" + both("full", ("dept", dref), ("cat", '"Pipeline"')) + "+" + both("full", ("dept", dref), ("cat", '"Overdue"')),
        5: f"={MO}!$O${pip}",
        6: f"={MO}!$O${fc}",
        7: f"={MO}!$O${bud}",
        8: f'=IF(G{r}=0,"No budget",F{r}-G{r})',
        9: f'=IF(G{r}=0,"No budget",(B{r}+C{r})/G{r})',
        10: f"=IF(C{r}+E{r}=0,0,E{r}/(C{r}+E{r}))",
        11: f"={MO}!$O${flo}",
        12: f"=F{r}-K{r}",
        13: "=" + both("fwd", ("dept", dref), ("wd", "1")),
        14: f"=F{r}-M{r}",
        15: f"=K{r}-(" + both("stress", ("dept", dref), ("wd", "1")) + ")",
        16: f'=IF(MAX({CR(ccol)})=0,"-",INDEX({CR("B")},MATCH(MAX({CR(ccol)}),{CR(ccol)},0)))',
        17: f"=IF(SUM({CR(ccol)})=0,0,MAX({CR(ccol)})/SUM({CR(ccol)}))",
        18: "=" + both("amt", ("dept", dref), ("cat", '"Overdue"')),
        19: f'=SUMIFS({Z("amt")},{Z("dept")},{dref},{Z("cat")},"Delivered",{Z("close")},">="&{FYSTART})',
        20: f'=IF(AND(B{r}=0,F{r}=0),"-",CHOOSE(MAX(V{r}:X{r})+1,"Low","Medium","High"))',
        22: f'=IF(G{r}=0,0,IF(F{r}<G{r}*(1-Settings!$C$54),2,IF(F{r}<G{r}*(1-Settings!$B$54),1,0)))',
        23: f"=IF(Q{r}>=Settings!$C$55,2,IF(Q{r}>=Settings!$B$55,1,0))",
        24: f"=IF(J{r}>=Settings!$C$56,2,IF(J{r}>=Settings!$B$56,1,0))",
    }
    for c, v in fm.items():
        fmt = PCT if c in (9, 10, 17) else (None if c in (1, 16, 20) else MONEY)
        cell = put(db, r, c, v, fmt, bold=(c in (1, 20)))
        if c >= 22:
            cell.font = Font(name=ARIAL, size=8, color=MUTED)
put(db, TT, 1, "Total", bold=True, fill=TOTAL)
for c in range(2, 21):
    L_ = CL(c)
    if c in (2, 3, 4, 5, 6, 7, 11, 12, 13, 14, 15, 18, 19):
        v = f"=SUM({L_}{T0}:{L_}{T1})"
    elif c == 8:
        v = f'=IF(G{TT}=0,"No budget",F{TT}-G{TT})'
    elif c == 9:
        v = f'=IF(G{TT}=0,"No budget",(B{TT}+C{TT})/G{TT})'
    elif c == 10:
        v = f"=IF(C{TT}+E{TT}=0,0,E{TT}/(C{TT}+E{TT}))"
    elif c == 16:
        v = "=B37"
    elif c == 17:
        v = "=F37"
    else:
        v = None
    put(db, TT, c, v, PCT if c in (9, 10, 17) else MONEY, bold=True, fill=TOTAL)
for col, txt in (("V", "Budget score"), ("W", "Concentration score"), ("X", "Pipeline score")):
    db[f"{col}{T0-1}"] = txt
    db[f"{col}{T0-1}"].font = Font(name=ARIAL, size=8, color=MUTED)
rng = f"T{T0}:T{T1}"
db.conditional_formatting.add(rng, FormulaRule(formula=[f'$T{T0}="High"'], fill=PatternFill("solid", fgColor="F8CBAD"), font=Font(name=ARIAL, bold=True, color="9C0006")))
db.conditional_formatting.add(rng, FormulaRule(formula=[f'$T{T0}="Medium"'], fill=PatternFill("solid", fgColor="FFE699")))
db.conditional_formatting.add(rng, FormulaRule(formula=[f'$T{T0}="Low"'], fill=PatternFill("solid", fgColor="C6EFCE")))

# Top 10 customers
put(db, 35, 1, "Top 10 customers by forward revenue", bold=True, color="FFFFFF", fill=HEAD)
head(db, 36, 1, ["Rank", "Customer", "Backlog, signed", "Pipeline, weighted", "Forward revenue",
                 "Share of all forward revenue", "Cumulative share", "Withdrawn in scenario"])
for k in range(1, 11):
    r = 36 + k
    m = f'MATCH(LARGE({CR("R")},{k}),{CR("R")},0)'
    fm = {1: k,
          2: f'=IFERROR(INDEX({CR("B")},{m}),"")',
          3: f'=IFERROR(INDEX({CR("D")},{m}),0)',
          4: f'=IFERROR(INDEX({CR("F")},{m}),0)',
          5: f'=IFERROR(INDEX({CR("G")},{m}),0)',
          6: f'=IFERROR(INDEX({CR("H")},{m}),0)',
          7: f"=SUM($F$37:F{r})",
          8: f'=IFERROR(INDEX({CR("J")},{m}),"")'}
    for c, v in fm.items():
        put(db, r, c, v, PCT if c in (6, 7) else (None if c in (1, 2, 8) else MONEY))

# Checks
put(db, 49, 1, "Data checks", bold=True, color="FFFFFF", fill=HEAD)
head(db, 50, 1, ["Check", "Result", "Status"])
zr = "'Zoho Dump'"
checks = [
    ("Zoho column headings found", f"=COUNTIF({zr}!$L$6:$R$6,\">0\")", "=IF(B{r}=6,\"OK\",\"A Zoho heading has changed\")"),
    ("Zoho deals read", f'=COUNTIF({Z("deal")},"?*")',
     f'=IFERROR(IF(B{{r}}=VALUE(TRIM(MID({zr}!$A$5,FIND(":",{zr}!$A$5)+1,20))),"OK, matches Zoho record count","Does not match Zoho record count"),"Check")'),
    ("Deals whose Zoho department is not mapped", f'=COUNTIF({Z("dept")},"UNMAPPED")', '=IF(B{r}=0,"OK","Add it on Settings")'),
    ("Deals whose stage is not in Settings", f'=COUNTIF({Z("stype")},"UNMAPPED")', '=IF(B{r}=0,"OK","Add it on Settings")'),
    ("Deals with no closing date", f'=COUNTIF({Z("cat")},"NO DATE")', '=IF(B{r}=0,"OK","Ask sales to fix in Zoho")'),
    ("Open deals with zero or blank amount", f'=COUNTIFS({Z("stype")},"Open",{Z("amt")},0)', '=IF(B{r}=0,"OK","Pipeline is understated")'),
    ("Open deals past their closing date", f'=COUNTIF({Z("cat")},"Overdue")', f'=IF(B{{r}}=0,"OK",IF({INCOVER}="Yes","Counted in forecast","Not counted. Ask sales to update"))'),
    ("P&L months picked in row 5", "=COUNT('P&L Dump'!$B$5:$Z$5)", '=IF(B{r}>0,"OK","No P&L pasted yet")'),
    ("Actual months with no P&L column", f"=SUMPRODUCT(({MONTH_RNG}<={ASAT})*(COUNTIF('P&L Dump'!$B$5:$Z$5,{MONTH_RNG})=0))", '=IF(B{r}=0,"OK","Paste or pick the missing months")'),
    ("P&L rows not in the Account Map", f"=SUM('P&L Dump'!$AE${PF}:$AE${PL})", '=IF(B{r}=0,"OK","Add the account to Account Map")'),
    ("Customers in the Zoho export", f"=MAX({Z('cno')})", f'=IF(B{{r}}<={CUST_ROWS},"OK","More than {CUST_ROWS}. Extend Customer Risk")'),
]
for i, (lab, v, s_) in enumerate(checks):
    r = 51 + i
    put(db, r, 1, lab)
    put(db, r, 2, v, "#,##0")
    put(db, r, 3, s_.replace("{r}", str(r)))
db.conditional_formatting.add("C51:C61", FormulaRule(formula=['LEFT($C51,2)<>"OK"'], font=Font(name=ARIAL, bold=True, color="C00000")))
widths(db, {"A": 30, "B": 24, **{CL(c): 13 for c in range(3, 21)}, "P": 26, "T": 11})
db.column_dimensions["B"].width = 26
for col in ("V", "W", "X"):
    db.column_dimensions[col].hidden = True
db.freeze_panes = "B1"

# ---------------------------------------------------------------- Read Me
rm = wb.create_sheet("Read Me")
title(rm, "How to use the Revenue Risk Profile")
lines = [
    ("What it shows", True),
    ("How much of this year's revenue is already earned, how much is signed, and how much still depends on pipeline. "
     "It shows what happens if pipeline does not come through, or if certain customers walk away. It does this for each department.", False),
    ("", False),
    ("Each month", True),
    ("1. Settings: change the last month of P&L actuals to the new month end.", False),
    ("2. P&L Dump: clear A7 to Z1006. Paste the Xero Profit and Loss from cell A7, account names in column A, months across. "
     "In row 5, pick the month for each pasted column.", False),
    ("3. Zoho Dump: clear columns A to J. Run the Zoho report All Deals by Stage and paste the whole export into cell A1, exactly as it comes out.", False),
    ("4. Adjustments: add signed or likely revenue that is not in Zoho, such as recurring support contracts.", False),
    ("5. Dashboard: check the Data checks box at the bottom. Every line should say OK.", False),
    ("6. Dashboard: pick customers to withdraw and set the stress case to show management the risk.", False),
    ("", False),
    ("Colours", True),
    ("Yellow cells with blue text are inputs. Everything else is a formula. Do not type over formulas.", False),
    ("", False),
    ("How the numbers are built", True),
    ("Actual: Income accounts in the P&L, by department, from the Account Map.", False),
    ("Backlog: Zoho Closed Won deals with a closing date after the last P&L month, in the month they close.", False),
    ("Pipeline: open Zoho deals. Full value, and weighted by the stage probability on Settings.", False),
    ("Base forecast: actuals for past months, then backlog plus weighted pipeline for future months.", False),
    ("If pipeline falls through: actuals plus backlog, plus pipeline at the stress case rate. At 0 percent, no open deal lands.", False),
    ("Customer withdrawal: removes that customer's backlog and weighted pipeline for the rest of the year.", False),
    ("Risk rating: the worst of three tests on Settings. Forecast against budget, largest customer share, and how much of forward revenue is still pipeline.", False),
    ("", False),
    ("Things to confirm", True),
    ("Stage probabilities and risk thresholds on Settings are placeholders. Management should set them.", False),
    ("Zoho department Support is mapped to Onsite, and Multi-Service to Unallocated. Both are assumptions. Please confirm.", False),
    ("Most income accounts have no department in the current account list, so they show as Unallocated. Assign them on the Account Map if they belong to one team.", False),
    ("It is not known whether Zoho amounts include GST. The P&L is ex GST. Please confirm with sales.", False),
    ("Revenue that is never entered in Zoho, like recurring contracts, only shows if it is added on Adjustments.", False),
]
for i, (t, bold) in enumerate(lines):
    c = rm.cell(row=3 + i, column=1, value=t)
    c.font = Font(name=ARIAL, size=11 if bold else 10, bold=bold, color=NAVY if bold else "000000")
    c.alignment = Alignment(wrap_text=True, vertical="top")
rm.column_dimensions["A"].width = 120

order = ["Read Me", "Dashboard", "Monthly Outlook", "Customer Risk", "Zoho Dump", "P&L Dump",
         "Adjustments", "Budget", "Account Map", "Settings"]
wb._sheets = [wb[n] for n in order]
for n in ("Zoho Dump", "P&L Dump", "Adjustments", "Budget"):
    wb[n].sheet_properties.tabColor = "FFC000"
wb["Dashboard"].sheet_properties.tabColor = NAVY
wb.active = 1
wb.save(OUT)
print("saved", OUT)
