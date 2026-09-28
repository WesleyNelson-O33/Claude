"""Builds the Revenue Risk Profile workbook from a (recalculated) FY27 Revenue Tracker.

Two data sheets hold pasted values from the tracker:
  Data - Revenue   <- Finance!A6:AB4506   (actual invoiced revenue)
  Data - Pipeline  <- Work Won!A4:R604    (won / lost / open deals from the Zoho dump)
Every other sheet is formulas off those two, so a monthly refresh is two copy / paste-values.

Usage: python build_risk_profile.py <tracker_recalculated.xlsx> <out.xlsx> <as_at yyyy-mm-dd>
"""
import datetime as dt
import sys

import openpyxl
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as CL
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.table import Table, TableStyleInfo

TRACKER, OUT, AS_AT = sys.argv[1:4]
AS_AT = dt.datetime.strptime(AS_AT, "%Y-%m-%d")

FONT = "Arial"
NAVY, SLATE, ACCENT, TINT, CALC, FIN = "1F3864", "1F4E79", "BF8F00", "D9E2F3", "F4F7FB", "FFF2CC"
TXT, MUTED, LINE = "262626", "595959", "C9D3E3"
# chart series: validated categorical slots 1-3 (blue, orange, aqua) - fixed order
S1, S2, S3 = "2A78D6", "EB6834", "1BAF7A"
# status colours - always paired with the word
ST = {"High": "D03B3B", "Medium": "EC835A", "Low": "FAB219", "OK": "0CA30C"}
MONEY = '$#,##0;[Red]-$#,##0;"-"'
PCT = '0.0%;-0.0%;"-"'
THIN = Side(style="thin", color=LINE)
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
ROWL = Border(bottom=Side(style="hair", color=LINE))
F = lambda **k: Font(name=FONT, size=k.pop("size", 10), color=k.pop("color", TXT), **k)
FILL = lambda c: PatternFill("solid", fgColor=c)

REV_N, PIPE_N = 4501, 601            # data rows incl. header
wb = openpyxl.Workbook()
wb.remove(wb.active)
wb.calculation.fullCalcOnLoad = True


def name(n, ref):
    wb.defined_names[n] = DefinedName(n, attr_text=ref)


def title(ws, t, sub, cols=14):
    for c in range(1, cols + 1):
        for r in (1, 2):
            ws.cell(r, c).fill = FILL(NAVY)
    ws["A1"], ws["A2"] = t, sub
    ws["A1"].font = F(size=18, bold=True, color="FFFFFF")
    ws["A2"].font = F(size=10, italic=True, color=TINT)
    ws.row_dimensions[1].height = 30
    ws.sheet_view.showGridLines = False


def hdr(ws, r, c, text, fill=SLATE):
    x = ws.cell(r, c, text)
    x.font = F(bold=True, color="FFFFFF")
    x.fill = FILL(fill)
    x.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    x.border = BOX
    return x


def body(x, fmt=None, bold=False, fill=None):
    x.font = F(bold=bold)
    x.border = ROWL
    if fmt:
        x.number_format = fmt
    if fill:
        x.fill = FILL(fill)
    return x


def section(ws, r, text, cols=14):
    ws.cell(r, 1, text).font = F(size=11, bold=True, color="FFFFFF")
    for c in range(1, cols + 1):
        ws.cell(r, c).fill = FILL(NAVY)
    ws.row_dimensions[r].height = 20


def rating_fmt(ws, rng):
    first = rng.split(":")[0]
    for lab, col in ST.items():
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{first}="{lab}"'], fill=FILL(col),
                                                       font=Font(name=FONT, bold=True, color="FFFFFF" if lab in ("High", "OK") else TXT)))


# ======================================================================= Settings
st = wb.create_sheet("Settings")
title(st, "Settings", "Thresholds the risk ratings use. Change the cream cells - every sheet follows.", 6)
settings = [
    ("set_AsAt", "Data as at", AS_AT, "dd-mmm-yy", "The date the tracker data was taken. Used to decide what is overdue or stale."),
    ("set_FYStart", "Financial year starts", dt.datetime(2026, 7, 1), "dd-mmm-yy", "First day of the financial year (1 July)."),
    ("set_DealShare", "Single deal - share of open pipeline", 0.10, "0%", "An open deal worth this share of all open pipeline or more is a concentration risk."),
    ("set_BigValue", "Large deal value", 50000, MONEY, "Deals at or above this value count as large."),
    ("set_EarlyProb", "Early stage - probability at or below", 0.25, "0%", "A large deal this early is a forecast risk."),
    ("set_ClientHigh", "Top client share of revenue - High", 0.20, "0%", "One client above this share of FYTD revenue = High concentration."),
    ("set_Top5High", "Top 5 clients share - High", 0.60, "0%", "Top five clients above this share of FYTD revenue = High concentration."),
    ("set_ScoreHigh", "Risk score for High", 5, "0", "Risk Register: score at or above this = High."),
    ("set_ScoreMed", "Risk score for Medium", 3, "0", "Score at or above this (and below High) = Medium. Anything above 0 = Low."),
]
hdr(st, 4, 1, "Setting")
hdr(st, 4, 2, "Value", ACCENT)
hdr(st, 4, 3, "What it does")
for i, (n, lab, v, fmt, note) in enumerate(settings):
    r = 5 + i
    body(st.cell(r, 1, lab), bold=True)
    body(st.cell(r, 2, v), fmt, fill=FIN)
    body(st.cell(r, 3, note))
    name(n, f"Settings!$B${r}")
st.cell(15, 1, "Probabilities by pipeline stage come from the tracker (Lists AA:AB) and arrive already applied in the "
               "Work Won data. They are assumptions, not Zoho data.").font = F(italic=True, color=MUTED)
st.column_dimensions["A"].width = 38
st.column_dimensions["B"].width = 14
st.column_dimensions["C"].width = 90

# =================================================================== Data sheets
src = openpyxl.load_workbook(TRACKER, data_only=True)


def data_sheet(nm, src_ws, top, n_rows, n_cols, helpers, tname):
    ws = wb.create_sheet(nm)
    for i, row in enumerate(src_ws.iter_rows(min_row=top, max_row=top + n_rows - 1, max_col=n_cols, values_only=True)):
        for j, v in enumerate(row):
            if v is not None and v != "":
                ws.cell(1 + i, 1 + j, v)
    for c in range(1, n_cols + 1):
        x = ws.cell(1, c)
        x.font = F(bold=True, color="FFFFFF")
        x.fill = FILL(ACCENT)
        ws.column_dimensions[CL(c)].width = 13
    h0 = n_cols + 2
    for k, (h, fx, fmt) in enumerate(helpers):
        col = h0 + k
        L = CL(col)
        hdr(ws, 1, col, h)
        ws.column_dimensions[L].width = 14
        for r in range(2, n_rows + 1):
            x = ws.cell(r, col, fx(r))
            x.font = F(size=9, color=MUTED)
            if fmt:
                x.number_format = fmt
    ws.freeze_panes = "A2"
    return ws, h0


def col_of(head_range, h, r, last):
    return f'INDEX($A{r}:${last}{r},MATCH("{h}",${head_range}$1:${last}$1,0))'


# --- revenue: Finance A6:AB
fin = src["Finance"]
RL = "AB"
rv = lambda h, r: f'IFERROR({col_of("A", h, r, RL)},"")'
rev_helpers = [
    ("Amount", lambda r: f'=IFERROR(N({rv("Xero Invoiced Ex GST", r)}),0)', MONEY),                          # AD
    ("Invoice Date", lambda r: f'={rv("Xero Invoice Date", r)}', "dd-mmm-yy"),                              # AE
    ("Counted", lambda r: f'=IF(AND({rv("Xero Invoice No", r)}&""<>"",ISNUMBER(AE{r}),AD{r}<>0),1,0)', "0"),  # AF
    ("Month", lambda r: f'=IF(AF{r}=1,EOMONTH(AE{r},0),"")', "mmm-yy"),                                     # AG
    ("Client", lambda r: f'=IF(AF{r}=1,UPPER(TRIM({rv("Client", r)}&"")),"")', None),                       # AH
    ("Cost Centre", lambda r: f'=IF(AF{r}=1,{rv("Cost Centre", r)}&"","")', None),                           # AI
    ("In this FY to date", lambda r: f'=IF(AND(AF{r}=1,AE{r}>=set_FYStart,AE{r}<=set_AsAt),1,0)', "0"),       # AJ
    ("New client no.", lambda r: f'=IF(AND(AJ{r}=1,AH{r}<>"",COUNTIFS($AH$2:AH{r},AH{r},$AJ$2:AJ{r},1)=1),N(AL{r - 1})+1,"")', "0"),  # AK
    ("Running", lambda r: f'=IF(AK{r}<>"",AK{r},N(AL{r - 1}))', "0"),                                       # AL
]
drev, _ = data_sheet("Data - Revenue", fin, 6, REV_N, 28, rev_helpers, "tbl_Rev")
for n, L in (("rev_Amt", "AD"), ("rev_Date", "AE"), ("rev_Cnt", "AF"), ("rev_Month", "AG"), ("rev_Client", "AH"),
             ("rev_CC", "AI"), ("rev_FY", "AJ"), ("rev_New", "AK")):
    name(n, f"'Data - Revenue'!${L}$2:${L}${REV_N}")

# --- pipeline: Work Won A4:R
won = src["Work Won"]
PL = "R"
pv = lambda h, r: f'IFERROR({col_of("A", h, r, PL)},"")'
pipe_helpers = [
    ("Status", lambda r: f'={pv("Status", r)}&""', None),                                                  # T
    ("Value", lambda r: f'=IFERROR(N({pv("Value Ex GST", r)}),0)', MONEY),                                  # U
    ("Has value", lambda r: f'=IF(AND(T{r}<>"",{pv("Value Ex GST", r)}&""<>""),1,0)', "0"),                 # V
    ("Probability", lambda r: f'=IFERROR(N({pv("Probability", r)}),0)', "0%"),                             # W
    ("Weighted", lambda r: f'=IF(T{r}="Open",U{r}*W{r},0)', MONEY),                                        # X
    ("Still to Invoice", lambda r: f'=IF(T{r}="Won",IFERROR(N({pv("Still to Invoice", r)}),0),0)', MONEY),  # Y
    ("Expected Month", lambda r: f'=IFERROR(EOMONTH({pv("Expected Invoice Month", r)},0),"")', "mmm-yy"),  # Z
    ("Cost Centre", lambda r: f'={pv("Cost Centre", r)}&""', None),                                        # AA
    ("Stage", lambda r: f'={pv("Pipeline Stage", r)}&""', None),                                            # AB
    ("Client", lambda r: f'={pv("Client", r)}&""', None),                                                   # AC
    ("Deal", lambda r: f'={pv("Description", r)}&""', None),                                                # AD
    ("Job", lambda r: f'={pv("Job Number", r)}&""', None),                                                  # AE
    # risk flags
    ("Stale", lambda r: f'=IF(AND(T{r}="Open",ISNUMBER(Z{r}),Z{r}<EOMONTH(set_AsAt,-1)+1),3,0)', "0"),       # AF
    ("Concentration", lambda r: f'=IF(AND(T{r}="Open",U{r}>0,U{r}>=set_DealShare*SUMIFS($U$2:$U${PIPE_N},$T$2:$T${PIPE_N},"Open")),3,0)', "0"),  # AG
    ("Big and early", lambda r: f'=IF(AND(T{r}="Open",U{r}>=set_BigValue,W{r}<=set_EarlyProb),2,0)', "0"),  # AH
    ("No value", lambda r: f'=IF(AND(OR(T{r}="Open",T{r}="Won"),V{r}=0),2,0)', "0"),                       # AI
    ("Won - billing overdue", lambda r: f'=IF(AND(T{r}="Won",Y{r}>0.5,ISNUMBER(Z{r}),Z{r}<EOMONTH(set_AsAt,-1)+1),3,0)', "0"),  # AJ
    ("Over-invoiced", lambda r: f'=IF(AND(T{r}="Won",Y{r}<-1),1,0)', "0"),                                 # AK
    ("Score", lambda r: f'=SUM(AF{r}:AK{r})', "0"),                                                         # AL
    ("Rating", lambda r: f'=IF(AL{r}>=set_ScoreHigh,"High",IF(AL{r}>=set_ScoreMed,"Medium",IF(AL{r}>0,"Low","OK")))', None),  # AM
    ("On register", lambda r: f'=IF(OR(T{r}="Open",AND(T{r}="Won",ABS(Y{r})>0.5)),1,0)', "0"),              # AN
    ("At risk $", lambda r: f'=IF(T{r}="Open",IF(V{r}=0,0,X{r}),IF(T{r}="Won",MAX(0,Y{r}),0))', MONEY),      # AO
    ("Sort key", lambda r: f'=IF(AN{r}=1,AL{r}*1E+9+U{r}+ROW()/1E+6,"")', "0"),                           # AP
    ("Flags", lambda r: (f'=MID(IF(AF{r}>0,"; Expected month passed - update Zoho","")&IF(AG{r}>0,"; "&TEXT(U{r}/MAX(1,SUMIFS($U$2:$U${PIPE_N},$T$2:$T${PIPE_N},"Open")),"0%")&" of open pipeline in one deal","")'
                         f'&IF(AH{r}>0,"; Large deal at an early stage","")&IF(AI{r}>0,"; No amount in Zoho","")'
                         f'&IF(AJ{r}>0,"; Won, invoice month passed, not fully invoiced","")&IF(AK{r}>0,"; Xero invoiced more than the deal value",""),3,300)'), None),  # AQ
]
dpipe, _ = data_sheet("Data - Pipeline", won, 4, PIPE_N, 18, pipe_helpers, "tbl_Pipe")
for n, L in (("p_Status", "T"), ("p_Value", "U"), ("p_Has", "V"), ("p_Prob", "W"), ("p_Weighted", "X"), ("p_Still", "Y"),
             ("p_Month", "Z"), ("p_CC", "AA"), ("p_Stage", "AB"), ("p_Client", "AC"), ("p_Deal", "AD"), ("p_Job", "AE"),
             ("p_Score", "AL"), ("p_Rating", "AM"), ("p_Reg", "AN"), ("p_Risk", "AO"), ("p_Key", "AP"), ("p_Flags", "AQ")):
    name(n, f"'Data - Pipeline'!${L}$2:${L}${PIPE_N}")

# ======================================================================= Monthly
mo = wb.create_sheet("Monthly", 0)
title(mo, "Monthly Revenue and Forecast", "Actual invoiced (Xero via Finance) plus won work still to invoice and the weighted open pipeline, by expected invoice month.", 8)
heads = ["Month", "Invoiced (actual)", "Won - still to invoice", "Open pipeline - weighted", "Forecast", "Open pipeline - full", "Cumulative actual", "Cumulative forecast"]
for i, h in enumerate(heads):
    hdr(mo, 4, 1 + i, h)
    mo.column_dimensions[CL(1 + i)].width = 16
mo.row_dimensions[4].height = 32
for k in range(12):
    r = 5 + k
    body(mo.cell(r, 1, f"=EOMONTH(set_FYStart,{k})"), "mmm-yy", bold=True)
    body(mo.cell(r, 2, f'=SUMIFS(rev_Amt,rev_Cnt,1,rev_Month,$A{r})'), MONEY)
    body(mo.cell(r, 3, f'=SUMIFS(p_Still,p_Status,"Won",p_Month,$A{r},p_Still,">0")'), MONEY)
    body(mo.cell(r, 4, f'=SUMIFS(p_Weighted,p_Status,"Open",p_Month,$A{r})'), MONEY)
    body(mo.cell(r, 5, f"=B{r}+C{r}+D{r}"), MONEY, bold=True)
    body(mo.cell(r, 6, f'=SUMIFS(p_Value,p_Status,"Open",p_Month,$A{r})'), MONEY)
    body(mo.cell(r, 7, f"=IF($A{r}<=EOMONTH(set_AsAt,0),SUM($B$5:B{r}),"")"), MONEY)
    body(mo.cell(r, 8, f"=SUM($E$5:E{r})"), MONEY)
body(mo.cell(17, 1, "TOTAL"), bold=True, fill=TINT)
for c in range(2, 7):
    body(mo.cell(17, c, f"=SUM({CL(c)}5:{CL(c)}16)"), MONEY, bold=True, fill=TINT)
mo.cell(19, 1, ("Won - still to invoice counts only positive balances. In months already gone it is billing to chase, not future revenue. "
                "Weighted = deal value x stage probability (tracker Lists AA:AB - assumptions). Open deals dated outside this FY are not in the grid.")).font = F(italic=True, color=MUTED, size=9)
mo.freeze_panes = "B5"

# ================================================================= Departments
dp = wb.create_sheet("Departments", 1)
title(dp, "By Cost Centre and Pipeline Stage", "Where the revenue comes from and where the pipeline sits.", 10)
section(dp, 4, "1.  BY COST CENTRE  -  financial year to date", 10)
heads = ["Cost Centre", "Invoiced FYTD", "Share", "Won (FY)", "Lost (FY)", "Win rate (value)", "Won - still to invoice", "Open pipeline", "Weighted pipeline", "Rest-of-year forecast"]
for i, h in enumerate(heads):
    hdr(dp, 5, 1 + i, h)
    dp.column_dimensions[CL(1 + i)].width = 15
dp.row_dimensions[5].height = 32
CCS = ["ONSITE", "PRODUCTION", "VIDEO", "INTEGRATION", "CONSULTING", "CTS"]
for k, cc in enumerate(CCS):
    r = 6 + k
    crit_r = f'rev_CC,$A{r}'
    crit_p = f'p_CC,$A{r}'
    body(dp.cell(r, 1, cc), bold=True)
    body(dp.cell(r, 2, f'=SUMIFS(rev_Amt,rev_FY,1,{crit_r})'), MONEY)
    body(dp.cell(r, 3, f'=IFERROR(B{r}/$B$13,0)'), PCT)
    body(dp.cell(r, 4, f'=SUMIFS(p_Value,p_Status,"Won",{crit_p})'), MONEY)
    body(dp.cell(r, 5, f'=SUMIFS(p_Value,p_Status,"Lost",{crit_p})'), MONEY)
    body(dp.cell(r, 6, f'=IFERROR(D{r}/(D{r}+E{r}),"")'), PCT)
    body(dp.cell(r, 7, f'=SUMIFS(p_Still,p_Status,"Won",p_Still,">0",{crit_p})'), MONEY)
    body(dp.cell(r, 8, f'=SUMIFS(p_Value,p_Status,"Open",{crit_p})'), MONEY)
    body(dp.cell(r, 9, f'=SUMIFS(p_Weighted,p_Status,"Open",{crit_p})'), MONEY)
    body(dp.cell(r, 10, f"=G{r}+I{r}"), MONEY)
# row 12: anything with another or no cost centre, so the totals always equal the whole
body(dp.cell(12, 1, "Other / not set"), bold=True)
tot = {2: 'SUMIFS(rev_Amt,rev_FY,1)', 4: 'SUMIFS(p_Value,p_Status,"Won")', 5: 'SUMIFS(p_Value,p_Status,"Lost")',
       7: 'SUMIFS(p_Still,p_Status,"Won",p_Still,">0")', 8: 'SUMIFS(p_Value,p_Status,"Open")', 9: 'SUMIFS(p_Weighted,p_Status,"Open")'}
for c, fx in tot.items():
    body(dp.cell(12, c, f"={fx}-SUM({CL(c)}6:{CL(c)}11)"), MONEY)
body(dp.cell(12, 3, "=IFERROR(B12/$B$13,0)"), PCT)
body(dp.cell(12, 6, '=IFERROR(D12/(D12+E12),"")'), PCT)
body(dp.cell(12, 10, "=G12+I12"), MONEY)
body(dp.cell(13, 1, "TOTAL"), bold=True, fill=TINT)
for c in (2, 4, 5, 7, 8, 9, 10):
    body(dp.cell(13, c, f"=SUM({CL(c)}6:{CL(c)}12)"), MONEY, bold=True, fill=TINT)
body(dp.cell(13, 3, "=SUM(C6:C12)"), PCT, bold=True, fill=TINT)
body(dp.cell(13, 6, '=IFERROR(D13/(D13+E13),"")'), PCT, bold=True, fill=TINT)

section(dp, 16, "2.  OPEN PIPELINE BY STAGE", 10)
for i, h in enumerate(["Stage", "Deals", "Full value", "Weighted value", "Probability", "Share of pipeline"]):
    hdr(dp, 17, 1 + i, h)
STAGES = ["Building Value Proposition", "Proposal Sent", "Negotiation/Review", "Verbal Approval"]
for k, s in enumerate(STAGES):
    r = 18 + k
    body(dp.cell(r, 1, s), bold=True)
    body(dp.cell(r, 2, f'=COUNTIFS(p_Status,"Open",p_Stage,$A{r})'), "0")
    body(dp.cell(r, 3, f'=SUMIFS(p_Value,p_Status,"Open",p_Stage,$A{r})'), MONEY)
    body(dp.cell(r, 4, f'=SUMIFS(p_Weighted,p_Status,"Open",p_Stage,$A{r})'), MONEY)
    body(dp.cell(r, 5, f'=IFERROR(D{r}/C{r},0)'), "0%")
    body(dp.cell(r, 6, f'=IFERROR(C{r}/$C$22,0)'), PCT)
body(dp.cell(22, 1, "TOTAL OPEN"), bold=True, fill=TINT)
for c, fmt in ((2, "0"), (3, MONEY), (4, MONEY), (6, PCT)):
    body(dp.cell(22, c, f"=SUM({CL(c)}18:{CL(c)}21)"), fmt, bold=True, fill=TINT)
body(dp.cell(22, 5, "=IFERROR(D22/C22,0)"), "0%", bold=True, fill=TINT)
dp.freeze_panes = "B6"

# ======================================================= Client concentration
cc = wb.create_sheet("Client Concentration", 2)
title(cc, "Client Concentration", "How dependent FYTD invoiced revenue is on a few clients. Client names are grouped as typed on the tracker (upper case, trimmed).", 9)
section(cc, 4, "1.  TOP CLIENTS  -  invoiced financial year to date", 9)
for i, h in enumerate(["Rank", "Client", "Invoiced FYTD", "Share", "Cumulative share", "Invoices"]):
    hdr(cc, 5, 1 + i, h)
cc.column_dimensions["A"].width = 7
cc.column_dimensions["B"].width = 42
for L in "CDEF":
    cc.column_dimensions[L].width = 15
TOPN = 20
for k in range(TOPN):
    r = 6 + k
    body(cc.cell(r, 1, k + 1), "0", bold=True)
    body(cc.cell(r, 2, f'=IFERROR(INDEX($L$6:$L$305,MATCH(LARGE($N$6:$N$305,{k + 1}),$N$6:$N$305,0)),"")'))
    body(cc.cell(r, 3, f'=IF(B{r}="","",SUMIFS(rev_Amt,rev_FY,1,rev_Client,B{r}))'), MONEY)
    body(cc.cell(r, 4, f'=IF(B{r}="","",IFERROR(C{r}/$C$28,0))'), PCT)
    body(cc.cell(r, 5, f'=IF(B{r}="","",SUM($D$6:D{r}))'), PCT)
    body(cc.cell(r, 6, f'=IF(B{r}="","",COUNTIFS(rev_FY,1,rev_Client,B{r}))'), "0")
body(cc.cell(27, 2, "All other clients"), bold=True)
body(cc.cell(27, 3, "=C28-SUM(C6:C25)"), MONEY)
body(cc.cell(27, 4, "=IFERROR(C27/C28,0)"), PCT)
body(cc.cell(28, 2, "TOTAL INVOICED FYTD"), bold=True, fill=TINT)
body(cc.cell(28, 3, "=SUMIFS(rev_Amt,rev_FY,1)"), MONEY, bold=True, fill=TINT)
section(cc, 30, "2.  CONCENTRATION MEASURES", 9)
meas = [
    ("Clients invoiced FYTD", "=MAX($M$6:$M$305)", "0", None),
    ("Largest client share", "=D6", PCT, '=IF(D6>=set_ClientHigh,"High",IF(D6>=set_ClientHigh*0.6,"Medium","Low"))'),
    ("Top 5 clients share", "=IFERROR(SUM(C6:C10)/C28,0)", PCT, '=IF(C33>=set_Top5High,"High",IF(C33>=set_Top5High*0.75,"Medium","Low"))'),
    ("Top 10 clients share", "=IFERROR(SUM(C6:C15)/C28,0)", PCT, None),
    ("Herfindahl index (0-10,000)", "=ROUND(SUM($P$6:$P$305)*10000,0)", "#,##0", '=IF(C35>=2500,"High",IF(C35>=1500,"Medium","Low"))'),
]
for i, (lab, fx, fmt, rate) in enumerate(meas):
    r = 31 + i
    body(cc.cell(r, 2, lab), bold=True)
    body(cc.cell(r, 3, fx), fmt, bold=True)
    if rate:
        x = cc.cell(r, 4, rate)
        x.font = F(bold=True)
        x.alignment = Alignment(horizontal="center")
rating_fmt(cc, "D31:D35")
cc.cell(37, 2, ("Herfindahl index = sum of each client's squared share x 10,000. Under 1,500 is spread out, 1,500-2,500 moderate, "
                "over 2,500 concentrated (the usual market-concentration bands). Thresholds for the other two lines are on Settings.")).font = F(italic=True, color=MUTED, size=9)
# helper block L:O - every FYTD client once
for i, h in enumerate(["Client (all)", "No.", "Sort key", "Share", "Share squared"]):
    hdr(cc, 5, 12 + i, h, MUTED)
    cc.column_dimensions[CL(12 + i)].width = 14 if i else 30
for k in range(300):
    r = 6 + k
    cc.cell(r, 12, f'=IFERROR(INDEX(rev_Client,MATCH({k + 1},rev_New,0)),"")').font = F(size=9, color=MUTED)
    cc.cell(r, 13, f'=IF(L{r}="","",{k + 1})').font = F(size=9, color=MUTED)
    cc.cell(r, 14, f'=IF(L{r}="","",SUMIFS(rev_Amt,rev_FY,1,rev_Client,L{r})+{k + 1}/1E+6)').font = F(size=9, color=MUTED)
    x = cc.cell(r, 15, f'=IF(L{r}="","",IFERROR(SUMIFS(rev_Amt,rev_FY,1,rev_Client,L{r})/$C$28,0))')
    x.font = F(size=9, color=MUTED)
    x.number_format = PCT
    cc.cell(r, 16, f'=IF(O{r}="",0,O{r}^2)').font = F(size=9, color=MUTED)

# ================================================================ Risk register
rg = wb.create_sheet("Risk Register", 3)
title(rg, "Risk Register", "Every open deal and every won deal not yet fully invoiced, worst first. Flags and scores are explained at the bottom; thresholds are on Settings.", 13)
heads = ["#", "Rating", "Score", "Job / Opp No", "Client", "Deal", "Cost Centre", "Status", "Stage", "Value", "Probability", "Revenue at risk", "Expected Month", "Why"]
widths = [5, 10, 7, 12, 26, 44, 13, 9, 22, 13, 11, 14, 11, 60]
for i, (h, w) in enumerate(zip(heads, widths)):
    hdr(rg, 4, 1 + i, h)
    rg.column_dimensions[CL(1 + i)].width = w
rg.row_dimensions[4].height = 30
REG = 250
for k in range(REG):
    r = 5 + k
    m = f"MATCH(LARGE(p_Key,{k + 1}),p_Key,0)"
    cells = [
        (1, f'=IF(D{r}="","",{k + 1})', "0"),
        (2, f'=IFERROR(INDEX(p_Rating,{m}),"")', None),
        (3, f'=IFERROR(INDEX(p_Score,{m}),"")', "0"),
        (4, f'=IFERROR(INDEX(p_Job,{m}),"")', None),
        (5, f'=IFERROR(INDEX(p_Client,{m}),"")', None),
        (6, f'=IFERROR(INDEX(p_Deal,{m}),"")', None),
        (7, f'=IFERROR(INDEX(p_CC,{m}),"")', None),
        (8, f'=IFERROR(INDEX(p_Status,{m}),"")', None),
        (9, f'=IFERROR(INDEX(p_Stage,{m}),"")', None),
        (10, f'=IFERROR(INDEX(p_Value,{m}),"")', MONEY),
        (11, f'=IFERROR(IF(H{r}="Open",INDEX(p_Prob,{m}),""),"")', "0%"),
        (12, f'=IFERROR(INDEX(p_Risk,{m}),"")', MONEY),
        (13, f'=IFERROR(INDEX(p_Month,{m}),"")', "mmm-yy"),
        (14, f'=IFERROR(INDEX(p_Flags,{m}),"")', None),
    ]
    for c, fx, fmt in cells:
        x = body(rg.cell(r, c, fx), fmt)
        if c == 2:
            x.alignment = Alignment(horizontal="center")
rating_fmt(rg, f"B5:B{4 + REG}")
nr = 6 + REG
notes = [
    "HOW THE SCORE WORKS (add the points; High / Medium / Low cut-offs are on Settings)",
    "3 - Expected month passed: an open deal whose expected invoice month has gone. Update Zoho or mark it lost.",
    "3 - Concentration: one open deal is at least the Settings share of all open pipeline. If it slips, the forecast moves a lot.",
    "2 - Large deal at an early stage: at or above the large-deal value with a probability at or below the early-stage setting.",
    "2 - No amount in Zoho: the deal cannot be forecast until it has a value.",
    "3 - Won, invoice month passed, not fully invoiced: work we have won but not billed on time. Revenue at risk = what is still to invoice.",
    "1 - Xero invoiced more than the deal value: check scope, variations or a missing credit note.",
    "Revenue at risk: open deals = weighted value; won deals = amount still to invoice.",
]
for i, t in enumerate(notes):
    rg.cell(nr + i, 1, t).font = F(bold=i == 0, italic=i > 0, color=NAVY if i == 0 else MUTED, size=10 if i == 0 else 9)
rg.freeze_panes = "E5"
rg.auto_filter.ref = f"A4:N{4 + REG}"

# ==================================================================== Dashboard
db = wb.create_sheet("Dashboard", 0)
title(db, "Revenue Risk Profile - FY27",
      '="Corporate Technology Services - data as at "&TEXT(set_AsAt,"d mmmm yyyy")&". Actual revenue from Xero via the revenue tracker; won, lost and pipeline from Zoho."', 18)
for c in range(1, 19):
    db.column_dimensions[CL(c)].width = 11
db.column_dimensions["A"].width = 2
tiles = [
    ("B", "Invoiced FYTD", "=Departments!B13", MONEY),
    ("E", "Won - still to invoice", "=Departments!G13", MONEY),
    ("H", "Weighted open pipeline", "=Departments!I13", MONEY),
    ("K", "Full-year forecast", "=Monthly!E17", MONEY),
    ("N", "Win rate (value)", "=Departments!F13", "0%"),
    ("Q", "Largest client share", "='Client Concentration'!C32", "0%"),
]
for col, lab, fx, fmt in tiles:
    c0 = openpyxl.utils.column_index_from_string(col)
    db.merge_cells(start_row=4, start_column=c0, end_row=4, end_column=c0 + 1)
    db.merge_cells(start_row=5, start_column=c0, end_row=5, end_column=c0 + 1)
    a = db.cell(4, c0, lab)
    a.font = F(size=9, bold=True, color=MUTED)
    a.alignment = Alignment(horizontal="center")
    v = db.cell(5, c0, fx)
    v.font = F(size=16, bold=True, color=NAVY)
    v.number_format = fmt
    v.alignment = Alignment(horizontal="center", vertical="center")
    for r in (4, 5):
        for cc_ in (c0, c0 + 1):
            db.cell(r, cc_).fill = FILL(TINT)
db.row_dimensions[5].height = 30
# risk summary tiles
db.cell(7, 2, "RISK REGISTER SUMMARY").font = F(bold=True, color=NAVY)
for i, (lab, col) in enumerate([("High", "D03B3B"), ("Medium", "EC835A"), ("Low", "FAB219")]):
    r = 8 + i
    x = db.cell(r, 2, lab)
    x.font = F(bold=True, color="FFFFFF" if lab == "High" else TXT)
    x.fill = FILL(col)
    x.alignment = Alignment(horizontal="center")
    body(db.cell(r, 3, f'=COUNTIFS(p_Reg,1,p_Rating,B{r})'), "0")
    body(db.cell(r, 4, f'=SUMIFS(p_Risk,p_Reg,1,p_Rating,B{r})'), MONEY)
    db.merge_cells(start_row=r, start_column=4, end_row=r, end_column=5)
db.cell(7, 3, "Items").font = F(size=9, color=MUTED)
db.cell(7, 4, "Revenue at risk").font = F(size=9, color=MUTED)
# top 5 risks
db.cell(7, 7, "TOP RISKS").font = F(bold=True, color=NAVY)
for i, (h, c) in enumerate([("Rating", 7), ("Deal", 8), ("Revenue at risk", 14), ("Why", 16)]):
    db.cell(7, c, h if i else "Rating").font = F(size=9, color=MUTED)
for k in range(5):
    r = 8 + k
    x = db.cell(r, 7, f"='Risk Register'!B{5 + k}")
    x.alignment = Alignment(horizontal="center")
    x.font = F(bold=True)
    db.cell(r, 8, f"='Risk Register'!E{5 + k}&\" - \"&'Risk Register'!F{5 + k}").font = F(size=9)
    db.merge_cells(start_row=r, start_column=8, end_row=r, end_column=13)
    m = db.cell(r, 14, f"='Risk Register'!L{5 + k}")
    m.number_format = MONEY
    m.font = F(size=9, bold=True)
    db.merge_cells(start_row=r, start_column=14, end_row=r, end_column=15)
    db.cell(r, 16, f"='Risk Register'!N{5 + k}").font = F(size=8, color=MUTED)
    db.merge_cells(start_row=r, start_column=16, end_row=r, end_column=18)
rating_fmt(db, "G8:G12")
db.cell(13, 2, ("Weighted pipeline uses assumed stage probabilities. Forecast = invoiced + won still to invoice + weighted pipeline. "
                "See Risk Register for every item and Settings for thresholds.")).font = F(size=8, italic=True, color=MUTED)


def bar(title_, cats, series, anchor, w=17, h=8.5, stacked=False, horiz=False, fmt='$#,##0', colors=(S1, S2, S3), legend=True):
    ch = BarChart()
    ch.type = "bar" if horiz else "col"
    ch.title = title_
    ch.style = 10
    if stacked:
        ch.grouping = "stacked"
        ch.overlap = 100
    ch.gapWidth = 60
    for i, ref in enumerate(series):
        ch.add_data(ref, titles_from_data=True)
    ch.set_categories(cats)
    for i, s in enumerate(ch.series):
        s.graphicalProperties.solidFill = colors[i]
        s.graphicalProperties.line.solidFill = "FFFFFF"
        s.graphicalProperties.line.width = 12700
    ch.y_axis.numFmt = fmt
    ch.y_axis.majorGridlines.spPr = None
    ch.y_axis.delete = False
    ch.x_axis.delete = False
    if horiz:
        ch.x_axis.scaling.orientation = "maxMin"
    ch.legend.position = "b"
    if not legend:
        ch.legend = None
    ch.width, ch.height = w, h
    db.add_chart(ch, anchor)
    return ch


# 1 monthly stacked
bar("Monthly revenue: actual, won to invoice and weighted pipeline",
    Reference(mo, min_col=1, min_row=5, max_row=16),
    [Reference(mo, min_col=c, min_row=4, max_row=16) for c in (2, 3, 4)], "B15", w=34, h=9, stacked=True)
# 2 top 10 clients
bar("Top 10 clients - invoiced FYTD",
    Reference(cc, min_col=2, min_row=6, max_row=15),
    [Reference(cc, min_col=3, min_row=5, max_row=15)], "B34", w=17, h=9, horiz=True, legend=False)
# 3 cost centre
bar("By cost centre: invoiced, won to invoice, weighted pipeline",
    Reference(dp, min_col=1, min_row=6, max_row=11),
    [Reference(dp, min_col=c, min_row=5, max_row=11) for c in (2, 7, 9)], "K34", w=17, h=9, stacked=True)
# 4 pipeline by stage
bar("Open pipeline by stage: full vs weighted",
    Reference(dp, min_col=1, min_row=18, max_row=21),
    [Reference(dp, min_col=c, min_row=17, max_row=21) for c in (3, 4)], "B53", w=17, h=9)
# 5 risk by rating
ch5 = bar("Revenue at risk by rating",
          Reference(db, min_col=2, min_row=8, max_row=10),
          [Reference(db, min_col=4, min_row=7, max_row=10)], "K53", w=17, h=9, legend=False)
from openpyxl.chart.series import DataPoint
for i, col in enumerate(("D03B3B", "EC835A", "FAB219")):
    pt = DataPoint(idx=i)
    pt.graphicalProperties.solidFill = col
    ch5.series[0].dPt.append(pt)
ch5.dataLabels = DataLabelList()
ch5.dataLabels.showVal = True
ch5.dataLabels.showSerName = False
ch5.dataLabels.showCatName = False
ch5.dataLabels.showLegendKey = False
ch5.dataLabels.numFmt = '$#,##0'
db.freeze_panes = "A4"

# ====================================================================== Read Me
rm = wb.create_sheet("Read Me", 1)
title(rm, "How this workbook works", "Revenue Risk Profile - built from the FY27 Revenue Tracker.", 3)
rm.column_dimensions["A"].width = 3
rm.column_dimensions["B"].width = 120
lines = [
    ("WHAT IT SHOWS", True),
    ("Dashboard - the headline numbers, the risk summary, the five biggest risks and five charts.", False),
    ("Monthly - actual invoiced revenue by month, plus won work still to invoice and weighted pipeline by expected invoice month, and the full-year forecast.", False),
    ("Departments - invoiced, won, lost, win rate, still to invoice and pipeline by cost centre, and the open pipeline by stage.", False),
    ("Client Concentration - top 20 clients by invoiced revenue, their share and cumulative share, and three concentration measures with a High/Medium/Low rating.", False),
    ("Risk Register - every open deal and every won deal not fully invoiced, scored and sorted worst first, with the reason.", False),
    ("Settings - the thresholds behind every rating. Change them there.", False),
    ("", False),
    ("REFRESH IT EACH MONTH  (two copy / paste-values)", True),
    ("1. In the revenue tracker, do the Zoho dump and make sure it has recalculated (open it in Excel and save).", False),
    ("2. Tracker Finance sheet: select A6:AB4506, Ctrl+C. Here: Data - Revenue, click A1, Paste Special - Values.", False),
    ("3. Tracker Work Won sheet: select A4:R604, Ctrl+C. Here: Data - Pipeline, click A1, Paste Special - Values.", False),
    ("4. Settings: change 'Data as at' to the date the tracker data is as at. Everything else, charts included, updates by itself.", False),
    ("Only paste into columns A:AB (Revenue) and A:R (Pipeline). The grey columns to the right are the workings.", False),
    ("", False),
    ("WHAT TO KEEP IN MIND", True),
    ("Weighted pipeline uses the stage probabilities from the tracker (Lists AA:AB). They are assumptions, not Zoho data - the forecast is only as good as they are.", False),
    ("Client concentration groups clients as typed on the tracker (for example PwC, PWC and Pwc are one client). Two spellings of the same client (CBA vs Commonwealth Bank) count as two - tidy the Client column on the department sheets if that matters.", False),
    ("Won - still to invoice matches the job number (and its V job) against Finance. A won deal invoiced under a different job number shows as still to invoice.", False),
    ("Open deals dated outside this financial year are in the Risk Register but not in the monthly grid.", False),
]
for i, (t, h) in enumerate(lines):
    x = rm.cell(4 + i, 2, t)
    x.font = F(bold=h, color=NAVY if h else TXT, size=11 if h else 10)
    x.alignment = Alignment(wrap_text=True, vertical="top")

wb.move_sheet("Settings", offset=-(wb.sheetnames.index("Settings") - 6))
wb.save(OUT)
print("saved", OUT, wb.sheetnames)
