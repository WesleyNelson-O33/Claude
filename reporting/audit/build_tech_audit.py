"""
Standalone audit report: discretionary technology spend, FY2027 to date.

Self-contained on purpose - no links to any other workbook, base data written
as values so nothing can break in transit, subtotals left as formulas so the
recipient can see each schedule foots.


Build order matters: build, then recalculate, THEN run
reporting/tools/fix_outline.py. LibreOffice drops the outline properties when it
recalculates, and without them Excel puts the collapse buttons on the wrong rows.
    python3 reporting/tools/fix_outline.py "<the .xlsx>" "Supplier Analysis,Account Analysis"
"""
import json
from pathlib import Path
from collections import defaultdict
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as CL

UP = Path("/root/.claude/uploads/a00ffe59-433c-507b-9006-917d8c60083a")
SRC = UP / "8084a881-PL_Analysis_-_Aug_2026_-_Xero.xlsm"
OUT = Path("/home/user/Claude/reporting/audit/CTS - Discretionary Technology Spend - FY2027 YTD.xlsx")

ACCOUNTS = ["IT Network Service & Support", "Dues & Subscriptions", "Instant Asset Write Off",
            "Internet", "Mobile Phones", "Office Phones", "1300 Number"]
GROUP = {"IT Network Service & Support": "Core technology", "Dues & Subscriptions": "Core technology",
         "Instant Asset Write Off": "Core technology", "Internet": "Telecommunications",
         "Mobile Phones": "Telecommunications", "Office Phones": "Telecommunications",
         "1300 Number": "Telecommunications"}
BUDGET = {"IT Network Service & Support": 149160.0, "Dues & Subscriptions": 86544.0,
          "Instant Asset Write Off": 60000.0, "Internet": 4814.0, "Mobile Phones": 3300.0,
          "Office Phones": 2400.0, "1300 Number": 780.0}

# ---- read the ledger ------------------------------------------------------
ws = load_workbook(SRC, data_only=True, read_only=True)["1. Xero GL Transactions"]
txns, all_rows, gl_total = [], 0, 0.0
for r in ws.iter_rows(min_row=2, max_col=18, values_only=True):
    if r[1] is None:
        continue
    all_rows += 1
    gl_total += (r[13] or 0)
    name = str(r[2]).strip() if r[2] else ""
    if name not in ACCOUNTS:
        continue
    txns.append({
        "date": r[4], "code": r[1], "account": name, "source": r[3],
        "supplier": (str(r[5]).strip() if r[5] else ""), "invoice": r[9], "reference": r[10],
        "description": (str(r[11]).strip() if r[11] else ""), "job": r[8],
        "cost_centre": r[12], "debit": r[6] or 0, "credit": r[7] or 0,
        "amount": -(r[13] or 0),                      # ledger holds credit-debit; show cost positive
        "month": r[15], "fy": r[14],
    })
txns.sort(key=lambda t: (t["account"], t["date"] or 0, t["supplier"]))
MONTHS = []
for t in txns:
    if t["month"] not in MONTHS: MONTHS.append(t["month"])
PERIODS = len(MONTHS)
print(f"ledger rows {all_rows:,} | in-scope lines {len(txns):,} | months {MONTHS} | "
      f"total ${sum(t['amount'] for t in txns):,.2f}")

# ---- styling --------------------------------------------------------------
TNR = "Times New Roman"
NAVY, MUTED, WHITE = "1F3864", "595959", "FFFFFF"
HEAD = PatternFill("solid", fgColor=NAVY)
BANDF = PatternFill("solid", fgColor="D9E2F3")
TOTF = PatternFill("solid", fgColor="EDEDED")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
TOPB = Border(top=Side(style="thin", color="000000"))
MONEY = '$#,##0.00;($#,##0.00);"-"'
MONEY0 = '$#,##0;($#,##0);"-"'
PCT = '0.0%;(0.0%);"-"'
DATEF = 'dd/mm/yyyy'

wb = Workbook()
def sheet(name):
    s = wb.create_sheet(name)
    s.sheet_view.showGridLines = False
    return s
def title(s, t, sub=None):
    s["A1"] = "Corporate Technology Services Pty Ltd"
    s["A1"].font = Font(name=TNR, size=11, color=MUTED)
    s["A2"] = t
    s["A2"].font = Font(name=TNR, size=15, bold=True, color=NAVY)
    if sub:
        s["A3"] = sub
        s["A3"].font = Font(name=TNR, size=10, italic=True, color=MUTED)
def header(s, row, labels, c0=1, h=28):
    for i, l in enumerate(labels):
        c = s.cell(row=row, column=c0 + i, value=l)
        c.font = Font(name=TNR, size=10, bold=True, color=WHITE)
        c.fill = HEAD; c.border = BOX
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    s.row_dimensions[row].height = h
def widths(s, spec):
    for k, v in spec.items(): s.column_dimensions[k].width = v
print("data loaded")

FY_LABEL = "FY2027"
PERIOD_LABEL = "1 July 2026 to 31 August 2026"
SUBTITLE = f"Discretionary technology spend  -  {FY_LABEL} year to date ({PERIOD_LABEL})"

# ===========================================================================
# 1. BASIS OF PREPARATION
# ===========================================================================
bp = sheet("Basis of Preparation")
title(bp, "Basis of Preparation", SUBTITLE)
widths(bp, {"A": 30, "B": 104})
r = 5
def para(r, head, body):
    c = bp.cell(row=r, column=1, value=head)
    c.font = Font(name=TNR, size=10, bold=True)
    c.alignment = Alignment(vertical="top")
    d = bp.cell(row=r, column=2, value=body)
    d.font = Font(name=TNR, size=10)
    d.alignment = Alignment(wrap_text=True, vertical="top")
    bp.row_dimensions[r].height = max(15, 13 * (1 + len(body) // 108))
    return r + 1

for h, b in [
    ("Entity", "Corporate Technology Services Pty Ltd."),
    ("Reporting period", f"{PERIOD_LABEL}. This is {PERIODS} months of the {FY_LABEL} financial "
                         "year (1 July 2026 to 30 June 2027). Actual amounts in this report are for "
                         f"those {PERIODS} months only and are not annual figures."),
    ("Source of actuals", "General ledger transaction export from Xero, being the same export used "
                          "for the entity's monthly management reporting. The export covers "
                          f"{all_rows:,} ledger lines for the period, of which {len(txns):,} lines "
                          "relate to the accounts within the scope of this report."),
    ("Accounting basis", "Accrual. Amounts are exclusive of GST, consistent with the general ledger."),
    ("Scope", "Seven general ledger accounts identified by management as discretionary technology "
              "spend, listed on the Summary schedule. No other accounts have been included, and no "
              "allocation, apportionment or estimate has been applied."),
    ("Presentation of amounts", "The general ledger records amounts as credit less debit. Costs are "
                                "therefore negative in the ledger and have been presented here as "
                                "positive expenditure. Debit and credit as recorded are shown "
                                "unchanged on the Transaction Listing so that each line can be "
                                "agreed back to the ledger."),
    ("Budget figures", "Budget amounts are management's FY2027 annual budget for each account, as "
                       "provided. They are unaudited and are included for comparison only. The "
                       "year-to-date budget shown is the annual budget apportioned evenly across "
                       f"twelve months and multiplied by {PERIODS}; no seasonality has been applied."),
    ("Completeness", "Every ledger line coded to the seven accounts in the period is listed on the "
                     "Transaction Listing. No line has been excluded, summarised or netted. The "
                     "Reconciliation schedule agrees the listing to the summary and to the ledger."),
    ("Matters noted", "Two matters are disclosed on the Summary schedule: an arithmetic overstatement "
                      "in the budget total as originally presented to us, and expenditure recorded "
                      "without a supplier name. Both are described where they arise."),
    ("Prepared", "Prepared from the general ledger export by the finance function. Figures are "
                 "unaudited."),
]:
    r = para(r, h, b)

# ===========================================================================
# 2. SUMMARY
# ===========================================================================
sm = sheet("Summary")
title(sm, "Summary by account", SUBTITLE)
widths(sm, {"A": 22, "B": 34, "C": 16, "D": 16, "E": 16, "F": 16, "G": 14, "H": 16, "I": 12})
header(sm, 5, ["Grouping", "General ledger account", "Annual budget\nFY2027",
               f"Budget\n{PERIODS} months", f"Actual\n{PERIODS} months", "Variance",
               "Lines", "Annualised\nrun rate", "Run rate vs\nbudget"])
actual = defaultdict(float); lines_ct = defaultdict(int)
for t in txns:
    actual[t["account"]] += t["amount"]; lines_ct[t["account"]] += 1
r = 6
grp_rows = defaultdict(list)
for a in ACCOUNTS:
    g = GROUP[a]
    grp_rows[g].append(r)
    vals = [g, a, BUDGET[a], round(BUDGET[a] / 12 * PERIODS, 2), round(actual[a], 2),
            None, lines_ct[a], None, None]
    for i, v in enumerate(vals):
        c = sm.cell(row=r, column=1 + i, value=v)
        c.font = Font(name=TNR, size=10)
        c.border = BOX
        if i in (2, 3, 4): c.number_format = MONEY
        if i == 6: c.alignment = Alignment(horizontal="center")
    sm.cell(row=r, column=6, value=f"=E{r}-D{r}").number_format = MONEY
    sm.cell(row=r, column=8, value=f"=E{r}/{PERIODS}*12").number_format = MONEY
    sm.cell(row=r, column=9, value=f'=IF(C{r}=0,"n/a",H{r}/C{r}-1)').number_format = PCT
    for col in (6, 8, 9):
        sm.cell(row=r, column=col).font = Font(name=TNR, size=10)
        sm.cell(row=r, column=col).border = BOX
    r += 1
sub = {}
for g in ["Core technology", "Telecommunications"]:
    rows = grp_rows[g]
    sm.cell(row=r, column=2, value=f"Subtotal - {g.lower()}").font = Font(name=TNR, size=10, bold=True)
    for col in (3, 4, 5, 6, 7, 8):
        c = sm.cell(row=r, column=col, value="=" + "+".join(f"{CL(col)}{x}" for x in rows))
        c.font = Font(name=TNR, size=10, bold=True); c.border = BOX; c.fill = TOTF
        c.number_format = MONEY if col != 7 else '#,##0'
    c = sm.cell(row=r, column=9, value=f'=IF(C{r}=0,"n/a",H{r}/C{r}-1)')
    c.font = Font(name=TNR, size=10, bold=True); c.border = BOX; c.fill = TOTF; c.number_format = PCT
    for col in (1, 2): sm.cell(row=r, column=col).fill = TOTF
    sub[g] = r; r += 1
sm.cell(row=r, column=2, value="Total discretionary technology spend").font = Font(
    name=TNR, size=11, bold=True, color=WHITE)
for col in (3, 4, 5, 6, 7, 8):
    c = sm.cell(row=r, column=col,
                value=f'={CL(col)}{sub["Core technology"]}+{CL(col)}{sub["Telecommunications"]}')
    c.font = Font(name=TNR, size=11, bold=True, color=WHITE)
    c.number_format = MONEY if col != 7 else '#,##0'
c = sm.cell(row=r, column=9, value=f'=IF(C{r}=0,"n/a",H{r}/C{r}-1)')
c.font = Font(name=TNR, size=11, bold=True, color=WHITE); c.number_format = PCT
for col in range(1, 10): sm.cell(row=r, column=col).fill = HEAD
TOTAL_R = r
r += 2

notes = [
    ("Note 1 - budget total",
     "The budget schedule provided to us showed a total of $602,702. That figure adds a 'standard' "
     "subtotal of $295,704 (the three core technology accounts) to a 'broad' subtotal of $306,998, "
     "where the broad subtotal already contains the standard subtotal. The three core accounts are "
     "therefore counted twice and the total is overstated by $295,704. The annual budget for the "
     "seven accounts, each counted once, is $306,998 as shown above."),
    ("Note 2 - expenditure with no supplier recorded",
     f"${sum(t['amount'] for t in txns if not t['supplier']):,.2f} of the expenditure in the period, "
     f"being {sum(t['amount'] for t in txns if not t['supplier'])/sum(t['amount'] for t in txns):.1%} "
     "of the total, is recorded on ledger lines that carry no contact or supplier name. These lines "
     "are listed in full on the Transaction Listing and are shown as a separate line on the Supplier "
     "Analysis schedule."),
    ("Note 3 - supplier concentration",
     "One supplier, First Focus IT Pty Ltd, accounts for $37,799.82 of expenditure in the period, "
     "being 64.7% of the total. That expenditure is recorded across two accounts, IT Network Service "
     "& Support and Dues & Subscriptions. See the Supplier Analysis schedule."),
    ("Note 4 - period covered",
     f"These figures cover July and August only, which is {PERIODS} months of a twelve month year. "
     "The annualised run rate column takes what has been spent so far, divides it by "
     f"{PERIODS} and multiplies by twelve, so it assumes the rest of the year looks like the first "
     f"{PERIODS} months. It makes no allowance for quieter or busier periods, for annual renewals "
     "that fall due later in the year, or for spending already committed but not yet incurred. "
     "Treat it as a rough indicator, not a forecast."),
]
for h, b in notes:
    c = sm.cell(row=r, column=1, value=h)
    c.font = Font(name=TNR, size=10, bold=True)
    c.alignment = Alignment(vertical="top")
    d = sm.cell(row=r, column=2, value=b)
    d.font = Font(name=TNR, size=10)
    d.alignment = Alignment(wrap_text=True, vertical="top")
    sm.merge_cells(start_row=r, start_column=2, end_row=r, end_column=9)
    sm.row_dimensions[r].height = max(15, 12.5 * (1 + len(b) // 118))
    r += 1
print("basis + summary built")

# ===========================================================================
# 3. SUPPLIER ANALYSIS  - drill-down: supplier, then account, then the lines
# ===========================================================================
from openpyxl.worksheet.properties import Outline
sp = sheet("Supplier Analysis")
title(sp, "Supplier analysis", SUBTITLE +
      "  -  use the + and - buttons in the left margin to open a supplier down to its ledger lines")
widths(sp, {"A": 52, "B": 15, "C": 11, "D": 8, "E": 42, "F": 15, "G": 20, "H": 24, "I": 13, "J": 17})
header(sp, 5, ["Supplier  /  account  /  transaction date", "Amount", "% of total", "Lines",
               "Description", "Invoice no.", "Reference", "Job number", "Cost centre", "Source"])
sp.freeze_panes = "A6"

by_sup = defaultdict(list)
for t in txns:
    by_sup[t["supplier"] or "(no supplier name recorded on the ledger line)"].append(t)
order = sorted(by_sup, key=lambda k: -sum(x["amount"] for x in by_sup[k]))

r = 6
SUP_ROWS = []
for sup in order:
    items = by_sup[sup]
    blank = sup.startswith("(no supplier")
    sup_row = r
    SUP_ROWS.append(sup_row)
    c = sp.cell(row=r, column=1, value=sup)
    c.font = Font(name=TNR, size=11, bold=True, color=("9C0006" if blank else NAVY))
    for col in range(1, 11):
        sp.cell(row=r, column=col).fill = BANDF
        sp.cell(row=r, column=col).border = BOX
    sp.cell(row=r, column=4, value=len(items)).alignment = Alignment(horizontal="center")
    sp.cell(row=r, column=4).font = Font(name=TNR, size=11, bold=True)
    r += 1
    acc_rows = []
    per_acc = defaultdict(list)
    for t in items:
        per_acc[t["account"]].append(t)
    for acc in sorted(per_acc, key=lambda a: -sum(x["amount"] for x in per_acc[a])):
        rows_a = per_acc[acc]
        acc_row = r
        acc_rows.append(acc_row)
        c = sp.cell(row=r, column=1, value=f"    {acc}")
        c.font = Font(name=TNR, size=10, bold=True)
        sp.cell(row=r, column=4, value=len(rows_a)).alignment = Alignment(horizontal="center")
        for col in range(1, 11):
            sp.cell(row=r, column=col).fill = TOTF
            sp.cell(row=r, column=col).border = BOX
        sp.row_dimensions[r].outlineLevel = 1
        sp.row_dimensions[r].collapsed = True
        r += 1
        line_rows = []
        for t in sorted(rows_a, key=lambda x: (x["date"] or 0)):
            line_rows.append(r)
            vals = [t["date"], t["amount"], None, None, t["description"], t["invoice"],
                    t["reference"], t["job"], t["cost_centre"], t["source"]]
            for i, v in enumerate(vals):
                cc = sp.cell(row=r, column=1 + i, value=v)
                cc.font = Font(name=TNR, size=9)
                cc.border = BOX
                if i == 0: cc.number_format = DATEF
                if i == 1: cc.number_format = MONEY
            sp.row_dimensions[r].outlineLevel = 2
            sp.row_dimensions[r].hidden = True
            r += 1
        cc = sp.cell(row=acc_row, column=2,
                     value="=" + "+".join(f"B{x}" for x in line_rows))
        cc.font = Font(name=TNR, size=10, bold=True); cc.number_format = MONEY; cc.border = BOX
    cc = sp.cell(row=sup_row, column=2, value="=" + "+".join(f"B{x}" for x in acc_rows))
    cc.font = Font(name=TNR, size=11, bold=True, color=("9C0006" if blank else "000000"))
    cc.number_format = MONEY; cc.border = BOX

SUP_TOT = r
sp.cell(row=r, column=1, value="Total").font = Font(name=TNR, size=11, bold=True, color=WHITE)
cc = sp.cell(row=r, column=2, value="=" + "+".join(f"B{x}" for x in SUP_ROWS))
cc.font = Font(name=TNR, size=11, bold=True, color=WHITE); cc.number_format = MONEY
cc = sp.cell(row=r, column=4, value=f"={len(txns)}")
cc.font = Font(name=TNR, size=11, bold=True, color=WHITE)
cc.alignment = Alignment(horizontal="center")
for col in range(1, 11): sp.cell(row=r, column=col).fill = HEAD
for x in SUP_ROWS:
    cc = sp.cell(row=x, column=3, value=f"=B{x}/$B${SUP_TOT}")
    cc.font = Font(name=TNR, size=11, bold=True); cc.number_format = PCT; cc.border = BOX
r += 2
n = sp.cell(row=r, column=1, value=(
    "Each supplier total is the sum of its account rows, and each account row is the sum of the "
    "ledger lines beneath it, so the schedule foots at every level. Open a supplier with the + button "
    "to see the lines that make up its total. The same lines appear on the Transaction Listing."))
n.font = Font(name=TNR, size=9, italic=True, color=MUTED)
n.alignment = Alignment(wrap_text=True, vertical="top")
sp.merge_cells(start_row=r, start_column=1, end_row=r + 1, end_column=10)

# ===========================================================================
# 4. ACCOUNT ANALYSIS  - the same drill-down the other way up
# ===========================================================================
an = sheet("Account Analysis")
title(an, "Account analysis", SUBTITLE +
      "  -  the supplier schedule the other way up: account, then supplier, then the ledger lines")
widths(an, {"A": 52, "B": 15, "C": 11, "D": 8, "E": 42, "F": 15, "G": 20, "H": 24, "I": 13, "J": 17})
header(an, 5, ["Account  /  supplier  /  transaction date", "Amount", "% of total", "Lines",
               "Description", "Invoice no.", "Reference", "Job number", "Cost centre", "Source"])
an.freeze_panes = "A6"

by_acc = defaultdict(list)
for t in txns:
    by_acc[t["account"]].append(t)
acc_order = sorted(by_acc, key=lambda k: -sum(x["amount"] for x in by_acc[k]))

r = 6
ACC_ROWS = []
for acc in acc_order:
    items = by_acc[acc]
    acc_row = r
    ACC_ROWS.append(acc_row)
    c = an.cell(row=r, column=1, value=acc)
    c.font = Font(name=TNR, size=11, bold=True, color=NAVY)
    for col in range(1, 11):
        an.cell(row=r, column=col).fill = BANDF
        an.cell(row=r, column=col).border = BOX
    an.cell(row=r, column=4, value=len(items)).alignment = Alignment(horizontal="center")
    an.cell(row=r, column=4).font = Font(name=TNR, size=11, bold=True)
    an.cell(row=r, column=1).border = BOX
    r += 1
    per_sup = defaultdict(list)
    for t in items:
        per_sup[t["supplier"] or "(no supplier name recorded on the ledger line)"].append(t)
    sup_rows = []
    for sup in sorted(per_sup, key=lambda x: -sum(y["amount"] for y in per_sup[x])):
        rows_s = per_sup[sup]
        blank = sup.startswith("(no supplier")
        sup_row = r
        sup_rows.append(sup_row)
        c = an.cell(row=r, column=1, value=f"    {sup}")
        c.font = Font(name=TNR, size=10, bold=True, color=("9C0006" if blank else "000000"))
        an.cell(row=r, column=4, value=len(rows_s)).alignment = Alignment(horizontal="center")
        for col in range(1, 11):
            an.cell(row=r, column=col).fill = TOTF
            an.cell(row=r, column=col).border = BOX
        an.row_dimensions[r].outlineLevel = 1
        an.row_dimensions[r].collapsed = True
        r += 1
        line_rows = []
        for t in sorted(rows_s, key=lambda x: (x["date"] or 0)):
            line_rows.append(r)
            vals = [t["date"], t["amount"], None, None, t["description"], t["invoice"],
                    t["reference"], t["job"], t["cost_centre"], t["source"]]
            for i, v in enumerate(vals):
                cc = an.cell(row=r, column=1 + i, value=v)
                cc.font = Font(name=TNR, size=9)
                cc.border = BOX
                if i == 0: cc.number_format = DATEF
                if i == 1: cc.number_format = MONEY
            an.row_dimensions[r].outlineLevel = 2
            an.row_dimensions[r].hidden = True
            r += 1
        cc = an.cell(row=sup_row, column=2, value="=" + "+".join(f"B{x}" for x in line_rows))
        cc.font = Font(name=TNR, size=10, bold=True); cc.number_format = MONEY; cc.border = BOX
    cc = an.cell(row=acc_row, column=2, value="=" + "+".join(f"B{x}" for x in sup_rows))
    cc.font = Font(name=TNR, size=11, bold=True); cc.number_format = MONEY; cc.border = BOX

ACC_TOT = r
an.cell(row=r, column=1, value="Total").font = Font(name=TNR, size=11, bold=True, color=WHITE)
cc = an.cell(row=r, column=2, value="=" + "+".join(f"B{x}" for x in ACC_ROWS))
cc.font = Font(name=TNR, size=11, bold=True, color=WHITE); cc.number_format = MONEY
cc = an.cell(row=r, column=4, value=f"={len(txns)}")
cc.font = Font(name=TNR, size=11, bold=True, color=WHITE)
cc.alignment = Alignment(horizontal="center")
for col in range(1, 11): an.cell(row=r, column=col).fill = HEAD
for x in ACC_ROWS:
    cc = an.cell(row=x, column=3, value=f"=B{x}/$B${ACC_TOT}")
    cc.font = Font(name=TNR, size=11, bold=True); cc.number_format = PCT; cc.border = BOX
r += 2
n = an.cell(row=r, column=1, value=(
    "The same ledger lines as the Supplier Analysis schedule, grouped account first instead of "
    "supplier first. Each account total is the sum of its supplier rows, and each supplier row is "
    "the sum of the lines beneath it, so the schedule foots at every level."))
n.font = Font(name=TNR, size=9, italic=True, color=MUTED)
n.alignment = Alignment(wrap_text=True, vertical="top")
an.merge_cells(start_row=r, start_column=1, end_row=r + 1, end_column=10)

# ===========================================================================
# 4. MONTHLY ANALYSIS
# ===========================================================================
mn = sheet("Monthly Analysis")
title(mn, "Monthly analysis by account", SUBTITLE)
widths(mn, {"A": 34})
for i in range(2, 2 + PERIODS + 1): mn.column_dimensions[CL(i)].width = 16
header(mn, 5, ["General ledger account"] + [f"{m} 2026" for m in MONTHS] + ["Total"])
bym = defaultdict(float)
for t in txns: bym[(t["account"], t["month"])] += t["amount"]
r = 6
M_R0 = r
for a in ACCOUNTS:
    c = mn.cell(row=r, column=1, value=a)
    c.font = Font(name=TNR, size=10); c.border = BOX
    for i, m in enumerate(MONTHS):
        cc = mn.cell(row=r, column=2 + i, value=round(bym[(a, m)], 2))
        cc.font = Font(name=TNR, size=10); cc.number_format = MONEY; cc.border = BOX
    tc = mn.cell(row=r, column=2 + PERIODS, value=f"=SUM(B{r}:{CL(1+PERIODS)}{r})")
    tc.font = Font(name=TNR, size=10, bold=True); tc.number_format = MONEY; tc.border = BOX
    tc.fill = TOTF
    r += 1
M_R1 = r - 1
mn.cell(row=r, column=1, value="Total").font = Font(name=TNR, size=11, bold=True)
for i in range(PERIODS + 1):
    c = mn.cell(row=r, column=2 + i, value=f"=SUM({CL(2+i)}{M_R0}:{CL(2+i)}{M_R1})")
    c.font = Font(name=TNR, size=11, bold=True); c.number_format = MONEY
    c.border = BOX; c.fill = TOTF
mn.cell(row=r, column=1).fill = TOTF
MN_TOT = r

# ===========================================================================
# 5. TRANSACTION LISTING
# ===========================================================================
tl = sheet("Transaction Listing")
title(tl, "Transaction listing", SUBTITLE +
      f"  -  every ledger line coded to the seven accounts, {len(txns):,} lines")
COLS = [("Date", "date", 12, DATEF), ("Acct code", "code", 11, '0'),
        ("General ledger account", "account", 30, None), ("Source", "source", 18, None),
        ("Supplier / contact", "supplier", 30, None), ("Description", "description", 44, None),
        ("Invoice no.", "invoice", 14, None), ("Reference", "reference", 20, None),
        ("Job number", "job", 24, None), ("Cost centre", "cost_centre", 14, None),
        ("Debit", "debit", 13, MONEY), ("Credit", "credit", 13, MONEY),
        ("Expenditure", "amount", 14, MONEY)]
header(tl, 5, [c[0] for c in COLS])
widths(tl, {CL(i + 1): c[2] for i, c in enumerate(COLS)})
tl.freeze_panes = "A6"
r = 6
T_R0 = r
for t in txns:
    blank = not t["supplier"]
    for i, (_h, key, _w, fmt) in enumerate(COLS):
        v = t[key]
        if key == "supplier" and blank:
            v = "(no supplier name recorded)"
        c = tl.cell(row=r, column=1 + i, value=v)
        c.font = Font(name=TNR, size=9, color=("9C0006" if blank and key == "supplier" else "000000"))
        c.border = BOX
        if fmt: c.number_format = fmt
    r += 1
T_R1 = r - 1
tl.cell(row=r, column=1, value="Total").font = Font(name=TNR, size=11, bold=True)
for col in (11, 12, 13):
    c = tl.cell(row=r, column=col, value=f"=SUM({CL(col)}{T_R0}:{CL(col)}{T_R1})")
    c.font = Font(name=TNR, size=11, bold=True); c.number_format = MONEY
    c.border = BOX; c.fill = TOTF
TL_TOT = r
tl.auto_filter.ref = f"A5:{CL(len(COLS))}{T_R1}"

# ===========================================================================
# 6. RECONCILIATION
# ===========================================================================
rc = sheet("Reconciliation")
title(rc, "Reconciliation", SUBTITLE)
widths(rc, {"A": 8, "B": 60, "C": 18, "D": 14, "E": 52})
header(rc, 5, ["Ref", "Reconciling item", "Amount", "Agrees", "Basis"])
CHECKS = [
    ("R1", "Total expenditure per the Transaction Listing",
     f"='Transaction Listing'!M{TL_TOT}", None,
     f"Sum of {len(txns):,} ledger lines."),
    ("R2", "Total expenditure per the Summary schedule", f"=Summary!E{TOTAL_R}", None,
     "Sum of the seven accounts."),
    ("R3", "Difference, R1 less R2", "=C6-C7", "eq0",
     "Must be nil. The two schedules are built from the same lines."),
    ("R4", "Total expenditure per the Supplier Analysis schedule", f"='Supplier Analysis'!B{SUP_TOT}",
     None, "Sum of all suppliers, including lines with no supplier recorded."),
    ("R5", "Difference, R1 less R4", "=C6-C9", "eq0", "Must be nil."),
    ("R6", "Total expenditure per the Monthly Analysis schedule",
     f"='Monthly Analysis'!{CL(2+PERIODS)}{MN_TOT}", None, "Sum of all months."),
    ("R7", "Difference, R1 less R6", "=C6-C11", "eq0", "Must be nil."),
    ("R8", "Line count per the Transaction Listing", str(len(txns)), None,
     f"Against {all_rows:,} lines in the full ledger export for the period."),
    ("R9", "Debits less credits per the Transaction Listing",
     f"='Transaction Listing'!K{TL_TOT}-'Transaction Listing'!L{TL_TOT}", None,
     "Equals total expenditure, confirming these are expense lines with no credits netted off."),
    ("R10", "Difference, R9 less R1", "=C14-C6", "eq0", "Must be nil."),
    ("R11", "Total expenditure per the Account Analysis schedule", f"='Account Analysis'!B{ACC_TOT}",
     None, "Sum of all accounts. Same lines as R4, grouped the other way up."),
    ("R12", "Difference, R1 less R11", "=C6-C16", "eq0", "Must be nil."),
]
r = 6
for ref, desc, amt, mode, basis in CHECKS:
    for i, v in enumerate([ref, desc, amt, None, basis]):
        c = rc.cell(row=r, column=1 + i, value=(float(v) if i == 2 and str(v).isdigit() else v))
        c.font = Font(name=TNR, size=10, bold=(i == 0))
        c.border = BOX
        c.alignment = Alignment(horizontal=("center" if i == 0 else "left"),
                                wrap_text=(i in (1, 4)), vertical="center")
        if i == 2: c.number_format = '#,##0' if ref == "R8" else MONEY
    if mode == "eq0":
        c = rc.cell(row=r, column=4, value=f'=IF(ROUND(C{r},2)=0,"Yes","NO")')
        c.font = Font(name=TNR, size=10, bold=True)
        c.alignment = Alignment(horizontal="center"); c.border = BOX
    rc.row_dimensions[r].height = 26
    r += 1
from openpyxl.formatting.rule import FormulaRule
rc.conditional_formatting.add("D6:D17", FormulaRule(formula=['$D6="Yes"'],
    fill=PatternFill("solid", fgColor="C6EFCE"), font=Font(color="006100", bold=True)))
rc.conditional_formatting.add("D6:D17", FormulaRule(formula=['$D6="NO"'],
    fill=PatternFill("solid", fgColor="FFC7CE"), font=Font(color="9C0006", bold=True)))

del wb["Sheet"]
wb._sheets = [wb[n] for n in ["Basis of Preparation", "Summary", "Supplier Analysis",
                              "Account Analysis", "Monthly Analysis", "Transaction Listing",
                              "Reconciliation"]]
for s in wb.worksheets:
    for row in s.iter_rows():
        for c in row:
            if c.value is not None or c.has_style:
                f = c.font
                if f.name != TNR:
                    c.font = Font(name=TNR, size=f.size or 10, bold=f.bold, italic=f.italic,
                                  color=f.color)
    s.page_setup.orientation = "landscape"
    s.page_setup.fitToWidth = 1
    s.sheet_properties.pageSetUpPr.fitToPage = True
    s.print_options.horizontalCentered = True
    s.oddFooter.center.text = "Corporate Technology Services Pty Ltd  -  Discretionary technology spend  -  FY2027 year to date  -  Page &P of &N"
    s.oddFooter.center.size = 8
wb["Transaction Listing"].print_title_rows = "5:5"
wb["Supplier Analysis"].print_title_rows = "5:5"
wb["Account Analysis"].print_title_rows = "5:5"
# summary row sits above its detail, so Excel must be told, or the +/- land wrong
for _n in ("Supplier Analysis", "Account Analysis"):
    wb[_n].sheet_properties.outlinePr = Outline(
        summaryBelow=False, summaryRight=False, applyStyles=False)
wb.active = 0
OUT.parent.mkdir(parents=True, exist_ok=True)
wb.save(OUT)
print("saved:", OUT)
print("sheets:", wb.sheetnames)
