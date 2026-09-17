"""Build the FY27 Revenue Tracker workbook.

Replaces FY27_Revenue_Tracker_v2.xlsm, which was slow for one reason: a macro.
Every cell edit fired Worksheet_Change, which copied 26 columns x 200 rows back
into hidden storage columns, and every change of the department dropdown rewrote
the widths, formats and validation for the whole sheet. It was also capped at
200 rows per department and 800 in total, so it could never have held years of
transactions.

So there are no macros here at all. Two rules drive the design:

  1. Columns A-S are identical on every department sheet. Finance stacks exactly
     those columns with one formula, so adding a department means copying a sheet
     and keeping the spine - nothing has to be rebuilt.
  2. Nothing volatile, nothing whole-column, no per-cell styling over empty rows.
     Excel only recalculates what actually changed.

Cost centres, job numbers and revenue GL codes come from the live Xero
organisation (Cost Centres and Job Numbers tracking categories, chart of
accounts), so the tracker ties to Xero without a mapping table.
"""
import datetime
import json
import re
from pathlib import Path

import openpyxl
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as gcl
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableFormula, TableStyleInfo

ROOT = Path("/home/user/Claude/reporting")
JOBS = json.loads((ROOT / "data/revenue_jobs.json").read_text())
MIG = json.loads((ROOT / "data/revenue_migration.json").read_text())
ACC = json.loads((ROOT / "data/revenue_accounts.json").read_text())
OUT = ROOT / "FY27_Revenue_Tracker_v3.xlsx"

# --------------------------------------------------------------------------
# openpyxl is not Excel: post-2007 functions must be persisted with their
# internal prefix or Excel reports the file as unreadable and drops the formula.
_XLFN = ["VSTACK", "HSTACK", "CHOOSECOLS", "UNIQUE", "LET", "XLOOKUP",
         "TEXTJOIN", "SEQUENCE", "IFNA", "TOCOL", "SORTBY", "TAKE", "DROP"]
_XLWS = ["FILTER", "SORT"]          # worksheet-scoped: _xlfn._xlws.NAME


def fx(formula):
    """Rewrite a readable formula into the form openpyxl must persist."""
    for name in _XLWS:
        formula = re.sub(r"(?<![A-Z0-9_.])" + name + r"\(",
                         "_xlfn._xlws." + name + "(", formula)
    for name in _XLFN:
        formula = re.sub(r"(?<![A-Z0-9_.])" + name + r"\(",
                         "_xlfn." + name + "(", formula)
    return re.sub(r"\s*\n\s*", "", formula)


NAVY, SLATE, LIGHT = "3E5066", "5B708A", "EDF1F6"
CALC_BG, BORDER = "F2F2F2", "B7B7B7"
WARN, GOOD, BAD = "FFF2CC", "E2EFDA", "FCE4E4"
CUR = "$#,##0.00;[Red]-$#,##0.00"
DATE, MON, PCT, TXT, NUM, INT = "dd-mmm-yy", "mmm-yy", "0.0%", "@", "#,##0.00", "#,##0"

THIN = Side(style="thin", color=BORDER)
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
INPUT_FILL = PatternFill("solid", fgColor="FFF9E6")
CALC_FILL = PatternFill("solid", fgColor=CALC_BG)

COST_CENTRES = ACC["cost_centres"]
HDR_ROW, DATA_ROW, DV_LAST, CF_LAST = 4, 5, 20000, 5000
MONTH_END_R0 = 8                       # first cost-centre row on Month-End
PRODUCTION_ROW = MONTH_END_R0 + COST_CENTRES.index("PRODUCTION")
VIDEO_ROW = MONTH_END_R0 + COST_CENTRES.index("VIDEO")

# --------------------------------------------------------------------------
# The spine. Columns 1-19, identical on all four department sheets.
# (header, width, number format, kind)  kind: in=typed, f=formula, dv=dropdown
CORE = [
    ("Date", 11, DATE, "in"), ("Month", 10, MON, "f"), ("Team", 13, TXT, "f"),
    ("Cost Centre", 14, TXT, "dv"), ("Client", 26, TXT, "in"),
    ("Job Number", 15, TXT, "in"), ("Job in Xero?", 12, TXT, "f"),
    ("Description", 46, TXT, "in"), ("Invoice Type", 16, TXT, "dv"),
    ("Xero Invoice No", 16, TXT, "in"), ("Tax Code", 12, TXT, "dv"),
    ("Ex GST", 14, CUR, "in"), ("GST", 12, CUR, "f"), ("Inc GST", 14, CUR, "f"),
    ("Revenue GL", 13, TXT, "dv"), ("Posted to Xero", 14, TXT, "dv"),
    ("To WIP", 9, TXT, "dv"), ("Status", 14, TXT, "dv"), ("Notes", 42, TXT, "in"),
]
NCORE = len(CORE)

# In-table formulas use {Column Name} placeholders, rendered to plain A1
# references for each row. Structured row references ([@Col]) are Excel-native
# but are not resolved by every engine that may open this file, and a formula
# that silently returns #N/A in a revenue tracker is worse than a verbose one.
CORE_FORMULAS = {
    "Month": '=IF({Date}="","",EOMONTH({Date},0))',
    "Job in Xero?": '=IF({Job Number}="","",IF(COUNTIF(lst_Jobs,{Job Number}&"")>0,"OK","CHECK"))',
    "GST": '=IF({Ex GST}="","",IF({Tax Code}="GST 10%",ROUND({Ex GST}*0.1,2),0))',
    "Inc GST": '=IF({Ex GST}="","",{Ex GST}+{GST})',
}

EXTRA = {
    "Onsite": [("PO / Reference", 18, TXT, "in"), ("Ariba Status", 15, TXT, "dv"),
               ("Billable Hours", 13, NUM, "in"), ("Approved By", 16, TXT, "in")],
    "Production": [
        # The video split sits first, immediately after the spine, because it is
        # what Month-End reads to separate the PRODUCTION and VIDEO cost centres.
        ("Video Revenue", 15, CUR, "in"), ("Production Revenue", 17, CUR, "f"),
        ("Video Split", 26, TXT, "f"), ("Video GST", 12, CUR, "f"),
        ("Client Email", 30, TXT, "in"), ("Event Date", 12, DATE, "in"),
        ("Current RMS No", 14, TXT, "in"), ("Zoho Number", 14, TXT, "in"),
        ("Job Closed", 11, TXT, "dv"), ("Discounts Included", 17, CUR, "in"),
        ("Discount %", 11, PCT, "f"), ("Cross Hire Expense", 17, CUR, "in"),
        ("Labour Expense (Internal)", 22, CUR, "in"), ("Net Total", 14, CUR, "f"),
        ("Margin", 14, CUR, "f"), ("Margin %", 10, PCT, "f"),
        ("Video Filming Hrs", 15, NUM, "in"), ("Video Editing Hrs", 15, NUM, "in"),
        ("Project Mgmt Hrs", 15, NUM, "in"), ("Video Project Mgmt Hrs", 19, NUM, "in"),
        ("Production Labour Hrs", 18, NUM, "in")],
    "Consulting": [
        ("Qwilr Link", 46, TXT, "in"), ("Opportunity No", 15, TXT, "in"),
        ("Labour Revenue", 15, CUR, "in"), ("Equipment Revenue", 17, CUR, "in"),
        ("Subscription Revenue", 19, CUR, "in"),
        ("Labour Expense (External)", 22, CUR, "in"),
        ("Equipment Expense (Internal)", 26, CUR, "in"),
        ("Subscription & Licences Expense", 28, CUR, "in"),
        ("Total Expense", 14, CUR, "f"), ("Margin", 14, CUR, "f"),
        ("Margin %", 10, PCT, "f"), ("Revenue Split Check", 18, TXT, "f")],
    "Other": [],
}

EXTRA_FORMULAS = {
    "Production": {
        "Production Revenue": '=IF({Ex GST}="","",{Ex GST}-N({Video Revenue}))',
        # GST follows the revenue split exactly, so Inc GST still equals Ex + GST
        # on every line of the Month-End breakdown.
        "Video GST": '=IF({Ex GST}="","",IF(N({Video Revenue})=0,0,'
                     'ROUND({GST}*N({Video Revenue})/{Ex GST},2)))',
        # Credit notes are negative, so the test is on magnitude and matching
        # sign, not on "video is bigger than the invoice".
        "Video Split":
            '=IF({Ex GST}="","",'
            'IF(SIGN(N({Video Revenue}))*SIGN({Ex GST})=-1,"CHECK - video sign",'
            'IF(ABS(N({Video Revenue}))>ABS({Ex GST}),"CHECK - video exceeds invoice",'
            'IF(AND({Cost Centre}="VIDEO",ROUND(N({Video Revenue}),2)<>ROUND({Ex GST},2)),'
            '"CHECK - VIDEO job not fully split",'
            'IF(N({Video Revenue})=0,"All production",'
            'IF(ROUND(N({Video Revenue}),2)=ROUND({Ex GST},2),"All video","Split"))))))',
        "Discount %": '=IFERROR({Discounts Included}/{Net Total},"")',
        "Net Total": '=IF({Ex GST}="","",{Ex GST}-N({Discounts Included}))',
        "Margin": '=IF({Ex GST}="","",{Ex GST}-N({Cross Hire Expense})'
                  '-N({Labour Expense (Internal)}))',
        "Margin %": '=IFERROR({Margin}/{Ex GST},"")',
    },
    "Consulting": {
        "Total Expense": '=IF({Ex GST}="","",N({Labour Expense (External)})'
                         '+N({Equipment Expense (Internal)})+N({Subscription & Licences Expense}))',
        "Margin": '=IF({Ex GST}="","",{Ex GST}-{Total Expense})',
        "Margin %": '=IFERROR({Margin}/{Ex GST},"")',
        "Revenue Split Check":
            '=IF({Ex GST}="","",IF(N({Labour Revenue})+N({Equipment Revenue})'
            '+N({Subscription Revenue})=0,"Not split",IF(ROUND(N({Labour Revenue})'
            '+N({Equipment Revenue})+N({Subscription Revenue}),2)'
            '=ROUND({Ex GST},2),"OK","MISMATCH")))',
    },
}


def render(template, headers, row):
    """Turn {Column Name} placeholders into plain A1 references for one row."""
    out = template
    for name in sorted(headers, key=len, reverse=True):
        out = out.replace("{" + name + "}", gcl(headers.index(name) + 1) + str(row))
    if "{" in out:
        raise ValueError("unresolved placeholder in: " + out)
    return out


WIP_COLS = [
    ("Month", 10, MON, "in"), ("Job Number", 15, TXT, "in"),
    ("Job in Xero?", 12, TXT, "f"), ("Client", 26, TXT, "in"),
    ("Cost Centre", 14, TXT, "dv"), ("Description", 52, TXT, "in"),
    ("Type", 26, TXT, "dv"), ("Xero Invoice No", 16, TXT, "in"),
    ("Amount", 15, CUR, "in"), ("GL Code", 11, TXT, "dv"),
    ("Journal Ref", 14, TXT, "in"), ("Posted to Xero", 14, TXT, "dv"),
    ("Notes", 42, TXT, "in"),
]
WIP_FORMULAS = {
    "Job in Xero?": '=IF({Job Number}="","",IF(COUNTIF(lst_Jobs,{Job Number}&"")>0,"OK","CHECK"))',
}

TABLES = {"Onsite": "tbl_Onsite", "Production": "tbl_Production",
          "Consulting": "tbl_Consulting", "Other": "tbl_Other"}
REV = list(TABLES.values())
DV_FOR = {"Cost Centre": "lst_CostCentre", "Invoice Type": "lst_InvoiceType",
          "Tax Code": "lst_TaxCode", "Revenue GL": "lst_RevGL",
          "Posted to Xero": "lst_YN", "To WIP": "lst_YN", "Status": "lst_Status",
          "Ariba Status": "lst_Ariba", "Job Closed": "lst_YN",
          "Type": "lst_WIPType", "GL Code": "lst_WIPGL", "Month": "lst_Months"}


def stack():
    """The four department tables, spine columns only, stacked.

    [#Data] is not decoration: a bare table name includes the header row in some
    engines, which would put four literal header rows into the Finance list.
    """
    cols = ",".join(str(i) for i in range(1, NCORE + 1))
    return "VSTACK(" + ",".join(f"CHOOSECOLS({t}[#Data],{cols})" for t in REV) + ")"


# A row is real if it carries a date, client, invoice number or amount. Testing
# the date alone silently hid every undated row, and those hold real money.
# LET variable names must not look like cell references. "f1".."f5" ARE cells
# F1:F5, and naming them that returns #VALUE! - which is why the whole Finance
# sheet came back empty. Names here are deliberately un-reference-like.
CRIT = ('keep,--((CHOOSECOLS(d,1)&CHOOSECOLS(d,5)&CHOOSECOLS(d,10)&CHOOSECOLS(d,12))<>""),'
        'fTeam,IF($A$5="All",1,--(CHOOSECOLS(d,3)=$A$5)),'
        'fCC,IF($C$5="All",1,--(CHOOSECOLS(d,4)=$C$5)),'
        'fMonth,IF($E$5="All",1,--(CHOOSECOLS(d,2)=$E$5)),'
        'fStat,IF($G$5="All",1,--(CHOOSECOLS(d,18)=$G$5)),'
        'fFind,IF($I$5="",1,--ISNUMBER(SEARCH($I$5,CHOOSECOLS(d,5)&"|"&CHOOSECOLS(d,6)'
        '&"|"&CHOOSECOLS(d,8)&"|"&CHOOSECOLS(d,10)))),'
        'k,keep*fTeam*fCC*fMonth*fStat*fFind,')


def across(val, pairs, tables=None):
    """SUMIFS the same question across several department tables."""
    return "+".join(
        f"SUMIFS({t}[{val}]," + ",".join(f"{t}[{c}],{v}" for c, v in pairs) + ")"
        for t in (tables or REV))


# Production carries both cost centres on one row, so its contribution comes from
# the split columns rather than from Cost Centre. Every other sheet is one row,
# one cost centre, and is summed the usual way.
OTHERS = [t for t in REV if t != "tbl_Production"]


def _prod(col):
    """A Production-sheet column for the month, limited to its two cost centres.

    Restricting to PRODUCTION and VIDEO matters: a Production row coded to
    anything else is picked up by the ordinary by-cost-centre sum below, so it
    lands on its own line instead of disappearing between the two.
    """
    return "+".join(
        f'SUMIFS(tbl_Production[{col}],tbl_Production[Month],$C$4,'
        f'tbl_Production[Cost Centre],"{cc}")' for cc in ("PRODUCTION", "VIDEO"))


def by_cost_centre(value, video_value, row):
    """Month-End revenue for one cost centre, in the selected month."""
    pairs = [("Month", "$C$4"), ("Cost Centre", f"$A{row}")]
    if row == PRODUCTION_ROW:
        return f"({_prod(value)})-({_prod(video_value)})+{across(value, pairs, OTHERS)}"
    if row == VIDEO_ROW:
        return f"{_prod(video_value)}+{across(value, pairs, OTHERS)}"
    # every other cost centre sums all four sheets, Production included
    return across(value, pairs, REV)


def count_across(pairs):
    return "+".join("COUNTIFS(" + ",".join(f"{t}[{c}],{v}" for c, v in pairs) + ")"
                    for t in REV)


def title_block(ws, title, sub):
    ws.sheet_view.showGridLines = False
    ws["A1"] = title
    ws["A1"].font = Font(size=16, bold=True, color=NAVY)
    ws["A2"] = sub
    ws["A2"].font = Font(size=10, italic=True, color="595959")
    ws.row_dimensions[1].height = 24
    ws.row_dimensions[2].height = 16
    ws.row_dimensions[3].height = 6


def band(ws, row, text, width=9, colour=NAVY):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=width)
    c = ws.cell(row, 1, text)
    c.font = Font(bold=True, size=11, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor=colour)
    c.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[row].height = 20


def col_heads(ws, row, labels, start=1):
    for i, label in enumerate(labels):
        c = ws.cell(row, start + i, label)
        c.font = Font(bold=True, size=9, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=SLATE)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BOX
    ws.row_dimensions[row].height = 30


# --------------------------------------------------------------------------
# Migration out of FY27_Revenue_Tracker_v2.xlsm. The old file kept each
# department in a hidden 26-column storage block; those blocks are what
# data/revenue_migration.json holds.
def _date(s):
    if isinstance(s, str) and len(s) == 10 and s[4] == "-":
        return datetime.datetime.strptime(s, "%Y-%m-%d")
    return s


def _num(v):
    return v if isinstance(v, (int, float)) and not isinstance(v, bool) else None


def _txt(v):
    return None if v in (None, "") else str(v)


def _month_end(d):
    if not isinstance(d, datetime.datetime):
        return None
    nxt = datetime.date(d.year + (d.month // 12), (d.month % 12) + 1, 1)
    return datetime.datetime.combine(nxt - datetime.timedelta(days=1), datetime.time())


def migrate_onsite():
    rows = []
    for dt, client, invno, job, desc, amt, notes in MIG["Support"]["rows"]:
        rows.append({
            "Date": _date(dt), "Team": "Onsite", "Cost Centre": "ONSITE",
            "Client": _txt(client), "Job Number": _txt(job), "Description": _txt(desc),
            "Invoice Type": "Contract", "Xero Invoice No": _txt(invno),
            "Tax Code": "GST 10%", "Ex GST": _num(amt), "Revenue GL": "41100",
            "Posted to Xero": "Y" if invno else "N", "To WIP": "N",
            "Status": "Invoiced" if invno else "To Invoice",
            "Notes": _txt(notes), "Ariba Status": "N/A"})
    return rows


CC_PROD = {"Production": "PRODUCTION", "Video": "VIDEO",
           "Production/Video": "PRODUCTION", "Other": "CTS", None: "PRODUCTION"}


def migrate_production():
    ix = {k: i for i, k in enumerate(MIG["Production"]["headers"])}
    rows = []
    for r in MIG["Production"]["rows"]:
        g = lambda k: r[ix[k]] if k in ix else None
        invno = _txt(g("Xero Invoice No."))
        dept = g("Department")
        rows.append({
            "Date": _date(g("Date")), "Team": "Production",
            "Cost Centre": CC_PROD.get(dept, "PRODUCTION"),
            "Client": _txt(g("Client")), "Job Number": _txt(g("Job Number")),
            "Description": _txt(g("Project Name")), "Invoice Type": "Project",
            "Xero Invoice No": invno, "Tax Code": "GST 10%",
            "Ex GST": _num(g("Invoice Value")),
            "Revenue GL": "42150" if dept == "Video" else "42100",
            "Posted to Xero": _txt(g("Invoice Posted to Xero")) or ("Y" if invno else "N"),
            "To WIP": "N", "Status": "Invoiced" if invno else "To Invoice",
            "Client Email": _txt(g("Client Email")), "Event Date": _date(g("Event Date")),
            "Current RMS No": _txt(g("Current Number")),
            "Zoho Number": _txt(g("Zoho Number")), "Job Closed": _txt(g("Closed")),
            "Video Revenue": _num(g("Video Total")),
            "Discounts Included": _num(g("Discounts Included")),
            "Cross Hire Expense": _num(g("Cross Hire Expense")),
            "Labour Expense (Internal)": _num(g("Labour Expense (Internal)")),
            "Video Filming Hrs": _num(g("Video Filming")),
            "Video Editing Hrs": _num(g("Video Editing")),
            "Project Mgmt Hrs": _num(g("Project Management")),
            "Video Project Mgmt Hrs": _num(g("Video Project Management")),
            "Production Labour Hrs": _num(g("Production Labour Hours"))})
    return rows


CC_CONS = {"Consulting": "CONSULTING", "Integration": "INTEGRATION", None: "CONSULTING"}
ST_CONS = {"In Progress": "To Invoice", "Completed": "Invoiced",
           "Closed": "Paid", "Cancelled": "Cancelled"}


def migrate_consulting():
    ix = {k: i for i, k in enumerate(MIG["Consulting"]["headers"])}
    rows = []
    for r in MIG["Consulting"]["rows"]:
        g = lambda k: r[ix[k]] if k in ix else None
        invno = _txt(g("Invoice Number"))
        dept = g("Department")
        rows.append({
            "Date": _date(g("Date")), "Team": "Consulting",
            "Cost Centre": CC_CONS.get(dept, "CONSULTING"),
            "Client": _txt(g("Client")), "Job Number": _txt(g("Job Number")),
            "Description": _txt(g("Project Name")), "Invoice Type": "Project",
            "Xero Invoice No": invno, "Tax Code": "GST 10%",
            "Ex GST": _num(g("Invoice Value")),
            "Revenue GL": "42800" if dept == "Consulting" else "42300",
            "Posted to Xero": "Y" if invno else "N", "To WIP": "N",
            "Status": ST_CONS.get(g("Status"), "To Invoice"), "Notes": _txt(g("Notes")),
            "Qwilr Link": _txt(g("Qwilr Link")),
            "Labour Revenue": _num(g("Labour Revenue")),
            "Equipment Revenue": _num(g("Equipment Revenue")),
            "Subscription Revenue": _num(g("Subscription Revenue")),
            "Labour Expense (External)": _num(g("Labour Expense (External)")),
            "Equipment Expense (Internal)": _num(g("Equipment Expense (Internal)")),
            "Subscription & Licences Expense": _num(g("Subscription & Licences Expense"))})
    return rows


CC_WIP = {"Production": "PRODUCTION", "Video": "VIDEO", "Integration": "INTEGRATION",
          "Consulting": "CONSULTING", "Support": "ONSITE", "Onsite": "ONSITE", "CTS": "CTS"}


def migrate_wip():
    rows = []
    for dt, job, dept, client, desc, amt in MIG["WIP"]["rows"]:
        amount = _num(amt)
        rows.append({
            "Month": _month_end(_date(dt)), "Job Number": _txt(job),
            "Client": _txt(client),
            "Cost Centre": CC_WIP.get(dept, dept if dept in COST_CENTRES else None),
            "Description": _txt(desc),
            "Type": ("Accrual - unbilled work" if (amount or 0) >= 0
                     else "Deferral - invoiced in advance"),
            "Amount": amount, "GL Code": "11300", "Posted to Xero": "Y",
            "Notes": "Migrated from FY27 Revenue Tracker v2 - confirm type/sign"})
    return rows


# --------------------------------------------------------------------------
def build_lists(wb):
    ws = wb.create_sheet("Lists")
    title_block(ws, "Reference Lists",
                "Every dropdown in this workbook reads from here. Add a value to the "
                "bottom of a list and it appears in the dropdowns - no macros, nothing to re-run.")
    simple = [
        ("A", "Cost Centre", COST_CENTRES, "lst_CostCentre"),
        ("B", "Invoice Type", ["Contract", "Chargeback", "Ad-hoc", "Project",
                               "Progress Claim", "Milestone", "Subscription",
                               "Equipment Sale", "Recharge", "Credit Note"], "lst_InvoiceType"),
        ("C", "Tax Code", ["GST 10%", "GST Free", "Export (0%)", "BAS Excluded"], "lst_TaxCode"),
        ("D", "Status", ["To Invoice", "Draft", "Awaiting Approval", "Invoiced",
                         "Sent", "Paid", "On Hold", "Cancelled", "Credited"], "lst_Status"),
        ("E", "Yes / No", ["Y", "N", "N/A"], "lst_YN"),
        ("F", "Ariba Status", ["N/A", "Awaiting PO", "To Upload", "Uploaded",
                               "Approved", "Rejected"], "lst_Ariba"),
        ("G", "WIP Movement Type", ["Accrual - unbilled work", "Reversal of prior accrual",
                                    "Deferral - invoiced in advance", "Release of deferral",
                                    "Adjustment / correction", "Migrated opening balance"],
         "lst_WIPType"),
        ("H", "WIP GL Code", ACC["wip_gl"], "lst_WIPGL"),
        ("I", "Production Cost Centre", ["PRODUCTION", "VIDEO"], "lst_CostCentrePrd"),
    ]
    for col, head, vals, name in simple:
        ws[f"{col}4"] = head
        for i, v in enumerate(vals):
            c = ws[f"{col}{5 + i}"]
            c.value, c.number_format, c.border = v, TXT, BOX
        ws.column_dimensions[col].width = max(
            14, len(head) + 4, max(len(str(v)) for v in vals) + 3)
        wb.defined_names.add(DefinedName(
            name, attr_text=f"Lists!${col}$5:${col}${4 + len(vals)}"))

    ws["J4"], ws["K4"] = "Revenue GL", "GL Account Name (from Xero)"
    for i, a in enumerate(ACC["revenue_gl"]):
        for col, val in ((10, a["code"]), (11, a["name"])):
            c = ws.cell(5 + i, col, val)
            c.number_format, c.border = TXT, BOX
    ws.column_dimensions["J"].width = 13
    ws.column_dimensions["K"].width = 36
    wb.defined_names.add(DefinedName(
        "lst_RevGL", attr_text=f"Lists!$J$5:$J${4 + len(ACC['revenue_gl'])}"))

    # FY24 to FY31, so nothing needs rebuilding for years
    ws["M4"] = "Month (period end)"
    months, y, m = [], 2023, 7
    while (y, m) <= (2031, 6):
        months.append(datetime.date(y + (m // 12), (m % 12) + 1, 1)
                      - datetime.timedelta(days=1))
        m += 1
        if m == 12:
            y, m = y + 1, 0
    for i, d in enumerate(months):
        c = ws.cell(5 + i, 13, d)
        c.number_format, c.border = MON, BOX
    ws.column_dimensions["M"].width = 18
    wb.defined_names.add(DefinedName(
        "lst_Months", attr_text=f"Lists!$M$5:$M${4 + len(months)}"))

    ws["O4"], ws["P4"], ws["Q4"] = "Job Number", "Job Name (from Xero)", "CC"
    dept_map = {"ONS": "ONSITE", "PRD": "PRODUCTION", "VID": "VIDEO",
                "INT": "INTEGRATION", "CONS": "CONSULTING", "CTS": "CTS", "": ""}
    for i, j in enumerate(sorted(JOBS, key=lambda x: str(x["job"]))):
        for col, val in ((15, str(j["job"])), (16, j["name"]),
                         (17, dept_map.get(j["dept"], j["dept"]))):
            c = ws.cell(5 + i, col, val)
            c.number_format, c.border = TXT, BOX
    for col, width in (("O", 16), ("P", 48), ("Q", 14)):
        ws.column_dimensions[col].width = width
    # generous: used by COUNTIF only, so trailing blanks are harmless
    for name, col in (("lst_Jobs", "O"), ("lst_JobName", "P"), ("lst_JobCC", "Q")):
        wb.defined_names.add(DefinedName(name, attr_text=f"Lists!${col}$5:${col}$1500"))

    for col in "ABCDEFGHIJKMOPQ":
        c = ws[f"{col}4"]
        c.font = Font(bold=True, color="FFFFFF", size=10)
        c.fill = PatternFill("solid", fgColor=SLATE)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BOX
    ws.row_dimensions[4].height = 30
    ws.freeze_panes = "A5"


# --------------------------------------------------------------------------
BLURB = {
    "Onsite": "Onsite / Support invoicing (ACC005 s.1). One row per Xero invoice. "
              "Fill left to right - grey columns calculate themselves. Never insert rows above row 5.",
    "Production": "Production & Video invoicing (ACC005 s.2). One row per Xero invoice. "
                  "Set Cost Centre to VIDEO for the video half of a split job so Finance can report the split.",
    "Consulting": "Consulting & Integration invoicing (ACC005 s.3). One row per Xero invoice. "
                  "Split the revenue across Labour / Equipment / Subscription - the split check flags any mismatch.",
    "Other": "Anything that is not Onsite, Production or Consulting - CTS internal recharges, "
             "miscellaneous income, credit notes that do not belong to a department.",
}


def build_entry_sheet(wb, name, cols, formulas, team, rows, tblname, blurb,
                      dv_override=None):
    ws = wb.create_sheet(name)
    title_block(ws, f"{name} - Revenue Entry" if team else name, blurb)
    heads = [c[0] for c in cols]
    for i, (h, w, fmt, kind) in enumerate(cols, start=1):
        c = ws.cell(HDR_ROW, i, h)
        c.font = Font(bold=True, color="FFFFFF", size=10)
        c.fill = PatternFill("solid", fgColor=(SLATE if kind == "f" else NAVY))
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BOX
        ws.column_dimensions[gcl(i)].width = w
    ws.row_dimensions[HDR_ROW].height = 34

    nrows = max(len(rows), 1)
    for ri in range(nrows):
        rec = rows[ri] if ri < len(rows) else {}
        r = DATA_ROW + ri
        for ci, (h, w, fmt, kind) in enumerate(cols, start=1):
            c = ws.cell(r, ci)
            c.number_format, c.border, c.font = fmt, BOX, Font(size=10)
            if kind == "f":
                c.fill = CALC_FILL
                c.font = Font(size=10, color="595959")
                if h == "Team":
                    c.value = f'="{team}"' if team else None
                elif h in formulas:
                    c.value = fx(render(formulas[h], heads, r))
            elif rec.get(h) is not None:
                c.value = rec[h]
        ws.row_dimensions[r].height = 15

    ref = f"A{HDR_ROW}:{gcl(len(cols))}{DATA_ROW + nrows - 1}"
    t = Table(displayName=tblname, ref=ref)
    t.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True,
                                      showColumnStripes=False, showFirstColumn=False,
                                      showLastColumn=False)
    ws.add_table(t)

    # Registering the calculated columns is what makes the formulas follow a new
    # row when someone tabs off the end of the table.
    calcs = dict(formulas)
    if team:
        calcs["Team"] = f'="{team}"'
    for col in t.tableColumns:
        if col.name in calcs:
            col.calculatedColumnFormula = TableFormula(
                attr_text=fx(render(calcs[col.name], heads, DATA_ROW)).lstrip("="))

    for ci, (h, w, fmt, kind) in enumerate(cols, start=1):
        source = (dv_override or {}).get(h, DV_FOR.get(h))
        if kind == "dv" and source:
            # formula1 holds the formula WITHOUT a leading "=" - Excel treats
            # "=name" as malformed and offers to repair the whole workbook.
            dv = DataValidation(
                type="list", formula1=source, allow_blank=True,
                showDropDown=False, errorStyle="warning", showErrorMessage=True,
                errorTitle="Not on the list",
                error=f'"{h}" is not on the list in the Lists sheet. '
                      "Add it there if it is genuinely new.")
            ws.add_data_validation(dv)
            dv.add(f"{gcl(ci)}{DATA_ROW}:{gcl(ci)}{DV_LAST}")

    for header, test, colour in (("Job in Xero?", '="CHECK"', WARN),
                                 ("Posted to Xero", '="N"', BAD),
                                 ("Revenue Split Check", '="MISMATCH"', BAD),
                                 ("Video Split", '<>""', None)):
        if header not in heads:
            continue
        c = gcl(heads.index(header) + 1)
        if header == "Video Split":
            # green for a clean split, red only when the numbers do not add up
            for expr, fill in ((f'LEFT({c}{DATA_ROW},5)="CHECK"', BAD),
                               (f'{c}{DATA_ROW}="Split"', WARN)):
                ws.conditional_formatting.add(
                    f"{c}{DATA_ROW}:{c}{CF_LAST}",
                    FormulaRule(formula=[expr],
                                fill=PatternFill("solid", fgColor=fill), stopIfTrue=False))
            continue
        ws.conditional_formatting.add(
            f"{c}{DATA_ROW}:{c}{CF_LAST}",
            FormulaRule(formula=[f"{c}{DATA_ROW}{test}"],
                        fill=PatternFill("solid", fgColor=colour), stopIfTrue=False))

    # The table supplies its own filter buttons; a second sheet-level autofilter
    # over the same range gives Excel two competing sets of dropdowns.
    ws.freeze_panes = f"F{DATA_ROW}"


# --------------------------------------------------------------------------
def build_finance(wb):
    ws = wb.create_sheet("Finance", 1)
    title_block(ws, "Finance - All Departments",
                "Live view of every department sheet. Change a filter and the list "
                "rebuilds instantly. Row 8 down is one single formula - do not type in it. "
                "Fix data on the department sheet it came from.")
    filters = [("Team", "A", "B", '"All,Onsite,Production,Consulting,Other"', "All", TXT),
               ("Cost Centre", "C", "D", "lst_CostCentre", "All", TXT),
               ("Month", "E", "F", "lst_Months", "All", MON),
               ("Status", "G", "H", "lst_Status", "All", TXT),
               ("Search text", "I", "K", None, "", TXT)]
    for label, c1, c2, dvf, default, fmt in filters:
        ws.merge_cells(f"{c1}4:{c2}4")
        ws.merge_cells(f"{c1}5:{c2}5")
        lc = ws[f"{c1}4"]
        lc.value = label
        lc.font = Font(bold=True, size=9, color="FFFFFF")
        lc.fill = PatternFill("solid", fgColor=SLATE)
        lc.alignment = Alignment(horizontal="center", vertical="center")
        vc = ws[f"{c1}5"]
        vc.value, vc.number_format = default, fmt
        vc.fill = INPUT_FILL
        vc.font = Font(bold=True, size=11)
        vc.alignment = Alignment(horizontal="center", vertical="center")
        for cx in (c1, c2):
            ws[f"{cx}4"].border = BOX
            ws[f"{cx}5"].border = BOX
        if dvf:
            dv = DataValidation(type="list", formula1=dvf, allow_blank=True,
                                errorStyle="warning", showErrorMessage=False)
            ws.add_data_validation(dv)
            dv.add(f"{c1}5")
    ws["L5"] = "<- type All to clear a filter"
    ws["L5"].font = Font(size=9, italic=True, color="808080")
    ws.row_dimensions[4].height = 20
    ws.row_dimensions[5].height = 20

    ws.merge_cells("A6:S6")
    ws["A6"] = fx(
        f"=LET(d,{stack()},{CRIT}"
        "cnt,SUM(k),tot,SUM(keep),"
        "ex,SUMPRODUCT(k,IFERROR(CHOOSECOLS(d,12)*1,0)),"
        "gs,SUMPRODUCT(k,IFERROR(CHOOSECOLS(d,13)*1,0)),"
        "ic,SUMPRODUCT(k,IFERROR(CHOOSECOLS(d,14)*1,0)),"
        '"Showing "&TEXT(cnt,"#,##0")&" of "&TEXT(tot,"#,##0")&" records"'
        '&"      Ex GST "&TEXT(ex,"$#,##0.00")'
        '&"      GST "&TEXT(gs,"$#,##0.00")'
        '&"      Inc GST "&TEXT(ic,"$#,##0.00"))')
    ws["A6"].font = Font(bold=True, size=11, color=NAVY)
    ws["A6"].fill = PatternFill("solid", fgColor=LIGHT)
    ws["A6"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws["A6"].border = BOX
    ws.row_dimensions[6].height = 22

    for i, (h, w, fmt, kind) in enumerate(CORE, start=1):
        c = ws.cell(7, i, h)
        c.font = Font(bold=True, color="FFFFFF", size=10)
        c.fill = PatternFill("solid", fgColor=NAVY)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BOX
        # Column-level formats: the whole column reads correctly with no stored
        # cells. Styling 8000 empty rows per column is its own kind of bloat.
        dim = ws.column_dimensions[gcl(i)]
        dim.width, dim.number_format, dim.font = w, fmt, Font(size=10)
    ws.row_dimensions[7].height = 34
    ws["A8"] = fx(
        f"=LET(d,{stack()},{CRIT}"
        'SORT(FILTER(d,k=1,"No records match these filters - widen them '
        'or check the department sheets"),1,-1))')
    ws.freeze_panes = "A8"


# --------------------------------------------------------------------------
CHECKS = [
    ("Invoice rows with an amount but no date",
     "+".join(f'SUMPRODUCT(--({t}[Date]=""),--({t}[Ex GST]<>""))' for t in REV)),
    ("Invoice rows with an amount but no job number",
     count_across([("Ex GST", '"<>"'), ("Job Number", '""')])),
    ("Job numbers not found in the Xero job list",
     "+".join(f'COUNTIF({t}[Job in Xero?],"CHECK")' for t in REV)),
    ("Marked posted to Xero but no Xero invoice number",
     count_across([("Posted to Xero", '"Y"'), ("Xero Invoice No", '""')])),
    ("Duplicate Xero invoice numbers across all departments",
     "IFERROR(LET(a,VSTACK(" + ",".join(f"{t}[Xero Invoice No]" for t in REV)
     + '),v,FILTER(a,a<>""),ROWS(v)-ROWS(UNIQUE(v))),0)'),
    ("Invoice rows with an amount but no tax code",
     count_across([("Ex GST", '"<>"'), ("Tax Code", '""')])),
    ("Invoice rows with an amount but no revenue GL code",
     count_across([("Ex GST", '"<>"'), ("Revenue GL", '""')])),
    ('Flagged "To WIP = Y" but no matching WIP movement',
     "+".join(f'SUMPRODUCT(--({t}[To WIP]="Y"),--({t}[Xero Invoice No]<>""),'
              f'--(COUNTIF(tbl_WIP[Xero Invoice No],{t}[Xero Invoice No]&"")=0))'
              for t in REV)),
    ("WIP movements with an amount but no month or no job",
     'COUNTIFS(tbl_WIP[Month],"",tbl_WIP[Amount],"<>")'
     '+COUNTIFS(tbl_WIP[Job Number],"",tbl_WIP[Amount],"<>")'),
    ("WIP job numbers not found in the Xero job list",
     'COUNTIF(tbl_WIP[Job in Xero?],"CHECK")'),
    ("Consulting revenue split does not equal the invoice",
     'COUNTIF(tbl_Consulting[Revenue Split Check],"MISMATCH")'),
    ("Production rows where the video split does not add up",
     'COUNTIF(tbl_Production[Video Split],"CHECK*")'),
    ("Production rows coded to neither PRODUCTION nor VIDEO",
     'SUMPRODUCT(--(tbl_Production[Cost Centre]<>"PRODUCTION"),'
     '--(tbl_Production[Cost Centre]<>"VIDEO"),--(tbl_Production[Cost Centre]<>""))'),
    ("Production rows with video hours but no video revenue",
     'SUMPRODUCT(--(IFERROR(tbl_Production[Video Filming Hrs]*1,0)'
     '+IFERROR(tbl_Production[Video Editing Hrs]*1,0)'
     '+IFERROR(tbl_Production[Video Project Mgmt Hrs]*1,0)>0),'
     '--(IFERROR(tbl_Production[Video Revenue]*1,0)=0))'),
    ('Still sitting at "To Invoice" for the selected month',
     count_across([("Month", "$C$4"), ("Status", '"To Invoice"')])),
]


def build_month_end(wb):
    ws = wb.create_sheet("Month-End", 2)
    title_block(ws, "Month-End Revenue Close",
                "Pick the month, type the Xero figures into the yellow cells, and "
                "work down. Everything white or grey calculates itself.")
    ws["A4"] = "Month being closed"
    ws["A4"].font = Font(bold=True, size=11)
    c = ws["C4"]
    c.value = datetime.datetime(2026, 8, 31)
    c.number_format, c.fill, c.border = MON, INPUT_FILL, BOX
    c.font = Font(bold=True, size=12, color=NAVY)
    c.alignment = Alignment(horizontal="center")
    dv = DataValidation(type="list", formula1="lst_Months", allow_blank=False,
                        errorStyle="warning", showErrorMessage=False)
    ws.add_data_validation(dv)
    dv.add("C4")
    ws["D4"] = "<- month end date"
    ws["D4"].font = Font(size=9, italic=True, color="808080")
    for col, width in zip("ABCDEFGHI", (22, 16, 14, 15, 16, 18, 17, 15, 20)):
        ws.column_dimensions[col].width = width

    band(ws, 6, "1.  REVENUE BY COST CENTRE  -  tracker vs Xero")
    col_heads(ws, 7, ["Cost Centre", "Invoiced Ex GST", "GST", "Inc GST",
                      "WIP Movement", "Revenue Recognised", "Xero Revenue (type in)",
                      "Variance", "Check"])
    r0 = MONTH_END_R0
    for i, cc in enumerate(COST_CENTRES):
        r = r0 + i
        ws.cell(r, 1, cc).font = Font(bold=True, size=10)
        ws.cell(r, 2).value = fx("=" + by_cost_centre("Ex GST", "Video Revenue", r))
        ws.cell(r, 3).value = fx("=" + by_cost_centre("GST", "Video GST", r))
        ws.cell(r, 4).value = f"=B{r}+C{r}"
        ws.cell(r, 5).value = fx(f"=SUMIFS(tbl_WIP[Amount],tbl_WIP[Month],$C$4,"
                                 f"tbl_WIP[Cost Centre],$A{r})")
        ws.cell(r, 6).value = f"=B{r}+E{r}"
        ws.cell(r, 8).value = f'=IF($G{r}="","",$F{r}-$G{r})'
        ws.cell(r, 9).value = (f'=IF($G{r}="","Enter Xero figure",IF(ROUND($H{r},2)=0,'
                               f'"Reconciled","CHECK - "&TEXT($H{r},"$#,##0.00")))')
    tot = r0 + len(COST_CENTRES)
    ws.cell(tot, 1, "TOTAL")
    for col in range(2, 9):
        ws.cell(tot, col).value = f"=SUM({gcl(col)}{r0}:{gcl(col)}{tot - 1})"
    ws.cell(tot, 9).value = (f'=IF(COUNT($G{r0}:$G{tot - 1})=0,"Enter Xero figures",'
                             f'IF(ROUND($H{tot},2)=0,"Reconciled",'
                             f'"CHECK - "&TEXT($H{tot},"$#,##0.00")))')
    for r in range(r0, tot + 1):
        for col in range(1, 10):
            cell = ws.cell(r, col)
            cell.border = BOX
            if 2 <= col <= 8:
                cell.number_format = CUR
            if col == 7:
                cell.fill = INPUT_FILL
            elif col in (4, 6, 8):
                cell.fill = CALC_FILL
            if r == tot:
                cell.font = Font(bold=True, size=10, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor=SLATE)
    for value, colour in (("CHECK", BAD), ("Reconciled", GOOD)):
        test = (f'LEFT(I{r0},5)="CHECK"' if value == "CHECK" else f'I{r0}="Reconciled"')
        ws.conditional_formatting.add(f"I{r0}:I{tot}", FormulaRule(
            formula=[test], fill=PatternFill("solid", fgColor=colour)))
    return ws, tot


VIDEO_PANEL = [
    ("Video revenue (ex GST)",
     'SUMIFS(tbl_Production[Video Revenue],tbl_Production[Month],$C$4)', CUR),
    ("Production revenue (ex GST)",
     'SUMIFS(tbl_Production[Production Revenue],tbl_Production[Month],$C$4)', CUR),
    ("Total Production sheet revenue (ex GST)", None, CUR),
    ("Video as a share of the Production sheet", None, PCT),
    # counts credit notes too, which a ">0" criterion would miss
    ("Jobs with a video component",
     'SUMPRODUCT(--(tbl_Production[Month]=$C$4),'
     '--(IFERROR(tbl_Production[Video Revenue]*1,0)<>0))', INT),
    ("Video hours (filming + editing + project management)",
     'SUMPRODUCT(--(tbl_Production[Month]=$C$4),'
     'IFERROR(tbl_Production[Video Filming Hrs]*1,0)'
     '+IFERROR(tbl_Production[Video Editing Hrs]*1,0)'
     '+IFERROR(tbl_Production[Video Project Mgmt Hrs]*1,0))', NUM),
]


def build_video_panel(ws, tot):
    """Video and production revenue for the month, off the split columns.

    The two revenue lines add back to the Production sheet total, so this is the
    same money as section 1 - just shown as the split the department books it in.
    """
    s = tot + 2
    band(ws, s, "2.  VIDEO / PRODUCTION SPLIT  -  from the Production sheet")
    col_heads(ws, s + 1, ["", "Amount", "", "", "", "", "", "", ""])
    b = s + 2
    for i, (label, formula, fmt) in enumerate(VIDEO_PANEL):
        r = b + i
        ws.cell(r, 1, label).font = Font(size=10, bold=formula is None)
        vc = ws.cell(r, 2)
        vc.number_format, vc.border = fmt, BOX
        if formula:
            vc.value, vc.fill = fx("=" + formula), CALC_FILL
        elif label.startswith("Total"):
            vc.value = f"=B{b}+B{b + 1}"
            vc.fill = PatternFill("solid", fgColor=LIGHT)
            vc.font = Font(bold=True, size=10, color=NAVY)
        else:
            vc.value = f'=IFERROR(B{b}/B{b + 2},"")'
            vc.fill = PatternFill("solid", fgColor=LIGHT)
            vc.font = Font(bold=True, size=10, color=NAVY)
    note = b + len(VIDEO_PANEL)
    ws.cell(note, 1,
            "Video revenue is typed on the Production sheet (column T). Production "
            "Revenue next to it is the remainder, so one figure does the whole split."
            ).font = Font(size=9, italic=True, color="808080")
    return note


def build_month_end_rest(ws, tot):
    tot = build_video_panel(ws, tot)
    # ---------------- 3. WIP reconciliation
    s2 = tot + 2
    band(ws, s2, "3.  WORK IN PROGRESS  -  GL 11300")
    col_heads(ws, s2 + 1, ["", "Amount", "", "", "", "", "", "", "Check"])
    b = s2 + 2
    # A blank Month would otherwise count as zero and fall into the opening
    # balance, so opening is guarded on the month being present.
    rows = [
        ("Opening WIP balance (all months before this one)",
         fx('=SUMIFS(tbl_WIP[Amount],tbl_WIP[Month],"<>",tbl_WIP[Month],"<"&$C$4)'), "calc"),
        ("Movement this month",
         fx("=SUMIFS(tbl_WIP[Amount],tbl_WIP[Month],$C$4)"), "calc"),
        ("Closing WIP balance per this tracker", None, "sum"),
        ("Closing balance per the WIP Schedule file", None, "input"),
        ("Variance - tracker vs WIP Schedule", None, "var1"),
        ("Closing balance per Xero GL 11300", None, "input"),
        ("Variance - tracker vs Xero", None, "var2"),
    ]
    for i, (label, formula, kind) in enumerate(rows):
        r = b + i
        ws.cell(r, 1, label).font = Font(size=10, bold=kind in ("sum", "var1", "var2"))
        vc = ws.cell(r, 2)
        vc.number_format, vc.border = CUR, BOX
        if kind == "calc":
            vc.value, vc.fill = formula, CALC_FILL
        elif kind == "sum":
            vc.value = f"=B{b}+B{b + 1}"
            vc.fill = PatternFill("solid", fgColor=LIGHT)
            vc.font = Font(bold=True, size=10, color=NAVY)
        elif kind == "input":
            vc.fill = INPUT_FILL
        else:
            src = b + 3 if kind == "var1" else b + 5
            note = ("Enter WIP Schedule figure" if kind == "var1"
                    else "Enter Xero 11300 balance")
            vc.value = f'=IF(B{src}="","",B{b + 2}-B{src})'
            vc.fill, vc.font = CALC_FILL, Font(bold=True, size=10)
            sc = ws.cell(r, 9)
            sc.value = f'=IF(B{src}="","{note}",IF(ROUND(B{r},2)=0,"Reconciled","CHECK"))'
            sc.font = Font(bold=True, size=9)
        ws.cell(r, 9).border = BOX
        ws.cell(r, 9).alignment = Alignment(horizontal="center")
    for value, colour in (("CHECK", BAD), ("Reconciled", GOOD)):
        ws.conditional_formatting.add(f"I{b}:I{b + 6}", FormulaRule(
            formula=[f'I{b}="{value}"'], fill=PatternFill("solid", fgColor=colour)))
    sign = b + 7
    ws.cell(sign, 1,
            "Sign rule: + = revenue recognised this month and WIP balance goes up.  "
            "- = revenue deferred out of this month and WIP balance goes down."
            ).font = Font(size=9, italic=True, color="808080")

    # ---------------- 3. data checks
    s3 = sign + 2
    band(ws, s3, "4.  DATA CHECKS  -  every count should be zero before you close")
    col_heads(ws, s3 + 1, ["Check", "Count", "", "", "", "", "", "", "Status"])
    cb = s3 + 2
    for i, (label, formula) in enumerate(CHECKS):
        r = cb + i
        ws.cell(r, 1, label).font = Font(size=10)
        vc = ws.cell(r, 2)
        vc.value = fx("=" + formula)
        vc.number_format, vc.border, vc.fill = INT, BOX, CALC_FILL
        vc.alignment = Alignment(horizontal="center")
        vc.font = Font(bold=True, size=10)
        sc = ws.cell(r, 9)
        sc.value = f'=IF(B{r}=0,"Clear","REVIEW")'
        sc.border = BOX
        sc.alignment = Alignment(horizontal="center")
        sc.font = Font(bold=True, size=9)
    # Undated rows hold real money ($93,346.99 at migration) and belong to no
    # month, so the count alone understates them - show the dollars.
    money = cb + len(CHECKS)
    ws.cell(money, 1, "Value of invoice rows with no date (excluded from every month)"
            ).font = Font(bold=True, size=10)
    vc = ws.cell(money, 2)
    vc.value = fx("=" + "+".join(
        f'SUMPRODUCT(--({t}[Date]=""),IFERROR({t}[Ex GST]*1,0))' for t in REV))
    vc.number_format, vc.border, vc.fill = CUR, BOX, CALC_FILL
    vc.alignment = Alignment(horizontal="center")
    vc.font = Font(bold=True, size=10)
    sc = ws.cell(money, 9)
    sc.value = f'=IF(ROUND(B{money},2)=0,"Clear","REVIEW")'
    sc.border = BOX
    sc.alignment = Alignment(horizontal="center")
    sc.font = Font(bold=True, size=9)
    for value, colour in (("REVIEW", BAD), ("Clear", GOOD)):
        ws.conditional_formatting.add(f"I{cb}:I{money}", FormulaRule(
            formula=[f'I{cb}="{value}"'], fill=PatternFill("solid", fgColor=colour)))
    ws.conditional_formatting.add(f"B{cb}:B{money - 1}", CellIsRule(
        operator="greaterThan", formula=["0"], fill=PatternFill("solid", fgColor=WARN)))

    # ---------------- 4. sign-off
    s4 = money + 2
    band(ws, s4, "5.  SIGN-OFF")
    for i, (label, who) in enumerate([("Prepared by", "Accounts Assistant"),
                                      ("Reviewed by", "Finance Operations Manager"),
                                      ("Date closed", ""),
                                      ("Notes / carried forward items", "")]):
        r = s4 + 1 + i
        ws.cell(r, 1, label).font = Font(bold=True, size=10)
        for col in range(2, 5):
            ws.cell(r, col).border = BOX
            ws.cell(r, col).fill = INPUT_FILL
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        if who:
            ws.cell(r, 5, who).font = Font(size=9, italic=True, color="808080")
    ws.freeze_panes = "A5"


def build_wip_summary(wb):
    ws = wb.create_sheet("WIP Summary", 3)
    title_block(ws, "WIP Balance by Job",
                "Job-level WIP for the month selected on the Month-End sheet. Paste the "
                "matching balance from the WIP Schedule file into column G and the "
                "variance column finds the breaks.")
    ws["A4"] = "Month"
    ws["A4"].font = Font(bold=True, size=11)
    c = ws["C4"]
    c.value = "='Month-End'!C4"
    c.number_format, c.fill, c.border = MON, CALC_FILL, BOX
    c.font = Font(bold=True, size=12, color=NAVY)
    c.alignment = Alignment(horizontal="center")
    ws["D4"] = "<- set this on the Month-End sheet"
    ws["D4"].font = Font(size=9, italic=True, color="808080")

    totals = [("Opening total", fx('=SUMIFS(tbl_WIP[Amount],tbl_WIP[Month],"<>",'
                                   'tbl_WIP[Month],"<"&$C$4)')),
              ("Movement this month", fx("=SUMIFS(tbl_WIP[Amount],tbl_WIP[Month],$C$4)")),
              ("Closing total", "=B6+B7")]
    for i, (label, formula) in enumerate(totals):
        r = 6 + i
        ws.cell(r, 1, label).font = Font(bold=True, size=10)
        vc = ws.cell(r, 2)
        vc.value, vc.number_format, vc.border = formula, CUR, BOX
        vc.fill = CALC_FILL if r < 8 else PatternFill("solid", fgColor=LIGHT)
        vc.font = Font(bold=True, size=10, color=NAVY)

    col_heads(ws, 10, ["Job Number", "Client", "Cost Centre", "Opening", "Movement",
                       "Closing", "Per WIP Schedule (paste in)", "Variance"])
    for i, w in enumerate((16, 26, 14, 15, 15, 15, 20, 14)):
        dim = ws.column_dimensions[gcl(i + 1)]
        dim.width, dim.number_format = w, (CUR if i >= 3 else TXT)
    # One formula spills the whole block, so the job list maintains itself.
    ws["A11"] = fx(
        '=LET(m,$C$4,'
        'j,SORT(UNIQUE(FILTER(tbl_WIP[Job Number],tbl_WIP[Job Number]<>"","No WIP rows"))),'
        'o,SUMIFS(tbl_WIP[Amount],tbl_WIP[Job Number],j,tbl_WIP[Month],"<>",'
        'tbl_WIP[Month],"<"&m),'
        'v,SUMIFS(tbl_WIP[Amount],tbl_WIP[Job Number],j,tbl_WIP[Month],m),'
        'HSTACK(j,XLOOKUP(j,tbl_WIP[Job Number],tbl_WIP[Client],""),'
        'XLOOKUP(j,tbl_WIP[Job Number],tbl_WIP[Cost Centre],""),o,v,o+v))')
    for r in range(11, 1211):
        g = ws.cell(r, 7)
        g.fill, g.number_format, g.border = INPUT_FILL, CUR, BOX
        h = ws.cell(r, 8)
        h.value = f'=IF($A{r}="","",IF($G{r}="","",$F{r}-$G{r}))'
        h.number_format, h.border, h.fill = CUR, BOX, CALC_FILL
    ws.conditional_formatting.add("H11:H1210", FormulaRule(
        formula=['AND(H11<>"",ROUND(H11,2)<>0)'],
        fill=PatternFill("solid", fgColor=BAD)))
    ws.freeze_panes = "A11"


# --------------------------------------------------------------------------
README = [
 ("T", "FY27 Revenue Tracker"),
 ("S", "Corporate Technology Services Pty Ltd  -  built to replace FY27_Revenue_Tracker_v2.xlsm"),
 ("B", ""),
 ("H", "WHY THIS VERSION IS NOT SLOW"),
 ("P", "The old file was slow because of the macro, not the data. Every time you typed in a cell, the macro "
       "copied 26 columns x 200 rows back into hidden storage columns, and every time you changed the department "
       "dropdown it rewrote the column widths, number formats and dropdowns for the whole sheet."),
 ("P", "It also had a hard ceiling: 200 rows per department, 800 rows in total. It could never have held years of data."),
 ("P", "This workbook has no macros at all. Each department is a real sheet with a real Excel Table. "
       "Finance pulls them together with one formula. Nothing is copied, nothing is rewritten, nothing runs when you type. "
       "It is saved as .xlsx, so there is no macro warning and no blocked-file problem on SharePoint."),
 ("P", "Row limit: about a million rows per sheet. Realistically this holds a decade of invoicing without slowing down."),
 ("B", ""),
 ("H", "HOW IT FITS TOGETHER"),
 ("R", "Onsite / Production / Consulting / Other",
       "Where the departments work. One row per Xero invoice. Type left to right. Grey columns are formulas - leave them alone."),
 ("R", "WIP Movements",
       "Every journal that moves revenue between the P&L and GL 11300 Work in Progress. One row per job per month."),
 ("R", "Finance",
       "Read-only. Stacks all four department sheets into one list of 19 columns. Filter by team, cost centre, "
       "month, status or free text. Rebuilds instantly - there is nothing to refresh."),
 ("R", "Month-End",
       "The close. Revenue by cost centre against Xero, the video/production split, the WIP reconciliation, sixteen data checks, and a sign-off box."),
 ("R", "WIP Summary",
       "WIP opening / movement / closing for every job, for the selected month, so you can tie job by job back to the WIP Schedule file."),
 ("R", "Lists",
       "Every dropdown reads from here. The cost centres, job numbers and revenue GL codes were pulled straight out of your Xero "
       "organisation, so they already match. Add a new value at the bottom of a list and it appears in the dropdowns straight away."),
 ("B", ""),
 ("H", "THE COLUMNS EVERY DEPARTMENT SHARES"),
 ("P", "Columns A to S are identical on all four department sheets - that is what lets Finance stack them. "
       "Department-specific columns start at column T and Finance ignores them. "
       "If you add a fifth department, copy any department sheet, keep columns A to S as they are, name the table, "
       "and add it to the Finance formula."),
 ("B", ""),
 ("H", "VIDEO REVENUE  -  how the Production sheet splits it"),
 ("P", "An invoice can carry both production and video work, and in Xero those are two different Cost "
       "Centres on the same invoice. So the Production sheet splits the money rather than the row."),
 ("P", "   Video Revenue (column T) is the only figure you type. It is the ex-GST portion of the invoice "
       "coded to VIDEO in Xero. Leave it blank for a job with no video."),
 ("P", "   Production Revenue (column U) is the remainder, worked out for you."),
 ("P", "   Video Split (column V) tells you what the row is: All production, Split, All video, or CHECK "
       "if the numbers do not add up. Amber means split, red means something is wrong."),
 ("P", "   Video GST (column W) apportions the GST the same way, so Inc GST still adds up on every line "
       "of the Month-End breakdown."),
 ("P", "For a pure video job set Cost Centre to VIDEO and put the full invoice in Video Revenue. The check "
       "flags it if you set VIDEO but do not split the whole amount. The Cost Centre dropdown on this sheet "
       "is limited to PRODUCTION and VIDEO - anything else belongs on another sheet."),
 ("P", "Month-End section 1 reads these two columns for the PRODUCTION and VIDEO lines, so video revenue "
       "reconciles against the VIDEO cost centre in Xero. The two lines always add back to the Production "
       "sheet total, so nothing can leak between them. Section 2 shows the split for the month on its own."),
 ("B", ""),
 ("H", "THE WIP SIGN RULE  -  read this once and it will always make sense"),
 ("P", "One signed Amount column does both jobs, because GL 11300 nets accrued and deferred revenue."),
 ("P", "   Positive  =  revenue recognised this month.  Work done but not yet invoiced (Dr 11300, Cr Revenue), "
       "or a deferral being released."),
 ("P", "   Negative  =  revenue pushed out of this month.  Invoiced in advance (Dr Revenue, Cr 11300)."),
 ("P", "So: Revenue recognised  =  Invoiced Ex GST  +  WIP movement.  And the running total of the Amount column "
       "is the GL 11300 balance. A negative closing balance means you are net deferred, which is what your current "
       "WIP Schedule shows (-128,544.93 at Aug-26)."),
 ("B", ""),
 ("H", "MONTH-END, IN ORDER"),
 ("P", "1.  Chase the departments until every invoice for the month is on their sheet and Posted to Xero is Y."),
 ("P", "2.  Open Month-End and set the month."),
 ("P", "3.  Work section 3 (Data checks) first. Every count must be zero. Fix on the department sheet, not here."),
 ("P", "4.  Section 1: run the Xero P&L for the month by Cost Centre tracking category and type the revenue into "
       "the yellow column. Each line should read Reconciled."),
 ("P", "5.  Section 2: type the WIP Schedule closing balance and the Xero GL 11300 balance. Both should read Reconciled."),
 ("P", "6.  Breaks in section 2 - open WIP Summary, paste the WIP Schedule balances into column G and the variance "
       "column shows you which job is out."),
 ("P", "7.  Sign off in section 4."),
 ("B", ""),
 ("H", "RULES THAT KEEP IT FAST"),
 ("P", "   Never insert or delete rows above the header row on a department sheet."),
 ("P", "   Add new rows by clicking the last cell of the table and pressing Tab - the table grows and the formulas follow."),
 ("P", "   Do not paste whole columns in. Paste values only, into the table."),
 ("P", "   Do not add conditional formatting over whole columns (A:A). That is the other classic way to make Excel crawl."),
 ("P", "   Leave calculation on Automatic. There is nothing volatile in here - no OFFSET, no INDIRECT, no TODAY in bulk."),
 ("P", "   Keep it as .xlsx. The moment someone saves it as .xlsm and adds a macro, you are back where you started."),
 ("B", ""),
 ("H", "WHAT WAS BROUGHT ACROSS FROM THE OLD FILE"),
 ("P", "   Onsite (was \"Support\"): 54 rows.   Production: 61 rows.   Consulting: 34 rows.   WIP Movements: 126 rows."),
 ("P", "   \"Support\" was renamed ONSITE to match the Cost Centre tracking category in Xero."),
 ("P", "   Department totals tie to the cent: Onsite 512,023.81, Production 460,883.46, Consulting 857,950.27, WIP 17,537.91."),
 ("P", "   Margins, GST, net totals and discount percentages were all recalculated by formula rather than copied, "
       "so the numbers are derived, not stale."),
 ("P", "   Revenue GL codes were assigned by rule on migration (Onsite 41100, Production 42100, Video 42150, "
       "Consulting 42800, Integration 42300). Spot-check these - they are a starting point, not gospel."),
 ("P", "   The 126 migrated WIP rows were given a Type based on the sign of the amount. Confirm those before you rely on them."),
 ("B", ""),
 ("W", "TWO THINGS TO CONFIRM BEFORE YOU RELY ON THIS"),
 ("P", "Production revenue basis. The old Dashboard reported Production on \"Net Total\", which is Invoice Value "
       "minus Discounts Included - $415,476.74 against an invoice value of $460,883.46, a gap of $45,406.72. "
       "Xero holds the invoice value, so this workbook reconciles on Ex GST (invoice value) and keeps Net Total "
       "as a separate management column. If the discount is applied inside the Xero invoice, Ex GST is the right "
       "basis and the old Dashboard was under-reporting Production by the discount. Confirm this."),
 ("P", "Undated rows. $93,346.99 across 21 migrated rows has an amount but no date, so it cannot belong to a month. "
       "They show on the department sheets and on Finance when the month filter is All, and Month-End reports "
       "the dollar value. Date them and the figure goes to zero."),
 ("P", "Separately: the old Dashboard only bucketed Jul-26 to Jun-27, so Consulting rows dated May-25 to Jun-26 were "
       "counted as \"undated\" - that is why its reconciliation showed $495,403.50 unallocated against Consulting. "
       "This workbook covers FY24 to FY31, so those rows land in their real month."),
]


def build_readme(wb):
    ws = wb.create_sheet("Read Me", 0)
    ws.sheet_view.showGridLines = False
    for col, width in (("A", 3), ("B", 30), ("C", 108)):
        ws.column_dimensions[col].width = width
    r = 2
    for item in README:
        kind, body = item[0], item[1]
        if kind == "T":
            ws.cell(r, 2, body).font = Font(size=20, bold=True, color=NAVY)
            ws.row_dimensions[r].height = 28
        elif kind == "S":
            ws.cell(r, 2, body).font = Font(size=10, italic=True, color="595959")
            ws.row_dimensions[r].height = 16
        elif kind == "B":
            ws.row_dimensions[r].height = 10
        elif kind in ("H", "W"):
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
            c = ws.cell(r, 2, body)
            c.font = Font(size=11, bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor=(NAVY if kind == "H" else "8B2B2B"))
            c.alignment = Alignment(vertical="center", indent=1)
            ws.row_dimensions[r].height = 22
        elif kind == "R":
            lc = ws.cell(r, 2, body)
            lc.font = Font(size=10, bold=True, color=NAVY)
            lc.alignment = Alignment(vertical="top")
            lc.fill = PatternFill("solid", fgColor=LIGHT)
            lc.border = BOX
            c = ws.cell(r, 3, item[2])
            c.font = Font(size=10)
            c.alignment = Alignment(wrap_text=True, vertical="top")
            c.border = BOX
            ws.row_dimensions[r].height = 15 * max(1, (len(item[2]) // 105) + 1)
        else:
            c = ws.cell(r, 3, body)
            c.font = Font(size=10)
            c.alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[r].height = 14 * max(1, (len(body) // 108) + 1)
        r += 1


# --------------------------------------------------------------------------
def main():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    build_lists(wb)

    data = {"Onsite": migrate_onsite(), "Production": migrate_production(),
            "Consulting": migrate_consulting(), "Other": []}
    for name in ("Onsite", "Production", "Consulting", "Other"):
        formulas = dict(CORE_FORMULAS)
        formulas.update(EXTRA_FORMULAS.get(name, {}))
        build_entry_sheet(wb, name, CORE + EXTRA[name], formulas, name,
                          data[name], TABLES[name], BLURB[name],
                          dv_override={"Cost Centre": "lst_CostCentrePrd"}
                          if name == "Production" else None)
    build_entry_sheet(
        wb, "WIP Movements", WIP_COLS, WIP_FORMULAS, None, migrate_wip(), "tbl_WIP",
        "Every journal that moves revenue between the P&L and GL 11300 Work in Progress. "
        "SIGN RULE: + = revenue recognised this month (WIP balance up). - = revenue "
        "deferred out of this month (WIP balance down). One row per job per month.")

    build_finance(wb)
    ws, tot = build_month_end(wb)
    build_month_end_rest(ws, tot)
    build_wip_summary(wb)
    build_readme(wb)

    colours = {"Read Me": "7F7F7F", "Finance": NAVY, "Month-End": "2E6B4F",
               "WIP Summary": "2E6B4F", "Onsite": SLATE, "Production": SLATE,
               "Consulting": SLATE, "Other": SLATE, "WIP Movements": "8B6A2B",
               "Lists": "A6A6A6"}
    order = ["Read Me", "Finance", "Month-End", "WIP Summary", "Onsite", "Production",
             "Consulting", "Other", "WIP Movements", "Lists"]
    for name, colour in colours.items():
        wb[name].sheet_properties.tabColor = colour
    wb._sheets = [wb[n] for n in order]
    wb.active = 0
    for sh in wb.worksheets:
        sh.sheet_view.zoomScale = 90
        sh.sheet_view.tabSelected = False
    wb["Read Me"].sheet_view.tabSelected = True
    wb.calculation.fullCalcOnLoad = True
    wb.save(OUT)
    print(f"wrote {OUT} ({OUT.stat().st_size:,} bytes)")


if __name__ == "__main__":
    main()
