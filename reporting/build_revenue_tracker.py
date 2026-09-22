"""Build the FY27 Revenue Tracker workbook.

Ownership is the point of this design. Finance raises every invoice (ACC005),
so Finance owns the revenue record and types on the Finance sheet only. Each
department owns its operational detail and types on its own sheet only. The two
meet on Job Number, which is a real Xero tracking category and already
validated, so there is no synthetic key to keep in step.

Two rules follow from that:

  1. Finance is one row per invoice LINE, coded to one Cost Centre, exactly as
     Xero holds it. A job invoiced part production and part video is two lines,
     so the VIDEO total is a plain sum rather than a split column that has to be
     kept honest.
  2. No dynamic array functions anywhere - no VSTACK, FILTER, SORT, UNIQUE,
     XLOOKUP. Finance is a real Excel Table with real filter buttons that can be
     sorted and pivoted, and every formula here works in any version of Excel.

Cost centres, job numbers and revenue GL codes come from the live Xero
organisation, so the tracker ties to Xero without a mapping table.
"""
import datetime
import json
import re
from pathlib import Path

import openpyxl
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import (Alignment, Border, Font, PatternFill,
                             Protection, Side)
from openpyxl.utils import get_column_letter as gcl
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableFormula, TableStyleInfo

ROOT = Path("/home/user/Claude/reporting")
JOBS = json.loads((ROOT / "data/revenue_jobs.json").read_text())
MIG = json.loads((ROOT / "data/revenue_migration.json").read_text())
ACC = json.loads((ROOT / "data/revenue_accounts.json").read_text())
OUT = ROOT / "FY27_Revenue_Tracker_v3.xlsx"

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
SPARE = 400                      # blank rows kept ready on each entry table

FIN = "tbl_Finance"
DEPTS = ["Onsite", "Production", "Consulting"]
DEPT_TABLE = {d: f"tbl_{d}" for d in DEPTS}


def balanced(formula):
    """True when brackets match outside of quoted text."""
    depth, i, n = 0, 0, len(formula)
    while i < n:
        ch = formula[i]
        if ch == '"':
            i += 1
            while i < n:
                if formula[i] == '"':
                    if i + 1 < n and formula[i + 1] == '"':
                        i += 2
                        continue
                    break
                i += 1
        elif ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
            if depth < 0:
                return False
        i += 1
    return depth == 0


A1_MODE = False          # True only for the test twin, which can be recalculated


def render(template, headers, row):
    """Resolve {Column Name} placeholders.

    Normally to a structured reference - [@[Ex GST]] - which is what Excel
    itself writes and what survives a column being moved or inserted. A1 mode
    produces the same logic with plain cell addresses, for the test copy.
    """
    out = template
    for name in sorted(headers, key=len, reverse=True):
        ref = (gcl(headers.index(name) + 1) + str(row)) if A1_MODE else f"[@[{name}]]"
        out = out.replace("{" + name + "}", ref)
    if "{" in out:
        raise ValueError("unresolved placeholder in: " + out)
    if not balanced(out):
        raise ValueError("unbalanced brackets in: " + out)
    return out


# ---------------------------------------------------------------- Finance
FIN_COLS = [
    ("Date", 11, DATE, "in"), ("Month", 10, MON, "f"),
    ("Xero Invoice No", 16, TXT, "in"), ("Client", 26, TXT, "in"),
    ("Job Number", 15, TXT, "in"), ("Job in Xero?", 12, TXT, "f"),
    ("Description", 44, TXT, "in"), ("Team", 13, TXT, "dv"),
    ("Cost Centre", 14, TXT, "dv"), ("Invoice Type", 15, TXT, "dv"),
    ("Tax Code", 12, TXT, "dv"), ("Ex GST", 14, CUR, "in"),
    ("GST", 12, CUR, "f"), ("Inc GST", 14, CUR, "f"),
    ("Revenue GL", 13, TXT, "dv"), ("Posted to Xero", 14, TXT, "dv"),
    ("To WIP", 9, TXT, "dv"),
    # A16: fill both and the deferral releases itself every month - see the
    # Deferred Revenue sheet. Leave both blank for a normal one-off invoice.
    ("Defer Start", 12, DATE, "in"), ("Defer End", 12, DATE, "in"),
    ("Status", 14, TXT, "dv"), ("Notes", 40, TXT, "in"), ("Issue", 34, TXT, "f"),
]
FIN_FORMULAS = {
    "Month": '=IF({Date}="","",EOMONTH({Date},0))',
    "Job in Xero?": '=IF({Job Number}="","",IF(COUNTIF(lst_Jobs,{Job Number}&"")>0,"OK","CHECK"))',
    "GST": '=IF({Ex GST}="","",IF({Tax Code}="GST 10%",ROUND({Ex GST}*0.1,2),0))',
    "Inc GST": '=IF({Ex GST}="","",{Ex GST}+{GST})',
    # One filterable column instead of a separate exceptions sheet: turn on the
    # filter for Issue and the list of what needs fixing is right here.
    "Issue":
        '=IF(COUNTA({Date},{Client},{Xero Invoice No},{Ex GST})=0,"",'
        'IF({Date}="","No date - sits outside every month",'
        'IF({Job Number}="","No job number",'
        'IF({Job in Xero?}="CHECK","Job number is not in the Xero job list",'
        'IF({Cost Centre}="","No cost centre",'
        'IF(AND({Posted to Xero}="Y",{Xero Invoice No}=""),'
        '"Marked posted to Xero but has no invoice number",'
        'IF(AND({Ex GST}<>"",{Tax Code}=""),"No tax code",'
        'IF(AND({Ex GST}<>"",{Revenue GL}=""),"No revenue GL code",""))))))))',
}

# ---------------------------------------------------------------- departments
# Columns 1-7 are the same on every department sheet: the job, and what Finance
# has invoiced against it. Everything from column 8 is that department's own.
# A12: invoice date and number come first. A2/A7/A8: both are pulled from the
# Finance sheet, matched on the job. A job invoiced more than once shows its
# most recent invoice, and the Invoices column says how many there are.
JOB_CORE = [
    ("Invoice Date", 12, DATE, "f"), ("Invoice No", 16, TXT, "f"),
    ("Job Number", 15, TXT, "in"), ("Job in Xero?", 12, TXT, "f"),
    ("Job Name (Xero)", 38, TXT, "f"), ("Client", 24, TXT, "f"),
    ("Invoices", 10, INT, "f"), ("Revenue Ex GST", 16, CUR, "f"),
    ("Cost Centres", 16, TXT, "f"),
]
JOB_CORE_FORMULAS = {
    "Invoice Date": '=IF({Job Number}="","",IFERROR(IF(SUMPRODUCT(MAX((' + FIN +
                    '[Job Number]={Job Number}&"")*' + FIN + '[Date]))=0,"",'
                    'SUMPRODUCT(MAX((' + FIN + '[Job Number]={Job Number}&"")*'
                    + FIN + '[Date]))),""))',
    "Invoice No": '=IF({Invoice Date}="","",IFERROR(LOOKUP(2,1/((' + FIN +
                  '[Job Number]={Job Number}&"")*(' + FIN + '[Date]={Invoice Date})),'
                  + FIN + '[Xero Invoice No]),""))',
    "Job in Xero?": '=IF({Job Number}="","",IF(COUNTIF(lst_Jobs,{Job Number}&"")>0,"OK","CHECK"))',
    "Job Name (Xero)": '=IF({Job Number}="","",IFERROR(INDEX(lst_JobName,'
                       'MATCH({Job Number}&"",lst_Jobs,0)),""))',
    "Client": '=IF({Job Number}="","",IFERROR(INDEX(' + FIN + '[Client],'
              'MATCH({Job Number}&"",' + FIN + '[Job Number],0)),""))',
    "Invoices": '=IF({Job Number}="","",COUNTIFS(' + FIN + '[Job Number],{Job Number}&"",'
                + FIN + '[Ex GST],"<>"))',
    "Revenue Ex GST": '=IF({Job Number}="","",SUMIFS(' + FIN + '[Ex GST],'
                      + FIN + '[Job Number],{Job Number}&""))',
    "Cost Centres": '=IF({Job Number}="","",IFERROR(INDEX(' + FIN + '[Cost Centre],'
                    'MATCH({Job Number}&"",' + FIN + '[Job Number],0)),"not invoiced yet"))',
}

DEPT_EXTRA = {
    # A6: Ariba Status removed. A10: Onsite gains a Qwilr quote, hyperlinked.
    "Onsite": [("PO / Reference", 18, TXT, "in"), ("Qwilr Quote", 40, TXT, "in"),
               ("Open Qwilr", 14, TXT, "f"), ("Billable Hours", 13, NUM, "in"),
               ("Approved By", 16, TXT, "in"), ("Notes", 38, TXT, "in")],
    # A3: Zoho is the job number, so it is a formula now, not something to type.
    # A5: the Current RMS number is hyperlinked.
    "Production": [
        ("Event Date", 12, DATE, "in"), ("Current RMS No", 14, TXT, "in"),
        ("Open in Current RMS", 18, TXT, "f"), ("Zoho Number", 14, TXT, "f"),
        ("Job Closed", 11, TXT, "dv"),
        ("Discounts Given", 16, CUR, "in"), ("Value Before Discount", 20, CUR, "f"),
        ("Discount %", 11, PCT, "f"), ("Cross Hire Expense", 17, CUR, "in"),
        ("Labour Expense (Internal)", 22, CUR, "in"), ("Total Expense", 14, CUR, "f"),
        ("Margin", 14, CUR, "f"), ("Margin %", 10, PCT, "f"),
        ("Video Filming Hrs", 15, NUM, "in"), ("Video Editing Hrs", 15, NUM, "in"),
        ("Project Mgmt Hrs", 15, NUM, "in"), ("Video Project Mgmt Hrs", 19, NUM, "in"),
        ("Production Labour Hrs", 18, NUM, "in"), ("Notes", 38, TXT, "in")],
    # A11 / A13: Opportunity No removed - it is the job number.
    "Consulting": [
        ("Qwilr Quote", 40, TXT, "in"), ("Open Qwilr", 14, TXT, "f"),
        ("Labour Revenue", 15, CUR, "in"), ("Equipment Revenue", 17, CUR, "in"),
        ("Subscription Revenue", 19, CUR, "in"), ("Revenue Split Check", 18, TXT, "f"),
        ("Labour Expense (External)", 22, CUR, "in"),
        ("Equipment Expense (Internal)", 26, CUR, "in"),
        ("Subscription & Licences Expense", 28, CUR, "in"),
        ("Total Expense", 14, CUR, "f"), ("Margin", 14, CUR, "f"),
        ("Margin %", 10, PCT, "f"), ("Notes", 38, TXT, "in")],
}
# A5 / A10: a quote reference becomes a clickable link. A pasted full URL is used
# as it stands; a bare number is appended to the base address on the Lists sheet.
QWILR = ('=IF({Qwilr Quote}="","",HYPERLINK(IF(LEFT({Qwilr Quote},4)="http",{Qwilr Quote},'
         'set_QwilrBase&{Qwilr Quote}),"Open quote"))')
RMS = ('=IF(OR({Current RMS No}="",set_RMSBase=""),"",'
       'HYPERLINK(set_RMSBase&{Current RMS No},"Open "&{Current RMS No}))')

DEPT_EXTRA_FORMULAS = {
    "Onsite": {"Open Qwilr": QWILR},
    "Production": {
        "Open in Current RMS": RMS,
        "Zoho Number": '=IF({Job Number}="","",{Job Number})',
        "Value Before Discount": '=IF({Job Number}="","",{Revenue Ex GST}+N({Discounts Given}))',
        "Discount %": '=IFERROR({Discounts Given}/{Value Before Discount},"")',
        "Total Expense": '=IF({Job Number}="","",N({Cross Hire Expense})'
                         '+N({Labour Expense (Internal)}))',
        "Margin": '=IF({Job Number}="","",{Revenue Ex GST}-{Total Expense})',
        "Margin %": '=IFERROR({Margin}/{Revenue Ex GST},"")',
    },
    "Consulting": {
        "Open Qwilr": QWILR,
        "Revenue Split Check":
            '=IF({Job Number}="","",IF(N({Labour Revenue})+N({Equipment Revenue})'
            '+N({Subscription Revenue})=0,"Not split",'
            'IF(ROUND(N({Labour Revenue})+N({Equipment Revenue})'
            '+N({Subscription Revenue}),2)=ROUND({Revenue Ex GST},2),"OK","MISMATCH")))',
        "Total Expense": '=IF({Job Number}="","",N({Labour Expense (External)})'
                         '+N({Equipment Expense (Internal)})'
                         '+N({Subscription & Licences Expense}))',
        "Margin": '=IF({Job Number}="","",{Revenue Ex GST}-{Total Expense})',
        "Margin %": '=IFERROR({Margin}/{Revenue Ex GST},"")',
    },
}

WIP_COLS = [
    ("Month", 10, MON, "in"), ("Invoice Date", 12, DATE, "f"),
    ("Xero Invoice No", 16, TXT, "in"), ("Job Number", 15, TXT, "in"),
    ("Job in Xero?", 12, TXT, "f"), ("Client", 26, TXT, "in"),
    ("Cost Centre", 14, TXT, "dv"), ("Description", 50, TXT, "in"),
    ("Type", 26, TXT, "dv"),
    ("Amount", 15, CUR, "in"), ("GL Code", 11, TXT, "dv"),
    ("Journal Ref", 14, TXT, "in"), ("Posted to Xero", 14, TXT, "dv"),
    ("Notes", 40, TXT, "in"),
]
WIP_FORMULAS = {
    "Job in Xero?": '=IF({Job Number}="","",IF(COUNTIF(lst_Jobs,{Job Number}&"")>0,"OK","CHECK"))',
    # A15: the invoice date follows the invoice number in from Finance
    "Invoice Date": '=IF({Xero Invoice No}="","",IFERROR(INDEX(' + FIN + '[Date],'
                    'MATCH({Xero Invoice No}&"",' + FIN + '[Xero Invoice No],0)),""))',
}

DV_FOR = {"Cost Centre": "lst_CostCentre", "Invoice Type": "lst_InvoiceType",
          "Tax Code": "lst_TaxCode", "Revenue GL": "lst_RevGL",
          "Posted to Xero": "lst_YN", "To WIP": "lst_YN", "Status": "lst_Status",
          "Ariba Status": "lst_Ariba", "Job Closed": "lst_YN",
          "Team": "lst_Team", "Type": "lst_WIPType", "GL Code": "lst_WIPGL",
          "Month": "lst_Months"}


# ---------------------------------------------------------------- migration
# The v2 workbook kept each department in a hidden 26-column storage block;
# those blocks are what data/revenue_migration.json holds. They were one row per
# invoice, so a production invoice that included video work carried a separate
# video amount. Here that becomes two Finance lines, which is how Xero holds it.
def _date(s):
    if isinstance(s, str) and len(s) == 10 and s[4] == "-":
        return datetime.datetime.strptime(s, "%Y-%m-%d")
    return s


def _num(v):
    return v if isinstance(v, (int, float)) and not isinstance(v, bool) else None


def _txt(v):
    return None if v in (None, "") else str(v).strip()


def _month_end(d):
    if not isinstance(d, datetime.datetime):
        return None
    nxt = datetime.date(d.year + (d.month // 12), (d.month % 12) + 1, 1)
    return datetime.datetime.combine(nxt - datetime.timedelta(days=1), datetime.time())


CC_PROD = {"Production": "PRODUCTION", "Video": "VIDEO",
           "Production/Video": "PRODUCTION", "Other": "CTS", None: "PRODUCTION"}
CC_CONS = {"Consulting": "CONSULTING", "Integration": "INTEGRATION", None: "CONSULTING"}
ST_CONS = {"In Progress": "To Invoice", "Completed": "Invoiced",
           "Closed": "Paid", "Cancelled": "Cancelled"}
CC_WIP = {"Production": "PRODUCTION", "Video": "VIDEO", "Integration": "INTEGRATION",
          "Consulting": "CONSULTING", "Support": "ONSITE", "Onsite": "ONSITE", "CTS": "CTS"}


def _idx(block):
    return {k: i for i, k in enumerate(MIG[block]["headers"])}


def migrate_finance():
    """Every invoice line, from all four v2 blocks."""
    out = []

    for dt, client, invno, job, desc, amt, notes in MIG["Support"]["rows"]:
        out.append({"Date": _date(dt), "Xero Invoice No": _txt(invno),
                    "Client": _txt(client), "Job Number": _txt(job),
                    "Description": _txt(desc), "Team": "Onsite",
                    "Cost Centre": "ONSITE", "Invoice Type": "Contract",
                    "Tax Code": "GST 10%", "Ex GST": _num(amt), "Revenue GL": "41100",
                    "Posted to Xero": "Y" if invno else "N", "To WIP": "N",
                    "Status": "Invoiced" if invno else "To Invoice", "Notes": _txt(notes)})

    ix = _idx("Production")
    for r in MIG["Production"]["rows"]:
        g = lambda k: r[ix[k]] if k in ix else None
        invno, dept = _txt(g("Xero Invoice No.")), g("Department")
        ex, video = _num(g("Invoice Value")), _num(g("Video Total"))
        base = {"Date": _date(g("Date")), "Xero Invoice No": invno,
                "Client": _txt(g("Client")), "Job Number": _txt(g("Job Number")),
                "Description": _txt(g("Project Name")), "Team": "Production",
                "Invoice Type": "Project", "Tax Code": "GST 10%",
                "Posted to Xero": _txt(g("Invoice Posted to Xero")) or ("Y" if invno else "N"),
                "To WIP": "N", "Status": "Invoiced" if invno else "To Invoice"}
        cc = CC_PROD.get(dept, "PRODUCTION")
        if video and ex is not None and 0 < abs(video) <= abs(ex) and cc != "VIDEO":
            # split the line the way Xero codes it: production part, video part
            out.append(dict(base, **{"Cost Centre": cc, "Ex GST": round(ex - video, 2),
                                     "Revenue GL": "42100",
                                     "Notes": "Production portion of this invoice"}))
            out.append(dict(base, **{"Cost Centre": "VIDEO", "Ex GST": video,
                                     "Revenue GL": "42150",
                                     "Notes": "Video portion of this invoice"}))
        else:
            out.append(dict(base, **{"Cost Centre": cc, "Ex GST": ex,
                                     "Revenue GL": "42150" if cc == "VIDEO" else "42100"}))

    ix = _idx("Consulting")
    for r in MIG["Consulting"]["rows"]:
        g = lambda k: r[ix[k]] if k in ix else None
        invno, dept = _txt(g("Invoice Number")), g("Department")
        out.append({"Date": _date(g("Date")), "Xero Invoice No": invno,
                    "Client": _txt(g("Client")), "Job Number": _txt(g("Job Number")),
                    "Description": _txt(g("Project Name")), "Team": "Consulting",
                    "Cost Centre": CC_CONS.get(dept, "CONSULTING"),
                    "Invoice Type": "Project", "Tax Code": "GST 10%",
                    "Ex GST": _num(g("Invoice Value")),
                    "Revenue GL": "42800" if dept == "Consulting" else "42300",
                    "Posted to Xero": "Y" if invno else "N", "To WIP": "N",
                    "Status": ST_CONS.get(g("Status"), "To Invoice"),
                    "Notes": _txt(g("Notes"))})
    return out


def _accumulate(block, key_col, first_cols, sum_cols):
    """Roll a v2 block up to one row per job: first value wins, amounts add."""
    ix = _idx(block)
    jobs = {}
    for r in MIG[block]["rows"]:
        g = lambda k: r[ix[k]] if k in ix else None
        job = _txt(g(key_col))
        if job is None:
            continue
        rec = jobs.setdefault(job, {"Job Number": job})
        for src, dst in first_cols.items():
            if rec.get(dst) in (None, "") and g(src) not in (None, ""):
                rec[dst] = _date(g(src)) if "Date" in dst else _txt(g(src))
        for src, dst in sum_cols.items():
            v = _num(g(src))
            if v:
                rec[dst] = round(rec.get(dst, 0) + v, 2)
    return list(jobs.values())


def migrate_jobs(dept):
    if dept == "Onsite":
        ix = _idx("Support")
        jobs = {}
        for dt, client, invno, job, desc, amt, notes in MIG["Support"]["rows"]:
            j = _txt(job)
            if j is None:
                continue
            rec = jobs.setdefault(j, {"Job Number": j})
            if not rec.get("Notes") and notes:
                rec["Notes"] = _txt(notes)
        return list(jobs.values())
    if dept == "Production":
        return _accumulate(
            "Production", "Job Number",
            {"Event Date": "Event Date", "Current Number": "Current RMS No",
             "Zoho Number": "Zoho Number", "Closed": "Job Closed"},
            {"Discounts Included": "Discounts Given",
             "Cross Hire Expense": "Cross Hire Expense",
             "Labour Expense (Internal)": "Labour Expense (Internal)",
             "Video Filming": "Video Filming Hrs", "Video Editing": "Video Editing Hrs",
             "Project Management": "Project Mgmt Hrs",
             "Video Project Management": "Video Project Mgmt Hrs",
             "Production Labour Hours": "Production Labour Hrs"})
    return _accumulate(
        "Consulting", "Job Number",
        {"Qwilr Link": "Qwilr Quote", "Notes": "Notes"},
        {"Labour Revenue": "Labour Revenue", "Equipment Revenue": "Equipment Revenue",
         "Subscription Revenue": "Subscription Revenue",
         "Labour Expense (External)": "Labour Expense (External)",
         "Equipment Expense (Internal)": "Equipment Expense (Internal)",
         "Subscription & Licences Expense": "Subscription & Licences Expense"})


def migrate_wip():
    out = []
    for dt, job, dept, client, desc, amt in MIG["WIP"]["rows"]:
        amount = _num(amt)
        out.append({"Month": _month_end(_date(dt)), "Job Number": _txt(job),
                    "Client": _txt(client),
                    "Cost Centre": CC_WIP.get(dept, dept if dept in COST_CENTRES else None),
                    "Description": _txt(desc),
                    "Type": ("Accrual - unbilled work" if (amount or 0) >= 0
                             else "Deferral - invoiced in advance"),
                    "Amount": amount, "GL Code": "11300", "Posted to Xero": "Y",
                    "Notes": "Migrated from FY27 Revenue Tracker v2 - confirm type/sign"})
    return out


# ---------------------------------------------------------------- helpers
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


def totals_strip(ws, cols, tblname, money_cols):
    """A14: the totals for this sheet, sitting above the table."""
    parts = []
    for name in money_cols:
        parts.append(f'"{name}  "&TEXT(SUBTOTAL(109,{tblname}[{name}]),"$#,##0.00")')
    ws.merge_cells(start_row=3, start_column=1, end_row=3,
                   end_column=min(len(cols), 12))
    c = ws.cell(3, 1)
    c.value = "=" + '&"      "&'.join(parts)
    c.font = Font(bold=True, size=11, color=NAVY)
    c.fill = PatternFill("solid", fgColor=LIGHT)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    c.border = BOX
    ws.row_dimensions[3].height = 22


def build_table(ws, cols, formulas, rows, tblname, spare=SPARE, dv_override=None):
    """One entry table: header row, migrated rows, then blank rows ready to use."""
    heads = [c[0] for c in cols]
    for i, (h, w, fmt, kind) in enumerate(cols, start=1):
        c = ws.cell(HDR_ROW, i, h)
        c.font = Font(bold=True, color="FFFFFF", size=10)
        c.fill = PatternFill("solid", fgColor=(SLATE if kind == "f" else NAVY))
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BOX
        ws.column_dimensions[gcl(i)].width = w
    ws.row_dimensions[HDR_ROW].height = 34

    nrows = len(rows) + spare
    for ri in range(nrows):
        rec = rows[ri] if ri < len(rows) else {}
        r = DATA_ROW + ri
        for ci, (h, w, fmt, kind) in enumerate(cols, start=1):
            c = ws.cell(r, ci)
            c.number_format, c.border, c.font = fmt, BOX, Font(size=10)
            if kind == "f":
                c.fill = CALC_FILL
                c.font = Font(size=10, color="595959")
                if h in formulas:
                    c.value = render(formulas[h], heads, r)
            else:
                # A4: only the cells a person is meant to fill stay unlocked
                c.protection = Protection(locked=False)
                if rec.get(h) is not None:
                    c.value = rec[h]
        ws.row_dimensions[r].height = 15

    last = DATA_ROW + nrows - 1
    t = Table(displayName=tblname, ref=f"A{HDR_ROW}:{gcl(len(cols))}{last}")
    t.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True,
                                      showColumnStripes=False, showFirstColumn=False,
                                      showLastColumn=False)
    ws.add_table(t)
    for col in t.tableColumns:
        if col.name in formulas:
            col.calculatedColumnFormula = TableFormula(
                attr_text=render(formulas[col.name], heads, DATA_ROW).lstrip("="))

    for ci, (h, w, fmt, kind) in enumerate(cols, start=1):
        source = (dv_override or {}).get(h, DV_FOR.get(h))
        if kind == "dv" and source:
            dv = DataValidation(
                type="list", formula1=source, allow_blank=True, showDropDown=False,
                errorStyle="warning", showErrorMessage=True, errorTitle="Not on the list",
                error=f'"{h}" is not on the list in the Lists sheet. '
                      "Add it there if it is genuinely new.")
            ws.add_data_validation(dv)
            dv.add(f"{gcl(ci)}{DATA_ROW}:{gcl(ci)}{DV_LAST}")

    # A4: the sheet is protected without a password - it stops a stray keystroke
    # landing in a formula, and anyone who needs to can turn it off in two clicks.
    ws.protection.sheet = True
    ws.protection.autoFilter = False
    ws.protection.sort = False
    ws.protection.formatCells = False
    ws.protection.formatColumns = False
    ws.protection.formatRows = False
    ws.protection.insertRows = False

    for header, value, colour in (("Job in Xero?", "CHECK", WARN),
                                  ("Posted to Xero", "N", BAD),
                                  ("Revenue Split Check", "MISMATCH", BAD)):
        if header in heads:
            c = gcl(heads.index(header) + 1)
            ws.conditional_formatting.add(f"{c}{DATA_ROW}:{c}{CF_LAST}", FormulaRule(
                formula=[f'{c}{DATA_ROW}="{value}"'],
                fill=PatternFill("solid", fgColor=colour), stopIfTrue=False))
    if "Issue" in heads:
        c = gcl(heads.index("Issue") + 1)
        ws.conditional_formatting.add(f"{c}{DATA_ROW}:{c}{CF_LAST}", FormulaRule(
            formula=[f'{c}{DATA_ROW}<>""'],
            fill=PatternFill("solid", fgColor=BAD), stopIfTrue=False))
    return heads, nrows


# ---------------------------------------------------------------- sheets
def build_finance(wb):
    ws = wb.create_sheet("Finance")
    title_block(ws, "Finance - Invoice Register",
                "One row per invoice line, coded to one cost centre, the same way Xero "
                "holds it. Finance types here and nowhere else. A job invoiced part "
                "production and part video is two lines. Filter the Issue column to see "
                "what needs fixing.")
    nrows = build_table(ws, FIN_COLS, FIN_FORMULAS, migrate_finance(), FIN)[1]
    totals_strip(ws, FIN_COLS, FIN, ["Ex GST", "GST", "Inc GST"])
    ws.freeze_panes = f"E{DATA_ROW}"
    return nrows


DEPT_BLURB = {
    "Onsite": "Onsite / Support jobs. One row per job number. Client, invoice count and "
              "revenue come from the Finance sheet - fill in the columns from PO / "
              "Reference onwards.",
    "Production": "Production and Video jobs. One row per job number. Revenue comes from "
                  "the Finance sheet, so margin is the whole job: every invoice against it "
                  "less the costs you enter here.",
    "Consulting": "Consulting and Integration jobs. One row per job number. Revenue comes "
                  "from the Finance sheet; split it across Labour / Equipment / "
                  "Subscription and the check column flags any mismatch.",
}


def build_dept(wb, dept):
    ws = wb.create_sheet(dept)
    title_block(ws, f"{dept} - Job Detail", DEPT_BLURB[dept])
    formulas = dict(JOB_CORE_FORMULAS)
    formulas.update(DEPT_EXTRA_FORMULAS.get(dept, {}))
    build_table(ws, JOB_CORE + DEPT_EXTRA[dept], formulas,
                migrate_jobs(dept), DEPT_TABLE[dept])
    totals_strip(ws, JOB_CORE, DEPT_TABLE[dept], ["Revenue Ex GST"])
    ws.freeze_panes = f"C{DATA_ROW}"


def build_wip(wb):
    ws = wb.create_sheet("WIP Movements")
    title_block(ws, "WIP Movements",
                "Every journal that moves revenue between the P&L and GL 11300 Work in "
                "Progress. SIGN RULE: + = revenue recognised this month (WIP balance up). "
                "- = revenue deferred out of this month (WIP balance down).")
    build_table(ws, WIP_COLS, WIP_FORMULAS, migrate_wip(), "tbl_WIP", spare=200)
    totals_strip(ws, WIP_COLS, "tbl_WIP", ["Amount"])
    ws.freeze_panes = f"D{DATA_ROW}"


CHECKS = [
    ("Invoice lines with an amount but no date",
     f'SUMPRODUCT(--({FIN}[Date]=""),--({FIN}[Ex GST]<>""))'),
    ("Invoice lines with an amount but no job number",
     f'COUNTIFS({FIN}[Ex GST],"<>",{FIN}[Job Number],"")'),
    ("Job numbers not found in the Xero job list",
     f'COUNTIF({FIN}[Job in Xero?],"CHECK")'),
    ("Marked posted to Xero but no invoice number",
     f'COUNTIFS({FIN}[Posted to Xero],"Y",{FIN}[Xero Invoice No],"")'),
    ("Invoice lines with an amount but no cost centre",
     f'COUNTIFS({FIN}[Ex GST],"<>",{FIN}[Cost Centre],"")'),
    ("Invoice lines with an amount but no tax code",
     f'COUNTIFS({FIN}[Ex GST],"<>",{FIN}[Tax Code],"")'),
    ("Invoice lines with an amount but no revenue GL code",
     f'COUNTIFS({FIN}[Ex GST],"<>",{FIN}[Revenue GL],"")'),
    ("Anything flagged in the Issue column",
     f'COUNTIF({FIN}[Issue],"?*")'),
    ('Flagged "To WIP = Y" but no matching WIP movement',
     f'SUMPRODUCT(--({FIN}[To WIP]="Y"),--({FIN}[Xero Invoice No]<>""),'
     f'--(COUNTIF(tbl_WIP[Xero Invoice No],{FIN}[Xero Invoice No]&"")=0))'),
    ("WIP movements with an amount but no month or no job",
     'COUNTIFS(tbl_WIP[Month],"",tbl_WIP[Amount],"<>")'
     '+COUNTIFS(tbl_WIP[Job Number],"",tbl_WIP[Amount],"<>")'),
    ("WIP job numbers not found in the Xero job list",
     'COUNTIF(tbl_WIP[Job in Xero?],"CHECK")'),
    ("Consulting revenue split does not equal the job revenue",
     'COUNTIF(tbl_Consulting[Revenue Split Check],"MISMATCH")'),
    ("Jobs on a department sheet with no invoice yet",
     "+".join(f'COUNTIFS({DEPT_TABLE[d]}[Job Number],"<>",{DEPT_TABLE[d]}[Invoices],0)'
              for d in DEPTS)),
    ("Deferred lines that also have a manual WIP journal (would double count)",
     f'SUMPRODUCT(--({FIN}[Defer Start]<>""),--({FIN}[Xero Invoice No]<>""),'
     f'--(COUNTIF(tbl_WIP[Xero Invoice No],{FIN}[Xero Invoice No]&"")>0))'),
    ("Deferred lines missing an end date",
     f'SUMPRODUCT(--({FIN}[Defer Start]<>""),--({FIN}[Defer End]=""))'),
    ('Still sitting at "To Invoice" for the selected month',
     f'COUNTIFS({FIN}[Month],$C$4,{FIN}[Status],"To Invoice")'),
]


def build_month_end(wb, dlast):
    ws = wb.create_sheet("Month-End")
    title_block(ws, "Month-End Revenue Close",
                "Pick the month, type the Xero figures into the yellow cells, and work "
                "down. Everything white or grey calculates itself.")
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

    # ---- 1. revenue by cost centre, straight off the Finance register
    band(ws, 6, "1.  REVENUE BY COST CENTRE  -  tracker vs Xero")
    col_heads(ws, 7, ["Cost Centre", "Invoiced Ex GST", "GST", "Inc GST",
                      "WIP Movement", "Revenue Recognised", "Xero Revenue (type in)",
                      "Variance", "Check"])
    r0 = 8
    for i, cc in enumerate(COST_CENTRES):
        r = r0 + i
        ws.cell(r, 1, cc).font = Font(bold=True, size=10)
        ws.cell(r, 2).value = f'=SUMIFS({FIN}[Ex GST],{FIN}[Month],$C$4,{FIN}[Cost Centre],$A{r})'
        ws.cell(r, 3).value = f'=SUMIFS({FIN}[GST],{FIN}[Month],$C$4,{FIN}[Cost Centre],$A{r})'
        ws.cell(r, 4).value = f"=B{r}+C{r}"
        # manual journals plus whatever the deferral engine releases this month
        ws.cell(r, 5).value = (
            '=SUMIFS(tbl_WIP[Amount],tbl_WIP[Month],$C$4,'
            f"tbl_WIP[Cost Centre],$A{r})"
            f'+SUMIF({defer_range("F", dlast)},$A{r},{defer_range("Q", dlast)})')
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
    ws.cell(tot + 1, 1, "Every line on the Finance sheet carries one cost centre, so "
                        "these add to the month's invoicing exactly.").font = \
        Font(size=9, italic=True, color="808080")
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

    # ---- 2. WIP
    s2 = tot + 3
    band(ws, s2, "2.  WORK IN PROGRESS  -  GL 11300")
    col_heads(ws, s2 + 1, ["", "Amount", "", "", "", "", "", "", "Check"])
    b = s2 + 2
    rows = [
        ("Opening WIP balance (all months before this one)",
         '=SUMIFS(tbl_WIP[Amount],tbl_WIP[Month],"<>",tbl_WIP[Month],"<"&$C$4)'
         f'+SUM({defer_range("P", dlast)})', "calc"),
        ("Movement this month",
         "=SUMIFS(tbl_WIP[Amount],tbl_WIP[Month],$C$4)"
         f'+SUM({defer_range("Q", dlast)})', "calc"),
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

    # ---- 3. data checks
    s3 = sign + 2
    band(ws, s3, "3.  DATA CHECKS  -  every count should be zero before you close")
    col_heads(ws, s3 + 1, ["Check", "Count", "", "", "", "", "", "", "Status"])
    cb = s3 + 2
    for i, (label, formula) in enumerate(CHECKS):
        r = cb + i
        ws.cell(r, 1, label).font = Font(size=10)
        vc = ws.cell(r, 2)
        vc.value = "=" + formula
        vc.number_format, vc.border, vc.fill = INT, BOX, CALC_FILL
        vc.alignment = Alignment(horizontal="center")
        vc.font = Font(bold=True, size=10)
        sc = ws.cell(r, 9)
        sc.value = f'=IF(B{r}=0,"Clear","REVIEW")'
        sc.border = BOX
        sc.alignment = Alignment(horizontal="center")
        sc.font = Font(bold=True, size=9)
    money = cb + len(CHECKS)
    ws.cell(money, 1, "Value of invoice lines with no date (excluded from every month)"
            ).font = Font(bold=True, size=10)
    vc = ws.cell(money, 2)
    vc.value = f'=SUMPRODUCT(--({FIN}[Date]=""),IFERROR({FIN}[Ex GST]*1,0))'
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

    # ---- 4. sign-off
    s4 = money + 2
    band(ws, s4, "4.  SIGN-OFF")
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



# A16: deferred revenue, released automatically.
#
# Put a start and an end date on a Finance line and the revenue spreads evenly
# across those months instead of landing all in the month invoiced. No monthly
# journal to remember and no rows to add.
#
#   in the month invoiced : recognise one month, defer the rest  (WIP down)
#   every month after     : recognise one more month             (WIP up)
#
# Each row mirrors one Finance line by position - INDEX into the table, nothing
# typed - so there is nothing here that can fall out of step.
DEFER_COLS = [
    ("Line", 7, INT, 'ROW()-{hdr}'),
    ("Date", 11, DATE, 'IFERROR(INDEX(F[Date],$A{r}),"")'),
    ("Xero Invoice No", 15, TXT, 'IFERROR(INDEX(F[Xero Invoice No],$A{r}),"")'),
    ("Client", 22, TXT, 'IFERROR(INDEX(F[Client],$A{r}),"")'),
    ("Job Number", 14, TXT, 'IFERROR(INDEX(F[Job Number],$A{r}),"")'),
    ("Cost Centre", 14, TXT, 'IFERROR(INDEX(F[Cost Centre],$A{r}),"")'),
    ("Ex GST", 14, CUR, 'IFERROR(INDEX(F[Ex GST],$A{r}),"")'),
    ("Defer Start", 12, DATE, 'IFERROR(INDEX(F[Defer Start],$A{r}),"")'),
    ("Defer End", 12, DATE, 'IFERROR(INDEX(F[Defer End],$A{r}),"")'),
    ("Months", 9, INT,
     'IF(OR($H{r}="",$I{r}="",$G{r}=""),"",'
     'MAX(1,(YEAR($I{r})-YEAR($H{r}))*12+MONTH($I{r})-MONTH($H{r})+1))'),
    ("Per Month", 14, CUR, 'IF($J{r}="","",$G{r}/$J{r})'),
    ("Recognised to date", 17, CUR,
     'IF($J{r}="",0,$K{r}*MAX(0,MIN(YEAR($C$3)*12+MONTH($C$3),'
     'YEAR($I{r})*12+MONTH($I{r}))-(YEAR($H{r})*12+MONTH($H{r}))+1))'),
    ("Recognised to prior", 18, CUR,
     'IF($J{r}="",0,$K{r}*MAX(0,MIN(YEAR($C$3)*12+MONTH($C$3)-1,'
     'YEAR($I{r})*12+MONTH($I{r}))-(YEAR($H{r})*12+MONTH($H{r}))+1))'),
    ("Invoiced to date", 16, CUR,
     'IF(OR($J{r}="",$B{r}=""),0,'
     'IF(YEAR($B{r})*12+MONTH($B{r})<=YEAR($C$3)*12+MONTH($C$3),$G{r},0))'),
    ("Invoiced to prior", 17, CUR,
     'IF(OR($J{r}="",$B{r}=""),0,'
     'IF(YEAR($B{r})*12+MONTH($B{r})<=YEAR($C$3)*12+MONTH($C$3)-1,$G{r},0))'),
    ("Opening WIP", 14, CUR, '$M{r}-$O{r}'),
    # movement = what is recognised this month, less what is invoiced this month
    ("Movement", 14, CUR, '($L{r}-$M{r})-($N{r}-$O{r})'),
    ("Closing WIP", 14, CUR, '$L{r}-$N{r}'),
]
DEFER_FIRST = 6          # first data row on the Deferred Revenue sheet


def build_deferred(wb, nlines):
    ws = wb.create_sheet("Deferred Revenue")
    title_block(ws, "Deferred Revenue - released automatically",
                "Every Finance line that carries a Defer Start and a Defer End. The "
                "revenue spreads evenly across those months and the WIP movement works "
                "itself out. Nothing is typed on this sheet - fill the two dates on "
                "Finance and this follows.")
    ws["A3"] = "Month"
    ws["A3"].font = Font(bold=True, size=11)
    c = ws["C3"]
    c.value = "='Month-End'!C4"
    c.number_format, c.fill, c.border = MON, CALC_FILL, BOX
    c.font = Font(bold=True, size=12, color=NAVY)
    c.alignment = Alignment(horizontal="center")
    ws["D3"] = "<- set on the Month-End sheet"
    ws["D3"].font = Font(size=9, italic=True, color="808080")

    hdr = DEFER_FIRST - 1
    for i, (h, w, fmt, _f) in enumerate(DEFER_COLS, start=1):
        cell = ws.cell(hdr, i, h)
        cell.font = Font(bold=True, color="FFFFFF", size=9)
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BOX
        ws.column_dimensions[gcl(i)].width = w
    ws.row_dimensions[hdr].height = 32

    for k in range(nlines):
        r = DEFER_FIRST + k
        for i, (h, w, fmt, f) in enumerate(DEFER_COLS, start=1):
            cell = ws.cell(r, i)
            cell.value = "=" + f.replace("F[", FIN + "[").format(r=r, hdr=hdr)
            cell.number_format, cell.border = fmt, BOX
            cell.font = Font(size=9, color="595959")
            cell.fill = CALC_FILL
    last = DEFER_FIRST + nlines - 1
    ws.auto_filter.ref = f"A{hdr}:{gcl(len(DEFER_COLS))}{last}"
    ws.freeze_panes = f"B{DEFER_FIRST}"
    return last


def defer_range(col, last):
    return f"'Deferred Revenue'!${col}${DEFER_FIRST}:${col}${last}"


def build_wip_summary(wb, dlast):
    """WIP by job, for the selected month.

    The job list is the Xero job list, fixed, so this is plain SUMIFS with no
    array formula to go wrong. Jobs on WIP Movements that Xero does not have are
    counted separately rather than silently dropped.
    """
    ws = wb.create_sheet("WIP Summary")
    title_block(ws, "WIP Balance by Job",
                "Job-level WIP for the month selected on the Month-End sheet. Paste the "
                "matching balance from the WIP Schedule file into column G and the "
                "variance column finds the breaks. Filter column F to hide the nil rows.")
    ws["A4"] = "Month"
    ws["A4"].font = Font(bold=True, size=11)
    c = ws["C4"]
    c.value = "='Month-End'!C4"
    c.number_format, c.fill, c.border = MON, CALC_FILL, BOX
    c.font = Font(bold=True, size=12, color=NAVY)
    c.alignment = Alignment(horizontal="center")
    ws["D4"] = "<- set this on the Month-End sheet"
    ws["D4"].font = Font(size=9, italic=True, color="808080")

    totals = [("Opening total",
               '=SUMIFS(tbl_WIP[Amount],tbl_WIP[Month],"<>",tbl_WIP[Month],"<"&$C$4)'
               f'+SUM({defer_range("P", dlast)})'),
              ("Movement this month", "=SUMIFS(tbl_WIP[Amount],tbl_WIP[Month],$C$4)"
               f'+SUM({defer_range("Q", dlast)})'),
              ("Closing total", "=B6+B7"),
              ("Of which is on a job Xero does not have",
               '=SUMIFS(tbl_WIP[Amount],tbl_WIP[Job in Xero?],"CHECK",'
               'tbl_WIP[Month],"<="&$C$4)')]
    for i, (label, formula) in enumerate(totals):
        r = 6 + i
        ws.cell(r, 1, label).font = Font(bold=True, size=10)
        vc = ws.cell(r, 2)
        vc.value, vc.number_format, vc.border = formula, CUR, BOX
        vc.fill = CALC_FILL if r != 8 else PatternFill("solid", fgColor=LIGHT)
        vc.font = Font(bold=True, size=10, color=NAVY)

    col_heads(ws, 11, ["Job Number", "Job Name (Xero)", "Opening", "Movement",
                       "Closing", "Any balance?", "Per WIP Schedule (paste in)",
                       "Variance"])
    for i, w in enumerate((16, 42, 15, 15, 15, 13, 20, 14)):
        ws.column_dimensions[gcl(i + 1)].width = w
    joblist = sorted(JOBS, key=lambda x: str(x["job"]))
    for i, j in enumerate(joblist):
        r = 12 + i
        ws.cell(r, 1, str(j["job"])).number_format = TXT
        ws.cell(r, 2, j["name"]).number_format = TXT
        ws.cell(r, 3).value = (f'=SUMIFS(tbl_WIP[Amount],tbl_WIP[Job Number],$A{r},'
                               f'tbl_WIP[Month],"<>",tbl_WIP[Month],"<"&$C$4)'
                               f'+SUMIF({defer_range("E", dlast)},$A{r},'
                               f'{defer_range("P", dlast)})')
        ws.cell(r, 4).value = (f'=SUMIFS(tbl_WIP[Amount],tbl_WIP[Job Number],$A{r},'
                               f"tbl_WIP[Month],$C$4)"
                               f'+SUMIF({defer_range("E", dlast)},$A{r},'
                               f'{defer_range("Q", dlast)})')
        ws.cell(r, 5).value = f"=C{r}+D{r}"
        ws.cell(r, 6).value = f'=IF(ROUND(C{r},2)+ROUND(D{r},2)=0,"","yes")'
        ws.cell(r, 8).value = f'=IF($G{r}="","",$E{r}-$G{r})'
        for col in range(1, 9):
            cell = ws.cell(r, col)
            cell.border = BOX
            cell.font = Font(size=10)
            if col in (3, 4, 5, 7, 8):
                cell.number_format = CUR
            if col == 7:
                cell.fill = INPUT_FILL
            elif col in (3, 4, 5, 6, 8):
                cell.fill = CALC_FILL
    last = 11 + len(joblist)
    ws.auto_filter.ref = f"A11:H{last}"
    ws.conditional_formatting.add(f"H12:H{last}", FormulaRule(
        formula=['AND(H12<>"",ROUND(H12,2)<>0)'],
        fill=PatternFill("solid", fgColor=BAD)))
    ws.freeze_panes = "A12"


def build_lists(wb):
    ws = wb.create_sheet("Lists")
    title_block(ws, "Reference Lists",
                "Every dropdown in this workbook reads from here. Add a value to the "
                "bottom of a list and it appears in the dropdowns - no macros, nothing "
                "to re-run.")
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
        ("I", "Team", DEPTS + ["Other"], "lst_Team"),
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

    # link bases used by the Qwilr / Current RMS hyperlink columns
    ws["T4"] = "Link settings"
    ws["T5"], ws["U5"] = "Qwilr base address", "https://cts.qwilr.com/"
    ws["T6"], ws["U6"] = "Current RMS base address", ""
    ws["T7"] = "Paste your Current RMS opportunity address above, ending with a slash, " \
               "and the Production links switch on."
    ws["T7"].font = Font(size=9, italic=True, color="808080")
    for r in (5, 6):
        ws.cell(r, 20).font = Font(size=10, bold=True)
        c = ws.cell(r, 21)
        c.fill, c.border, c.number_format = INPUT_FILL, BOX, TXT
    ws.column_dimensions["T"].width = 26
    ws.column_dimensions["U"].width = 44
    wb.defined_names.add(DefinedName("set_QwilrBase", attr_text="Lists!$U$5"))
    wb.defined_names.add(DefinedName("set_RMSBase", attr_text="Lists!$U$6"))

    ws["K4"], ws["L4"] = "Revenue GL", "GL Account Name (from Xero)"
    for i, a in enumerate(ACC["revenue_gl"]):
        for col, val in ((11, a["code"]), (12, a["name"])):
            c = ws.cell(5 + i, col, val)
            c.number_format, c.border = TXT, BOX
    ws.column_dimensions["K"].width = 13
    ws.column_dimensions["L"].width = 36
    wb.defined_names.add(DefinedName(
        "lst_RevGL", attr_text=f"Lists!$K$5:$K${4 + len(ACC['revenue_gl'])}"))

    ws["N4"] = "Month (period end)"
    months, y, m = [], 2023, 7
    while (y, m) <= (2031, 6):
        months.append(datetime.date(y + (m // 12), (m % 12) + 1, 1)
                      - datetime.timedelta(days=1))
        m += 1
        if m == 12:
            y, m = y + 1, 0
    for i, d in enumerate(months):
        c = ws.cell(5 + i, 14, d)
        c.number_format, c.border = MON, BOX
    ws.column_dimensions["N"].width = 18
    wb.defined_names.add(DefinedName(
        "lst_Months", attr_text=f"Lists!$N$5:$N${4 + len(months)}"))

    ws["P4"], ws["Q4"], ws["R4"] = "Job Number", "Job Name (from Xero)", "CC"
    dept_map = {"ONS": "ONSITE", "PRD": "PRODUCTION", "VID": "VIDEO",
                "INT": "INTEGRATION", "CONS": "CONSULTING", "CTS": "CTS", "": ""}
    for i, j in enumerate(sorted(JOBS, key=lambda x: str(x["job"]))):
        for col, val in ((16, str(j["job"])), (17, j["name"]),
                         (18, dept_map.get(j["dept"], j["dept"]))):
            c = ws.cell(5 + i, col, val)
            c.number_format, c.border = TXT, BOX
    for col, width in (("P", 16), ("Q", 48), ("R", 14)):
        ws.column_dimensions[col].width = width
    # generous: used by COUNTIF and MATCH, so trailing blanks are harmless
    for name, col in (("lst_Jobs", "P"), ("lst_JobName", "Q"), ("lst_JobCC", "R")):
        wb.defined_names.add(DefinedName(name, attr_text=f"Lists!${col}$5:${col}$1500"))

    for col in "ABCDEFGHIKLNPQRT":
        c = ws[f"{col}4"]
        c.font = Font(bold=True, color="FFFFFF", size=10)
        c.fill = PatternFill("solid", fgColor=SLATE)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BOX
    ws.row_dimensions[4].height = 30
    ws.freeze_panes = "A5"


README = [
 ("T", "FY27 Revenue Tracker"),
 ("S", "Corporate Technology Services Pty Ltd  -  replaces FY27_Revenue_Tracker_v2.xlsm"),
 ("B", ""),
 ("H", "WHO TYPES WHERE"),
 ("P", "Finance types on the Finance sheet. Nowhere else. Every invoice line, every month."),
 ("P", "Each department types on its own sheet. Nowhere else. One row per job number."),
 ("P", "The two meet on Job Number, which is a real Xero tracking category and is checked "
       "against the live job list as you type. Nothing has to be kept lined up by hand."),
 ("B", ""),
 ("H", "HOW IT FITS TOGETHER"),
 ("R", "Finance",
       "The invoice register, and the only source of revenue. One row per invoice LINE, "
       "coded to one cost centre, exactly as Xero holds it. Filter the Issue column to see "
       "everything that needs fixing."),
 ("R", "Onsite / Production / Consulting",
       "One row per job. Client, invoice count and revenue come from Finance automatically; "
       "the department fills in its own costs, hours and references."),
 ("R", "WIP Movements",
       "Manual journals between the P&L and GL 11300 Work in Progress. Anything with a "
       "defer start and end on Finance does not belong here - it is handled automatically."),
 ("R", "Deferred Revenue",
       "The workings behind the automatic deferrals. Nothing is typed here; it mirrors the "
       "Finance lines that carry a defer start and end."),
 ("R", "Month-End",
       "The close: revenue by cost centre against Xero, the WIP reconciliation, fifteen "
       "data checks and a sign-off box."),
 ("R", "WIP Summary",
       "WIP opening / movement / closing for every job, to tie job by job back to the WIP "
       "Schedule file."),
 ("R", "Lists",
       "Every dropdown. Cost centres, job numbers and revenue GL codes came straight out of "
       "your Xero organisation, so they already match."),
 ("B", ""),
 ("H", "DEFERRED REVENUE  -  it releases itself"),
 ("P", "Put a Defer Start and a Defer End on a Finance line and the revenue spreads evenly "
       "across those months instead of landing all in the month you invoiced. Leave both "
       "blank for an ordinary one-off invoice."),
 ("P", "   In the month invoiced: one month is recognised, the rest goes to WIP."),
 ("P", "   Every month after: one more month is released, automatically."),
 ("P", "Worked example - $12,000 invoiced 15 Sep 2025 for Sep 2025 to Aug 2026. September "
       "recognises $1,000 and defers $11,000. Each month after releases $1,000. By August "
       "2026 the balance is nil. Nothing is typed each month and no journal is needed."),
 ("P", "The Deferred Revenue sheet shows the workings line by line, and Month-End and WIP "
       "Summary already include it. One warning: a deferred line must not ALSO have a "
       "manual row on WIP Movements or it counts twice - Month-End checks for that."),
 ("B", ""),
 ("H", "VIDEO REVENUE"),
 ("P", "There is no video split column any more, and nothing to keep honest. In Xero a "
       "cost centre sits on the invoice LINE, so a job invoiced part production and part "
       "video is simply two lines on the Finance sheet - one coded PRODUCTION, one coded "
       "VIDEO. The VIDEO figure at month-end is then a plain sum of the lines coded VIDEO, "
       "and it ties to the Xero VIDEO cost centre without any adjustment."),
 ("P", "The 26 migrated invoices that carried a video amount were split into two lines "
       "each on the way across. Ex GST in total is unchanged."),
 ("B", ""),
 ("H", "WHY THIS VERSION IS NOT SLOW, AND WHY IT OPENS"),
 ("P", "The v2 file was slow because of its macro: every keystroke copied 26 columns by 200 "
       "rows into hidden storage, and it was capped at 200 rows per department. There are "
       "no macros here at all, and it is a .xlsx, so no macro warning and nothing blocked "
       "on SharePoint."),
 ("P", "There are also no dynamic array formulas - no VSTACK, FILTER, SORT, UNIQUE or "
       "XLOOKUP. Every sheet is a real Excel Table or a plain SUMIFS grid, which means the "
       "filter buttons, sorting and PivotTables all work normally, and every formula runs "
       "in any version of Excel."),
 ("B", ""),
 ("H", "MONTH-END, IN ORDER"),
 ("P", "1.  Chase the departments until every invoice for the month is on the Finance sheet "
       "and Posted to Xero is Y."),
 ("P", "2.  Open Month-End and set the month."),
 ("P", "3.  Work section 3 first. Every count must be zero. Fix on the Finance sheet, or "
       "filter the Issue column there, which lists the same problems row by row."),
 ("P", "4.  Section 1: run the Xero P&L for the month by Cost Centre and type the revenue "
       "into the yellow column. Each line should read Reconciled."),
 ("P", "5.  Section 2: type the WIP Schedule closing balance and the Xero GL 11300 balance."),
 ("P", "6.  Breaks in section 2 - open WIP Summary, paste the WIP Schedule balances into "
       "column G and the variance column shows which job is out."),
 ("P", "7.  Sign off in section 4."),
 ("B", ""),
 ("H", "THE WIP SIGN RULE"),
 ("P", "One signed Amount column does both jobs, because GL 11300 nets accrued and deferred "
       "revenue. Positive = revenue recognised this month (work done but not yet invoiced, "
       "or a deferral released). Negative = revenue pushed out of this month (invoiced in "
       "advance). So revenue recognised = invoiced Ex GST + WIP movement, and the running "
       "total of the Amount column is the GL 11300 balance."),
 ("B", ""),
 ("H", "LOCKED CELLS, LINKS AND TOTALS"),
 ("P", "Every sheet is protected with no password. Only the cells you are meant to fill are "
       "open; the grey calculated ones are locked so a stray keystroke cannot wipe a "
       "formula. To turn it off: Review, Unprotect Sheet. Sorting and filtering still work."),
 ("P", "Qwilr quotes and Current RMS numbers are clickable. Paste a full web address and it "
       "is used as it stands; type a bare number and it is added to the base address on the "
       "Lists sheet. The Qwilr base is already set. Paste your Current RMS address into "
       "Lists column U, row 6, and the Production links switch on."),
 ("P", "Each sheet shows its total ex-GST at the top, and it follows the filter - filter to "
       "one client or one month and the total follows. Finance shows Ex GST, GST and Inc GST."),
 ("B", ""),
 ("H", "RULES THAT KEEP IT FAST"),
 ("P", "   Never insert or delete rows above the header row."),
 ("P", "   Add rows by typing in the next blank row of the table, or press Tab at the end "
       "of the last row and the table grows with the formulas."),
 ("P", "   Paste values only, never whole columns."),
 ("P", "   Leave calculation on Automatic. Nothing here is volatile - no OFFSET, no "
       "INDIRECT, no TODAY in bulk."),
 ("P", "   Keep it as .xlsx. The moment someone saves it as .xlsm and adds a macro you are "
       "back where you started."),
 ("B", ""),
 ("H", "WHAT CAME ACROSS FROM v2"),
 ("P", "   Finance: 175 invoice lines from the 149 v2 records, the difference being the 26 "
       "video splits now carried as their own line."),
 ("P", "   Department sheets: one row per job - Onsite 28, Production 58, Consulting 28."),
 ("P", "   WIP Movements: 126 rows."),
 ("P", "   Totals tie to the cent: Onsite 512,023.81, Production 460,883.46, Consulting "
       "857,950.27, WIP 17,537.91."),
 ("P", "   Revenue GL codes were assigned by rule (Onsite 41100, Production 42100, Video "
       "42150, Consulting 42800, Integration 42300). Spot-check them."),
 ("P", "   The 126 WIP rows were given a Type from the sign of the amount. Confirm before "
       "relying on them."),
 ("B", ""),
 ("W", "STILL TO CONFIRM"),
 ("P", "Discounts. Checked against Xero: for all 17 August production invoices carrying a "
       "discount, the ex-GST total on the Xero invoice equals Ex GST here exactly, and not "
       "one matched Ex GST minus the discount. So the discount comes off before the invoice "
       "is raised and Ex GST is what reconciles. The old \"Net Total\" is gone; Discounts "
       "Given now sits on the Production sheet at job level with Value Before Discount "
       "beside it. ASSUMPTION: it records what was given away off standard rates."),
 ("P", "23 migrated records have no invoice number and 21 have no date - $93,346.99 of "
       "revenue that belongs to no month. They are on the Finance sheet with the Issue "
       "column filled in. Filter that column and work the list to nothing."),
]


def build_readme(wb):
    ws = wb.create_sheet("Read Me")
    ws.sheet_view.showGridLines = False
    for col, width in (("A", 3), ("B", 30), ("C", 106)):
        ws.column_dimensions[col].width = width
    r = 2
    for item in README:
        kind, body = item[0], item[1]
        if kind == "T":
            ws.cell(r, 2, body).font = Font(size=20, bold=True, color=NAVY)
            ws.row_dimensions[r].height = 28
        elif kind == "S":
            ws.cell(r, 2, body).font = Font(size=10, italic=True, color="595959")
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
            ws.row_dimensions[r].height = 15 * max(1, (len(item[2]) // 103) + 1)
        else:
            c = ws.cell(r, 3, body)
            c.font = Font(size=10)
            c.alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[r].height = 14 * max(1, (len(body) // 106) + 1)
        r += 1


def main():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    build_lists(wb)
    nlines = build_finance(wb)
    for d in DEPTS:
        build_dept(wb, d)
    build_wip(wb)
    dlast = build_deferred(wb, nlines)
    build_month_end(wb, dlast)
    build_wip_summary(wb, dlast)
    build_readme(wb)

    colours = {"Read Me": "7F7F7F", "Finance": NAVY, "Month-End": "2E6B4F",
               "WIP Summary": "2E6B4F", "Deferred Revenue": "8B6A2B",
               "Onsite": SLATE, "Production": SLATE,
               "Consulting": SLATE, "WIP Movements": "8B6A2B", "Lists": "A6A6A6"}
    order = ["Read Me", "Finance", "Month-End", "WIP Summary", "Deferred Revenue",
             "Onsite", "Production", "Consulting", "WIP Movements", "Lists"]
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
