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
import os
import re
from pathlib import Path

import openpyxl
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import (Alignment, Border, Font, PatternFill,
                             Protection, Side)
from openpyxl.utils import get_column_letter as gcl
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.workbook.protection import WorkbookProtection
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableFormula, TableStyleInfo

ROOT = Path("/home/user/Claude/reporting")
JOBS = json.loads((ROOT / "data/revenue_jobs.json").read_text())
MIG = json.loads((ROOT / "data/revenue_migration.json").read_text())
ACC = json.loads((ROOT / "data/revenue_accounts.json").read_text())
# BLANK=1 builds the identical workbook with no transactions in it - every
# formula, dropdown, lock and check intact, nothing to delete before testing.
# The Xero job list, cost centres and GL codes stay: they are reference data,
# not transactions, and the dropdowns and job checks need them.
BLANK = os.environ.get("BLANK") == "1"
OUT = ROOT / ("FY27_Revenue_Tracker_v3_BLANK.xlsx" if BLANK
              else "FY27_Revenue_Tracker_v3.xlsx")

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
SHEET_PASSWORD = "CTS1234"       # Review, Unprotect Sheet - then edit anything

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


A1_MODE = True           # A1 addresses: the form proven to open in Excel


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
    # Columns above are the invoice LINE, which is what Xero holds and what the
    # cost centres reconcile on. Columns below are the whole JOB that line
    # belongs to, so a job split over several lines reads as one job. They
    # repeat on each of its lines - never add them up.
    ("Job Name (Xero)", 34, TXT, "f"), ("Lines on this Job", 12, INT, "f"),
    ("Job Total Ex GST", 15, CUR, "f"), ("Job Invoiced", 14, CUR, "f"),
    ("Job To Invoice", 14, CUR, "f"), ("Job Quote / Ref", 34, TXT, "f"),
    ("Open Quote", 13, TXT, "f"), ("Job Expense", 14, CUR, "f"),
    ("Job Margin", 14, CUR, "f"), ("Job Margin %", 12, PCT, "f"),
    ("Dept Notes", 34, TXT, "f"),
]


def _dept_pick(col_by_team, default='""'):
    """Read a job-level column from whichever department sheet owns the job."""
    out = default
    for team in reversed(DEPTS):
        col = col_by_team.get(team)
        if col is None:
            continue
        if isinstance(col, tuple):
            # Try the first column, fall back to the second when it is blank.
            first, second = (
                f'INDEX(tbl_{team}[{c}],MATCH({{Job Number}}&"",'
                f'tbl_{team}[Job Number],0))' for c in col)
            pick = f'IF({first}="",{second},{first})'
        else:
            pick = (f'INDEX(tbl_{team}[{col}],'
                    f'MATCH({{Job Number}}&"",tbl_{team}[Job Number],0))')
        out = (f'IF({{Team}}="{team}",IFERROR({pick},{default}),{out})')
    return out
FIN_FORMULAS = {
    "Month": '=IF({Date}="","",EOMONTH({Date},0))',
    "Job in Xero?": '=IF({Job Number}="","",IF(COUNTIF(lst_Jobs,{Job Number}&"")>0,"OK","CHECK"))',
    "GST": '=IF({Ex GST}="","",IF({Tax Code}="GST 10%",'
           'ROUND({Ex GST}*set_GSTRate,2),0))',
    "Inc GST": '=IF({Ex GST}="","",{Ex GST}+{GST})',
    # One filterable column instead of a separate exceptions sheet: turn on the
    # filter for Issue and the list of what needs fixing is right here.
    "Job Name (Xero)": '=IF({Job Number}="","",IFERROR(INDEX(lst_JobName,'
                       'MATCH({Job Number}&"",lst_Jobs,0)),""))',
    "Lines on this Job": '=IF({Job Number}="","",COUNTIFS(' + FIN + '[Job Number],'
                         '{Job Number}&"",' + FIN + '[Ex GST],"<>"))',
    "Job Total Ex GST": '=IF({Job Number}="","",SUMIFS(' + FIN + '[Ex GST],'
                        + FIN + '[Job Number],{Job Number}&""))',
    "Job Invoiced": '=IF({Job Number}="","",SUMIFS(' + FIN + '[Ex GST],'
                    + FIN + '[Job Number],{Job Number}&"",'
                    + FIN + '[Xero Invoice No],"<>"))',
    "Job To Invoice": '=IF({Job Number}="","",{Job Total Ex GST}-{Job Invoiced})',
    "Job Quote / Ref": None,        # filled in below - needs _dept_pick
    "Open Quote": None,
    "Job Expense": None,
    "Job Margin": '=IF({Job Number}="","",{Job Total Ex GST}-N({Job Expense}))',
    "Job Margin %": '=IFERROR({Job Margin}/{Job Total Ex GST},"")',
    "Dept Notes": None,
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
    ("Lines on Job", 11, INT, "f"), ("Revenue Ex GST", 16, CUR, "f"),
    ("Invoiced", 14, CUR, "f"), ("To Invoice", 14, CUR, "f"),
    ("Cost Centres", 16, TXT, "f"),
]
JOB_CORE_FORMULAS = {
    "Invoice Date": '=IF({Job Number}="","",IFERROR(IF(SUMPRODUCT(MAX((tbl_Finance[Job Number]={Job Number}&"")*(tbl_Finance[Xero Invoice No]<>"")*tbl_Finance[Date]))=0,"",'
                    'SUMPRODUCT(MAX((tbl_Finance[Job Number]={Job Number}&"")*(tbl_Finance[Xero Invoice No]<>"")*tbl_Finance[Date]))),""))',
    # The most recent invoice number on the job. If the job carries more than
    # one, say so rather than quietly showing just the last.
    "Invoice No": '=IF({Job Number}="","",IF(COUNTIFS(tbl_Finance[Job Number],{Job Number}&"",tbl_Finance[Xero Invoice No],"<>")=0,"",'
                  'IFERROR(IF(COUNTIFS(tbl_Finance[Job Number],{Job Number}&"",tbl_Finance[Xero Invoice No],LOOKUP(2,1/((tbl_Finance[Job Number]={Job Number}&"")*(tbl_Finance[Xero Invoice No]<>"")),tbl_Finance[Xero Invoice No]))=COUNTIFS(tbl_Finance[Job Number],{Job Number}&"",tbl_Finance[Xero Invoice No],"<>"),LOOKUP(2,1/((tbl_Finance[Job Number]={Job Number}&"")*(tbl_Finance[Xero Invoice No]<>"")),tbl_Finance[Xero Invoice No]),'
                  'LOOKUP(2,1/((tbl_Finance[Job Number]={Job Number}&"")*(tbl_Finance[Xero Invoice No]<>"")),tbl_Finance[Xero Invoice No])&" +more"),"")))',
    "Job in Xero?": '=IF({Job Number}="","",IF(COUNTIF(lst_Jobs,{Job Number}&"")>0,"OK","CHECK"))',
    "Job Name (Xero)": '=IF({Job Number}="","",IFERROR(INDEX(lst_JobName,'
                       'MATCH({Job Number}&"",lst_Jobs,0)),""))',
    "Client": '=IF({Job Number}="","",IFERROR(INDEX(' + FIN + '[Client],'
              'MATCH({Job Number}&"",' + FIN + '[Job Number],0)),""))',
    "Lines on Job": '=IF({Job Number}="","",COUNTIFS(' + FIN + '[Job Number],'
                    '{Job Number}&"",' + FIN + '[Ex GST],"<>"))',
    # A job's revenue is not all invoiced yet, so say which part is which
    "Invoiced": '=IF({Job Number}="","",SUMIFS(' + FIN + '[Ex GST],'
                + FIN + '[Job Number],{Job Number}&"",'
                + FIN + '[Xero Invoice No],"<>"))',
    "To Invoice": '=IF({Job Number}="","",{Revenue Ex GST}-{Invoiced})',
    "Revenue Ex GST": '=IF({Job Number}="","",SUMIFS(' + FIN + '[Ex GST],'
                      + FIN + '[Job Number],{Job Number}&""))',
    # A job can be invoiced across several cost centres - list every one
    # of them, not just the first line's.
    "Cost Centres": '=IF({Job Number}="","",IF({Lines on Job}=0,"not invoiced yet",'
                    'SUBSTITUTE(TRIM(IF(COUNTIFS(tbl_Finance[Job Number],{Job Number}&"",tbl_Finance[Cost Centre],"ONSITE")>0,"ONSITE ","")&IF(COUNTIFS(tbl_Finance[Job Number],{Job Number}&"",tbl_Finance[Cost Centre],"PRODUCTION")>0,"PRODUCTION ","")&IF(COUNTIFS(tbl_Finance[Job Number],{Job Number}&"",tbl_Finance[Cost Centre],"VIDEO")>0,"VIDEO ","")&IF(COUNTIFS(tbl_Finance[Job Number],{Job Number}&"",tbl_Finance[Cost Centre],"INTEGRATION")>0,"INTEGRATION ","")&IF(COUNTIFS(tbl_Finance[Job Number],{Job Number}&"",tbl_Finance[Cost Centre],"CONSULTING")>0,"CONSULTING ","")&IF(COUNTIFS(tbl_Finance[Job Number],{Job Number}&"",tbl_Finance[Cost Centre],"CTS")>0,"CTS ",""))," ",", ")))',
}

DEPT_EXTRA = {
    # A6: Ariba Status removed. A10: Onsite gains a Qwilr quote, hyperlinked.
    "Onsite": [("PO / Reference", 18, TXT, "in"), ("Qwilr Quote", 40, TXT, "in"),
               ("Open Qwilr", 14, TXT, "f"), ("Billable Hours", 13, NUM, "in"),
               ("Approved By", 16, TXT, "in"), ("Notes", 38, TXT, "in")],
    # A3: Zoho is the job number, so it is a formula now, not something to type.
    # A5: the Current RMS number is hyperlinked. Production carries a Qwilr
    # quote as well, linked exactly the way Consulting's is.
    "Production": [
        ("Total Inc GST", 14, CUR, "f"),
        ("Company", 18, TXT, "in"), ("Event Name", 34, TXT, "f"),
        ("Client Email", 30, TXT, "in"), ("Event Grouping", 20, TXT, "in"),
        ("Event Date", 12, DATE, "in"), ("Current RMS No", 14, TXT, "in"),
        ("Open in Current RMS", 18, TXT, "f"),
        ("Qwilr Quote", 40, TXT, "in"), ("Open Qwilr", 14, TXT, "f"),
        ("Zoho Number", 14, TXT, "f"),
        ("Job Closed", 11, TXT, "dv"),
        ("Discounts Given", 16, CUR, "in"), ("Value Before Discount", 20, CUR, "f"),
        ("Discount %", 11, PCT, "f"), ("Cross Hire Expense", 17, CUR, "in"),
        # v2 had these two on screen but its macro never saved them, so they
        # come across empty - and v2's margin never counted them either.
        ("Conf. Call Costs", 15, CUR, "in"), ("Transcription Costs", 18, CUR, "in"),
        ("Labour Expense (Internal)", 22, CUR, "in"),
        # The video slice of the job, straight off the Finance lines coded
        # VIDEO. Revenue, not a cost - it stays out of Total Expense.
        ("Video Revenue", 14, CUR, "f"), ("Video % of Revenue", 16, PCT, "f"),
        ("Total Expense", 14, CUR, "f"),
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
RMS = ('=IF({Current RMS No}="","",IF(AND(LEFT({Current RMS No},4)<>"http",'
       'set_RMSBase=""),"",HYPERLINK(IF(LEFT({Current RMS No},4)="http",'
       '{Current RMS No},set_RMSBase&{Current RMS No}),"Open opportunity")))')

DEPT_EXTRA_FORMULAS = {
    "Onsite": {"Open Qwilr": QWILR},
    "Production": {
        "Open in Current RMS": RMS,
        "Open Qwilr": QWILR,
        "Video Revenue": '=IF({Job Number}="","",SUMIFS(' + FIN + '[Ex GST],'
                         + FIN + '[Job Number],{Job Number}&"",'
                         + FIN + '[Cost Centre],"VIDEO"))',
        "Zoho Number": '=IF({Job Number}="","",{Job Number})',
        "Value Before Discount": '=IF({Job Number}="","",{Revenue Ex GST}+N({Discounts Given}))',
        "Discount %": '=IFERROR({Discounts Given}/{Value Before Discount},"")',
        "Total Expense": '=IF({Job Number}="","",N({Cross Hire Expense})'
                         '+N({Conf. Call Costs})+N({Transcription Costs})'
                         '+N({Labour Expense (Internal)}))',
        "Video % of Revenue": '=IFERROR({Video Revenue}/{Revenue Ex GST},"")',
        "Total Inc GST": '=IF({Job Number}="","",SUMIFS(' + FIN + '[Inc GST],'
                         + FIN + '[Job Number],{Job Number}&""))',
        # The crew's own name for the event, as Finance typed it on the invoice.
        "Event Name": '=IF({Job Number}="","",IFERROR(INDEX(' + FIN + '[Description],'
                      'MATCH({Job Number}&"",' + FIN + '[Job Number],0)),""))',
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
    ("Type", 30, TXT, "dv"),
    # A19: which side of the P&L this belongs to, and the dates if it spreads.
    ("Revenue or Cost", 15, TXT, "dv"),
    ("Defer Start", 12, DATE, "in"), ("Defer End", 12, DATE, "in"),
    ("Spread", 30, TXT, "f"),
    ("Amount", 15, CUR, "in"), ("P&L Account", 30, TXT, "in"),
    # A17: say which account to debit, which to credit and for how much, so
    # the journal can be keyed straight into Xero. GL Code is gone - it said
    # 11300 on every row and the two columns below replace it.
    ("Debit Account", 30, TXT, "f"), ("Credit Account", 30, TXT, "f"),
    ("Journal Amount", 14, CUR, "f"), ("Journal Narration", 52, TXT, "f"),
    ("Journal Ref", 14, TXT, "in"), ("Posted to Xero", 14, TXT, "dv"),
    ("Notes", 40, TXT, "in"),
]
WIP_FORMULAS = {
    # A positive amount puts revenue back into the month, so WIP goes up:
    # debit the balance sheet, credit the P&L. A negative amount reverses it.
    # A revenue row credits 44200. A cost row credits the expense account the
    # bill came from, which is typed in P&L Account - Xero puts both in 11300
    # but they are not the same thing.
    "Debit Account": '=IF(N({Amount})=0,"",IF({Amount}>0,set_WIPAsset,IF({Revenue or Cost}="Cost",IF({P&L Account}="","(set the P&L Account)",{P&L Account}),IF({P&L Account}="",set_WIPIncome,{P&L Account}))))',
    "Credit Account": '=IF(N({Amount})=0,"",IF({Amount}>0,IF({Revenue or Cost}="Cost",IF({P&L Account}="","(set the P&L Account)",{P&L Account}),IF({P&L Account}="",set_WIPIncome,{P&L Account})),set_WIPAsset))',
    # A row that spreads is posted monthly, not in one hit, so its journal
    # amount comes off the Deferred Revenue sheet and the Month-End journal.
    "Journal Amount": '=IF(OR(N({Amount})=0,{Defer Start}<>""),"",ABS({Amount}))',
    "Spread": '=IF({Amount}="","",IF(OR({Defer Start}="",{Defer End}=""),'
              '"one month",TEXT(MAX(1,(YEAR({Defer End})-YEAR({Defer Start}))*12'
              '+MONTH({Defer End})-MONTH({Defer Start})+1),"0")'
              '&" months - see Deferred Revenue"))',
    "Journal Narration": '=IF(N({Amount})=0,"",IF({Defer Start}<>"",'
                         '"spreads monthly - see Deferred Revenue","WIP "&{Type}'
                         '&IF({Job Number}="",""," - "'
                         '&{Job Number})&IF({Client}="",""," "&{Client})'
                         '&IF({Description}="",""," - "&{Description})))',
    "Job in Xero?": '=IF({Job Number}="","",IF(COUNTIF(lst_Jobs,{Job Number}&"")>0,"OK","CHECK"))',
    # A15: the invoice date follows the invoice number in from Finance
    "Invoice Date": '=IF({Xero Invoice No}="","",IFERROR(INDEX(' + FIN + '[Date],'
                    'MATCH({Xero Invoice No}&"",' + FIN + '[Xero Invoice No],0)),""))',
}

FIN_FORMULAS["Job Quote / Ref"] = '=IF({Job Number}="","",' + _dept_pick(
    {"Onsite": "Qwilr Quote", "Consulting": "Qwilr Quote",
     "Production": ("Current RMS No", "Qwilr Quote")}) + ')'
FIN_FORMULAS["Open Quote"] = (
    '=IF({Job Quote / Ref}="","",IF(AND(LEFT({Job Quote / Ref},4)<>"http",'
    'IF({Team}="Production",set_RMSBase,set_QwilrBase)=""),"",'
    'HYPERLINK(IF(LEFT({Job Quote / Ref},4)="http",'
    '{Job Quote / Ref},IF({Team}="Production",set_RMSBase,set_QwilrBase)'
    '&{Job Quote / Ref}),"Open")))')
FIN_FORMULAS["Job Expense"] = '=IF({Job Number}="","",' + _dept_pick(
    {"Production": "Total Expense", "Consulting": "Total Expense"}, default="0") + ')'
FIN_FORMULAS["Dept Notes"] = '=IF({Job Number}="","",' + _dept_pick(
    {"Onsite": "Notes", "Production": "Notes", "Consulting": "Notes"}) + ')'

DV_FOR = {"Revenue or Cost": "lst_RevCost", "Source": "lst_WonSource",
          "Cost Centre": "lst_CostCentre", "Invoice Type": "lst_InvoiceType",
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
# v2 wrote the department on a WIP row several ways - "Production", "PRD",
# "VID " with a trailing space. Match on the trimmed upper-case value so every
# row lands on a cost centre; 9 August rows worth $58.73 net came across blank.
CC_WIP = {"PRODUCTION": "PRODUCTION", "PRD": "PRODUCTION",
          "VIDEO": "VIDEO", "VID": "VIDEO",
          "INTEGRATION": "INTEGRATION", "INT": "INTEGRATION",
          "CONSULTING": "CONSULTING", "CONS": "CONSULTING",
          "SUPPORT": "ONSITE", "ONSITE": "ONSITE", "ONS": "ONSITE", "CTS": "CTS"}


def _idx(block):
    return {k: i for i, k in enumerate(MIG[block]["headers"])}


def migrate_finance():
    if BLANK:
        return []
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
        # The residual is whatever is not video, so it is never VIDEO itself.
        # A job the crew tagged Video can still carry production labour: Xero
        # codes $380.00 of INV-10600 to 42100 and v2's own Video Total agreed.
        rest = "PRODUCTION" if cc == "VIDEO" else cc
        if video and ex is not None and 0 < abs(video) < abs(ex):
            # split the line the way Xero codes it: production part, video part
            out.append(dict(base, **{"Cost Centre": rest, "Ex GST": round(ex - video, 2),
                                     "Revenue GL": "42100",
                                     "Notes": "Production portion of this invoice"}))
            out.append(dict(base, **{"Cost Centre": "VIDEO", "Ex GST": video,
                                     "Revenue GL": "42150",
                                     "Notes": "Video portion of this invoice"}))
        elif video and ex is not None and abs(video) == abs(ex):
            out.append(dict(base, **{"Cost Centre": "VIDEO", "Ex GST": ex,
                                     "Revenue GL": "42150"}))
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
    if BLANK:
        return []
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
             "Zoho Number": "Zoho Number", "Closed": "Job Closed",
             "Client Email": "Client Email"},
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
    if BLANK:
        return []
    out = []
    for dt, job, dept, client, desc, amt in MIG["WIP"]["rows"]:
        amount = _num(amt)
        out.append({"Month": _month_end(_date(dt)), "Job Number": _txt(job),
                    "Client": _txt(client),
                    "Cost Centre": CC_WIP.get(str(dept or "").strip().upper()),
                    "Description": _txt(desc),
                    "Type": ("Accrual - unbilled work" if (amount or 0) >= 0
                             else "Deferral - invoiced in advance"),
                    "Amount": amount, "Revenue or Cost": "Revenue", "Posted to Xero": "Y",
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
    # Deliberately no calculatedColumnFormula: the version that opened cleanly in
    # Excel had none, and every row including the spare ones already carries the
    # formula, so there is nothing to propagate until the table is extended.

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

    # A4 / A20: protected with SHEET_PASSWORD. It stops a stray keystroke landing
    # in a formula; Review, Unprotect Sheet and the password turns it off.
    ws.protection.sheet = True
    ws.protection.password = SHEET_PASSWORD
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
                "Progress. SIGN RULE: + = the P&L gives up the amount this month and "
                "the 11300 balance goes up. - = the other way round. Set Revenue or "
                "Cost on every row - 11300 nets both and only this column tells them "
                "apart. Put a Defer Start and End on it and it spreads by month.")
    nrows = build_table(ws, WIP_COLS, WIP_FORMULAS, migrate_wip(), "tbl_WIP",
                        spare=200)[1]
    totals_strip(ws, WIP_COLS, "tbl_WIP", ["Amount"])
    ws.freeze_panes = f"D{DATA_ROW}"
    return nrows


# A19: what was won, from ZOHO, Current RMS and Qwilr, as a sense check against
# what was actually invoiced. Paste the month's export in and Month-End compares
# the two. Nothing here feeds revenue - it is a cross-check, not a source.
WON_COLS = [
    ("Month", 10, MON, "in"), ("Source", 15, TXT, "dv"),
    ("Reference", 20, TXT, "in"), ("Client", 26, TXT, "in"),
    ("Job Number", 15, TXT, "in"), ("Job in Xero?", 12, TXT, "f"),
    ("Cost Centre", 14, TXT, "dv"), ("Description", 44, TXT, "in"),
    ("Value Ex GST", 15, CUR, "in"), ("Status", 12, TXT, "dv"),
    ("Date Won", 12, DATE, "in"), ("Expected Invoice Month", 20, MON, "in"),
    ("On the Finance Sheet", 19, CUR, "f"), ("Still to Invoice", 15, CUR, "f"),
    ("Notes", 38, TXT, "in"),
]
WON_FORMULAS = {
    "Job in Xero?": '=IF({Job Number}="","",IF(COUNTIF(lst_Jobs,{Job Number}&"")>0,'
                    '"OK","CHECK"))',
    # Everything Finance holds against that job, invoiced or not. Subtracting
    # all of it is what stops the forecast counting the same work twice.
    "On the Finance Sheet": '=IF({Job Number}="","",SUMIFS(' + FIN + '[Ex GST],'
                            + FIN + '[Job Number],{Job Number}&""))',
    # No job number yet means nothing has been raised against it at all
    "Still to Invoice": '=IF({Status}<>"Won","",IF({Job Number}="",N({Value Ex GST}),'
                        '{Value Ex GST}-{On the Finance Sheet}))',
}


def build_won(wb):
    ws = wb.create_sheet("Work Won")
    title_block(ws, "Work Won  -  ZOHO, Current RMS and Qwilr",
                "Paste the month's won opportunities here, one row each, and Month-End "
                "compares the total against what was invoiced. This is a sense check "
                "only - nothing on this sheet feeds revenue. Put the job number on a "
                "row and it will tell you what has been invoiced against it.")
    build_table(ws, WON_COLS, WON_FORMULAS, [], "tbl_Won", spare=600,
                dv_override={"Status": "lst_WonStatus"})
    totals_strip(ws, WON_COLS, "tbl_Won", ["Value Ex GST", "On the Finance Sheet",
                                           "Still to Invoice"])
    ws.freeze_panes = f"C{DATA_ROW}"


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
     "+".join(f'COUNTIFS({DEPT_TABLE[d]}[Job Number],"<>",{DEPT_TABLE[d]}[Lines on Job],0)'
              for d in DEPTS)),
    ("Deferred lines that also have a manual WIP journal (would double count)",
     f'SUMPRODUCT(--({FIN}[Defer Start]<>""),--({FIN}[Xero Invoice No]<>""),'
     f'--(COUNTIF(tbl_WIP[Xero Invoice No],{FIN}[Xero Invoice No]&"")>0))'),
    ("Deferred lines missing an end date",
     f'SUMPRODUCT(--({FIN}[Defer Start]<>""),--({FIN}[Defer End]=""))'
     '+SUMPRODUCT(--(tbl_WIP[Defer Start]<>""),--(tbl_WIP[Defer End]=""))'),
    # 11300 nets revenue against cost, so an unset row lands on the wrong side
    ("WIP rows with an amount but Revenue or Cost not set",
     'SUMPRODUCT(--(tbl_WIP[Amount]<>""),--(tbl_WIP[Revenue or Cost]=""))'),
    ("Deferred cost rows with no P&L Account to credit",
     'COUNTIFS(tbl_WIP[Revenue or Cost],"Cost",tbl_WIP[P&L Account],"",'
     'tbl_WIP[Amount],"<>")'),
    ('Still sitting at "To Invoice" for the selected month',
     f'COUNTIFS({FIN}[Month],$C$4,{FIN}[Status],"To Invoice")'),
    # These carry revenue into the month that Xero has not raised, so section 1
    # will not tie until they are invoiced or moved out. They are the forecast.
    ("Revenue dated this month with no invoice number - it is in the forecast",
     f'COUNTIFS({FIN}[Month],$C$4,{FIN}[Xero Invoice No],"",{FIN}[Ex GST],"<>")'),
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
        # Revenue only. A deferred cost sits in the same GL but it is not
        # revenue, and a row that spreads is counted on the Deferred sheet.
        ws.cell(r, 5).value = (
            '=SUMIFS(tbl_WIP[Amount],tbl_WIP[Month],$C$4,'
            f"tbl_WIP[Cost Centre],$A{r},"
            'tbl_WIP[Revenue or Cost],"<>Cost",tbl_WIP[Defer Start],"")'
            f'+SUMIFS({defer_range("Q", dlast)},{defer_range("F", dlast)},$A{r},'
            f'{defer_range("T", dlast)},"<>Cost")')
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
    col_heads(ws, s2 + 1, ["", "Total", "Deferred / accrued revenue",
                           "Deferred / accrued cost", "", "", "", "", "Check"])
    b = s2 + 2

    def wip_side(when, dcol, side):
        """One side of GL 11300 - revenue or cost - for a point in time."""
        crit = '"Cost"' if side == "cost" else '"<>Cost"'
        return (f'=SUMIFS(tbl_WIP[Amount],{when},'
                f'tbl_WIP[Revenue or Cost],{crit},tbl_WIP[Defer Start],"")'
                f'+SUMIFS({defer_range(dcol, dlast)},'
                f'{defer_range("T", dlast)},{crit})')

    OPEN_W = 'tbl_WIP[Month],"<>",tbl_WIP[Month],"<"&$C$4'
    MOVE_W = "tbl_WIP[Month],$C$4"
    rows = [
        ("Opening WIP balance (all months before this one)",
         (wip_side(OPEN_W, "P", "rev"), wip_side(OPEN_W, "P", "cost")), "calc"),
        ("Movement this month",
         (wip_side(MOVE_W, "Q", "rev"), wip_side(MOVE_W, "Q", "cost")), "calc"),
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
            # C is revenue, D is cost, B adds them back to the GL balance
            for off, f in enumerate(formula):
                sc = ws.cell(r, 3 + off)
                sc.value, sc.fill = f, CALC_FILL
                sc.number_format, sc.border = CUR, BOX
                sc.font = Font(size=10)
            vc.value, vc.fill = f"=C{r}+D{r}", CALC_FILL
            vc.font = Font(bold=True, size=10)
        elif kind == "sum":
            vc.value = f"=B{b}+B{b + 1}"
            vc.fill = PatternFill("solid", fgColor=LIGHT)
            vc.font = Font(bold=True, size=10, color=NAVY)
            for off in (0, 1):
                sc = ws.cell(r, 3 + off)
                sc.value = f"={gcl(3 + off)}{b}+{gcl(3 + off)}{b + 1}"
                sc.number_format, sc.border = CUR, BOX
                sc.fill = PatternFill("solid", fgColor=LIGHT)
                sc.font = Font(bold=True, size=10, color=NAVY)
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
            "GL 11300 nets deferred revenue against deferred cost, so Xero shows one "
            "figure and cannot split it. The Revenue or Cost column on WIP Movements "
            "splits it here. Only the Total column should be compared to Xero."
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

    # ---- 4. the journal to post. Section 1 already worked out the movement
    # by cost centre, so this reads straight off it and cannot disagree.
    sj = money + 2
    band(ws, sj, "4.  WIP JOURNAL FOR THE MONTH  -  post this one journal in Xero")
    col_heads(ws, sj + 1, ["Account", "Cost Centre", "Debit", "Credit"])
    jr = sj + 2
    whole = ('(SUMIFS(tbl_WIP[Amount],tbl_WIP[Month],$C$4,'
             'tbl_WIP[Defer Start],"",tbl_WIP[Revenue or Cost],"<>Cost")'
             f'+SUMIFS({defer_range("Q", dlast)},'
             f'{defer_range("T", dlast)},"<>Cost"))')
    costmv = ('(SUMIFS(tbl_WIP[Amount],tbl_WIP[Month],$C$4,'
              'tbl_WIP[Defer Start],"",tbl_WIP[Revenue or Cost],"Cost")'
              f'+SUMIFS({defer_range("Q", dlast)},'
              f'{defer_range("T", dlast)},"Cost"))')
    lines = [("=set_WIPAsset", '"(no tracking)"', whole, True)]
    for i, cc in enumerate(COST_CENTRES):
        lines.append(("=set_WIPIncome", f'$A${r0 + i}', f"$E${r0 + i}", False))
    # whatever is left over has no cost centre on it - never let it disappear
    lines.append(("=set_WIPIncome", '"(cost centre not set - fix it)"',
                  f"({whole}-$E${tot})", False))
    for i, (acct, cc, src, is_bs) in enumerate(lines):
        r = jr + i
        ws.cell(r, 1).value = acct
        ws.cell(r, 2).value = "=" + cc
        # the balance sheet takes the debit when WIP goes up; the P&L takes it
        # when WIP goes down. One signed movement, two sides, always equal.
        if is_bs:
            ws.cell(r, 3).value = f'=IF(ROUND({src},2)>0,{src},"")'
            ws.cell(r, 4).value = f'=IF(ROUND({src},2)<0,-{src},"")'
        else:
            ws.cell(r, 3).value = f'=IF(ROUND({src},2)<0,-{src},"")'
            ws.cell(r, 4).value = f'=IF(ROUND({src},2)>0,{src},"")'
    jtot = jr + len(lines)
    ws.cell(jtot, 1, "TOTAL").font = Font(bold=True, size=10)
    for col in (3, 4):
        ws.cell(jtot, col).value = f"=SUM({gcl(col)}{jr}:{gcl(col)}{jtot - 1})"
    ws.cell(jtot, 5).value = (f'=IF(ROUND($C${jtot},2)=ROUND($D${jtot},2),'
                              f'"Balanced","CHECK - does not balance")')
    ws.cell(jtot, 5).font = Font(bold=True, size=10)
    for r in range(jr, jtot + 1):
        for col in range(1, 5):
            cell = ws.cell(r, col)
            cell.border, cell.font = BOX, Font(size=10, bold=(r == jtot))
            cell.fill = CALC_FILL
            if col in (3, 4):
                cell.number_format = CUR
    ws.conditional_formatting.add(f"E{jtot}", FormulaRule(
        formula=[f'LEFT($E${jtot},5)="CHECK"'], fill=PatternFill("solid", fgColor=BAD)))
    # a leftover line with anything on it means a row is missing its cost centre
    ws.conditional_formatting.add(f"A{jtot - 1}:D{jtot - 1}", FormulaRule(
        formula=[f'ROUND(N($C${jtot - 1})+N($D${jtot - 1}),2)<>0'],
        fill=PatternFill("solid", fgColor=BAD)))

    made = jtot + 1
    for i, (label, formula) in enumerate([
            ("Narration to use", '="WIP movement "&TEXT($C$4,"mmmm yyyy")'),
            ("Made up of - manual revenue rows on WIP Movements",
             '=SUMIFS(tbl_WIP[Amount],tbl_WIP[Month],$C$4,tbl_WIP[Defer Start],"",'
             'tbl_WIP[Revenue or Cost],"<>Cost")'),
            ("Made up of - revenue that spreads (Deferred Revenue sheet)",
             f'=SUMIFS({defer_range("Q", dlast)},'
             f'{defer_range("T", dlast)},"<>Cost")'),
            ("Total revenue movement  -  section 2, revenue column",
             "=" + whole),
            ("Deferred / accrued cost movement  -  post these line by line",
             "=" + costmv)]):
        r = made + i
        ws.cell(r, 1, label).font = Font(size=9, italic=True, color="808080")
        c2 = ws.cell(r, 3)
        c2.value, c2.border, c2.fill = formula, BOX, CALC_FILL
        c2.font = Font(size=10, bold=(i in (3, 4)))
        if i:
            c2.number_format = CUR
        else:
            ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
    ws.cell(made + 5, 1, "Nothing here posts itself. The tracker works the movement "
                         "out; you key this journal into Xero and put the reference "
                         "back on the WIP Movements rows. The cost line above is not "
                         "in this journal: every deferred cost credits the expense "
                         "account its bill came from, so post those from WIP Movements "
                         "with Revenue or Cost filtered to Cost.").font = \
        Font(size=9, italic=True, color="808080")

    # ---- 5. what was won against what was invoiced
    sw = made + 7
    band(ws, sw, "5.  WORK WON  -  ZOHO / Current RMS / Qwilr vs what was invoiced")
    col_heads(ws, sw + 1, ["Source", "Opportunities", "Value Ex GST", "", "", "", "",
                           "", "Check"])
    w0 = sw + 2
    for i, src in enumerate(["ZOHO", "Current RMS", "Qwilr", "Other"]):
        r = w0 + i
        ws.cell(r, 1, src).font = Font(bold=True, size=10)
        ws.cell(r, 2).value = ('=COUNTIFS(tbl_Won[Month],$C$4,tbl_Won[Source],$A'
                               f'{r},tbl_Won[Status],"Won")')
        ws.cell(r, 3).value = ('=SUMIFS(tbl_Won[Value Ex GST],tbl_Won[Month],$C$4,'
                               f'tbl_Won[Source],$A{r},tbl_Won[Status],"Won")')
    wtot = w0 + 4
    ws.cell(wtot, 1, "TOTAL WON THIS MONTH").font = Font(bold=True, size=10)
    for col in (2, 3):
        ws.cell(wtot, col).value = f"=SUM({gcl(col)}{w0}:{gcl(col)}{wtot - 1})"
    extra = [("Invoiced this month (Finance, Ex GST)", f"=$B${tot}"),
             ("Won less invoiced", f"=$C${wtot}-$C${wtot + 1}"),
             ("Won this month, still not invoiced against the job",
              '=SUMIFS(tbl_Won[Still to Invoice],tbl_Won[Month],$C$4,'
              'tbl_Won[Status],"Won")'),
             ("Won with a job number Xero does not have",
              '=COUNTIFS(tbl_Won[Month],$C$4,tbl_Won[Status],"Won",'
              'tbl_Won[Job in Xero?],"CHECK")')]
    for i, (label, formula) in enumerate(extra):
        r = wtot + 1 + i
        ws.cell(r, 1, label).font = Font(size=10, bold=(i == 1))
        vc = ws.cell(r, 3)
        vc.value, vc.fill = formula, CALC_FILL
        vc.number_format = "#,##0" if i == 3 else CUR
    for r in range(w0, wtot + 5):
        for col in range(1, 4):
            cell = ws.cell(r, col)
            cell.border = BOX
            if col == 2:
                cell.number_format = "#,##0"
            elif col == 3:
                cell.number_format = CUR
            if col > 1 and not cell.fill.fgColor.rgb.endswith(INPUT_FILL.fgColor.rgb[-6:]):
                cell.fill = CALC_FILL
    ws.cell(wtot + 5, 1,
            "A sense check, not a source. Revenue only ever comes off the Finance "
            "sheet. A big gap either way is worth a look: work won and never "
            "invoiced, or invoiced with no opportunity behind it."
            ).font = Font(size=9, italic=True, color="808080")

    # ---- 6. forecast: everything with no invoice raised against it yet
    sf = wtot + 7
    band(ws, sf, "6.  FORECAST  -  work with no invoice against it")
    col_heads(ws, sf + 1, ["Cost Centre", "Finance: to invoice", "Won, not invoiced",
                           "Total forecast", "", "", "", "", ""])
    f0 = sf + 2
    for i, cc in enumerate(COST_CENTRES):
        r = f0 + i
        ws.cell(r, 1, cc).font = Font(bold=True, size=10)
        # a Finance line with a value but no invoice number is work done or
        # agreed that nobody has billed yet, whatever month it is sitting in
        ws.cell(r, 2).value = (f'=SUMIFS({FIN}[Ex GST],{FIN}[Cost Centre],$A{r},'
                               f'{FIN}[Xero Invoice No],"")')
        ws.cell(r, 3).value = ('=SUMIFS(tbl_Won[Still to Invoice],tbl_Won[Cost Centre],'
                               f'$A{r},tbl_Won[Status],"Won")')
        ws.cell(r, 4).value = f"=B{r}+C{r}"
    ftot = f0 + len(COST_CENTRES)
    ws.cell(ftot, 1, "TOTAL FORECAST").font = Font(bold=True, size=10)
    for col in range(2, 5):
        ws.cell(ftot, col).value = f"=SUM({gcl(col)}{f0}:{gcl(col)}{ftot - 1})"
    for r in range(f0, ftot + 1):
        for col in range(1, 5):
            cell = ws.cell(r, col)
            cell.border, cell.fill = BOX, CALC_FILL
            cell.font = Font(size=10, bold=(r == ftot))
            if col > 1:
                cell.number_format = CUR
    memo = [("Not in the forecast above  -  open opportunities, not yet won",
             '=SUMIFS(tbl_Won[Value Ex GST],tbl_Won[Status],"Open")'),
            ("Not in the forecast above  -  deferred revenue still to release "
             "(already invoiced)",
             f'=-SUMIFS({defer_range("R", dlast)},{defer_range("T", dlast)},"<>Cost")'),
            ("Jobs on a department sheet with no invoice line at all (count)",
             "+".join(f'COUNTIFS({DEPT_TABLE[d]}[Job Number],"<>",'
                      f'{DEPT_TABLE[d]}[Lines on Job],0)' for d in DEPTS))]
    for i, (label, formula) in enumerate(memo):
        r = ftot + 1 + i
        ws.cell(r, 1, label).font = Font(size=10)
        vc = ws.cell(r, 4)
        vc.value = formula if formula.startswith("=") else "=" + formula
        vc.fill, vc.border = CALC_FILL, BOX
        vc.number_format = "#,##0" if i == 2 else CUR
    ws.cell(ftot + 4, 1,
            "Nothing here is revenue yet. Finance: to invoice is a line already on the "
            "Finance sheet with no invoice number on it. Won, not invoiced is the value "
            "won less everything Finance already holds against that job, so the two "
            "columns add up without counting the same work twice."
            ).font = Font(size=9, italic=True, color="808080")

    # ---- 7. sign-off
    s4 = ftot + 6
    band(ws, s4, "7.  SIGN-OFF")
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
    ("Source", 16, TXT, '"Finance invoice"'),
    ("Rev/Cost", 11, TXT, '"Revenue"'),
]

# A19: the same engine, fed from a WIP Movements row instead of a Finance line,
# so a supplier bill spread over twelve months behaves exactly like revenue
# invoiced in advance. Only the first nine columns differ; everything from
# Months onwards reads $B, $G, $H and $I and does not care where they came from.
#
# A cost deferral is the mirror image of a revenue one - money out instead of
# money in - so the amount is fed in negative and every formula below works
# unchanged, leaving a debit balance in 11300 the way a prepayment should.
DEFER_WIP_COLS = [
    ("Line", 7, INT, 'ROW()-{hdr}'),
    ("Date", 11, DATE, 'IFERROR(INDEX(W[Month],$A{r}),"")'),
    ("Xero Invoice No", 15, TXT, 'IFERROR(INDEX(W[Xero Invoice No],$A{r}),"")'),
    ("Client", 22, TXT, 'IFERROR(INDEX(W[Client],$A{r}),"")'),
    ("Job Number", 14, TXT, 'IFERROR(INDEX(W[Job Number],$A{r}),"")'),
    ("Cost Centre", 14, TXT, 'IFERROR(INDEX(W[Cost Centre],$A{r}),"")'),
    ("Ex GST", 14, CUR, 'IFERROR(IF(INDEX(W[Defer Start],$A{r})="","",'
     'IF(INDEX(W[Revenue or Cost],$A{r})="Cost",-1,1)*INDEX(W[Amount],$A{r})),"")'),
    ("Defer Start", 12, DATE, 'IFERROR(INDEX(W[Defer Start],$A{r}),"")'),
    ("Defer End", 12, DATE, 'IFERROR(INDEX(W[Defer End],$A{r}),"")'),
]
DEFER_FIRST = 6          # first data row on the Deferred Revenue sheet
DEFER_WIP_COLS += DEFER_COLS[9:-2] + [
    ("Source", 16, TXT, '"WIP journal"'),
    ("Rev/Cost", 11, TXT, 'IFERROR(INDEX(W[Revenue or Cost],$A{r}),"")'),
]


def build_deferred(wb, nlines, wlines):
    ws = wb.create_sheet("Deferred Revenue")
    title_block(ws, "Deferrals - revenue and cost, released automatically",
                "Every line that carries a Defer Start and a Defer End - Finance "
                "invoices in the top block, WIP Movements rows in the block below it. "
                "The amount spreads evenly across those months and the WIP movement "
                "works itself out. Nothing is typed here.")
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

    def block(first, count, cols, off):
        for k in range(count):
            r = first + k
            for i, (h, w, fmt, f) in enumerate(cols, start=1):
                cell = ws.cell(r, i)
                cell.value = "=" + (f.replace("F[", FIN + "[").replace("W[", "tbl_WIP[")
                                    .format(r=r, hdr=off))
                cell.number_format, cell.border = fmt, BOX
                cell.font = Font(size=9, color="595959")
                cell.fill = CALC_FILL
        return first + count - 1

    block(DEFER_FIRST, nlines, DEFER_COLS, hdr)

    # Second block: the WIP Movements rows that carry a Defer Start and End.
    # A deferred cost never touched this sheet before - it sat as one lump in
    # the month it was billed, which is not what the ledger should show.
    gap = DEFER_FIRST + nlines + 1
    ws.cell(gap, 1, "WIP Movements rows that spread over months  -  deferred cost "
                    "and manual deferred revenue").font = Font(bold=True, size=10,
                                                               color=NAVY)
    ws.cell(gap, 1).fill = PatternFill("solid", fgColor=LIGHT)
    wfirst = gap + 1
    last = block(wfirst, wlines, DEFER_WIP_COLS, gap)
    ws.auto_filter.ref = f"A{hdr}:{gcl(len(DEFER_COLS))}{DEFER_FIRST + nlines - 1}"
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
               '=SUMIFS(tbl_WIP[Amount],tbl_WIP[Month],"<>",tbl_WIP[Month],"<"&$C$4,'
               'tbl_WIP[Defer Start],"")'
               f'+SUM({defer_range("P", dlast)})'),
              ("Movement this month",
               '=SUMIFS(tbl_WIP[Amount],tbl_WIP[Month],$C$4,tbl_WIP[Defer Start],"")'
               f'+SUM({defer_range("Q", dlast)})'),
              ("Closing total", "=B6+B7"),
              # A WIP row on a job number Xero does not have cannot appear in
              # the job list below, so say what it is worth and what the list
              # below therefore adds up to. Otherwise the job-by-job total
              # quietly disagrees with the closing total.
              ("Of which is on a job Xero does not have - fix these",
               '=SUMIFS(tbl_WIP[Amount],tbl_WIP[Job in Xero?],"CHECK",'
               'tbl_WIP[Month],"<="&$C$4,tbl_WIP[Defer Start],"")'),
              ("Closing listed below, job by job", "=B8-B9")]
    for i, (label, formula) in enumerate(totals):
        r = 6 + i
        ws.cell(r, 1, label).font = Font(bold=True, size=10)
        vc = ws.cell(r, 2)
        vc.value, vc.number_format, vc.border = formula, CUR, BOX
        vc.fill = CALC_FILL if r not in (8, 10) else PatternFill("solid", fgColor=LIGHT)
        vc.font = Font(bold=True, size=10, color=NAVY)

    ws.conditional_formatting.add("B9", FormulaRule(
        formula=['ROUND(B9,2)<>0'], fill=PatternFill("solid", fgColor=BAD)))

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
                               f'tbl_WIP[Month],"<>",tbl_WIP[Month],"<"&$C$4,tbl_WIP[Defer Start],"")'
                               f'+SUMIF({defer_range("E", dlast)},$A{r},'
                               f'{defer_range("P", dlast)})')
        ws.cell(r, 4).value = (f'=SUMIFS(tbl_WIP[Amount],tbl_WIP[Job Number],$A{r},'
                               f'tbl_WIP[Month],$C$4,tbl_WIP[Defer Start],"")'
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
    # The list is every Xero job in number order, so the top of it is nearly all
    # nil. Say how many actually carry a balance, or it looks like nothing came
    # through at all.
    jc = ws.cell(10, 4)
    jc.value = f'=COUNTIF($F$12:$F${last},"yes")&" jobs carry a balance  -  filter '
    jc.value += 'column F to \'yes\' to see only those"'
    jc.font = Font(bold=True, size=10, color=NAVY)
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
                                    "Deferred cost - paid in advance",
                                    "Release of deferred cost",
                                    "Accrued cost - incurred not billed",
                                    "Adjustment / correction", "Migrated opening balance"],
         "lst_WIPType"),
        ("H", "WIP GL Code", ACC["wip_gl"], "lst_WIPGL"),
        ("I", "Team", DEPTS + ["Other"], "lst_Team"),
        # GL 11300 nets deferred revenue and deferred cost. Xero cannot tell them
        # apart; this column can, so the two are reported separately.
        ("J", "Revenue or Cost", ["Revenue", "Cost"], "lst_RevCost"),
        ("M", "Work Won Source", ["ZOHO", "Current RMS", "Qwilr", "Other"], "lst_WonSource"),
        ("O", "Work Won Status", ["Won", "Open", "Lost", "Cancelled"], "lst_WonStatus"),
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
    ws["T4"] = "Settings"
    ws["T8"], ws["U8"] = "GST rate", 0.1
    ws.cell(8, 20).font = Font(size=10, bold=True)
    gc = ws.cell(8, 21)
    gc.fill, gc.border, gc.number_format = INPUT_FILL, BOX, "0.0%"
    wb.defined_names.add(DefinedName("set_GSTRate", attr_text="Lists!$U$8"))
    ws["T9"], ws["U9"] = "WIP balance sheet account", "11300 Work in Progress"
    ws["T10"], ws["U10"] = "WIP P&L account", "44200 Closing Work in Progress"
    ws["T11"] = "The two accounts the WIP journal uses. Change them here and every " \
                "Debit / Credit on WIP Movements and Month-End follows."
    ws["T11"].font = Font(size=9, italic=True, color="808080")
    for cell in ("U9", "U10"):
        ws[cell].fill, ws[cell].border = INPUT_FILL, BOX
    ws["T5"], ws["U5"] = "Qwilr base address", "https://cts.qwilr.com/"
    ws["T6"], ws["U6"] = "Current RMS base address", ""
    ws["T7"] = "Optional. Paste the whole Current RMS address straight onto the " \
               "Production sheet and it links without this. Fill this in only if " \
               "you would rather type the bare opportunity number - end it with a slash."
    ws["T7"].font = Font(size=9, italic=True, color="808080")
    for r in (5, 6):
        ws.cell(r, 20).font = Font(size=10, bold=True)
        c = ws.cell(r, 21)
        c.fill, c.border, c.number_format = INPUT_FILL, BOX, TXT
    ws.column_dimensions["T"].width = 26
    ws.column_dimensions["U"].width = 44
    wb.defined_names.add(DefinedName("set_QwilrBase", attr_text="Lists!$U$5"))
    wb.defined_names.add(DefinedName("set_RMSBase", attr_text="Lists!$U$6"))
    wb.defined_names.add(DefinedName("set_WIPAsset", attr_text="Lists!$U$9"))
    wb.defined_names.add(DefinedName("set_WIPIncome", attr_text="Lists!$U$10"))

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
 ("P", "A cost can be deferred the same way, and it belongs on WIP Movements rather "
       "than Finance because it comes off a supplier bill, not an invoice. Type the "
       "bill amount as a positive, set Revenue or Cost to Cost, fill the two defer "
       "dates and put the expense account in P&L Account. It then spreads by month in "
       "the second block of the Deferred Revenue sheet exactly like revenue does."),
 ("P", "GL 11300 nets the two against each other, so Xero shows one figure and cannot "
       "tell you what it is made of. Month-End section 2 splits it: Total, Deferred / "
       "accrued revenue, Deferred / accrued cost. Compare only the Total column to "
       "Xero. Deferred cost never counts as revenue - section 1 leaves it out."),
 ("B", ""),
 ("H", "VIDEO REVENUE"),
 ("P", "There is no video split column any more, and nothing to keep honest. In Xero a "
       "cost centre sits on the invoice LINE, so a job invoiced part production and part "
       "video is simply two lines on the Finance sheet - one coded PRODUCTION, one coded "
       "VIDEO. The VIDEO figure at month-end is then a plain sum of the lines coded VIDEO, "
       "and it ties to the Xero VIDEO cost centre without any adjustment."),
 ("P", "The 26 migrated invoices that carried a video amount were split into two lines "
       "each on the way across. Ex GST in total is unchanged."),
 ("P", "The Production sheet carries a Video Revenue column beside the expenses, and Video "
       "% of Revenue next to it. Both read the Finance lines coded VIDEO, so nothing is "
       "typed. Job by job they reproduce the old Video Total column to the cent - 38,484.75 "
       "in total across 28 jobs."),
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
 ("H", "WORK WON  -  ZOHO, CURRENT RMS AND QWILR"),
 ("P", "Paste the month's won opportunities onto the Work Won sheet, one row each, with "
       "the source, the value and the job number where you have it. Month-End section 5 "
       "totals them by source and puts them next to what was actually invoiced."),
 ("P", "It is a sense check and nothing else. No revenue ever comes off this sheet. What "
       "it catches is work won and never invoiced, and invoicing with no opportunity "
       "behind it. Put a job number on a row and it also tells you what has been raised "
       "against that job so far."),
 ("P", "Month-End section 6 turns the same sheet into the forecast. It adds two "
       "things that have no invoice against them: a Finance line carrying a value but "
       "no invoice number, and work marked Won less everything Finance already holds "
       "against that job. Subtracting all of it is what stops the same work being "
       "counted twice. Open opportunities and deferred revenue still to release sit "
       "underneath as memos, deliberately outside the total."),
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
 ("P", "7.  Section 4 is the journal. It names the account to debit, the account to "
       "credit and the amount, split by cost centre. Key it into Xero as one manual "
       "journal, then put the journal reference back on the WIP Movements rows."),
 ("P", "8.  Section 5 is the Work Won sense check against ZOHO, Current RMS and Qwilr."),

 ("B", ""),
 ("H", "THE WIP SIGN RULE"),
 ("P", "One signed Amount column does both jobs, because GL 11300 nets accrued and deferred "
       "revenue. Positive = revenue recognised this month (work done but not yet invoiced, "
       "or a deferral released). Negative = revenue pushed out of this month (invoiced in "
       "advance). So revenue recognised = invoiced Ex GST + WIP movement, and the running "
       "total of the Amount column is the GL 11300 balance."),
 ("P", "Every row on WIP Movements now says which account to debit, which to credit and "
       "for how much, and gives you a narration to paste. A positive amount puts revenue "
       "back into the month, so WIP goes up: debit 11300 Work in Progress, credit 44200 "
       "Closing Work in Progress. A negative amount is the other way round. The old GL "
       "Code column is gone - it said 11300 on every row and these two columns replace it."),
 ("P", "Month-End section 4 adds those rows up with the automatic deferrals and gives you "
       "the whole month as one journal, split by cost centre so it can carry tracking in "
       "Xero. Debit and credit must agree and the sheet says Balanced when they do. Nothing "
       "posts itself - you key it in and put the reference back on the rows."),
 ("P", "The two accounts are set on the Lists sheet, column U rows 9 and 10. Change them "
       "there and every Debit / Credit follows."),
 ("B", ""),
 ("H", "LOCKED CELLS, LINKS AND TOTALS"),
 ("P", "Every sheet is protected with the password CTS1234. Only the cells you are meant "
       "to fill are open; the grey calculated ones are locked so a stray keystroke "
       "cannot wipe a formula. To edit a locked cell: Review, Unprotect Sheet, type "
       "CTS1234. Sorting and filtering work without unprotecting anything. The "
       "workbook structure uses the same password, so sheets cannot be renamed, "
       "reordered or deleted by accident either."),
 ("P", "Qwilr quotes and Current RMS numbers are clickable. Paste a full web address and it "
       "is used as it stands; type a bare number and it is added to the base address on the "
       "Lists sheet. The Qwilr base is already set, and Current RMS on the Production "
       "sheet now works exactly the same way - paste the whole Current RMS address into "
       "the Current RMS No cell and it links straight away. Only fill in Lists column U "
       "row 6 if you would rather type the bare opportunity number."),
 ("P", "Each sheet shows its total ex-GST at the top, and it follows the filter - filter to "
       "one client or one month and the total follows. Finance shows Ex GST, GST and Inc GST."),
 ("P", "The GST rate is no longer buried in the formula. It sits on the Lists sheet, column U "
       "row 8, and every GST cell reads it from there."),
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
 ("P", "   Department sheets: one row per job - Onsite 27, Production 58, Consulting 28."),
 ("P", "   WIP Movements: 126 rows."),
 ("P", "   Totals tie to the cent: Onsite 512,023.81, Production 460,883.46, Consulting "
       "857,950.27, WIP 17,537.91."),
 ("P", "   Revenue GL codes were assigned by rule (Onsite 41100, Production 42100, Video "
       "42150, Consulting 42800, Integration 42300). Spot-check them."),
 ("P", "   The 126 WIP rows were given a Type from the sign of the amount. Confirm before "
       "relying on them."),
 ("P", "   Production keeps its v2 columns. Client Email came across (57 of 58 jobs). "
       "Company, Event Grouping, Conf. Call Costs and Transcription Costs are there but "
       "empty - v2 showed them on screen and its macro never saved them, so there is "
       "nothing to bring. Conf. Call and Transcription now feed Total Expense, which v2 "
       "never did."),
 ("P", "   Job Type is gone - Cost Centres says the same thing and fills itself. Net Total "
       "and Discount as % of Net Total are gone for the reason under STILL TO CONFIRM."),
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


# Words that must keep their capitals when the easy-read version drops SHOUTING.
# Real acronyms - these keep their capitals even in a heading.
KEEP_CAPS = {"GST", "WIP", "GL", "PO", "P&L", "VBA", "CSV", "FY27", "CTS"}
# Shouted words inside body text that are only there for emphasis. Cost centre
# codes (PRODUCTION, VIDEO) and function names (VSTACK) are left alone - they
# are literal values you will see in the sheet.
DESHOUT = {"LINE": "line", "ALSO": "also", "LINES": "lines",
           "ASSUMPTION": "Assumption", "NOT": "not"}

ABBREV = {"no", "mr", "dr", "eg", "ie", "vs", "st", "ltd", "pty", "inc"}


def sentence_case(text):
    """WHO TYPES WHERE -> Who types where, without flattening real acronyms."""
    out = []
    for word in text.split(" "):
        core = word.strip(",.:;()-")
        if core and core.isupper() and core not in KEEP_CAPS and not core.isdigit():
            word = word.replace(core, core.lower())
        out.append(word)
    text = " ".join(out)
    for i, ch in enumerate(text):
        if ch.isalpha():
            return text[:i] + ch.upper() + text[i + 1:]
    return text


def split_sentences(text):
    """One sentence per line reads far easier. Split only where it is safe."""
    parts, start = [], 0
    for m in re.finditer(r"\.\s+", text):
        before = text[:m.start()]
        words = re.findall(r"[A-Za-z0-9&%]+", before)
        word = words[-1] if words else ""
        after = text[m.end():m.end() + 1]
        if not before or before[-1] not in (
                "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789)%"):
            continue
        if len(word) < 2 or word.lower() in ABBREV:
            continue
        if word.isdigit() and not before[:m.start()].strip():
            continue
        if not (after.isupper() or after.isdigit() or after in "$"):
            continue
        parts.append(text[start:m.start() + 1].strip())
        start = m.end()
    tail = text[start:].strip()
    if tail:
        parts.append(tail)
    return parts or [text]


def deshout(text):
    for word, plain in DESHOUT.items():
        text = re.sub(r"\b" + word + r"\b", plain, text)
    return text


# Two ways to read the same words. "normal" is the compact original.
# "easy" follows the British Dyslexia Association style guide: sans-serif
# Verdana, 13pt, off-white background, dark grey rather than black, short
# lines, one sentence per line, 1.5 line spacing, no italics, no SHOUTING.
STYLE = {
    "normal": dict(sheet="Read Me", font=None, size=10, head=11, title=20,
                   sub=10, ink="000000", paper=None, widths=(3, 30, 106),
                   chars=106, line=14, gap=1.0, italic_sub=True, split=False),
    "easy":   dict(sheet="Read Me (Easy Read)", font="Verdana", size=13,
                   head=14, title=22, sub=13, ink="1F1F1F", paper="FFFBF0",
                   widths=(4, 34, 92), chars=62, line=17, gap=1.55,
                   italic_sub=False, split=True),
}

POINTER = {
    "normal": "Bigger text, shorter lines, one sentence per line: open the "
              "Read Me (Easy Read) tab. Same words, easier on the eye.",
    "easy":   "This is the easy-read version. The standard, more compact "
              "version is on the Read Me tab.",
}


if BLANK:
    # No migration to describe, and no migrated exceptions to chase.
    _a = README.index(("H", "WHAT CAME ACROSS FROM v2"))
    _b = README.index(("W", "STILL TO CONFIRM"))
    README[_a:_b] = [
        ("H", "THIS IS THE EMPTY TEST COPY"),
        ("P", "Nothing has been migrated into this file. Every sheet is empty and ready to "
              "type into, and every formula, dropdown, lock, hyperlink and month-end check "
              "is exactly the same as the live workbook."),
        ("P", "What is still here is reference data, because the dropdowns and the job "
              "checks need it: the Xero job list, the cost centres, the revenue GL codes "
              "and the month list on the Lists sheet."),
        ("P", "To test it end to end: type an invoice line on the Finance sheet, then type "
              "the same job number on the matching department sheet, then open Month-End, "
              "set the month and work down. Section 3 should go to nil and section 1 should "
              "add up to what you typed."),
        ("B", ""),
    ]
    README[:] = [r for r in README
                 if not (r[0] == "P" and r[1].startswith("23 migrated records"))]


def build_readme(wb, mode="normal"):
    st = STYLE[mode]
    ws = wb.create_sheet(st["sheet"])
    ws.sheet_view.showGridLines = False
    for col, width in zip(("A", "B", "C"), st["widths"]):
        ws.column_dimensions[col].width = width

    def font(size=None, bold=False, italic=False, colour=None):
        kw = dict(size=size or st["size"], bold=bold, italic=italic,
                  color=colour or st["ink"])
        if st["font"]:
            kw["name"] = st["font"]
        return Font(**kw)

    def height(text, size=None):
        per = st["chars"] * (st["size"] / (size or st["size"]))
        lines = max(1, int(len(text) // per) + 1)
        return st["line"] * st["gap"] * lines

    r = 2
    for item in README:
        kind, body = item[0], item[1]
        if kind == "T":
            ws.cell(r, 2, body).font = font(st["title"], bold=True, colour=NAVY)
            ws.row_dimensions[r].height = st["title"] * 1.4
            r += 1
            if mode == "easy":
                continue
        elif kind == "S":
            ws.cell(r, 2, body).font = font(st["sub"], italic=st["italic_sub"],
                                            colour="595959")
        elif kind == "B":
            ws.row_dimensions[r].height = 10 * st["gap"]
        elif kind in ("H", "W"):
            text = sentence_case(body) if mode == "easy" else body
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
            c = ws.cell(r, 2, text)
            c.font = font(st["head"], bold=True, colour="FFFFFF")
            c.fill = PatternFill("solid", fgColor=(NAVY if kind == "H" else "8B2B2B"))
            c.alignment = Alignment(vertical="center", indent=1)
            ws.row_dimensions[r].height = st["head"] * 2.0
        elif kind == "R":
            label = item[1]
            text = deshout(item[2]) if mode == "easy" else item[2]
            lc = ws.cell(r, 2, label)
            lc.font = font(bold=True, colour=NAVY)
            lc.alignment = Alignment(vertical="top", wrap_text=True, indent=1)
            lc.fill = PatternFill("solid", fgColor=(LIGHT if mode == "normal" else "F4EDDD"))
            lc.border = BOX
            if mode == "easy":
                text = "\n".join(split_sentences(text))
            c = ws.cell(r, 3, text)
            c.font = font()
            c.alignment = Alignment(wrap_text=True, vertical="top", indent=1)
            c.border = BOX
            ws.row_dimensions[r].height = height(text) + (
                st["line"] * st["gap"] * text.count("\n"))
        else:
            if mode == "easy":
                lines = split_sentences(deshout(body))
                for line in lines:
                    c = ws.cell(r, 3, line)
                    c.font = font()
                    c.alignment = Alignment(wrap_text=True, vertical="top")
                    ws.row_dimensions[r].height = height(line)
                    r += 1
                ws.row_dimensions[r].height = 6
                r += 1
                continue
            c = ws.cell(r, 3, body)
            c.font = font()
            c.alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[r].height = height(body)
        r += 1

    # Cross-reference so nobody has to be told the other version exists.
    r += 1
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
    c = ws.cell(r, 2, POINTER[mode])
    c.font = font(bold=True, colour=NAVY)
    c.fill = PatternFill("solid", fgColor=(LIGHT if mode == "normal" else "F4EDDD"))
    c.alignment = Alignment(vertical="center", wrap_text=True, indent=1)
    c.border = BOX
    ws.row_dimensions[r].height = st["line"] * st["gap"] * 2

    if st["paper"]:
        paper = PatternFill("solid", fgColor=st["paper"])
        for row in ws.iter_rows(min_row=1, max_row=r + 2, min_col=1, max_col=4):
            for cell in row:
                if cell.fill is None or cell.fill.fgColor.rgb in (None, "00000000"):
                    cell.fill = paper


def main():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    build_lists(wb)
    nlines = build_finance(wb)
    for d in DEPTS:
        build_dept(wb, d)
    wlines = build_wip(wb)
    dlast = build_deferred(wb, nlines, wlines)
    build_month_end(wb, dlast)
    build_wip_summary(wb, dlast)
    build_won(wb)
    build_readme(wb, "normal")
    build_readme(wb, "easy")

    colours = {"Read Me": "7F7F7F", "Read Me (Easy Read)": "A6A6A6", "Finance": NAVY, "Month-End": "2E6B4F",
               "WIP Summary": "2E6B4F", "Deferred Revenue": "8B6A2B",
               "Work Won": "6B4E7A", "Onsite": SLATE, "Production": SLATE,
               "Consulting": SLATE, "WIP Movements": "8B6A2B", "Lists": "A6A6A6"}
    order = ["Read Me", "Read Me (Easy Read)", "Finance", "Month-End", "WIP Summary", "Deferred Revenue",
             "Onsite", "Production", "Consulting", "WIP Movements", "Work Won", "Lists"]
    for name, colour in colours.items():
        wb[name].sheet_properties.tabColor = colour
    wb._sheets = [wb[n] for n in order]
    wb.active = 0
    for sh in wb.worksheets:
        sh.sheet_view.zoomScale = 90
        sh.sheet_view.tabSelected = False
    wb["Read Me"].sheet_view.tabSelected = True
    wb.security = WorkbookProtection(lockStructure=True)
    wb.security.workbookPassword = SHEET_PASSWORD
    wb.calculation.fullCalcOnLoad = True
    wb.save(OUT)
    print(f"wrote {OUT} ({OUT.stat().st_size:,} bytes)")


if __name__ == "__main__":
    main()
