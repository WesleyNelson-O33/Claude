"""Build the CTS Financial Controller Pack.

One paste of the Xero account transactions export drives the departmental P&L,
the budget variance and the summary. The department comes from the bracket tag
on each GL line, so it is per transaction rather than per account.
"""
import json
import re
from datetime import date

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

GL_ROWS = 20000          # about a full financial year at CTS volume
FONT = "Arial"

NAVY = "1F3864"
YELLOW = "FFFF00"
GREY = "F2F2F2"
BAND = "DDEBF7"

INPUT_FONT = Font(name=FONT, size=10, color="0000FF")
LINK_FONT = Font(name=FONT, size=10, color="008000")
BODY = Font(name=FONT, size=10)
BOLD = Font(name=FONT, size=10, bold=True)
TITLE = Font(name=FONT, size=14, bold=True, color=NAVY)
HEAD = Font(name=FONT, size=10, bold=True, color="FFFFFF")

HEAD_FILL = PatternFill("solid", fgColor=NAVY)
YELLOW_FILL = PatternFill("solid", fgColor=YELLOW)
GREY_FILL = PatternFill("solid", fgColor=GREY)
BAND_FILL = PatternFill("solid", fgColor=BAND)

MONEY = '$#,##0;($#,##0);"-"'
PCT = '0.0%;(0.0%);"-"'
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# Department tag on the GL line maps to the reporting department.
DEPT_TAGS = [
    ("ONSITE", "ONSITE"),
    ("PRODUCTION", "PRODUCTION"),
    ("VIDEO", "VIDEO"),
    ("INTEGRATION", "INTEGRATION"),
    ("CONSULTING", "CONSULTING"),
    ("CTS", "ADMIN"),
    ("ONS", "ONSITE"),
    ("PRD", "PRODUCTION"),
    ("VID", "VIDEO"),
    ("INT", "INTEGRATION"),
    ("CONS", "CONSULTING"),
    ("ADMIN", "ADMIN"),
]
DEPTS = ["ONSITE", "PRODUCTION", "VIDEO", "INTEGRATION", "CONSULTING", "ADMIN"]

CATEGORY_ORDER = ["Income", "Cost of Sales", "Expenses", "Other Income", "Other Expenses"]

FY_START_YEAR = 2026     # FY27 runs July 2026 to June 2027
MONTHS = [(FY_START_YEAR + (0 if m >= 7 else 1), m) for m in list(range(7, 13)) + list(range(1, 7))]


def month_label(y, m):
    return date(y, m, 1).strftime("%b-%y")


def load_accounts(path):
    rows = json.load(open(path))
    seen, out = set(), []
    for a in rows:
        code = re.sub(r"[^0-9]", "", a["code"])
        if not code or code in seen:
            continue
        seen.add(code)
        out.append(dict(code=code, name=a["name"], cat=a["cat"], sub=a["sub"], typ=a["typ"]))
    return out


def subcategories(accounts):
    """Subcategories in category order, de-duplicated, keeping first appearance."""
    out = []
    for cat in CATEGORY_ORDER:
        for a in accounts:
            if a["cat"] == cat and (cat, a["sub"]) not in out:
                out.append((cat, a["sub"]))
    return out


def style_title(ws, text, note=None):
    ws["A1"] = text
    ws["A1"].font = TITLE
    if note:
        ws["A2"] = note
        ws["A2"].font = Font(name=FONT, size=9, italic=True, color="595959")
    ws.sheet_view.showGridLines = False


def build_setup(wb):
    ws = wb.create_sheet("Setup")
    style_title(ws, "Setup and control panel",
                "The yellow cell is the only one you type into. Everything else follows from it.")
    ws.column_dimensions["A"].width = 38
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 60

    rows = [
        ("Entity name", "Corporate Technology Services Pty Ltd", None),
        ("Reporting month", None, "Pick from the list. This drives the whole pack."),
        ("Financial year starts in month", 7, "July"),
        ("Financial year (the year it ends)", "=IF(MONTH($B$5)>=$B$6,YEAR($B$5)+1,YEAR($B$5))", None),
        ("Period number in the financial year", "=MOD(MONTH($B$5)-$B$6,12)+1", None),
        ("Prior financial year", "=$B$7-1", None),
        ("Comparative month last year", "=EDATE($B$5,-12)", None),
        ("Hours in a working day", 8, "Used by the utilisation block."),
    ]
    r = 4
    ws[f"A{r}"] = "1.  What am I reporting on"
    ws[f"A{r}"].font = BOLD
    r += 1
    for label, value, note in rows:
        ws[f"A{r}"] = label
        ws[f"A{r}"].font = BODY
        if value is not None:
            ws[f"B{r}"] = value
        ws[f"B{r}"].font = INPUT_FONT if not isinstance(value, str) or not str(value).startswith("=") else BODY
        if note:
            ws[f"C{r}"] = note
            ws[f"C{r}"].font = Font(name=FONT, size=9, italic=True, color="595959")
        r += 1

    ws["B5"] = date(2026, 8, 1)
    ws["B5"].number_format = "mmm yyyy"
    ws["B5"].fill = YELLOW_FILL
    ws["B5"].font = INPUT_FONT
    ws["B5"].border = BOX
    ws["B10"].number_format = "mmm yyyy"

    dv = DataValidation(type="list", formula1="=Months", allow_blank=False, showDropDown=False)
    ws.add_data_validation(dv)
    dv.add(ws["B5"])

    r += 1
    ws[f"A{r}"] = "2.  Data status  (calculated)"
    ws[f"A{r}"].font = BOLD
    r += 1
    status = [
        ("GL lines pasted", "=COUNTA(GL_Paste!$B$2:$B$%d)" % (GL_ROWS + 1)),
        ("GL capacity", GL_ROWS),
        ("Lines with no department tag", '=COUNTIF(Cleanup!$D$2:$D$%d,"UNALLOCATED")' % (GL_ROWS + 1)),
        ("Lines with an unknown account", '=COUNTIF(Cleanup!$E$2:$E$%d,"UNKNOWN")' % (GL_ROWS + 1)),
        ("Earliest GL date", "=IFERROR(MIN(GL_Paste!$E$2:$E$%d),0)" % (GL_ROWS + 1)),
        ("Latest GL date", "=IFERROR(MAX(GL_Paste!$E$2:$E$%d),0)" % (GL_ROWS + 1)),
    ]
    for label, formula in status:
        ws[f"A{r}"] = label
        ws[f"A{r}"].font = BODY
        ws[f"B{r}"] = formula
        ws[f"B{r}"].font = BODY
        if "date" in label:
            ws[f"B{r}"].number_format = "dd mmm yyyy"
        else:
            ws[f"B{r}"].number_format = "#,##0"
        r += 1
    return ws


def build_lists(wb, accounts):
    ws = wb.create_sheet("Lists")
    style_title(ws, "Reference lists",
                "The account master, the department tags and the month list. Add a new Xero account here or its GL lines are rejected.")
    for col, width in zip("ABCDEFGHIJ", (12, 38, 16, 30, 32, 4, 14, 16, 4, 14)):
        ws.column_dimensions[col].width = width

    ws["A4"] = "Account master"
    ws["A4"].font = BOLD
    for i, h in enumerate(["Code", "Account name", "Category", "Subcategory", "Type"]):
        c = ws.cell(row=5, column=1 + i, value=h)
        c.font = HEAD
        c.fill = HEAD_FILL
    for i, a in enumerate(accounts):
        r = 6 + i
        ws.cell(row=r, column=1, value=a["code"]).font = BODY
        ws.cell(row=r, column=2, value=a["name"]).font = BODY
        ws.cell(row=r, column=3, value=a["cat"]).font = BODY
        ws.cell(row=r, column=4, value=a["sub"]).font = BODY
        ws.cell(row=r, column=5, value=a["typ"]).font = BODY
    last_acct = 5 + len(accounts)

    ws["G4"] = "Department tags"
    ws["G4"].font = BOLD
    ws["G5"] = "Tag on GL line"
    ws["H5"] = "Department"
    for c in ("G5", "H5"):
        ws[c].font = HEAD
        ws[c].fill = HEAD_FILL
    for i, (tag, dept) in enumerate(DEPT_TAGS):
        ws.cell(row=6 + i, column=7, value=tag).font = BODY
        ws.cell(row=6 + i, column=8, value=dept).font = BODY
    last_tag = 5 + len(DEPT_TAGS)

    ws["J4"] = "Months in this financial year"
    ws["J4"].font = BOLD
    ws["J5"] = "Month"
    ws["J5"].font = HEAD
    ws["J5"].fill = HEAD_FILL
    for i, (y, m) in enumerate(MONTHS):
        c = ws.cell(row=6 + i, column=10, value=date(y, m, 1))
        c.number_format = "mmm yyyy"
        c.font = BODY

    ws["L4"] = "Public holidays"
    ws["L4"].font = BOLD
    ws["L5"] = "Date"
    ws["L5"].font = HEAD
    ws["L5"].fill = HEAD_FILL
    ws["M5"] = "Name"
    ws["M5"].font = HEAD
    ws["M5"].fill = HEAD_FILL
    ws.column_dimensions["L"].width = 14
    ws.column_dimensions["M"].width = 28
    ws["N4"] = "Type the public holidays here. The utilisation block uses them."
    ws["N4"].font = Font(name=FONT, size=9, italic=True, color="595959")
    for i in range(30):
        c = ws.cell(row=6 + i, column=12)
        c.fill = YELLOW_FILL
        c.number_format = "dd mmm yyyy"
        c.font = INPUT_FONT
        ws.cell(row=6 + i, column=13).fill = YELLOW_FILL

    defs = {
        "AcctCode": f"Lists!$A$6:$A${last_acct}",
        "AcctCat": f"Lists!$C$6:$C${last_acct}",
        "AcctSub": f"Lists!$D$6:$D${last_acct}",
        "DeptTag": f"Lists!$G$6:$G${last_tag}",
        "DeptName": f"Lists!$H$6:$H${last_tag}",
        "Months": f"Lists!$J$6:$J${5 + len(MONTHS)}",
        "Holidays": "Lists!$L$6:$L$35",
    }
    return ws, defs


GL_HEADERS = ["Index", "Account Code", "Account Name", "Source", "Date", "Contact",
              "Debit", "Credit", "Job Numbers", "Invoice Number", "Reference",
              "Description", "Cost Centres"]


def build_gl_paste(wb):
    ws = wb.create_sheet("GL_Paste")
    for i, h in enumerate(GL_HEADERS):
        c = ws.cell(row=1, column=1 + i, value=h)
        c.font = HEAD
        c.fill = HEAD_FILL
        c.border = BOX
    widths = (8, 12, 34, 14, 12, 28, 12, 12, 34, 14, 16, 40, 34)
    for i, w in enumerate(widths):
        ws.column_dimensions[get_column_letter(1 + i)].width = w
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False
    return ws


def build_cleanup(wb):
    """One row per GL line. Adds the four things the GL export does not carry:
    amount, month, department and subcategory."""
    ws = wb.create_sheet("Cleanup")
    style_title(ws, "Cleanup  -  calculated, never type here",
                "Amount is Credit less Debit, so income is positive and costs are negative. "
                "Department comes from the Cost Centres column, falling back to the bracket tag on the job number.")
    heads = ["Amount", "Month", "Tag", "Department", "Category", "Subcategory"]
    for i, h in enumerate(heads):
        c = ws.cell(row=4, column=1 + i, value=h)
        c.font = HEAD
        c.fill = HEAD_FILL
    for i, w in enumerate((14, 12, 10, 16, 16, 32)):
        ws.column_dimensions[get_column_letter(1 + i)].width = w
    ws.freeze_panes = "A5"
    ws.sheet_view.showGridLines = False

    # Row 5 lines up with GL_Paste row 2, the first pasted line.
    for i in range(GL_ROWS):
        r = 5 + i
        g = 2 + i
        blank = f'GL_Paste!$B{g}=""'
        tag_from = lambda col: (
            f'IFERROR(MID(GL_Paste!${col}{g},FIND("[",GL_Paste!${col}{g})+1,'
            f'FIND("]",GL_Paste!${col}{g})-FIND("[",GL_Paste!${col}{g})-1),"")'
        )
        ws.cell(row=r, column=1,
                value=f'=IF({blank},"",N(GL_Paste!$H{g})-N(GL_Paste!$G{g}))')
        ws.cell(row=r, column=2,
                value=f'=IF({blank},"",DATE(YEAR(GL_Paste!$E{g}),MONTH(GL_Paste!$E{g}),1))')
        ws.cell(row=r, column=3,
                value=(f'=IF({blank},"",IF(TRIM(GL_Paste!$M{g})<>"",TRIM(GL_Paste!$M{g}),'
                       f'{tag_from("I")}))'))
        ws.cell(row=r, column=4,
                value=f'=IF({blank},"",IFERROR(INDEX(DeptName,MATCH($C{r},DeptTag,0)),"UNALLOCATED"))')
        ws.cell(row=r, column=5,
                value=f'=IF({blank},"",IFERROR(INDEX(AcctCat,MATCH(GL_Paste!$B{g}&"",AcctCode,0)),"UNKNOWN"))')
        ws.cell(row=r, column=6,
                value=f'=IF({blank},"",IFERROR(INDEX(AcctSub,MATCH(GL_Paste!$B{g}&"",AcctCode,0)),"UNKNOWN"))')
        ws.cell(row=r, column=1).number_format = MONEY
        ws.cell(row=r, column=2).number_format = "mmm yyyy"
    return ws


def build_engine(wb, subs):
    """Subcategory by department by month. Everything downstream reads this."""
    ws = wb.create_sheet("Engine")
    style_title(ws, "Engine  -  calculated, never type here",
                "One row per department and subcategory, twelve months across. Read by the P&L, the Summary and the charts.")
    heads = ["Key", "Department", "Category", "Subcategory"]
    for i, h in enumerate(heads):
        c = ws.cell(row=4, column=1 + i, value=h)
        c.font = HEAD
        c.fill = HEAD_FILL
    for i, (y, m) in enumerate(MONTHS):
        c = ws.cell(row=4, column=5 + i, value=date(y, m, 1))
        c.font = HEAD
        c.fill = HEAD_FILL
        c.number_format = "mmm-yy"
    ws.cell(row=4, column=17, value="FY total").font = HEAD
    ws.cell(row=4, column=17).fill = HEAD_FILL

    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 15
    ws.column_dimensions["C"].width = 15
    ws.column_dimensions["D"].width = 32
    for i in range(len(MONTHS) + 1):
        ws.column_dimensions[get_column_letter(5 + i)].width = 12
    ws.freeze_panes = "E5"
    ws.sheet_view.showGridLines = False

    last = GL_ROWS + 4
    r = 5
    index = {}
    for dept in DEPTS:
        for cat, sub in subs:
            ws.cell(row=r, column=1, value=f"{dept}|{sub}").font = BODY
            ws.cell(row=r, column=2, value=dept).font = BODY
            ws.cell(row=r, column=3, value=cat).font = BODY
            ws.cell(row=r, column=4, value=sub).font = BODY
            for i in range(len(MONTHS)):
                col = get_column_letter(5 + i)
                f = (f"=SUMIFS(Cleanup!$A$5:$A${last},"
                     f"Cleanup!$D$5:$D${last},$B{r},"
                     f"Cleanup!$F$5:$F${last},$D{r},"
                     f"Cleanup!$B$5:$B${last},{col}$4)")
                c = ws.cell(row=r, column=5 + i, value=f)
                c.number_format = MONEY
                c.font = BODY
            c = ws.cell(row=r, column=17,
                        value=f"=SUM(E{r}:{get_column_letter(4 + len(MONTHS))}{r})")
            c.number_format = MONEY
            c.font = BOLD
            index[(dept, sub)] = r
            r += 1
    return ws, index, r - 1


