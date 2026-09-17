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
    ws.column_dimensions["C"].width = 62

    def put(r, label, value, note=None, pct=False, money=False):
        ws[f"A{r}"] = label
        ws[f"A{r}"].font = BODY
        ws[f"B{r}"] = value
        calc = isinstance(value, str) and value.startswith("=")
        ws[f"B{r}"].font = BODY if calc else INPUT_FONT
        if not calc:
            ws[f"B{r}"].fill = YELLOW_FILL
            ws[f"B{r}"].border = BOX
        if pct:
            ws[f"B{r}"].number_format = "0%"
        if money:
            ws[f"B{r}"].number_format = "$#,##0"
        if note:
            ws[f"C{r}"] = note
            ws[f"C{r}"].font = Font(name=FONT, size=9, italic=True, color="595959")

    ws["A4"] = "1.  What am I reporting on"
    ws["A4"].font = BOLD
    put(5, "Reporting month", date(2026, 8, 1), "The one control you change each month.")
    ws["B5"].number_format = "mmm yyyy"
    put(6, "Financial year starts in month", 7, "July")
    put(7, "Financial year (the year it ends)", "=IF(MONTH($B$5)>=$B$6,YEAR($B$5)+1,YEAR($B$5))")
    put(8, "Period number in the financial year", "=MOD(MONTH($B$5)-$B$6,12)+1")
    put(9, "Prior financial year", "=$B$7-1")
    put(10, "Comparative month last year", "=EDATE($B$5,-12)")
    ws["B10"].number_format = "mmm yyyy"
    put(11, "Hours in a working day", 8, "Used by the utilisation sheet.")

    ws["A13"] = "2.  Risk flag thresholds"
    ws["A13"].font = BOLD
    put(14, "Materiality floor ($)", 5000, "Below this, a variance is always Low.", money=True)
    put(15, "Medium risk threshold (% variance)", 0.10, None, pct=True)
    put(16, "High risk threshold (% variance)", 0.25, None, pct=True)
    put(17, "High risk threshold ($ variance)", 50000, None, money=True)

    dv = DataValidation(type="list", formula1="=Months", allow_blank=False, showDropDown=False)
    ws.add_data_validation(dv)
    dv.add(ws["B5"])

    ws["A19"] = "3.  Data status  (calculated)"
    ws["A19"].font = BOLD
    status = [
        ("GL lines pasted", "=COUNTA(GL_Paste!$B$2:$B$%d)" % (GL_ROWS + 1), "#,##0"),
        ("GL capacity", GL_ROWS, "#,##0"),
        ("Lines with no department tag", '=COUNTIF(Cleanup!$D$5:$D$%d,"UNALLOCATED")' % (GL_ROWS + 4), "#,##0"),
        ("Lines with an unknown account", '=COUNTIF(Cleanup!$E$5:$E$%d,"UNKNOWN")' % (GL_ROWS + 4), "#,##0"),
        ("Earliest GL date", "=IFERROR(MIN(GL_Paste!$E$2:$E$%d),0)" % (GL_ROWS + 1), "dd mmm yyyy"),
        ("Latest GL date", "=IFERROR(MAX(GL_Paste!$E$2:$E$%d),0)" % (GL_ROWS + 1), "dd mmm yyyy"),
    ]
    r = 20
    for label, formula, fmt in status:
        ws[f"A{r}"] = label
        ws[f"A{r}"].font = BODY
        ws[f"B{r}"] = formula
        ws[f"B{r}"].font = BODY
        ws[f"B{r}"].number_format = fmt
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



# ----------------------------------------------------------------------------
# One row skeleton shared by the P&L, both budget sheets, the prior year and
# the Summary, so a paste into any of them lines up with the rest.
# ----------------------------------------------------------------------------

def layout(subs):
    """(kind, label, dept, cat, sub) per row, in display order."""
    out = []
    for dept in DEPTS:
        out.append(("dept", dept, dept, None, None))
        for cat in CATEGORY_ORDER:
            cat_subs = [x for c, x in subs if c == cat]
            if not cat_subs:
                continue
            out.append(("cat", "    " + cat, dept, cat, None))
            for sub in cat_subs:
                out.append(("sub", "        " + sub, dept, cat, sub))
        out.append(("blank", "", None, None, None))
    return out


ROW0 = 6  # first body row on every reporting sheet


def row_map(rows):
    """dept row, category rows and their member subcategory rows, by department."""
    info = {}
    for i, (kind, _, dept, cat, sub) in enumerate(rows):
        r = ROW0 + i
        if kind == "dept":
            info[dept] = {"row": r, "cats": {}}
        elif kind == "cat":
            info[dept]["cats"][cat] = {"row": r, "subs": []}
        elif kind == "sub":
            info[dept]["cats"][cat]["subs"].append(r)
    return info


def paint(ws, rows, ncols):
    for i, (kind, label, *_rest) in enumerate(rows):
        r = ROW0 + i
        if kind == "blank":
            continue
        ws.cell(row=r, column=1, value=label)
        for c in range(1, ncols + 1):
            cell = ws.cell(row=r, column=c)
            if kind == "dept":
                cell.fill = HEAD_FILL
                cell.font = Font(name=FONT, size=10, bold=True, color="FFFFFF")
            elif kind == "cat":
                cell.fill = BAND_FILL
                cell.font = BOLD
            else:
                cell.font = BODY
        if kind == "sub":
            ws.row_dimensions[r].outlineLevel = 2
            ws.row_dimensions[r].hidden = True
        elif kind == "cat":
            ws.row_dimensions[r].outlineLevel = 1
    ws.sheet_properties.outlinePr.summaryBelow = False


def roll_up(ws, info, cols):
    """Category rows sum their subcategories; department rows sum their categories."""
    for dept, d in info.items():
        for cat, c in d["cats"].items():
            subs = c["subs"]
            for col in cols:
                ws[f"{col}{c['row']}"] = f"=SUM({col}{subs[0]}:{col}{subs[-1]})"
        for col in cols:
            parts = "+".join(f"{col}{c['row']}" for c in d["cats"].values())
            ws[f"{col}{d['row']}"] = "=" + parts


def risk(r, var_d, var_p, a, b):
    return (f'=IF(AND({a}{r}=0,{b}{r}=0),"No Activity",'
            f'IF(ABS({var_d}{r})<Setup!$B$14,"Low",'
            f'IF(OR(ABS({var_d}{r})>=Setup!$B$17,ABS({var_p}{r})>=Setup!$B$16),"High",'
            f'IF(ABS({var_p}{r})>=Setup!$B$15,"Medium","Low"))))')


def header(ws, labels, month_cols=True, widths=None):
    if month_cols:
        for i, (y, m) in enumerate(MONTHS):
            c = ws.cell(row=5, column=2 + i, value=date(y, m, 1))
            c.number_format = "mmm-yy"
    ws.cell(row=5, column=1, value="Line")
    for ref, text in labels:
        ws[f"{ref}5"] = text
    for c in ws[5]:
        if c.value is not None:
            c.font = HEAD
            c.fill = HEAD_FILL
            c.alignment = Alignment(horizontal="center", wrap_text=True)
    ws.column_dimensions["A"].width = 44
    for ref, w in (widths or []):
        ws.column_dimensions[ref].width = w
    ws.freeze_panes = "B6"
    ws.sheet_view.showGridLines = False


def titles(ws, note=None):
    ws["A1"] = "Profit & Loss"
    ws["A1"].font = TITLE
    ws["A2"] = ('=TEXT(Setup!$B$5,"mmm yyyy")&"   |   click the + and - buttons on the left to '
                'open a department, then a category, down to subcategory detail"')
    ws["A2"].font = Font(name=FONT, size=9, italic=True, color="595959")
    if note:
        ws["A3"] = note
        ws["A3"].font = Font(name=FONT, size=9, italic=True, color="C00000")


MONTH_COLS = [get_column_letter(2 + i) for i in range(12)]
FY_COL = get_column_letter(14)          # N


def build_paste_sheet(wb, title, subs, note):
    """Budget FY27, Budget FY26 and Actual FY26 - same skeleton, yellow months."""
    ws = wb.create_sheet(title)
    titles(ws, note)
    header(ws, [(FY_COL, "FY total")],
           widths=[(c, 12) for c in MONTH_COLS] + [(FY_COL, 13)])
    rows = layout(subs)
    paint(ws, rows, 14)
    info = row_map(rows)
    for i, (kind, _, dept, cat, sub) in enumerate(rows):
        r = ROW0 + i
        if kind != "sub":
            continue
        for col in MONTH_COLS:
            c = ws[f"{col}{r}"]
            c.fill = YELLOW_FILL
            c.font = INPUT_FONT
            c.number_format = MONEY
    roll_up(ws, info, MONTH_COLS)
    for i, (kind, *_x) in enumerate(rows):
        r = ROW0 + i
        if kind == "blank":
            continue
        ws[f"{FY_COL}{r}"] = f"=SUM(B{r}:M{r})"
        for col in MONTH_COLS + [FY_COL]:
            ws[f"{col}{r}"].number_format = MONEY
    return ws


def sum_months(sheet, r, lo, hi):
    """Sum a row's month cells between two month indexes, inclusive, 0 based."""
    cols = MONTH_COLS[lo:hi + 1]
    if not cols:
        return "0"
    return f"SUM('{sheet}'!{cols[0]}{r}:{cols[-1]}{r})"


def ytd_expr(sheet, r):
    """Year to date is months 1 to the reporting period. SUMIF over the header row."""
    return (f"SUMIF('{sheet}'!$B$5:$M$5,\"<=\"&Setup!$B$5,'{sheet}'!$B{r}:$M{r})")


def qtr_expr(sheet, r):
    return (f"SUMIFS('{sheet}'!$B{r}:$M{r},'{sheet}'!$B$5:$M$5,\"<=\"&Setup!$B$5,"
            f"'{sheet}'!$B$5:$M$5,\">=\"&DATE(YEAR(Setup!$B$5),FLOOR(MONTH(Setup!$B$5)-1,3)+1,1))")


def build_pl(wb, subs, engine_index):
    """The P&L. Actual to the reporting month, budget after it, then her
    comparison blocks: budget, prior year, year on year and risk."""
    ws = wb.create_sheet("P&L FY27")
    titles(ws)
    labels = [(FY_COL, "FY total"), ("O", "YTD"), ("P", "YTD vs Budget"), ("Q", "Var $"),
              ("R", "Var %"), ("T", "PY"), ("U", "PY vs PY budget"), ("V", "Var $"),
              ("W", "Var %"), ("Y", "YOY"), ("Z", "YOY Budget"), ("AA", "Var $"),
              ("AB", "Var %"), ("AC", "Risk")]
    header(ws, labels,
           widths=[(c, 12) for c in MONTH_COLS] + [(FY_COL, 13)]
                  + [(r, 13) for r, _ in labels[1:]] + [("S", 3), ("X", 3), ("AC", 12)])
    rows = layout(subs)
    paint(ws, rows, 29)
    info = row_map(rows)

    for i, (kind, _, dept, cat, sub) in enumerate(rows):
        r = ROW0 + i
        if kind != "sub":
            continue
        eng = engine_index[(dept, sub)]
        for j, col in enumerate(MONTH_COLS):
            ecol = get_column_letter(5 + j)
            ws[f"{col}{r}"] = (f"=IF({col}$5<=Setup!$B$5,Engine!${ecol}{eng},"
                               f"'Budget FY27'!{col}{r})")
    roll_up(ws, info, MONTH_COLS)

    for i, (kind, *_x) in enumerate(rows):
        r = ROW0 + i
        if kind == "blank":
            continue
        ws[f"{FY_COL}{r}"] = f"=SUM(B{r}:M{r})"
        ws[f"O{r}"] = "=" + ytd_expr("P&L FY27", r)
        ws[f"P{r}"] = "=" + ytd_expr("Budget FY27", r)
        ws[f"Q{r}"] = f"=O{r}-P{r}"
        ws[f"R{r}"] = f"=IFERROR(Q{r}/ABS(P{r}),0)"
        ws[f"T{r}"] = "=" + ytd_expr("Actual FY26", r)
        ws[f"U{r}"] = "=" + ytd_expr("Budget FY26", r)
        ws[f"V{r}"] = f"=T{r}-U{r}"
        ws[f"W{r}"] = f"=IFERROR(V{r}/ABS(U{r}),0)"
        ws[f"Y{r}"] = f"=O{r}-T{r}"
        ws[f"Z{r}"] = f"=P{r}-U{r}"
        ws[f"AA{r}"] = f"=Y{r}-Z{r}"
        ws[f"AB{r}"] = f"=IFERROR(AA{r}/ABS(Z{r}),0)"
        ws[f"AC{r}"] = risk(r, "Q", "R", "O", "P")
        for col in MONTH_COLS + [FY_COL, "O", "P", "Q", "T", "U", "V", "Y", "Z", "AA"]:
            ws[f"{col}{r}"].number_format = MONEY
        for col in ("R", "W", "AB"):
            ws[f"{col}{r}"].number_format = PCT
        ws[f"AC{r}"].alignment = Alignment(horizontal="center")
    return ws


def build_summary(wb, subs):
    """Her Summary: three rolling months, YTD, quarter, year on year, risk."""
    ws = wb.create_sheet("Summary")
    titles(ws)
    labels = [("E", "YTD"), ("F", "YTD vs Rolling Forecast"), ("G", "YTD vs Budget"),
              ("H", "Var $"), ("I", "Var %"),
              ("K", "Quarterly"), ("L", "QTR vs QTR budget"), ("M", "Var $"), ("N", "Var %"),
              ("P", "YOY"), ("Q", "YOY Budget"), ("R", "Var $"), ("S", "Var %"),
              ("T", "Risk")]
    header(ws, labels, month_cols=False,
           widths=[("B", 13), ("C", 13), ("D", 13)]
                  + [(r, 14) for r, _ in labels] + [("J", 3), ("O", 3), ("T", 12)])
    # The three rolling months are the quarter the reporting month sits in.
    for col, off in (("B", 0), ("C", 1), ("D", 2)):
        c = ws[f"{col}5"]
        c.value = (f'=DATE(YEAR(Setup!$B$5),FLOOR(MONTH(Setup!$B$5)-1,3)+1+{off},1)')
        c.number_format = "mmm yyyy"
        c.font = HEAD
        c.fill = HEAD_FILL
        c.alignment = Alignment(horizontal="center")

    rows = layout(subs)
    paint(ws, rows, 20)
    info = row_map(rows)

    for i, (kind, *_x) in enumerate(rows):
        r = ROW0 + i
        if kind == "blank":
            continue
        for col in ("B", "C", "D"):
            ws[f"{col}{r}"] = (f"=SUMIF('P&L FY27'!$B$5:$M$5,{col}$5,'P&L FY27'!$B{r}:$M{r})")
        ws[f"E{r}"] = "=" + ytd_expr("P&L FY27", r)
        ws[f"F{r}"] = "=" + ytd_expr("Budget FY27", r)
        ws[f"G{r}"] = "=" + ytd_expr("Budget FY27", r)
        ws[f"H{r}"] = f"=E{r}-G{r}"
        ws[f"I{r}"] = f"=IFERROR(H{r}/ABS(G{r}),0)"
        ws[f"K{r}"] = "=" + qtr_expr("P&L FY27", r)
        ws[f"L{r}"] = "=" + qtr_expr("Budget FY27", r)
        ws[f"M{r}"] = f"=K{r}-L{r}"
        ws[f"N{r}"] = f"=IFERROR(M{r}/ABS(L{r}),0)"
        ws[f"P{r}"] = "=E%d-'Actual FY26'!O%d" % (r, r) if False else f"=E{r}-'Actual FY26'!O{r}"
        ws[f"Q{r}"] = f"=G{r}-'Budget FY26'!O{r}"
        ws[f"R{r}"] = f"=P{r}-Q{r}"
        ws[f"S{r}"] = f"=IFERROR(R{r}/ABS(Q{r}),0)"
        ws[f"T{r}"] = risk(r, "H", "I", "E", "G")
        for col in ("B", "C", "D", "E", "F", "G", "H", "K", "L", "M", "P", "Q", "R"):
            ws[f"{col}{r}"].number_format = MONEY
        for col in ("I", "N", "S"):
            ws[f"{col}{r}"].number_format = PCT
        ws[f"T{r}"].alignment = Alignment(horizontal="center")
    return ws


UTIL_DEPTS = ["Onsite", "Production", "Video", "Integration", "Consulting", "CTS"]
HOUR_TYPES = ["Chargeable", "Non-chargeable", "Leave"]
UCOLS = [get_column_letter(4 + i) for i in range(12)]   # D..O


def build_utilisation(wb):
    ws = wb.create_sheet("Utilisation")
    ws.sheet_view.showGridLines = False
    ws["A1"] = "Utilisation"
    ws["A1"].font = TITLE
    ws["A2"] = ("Working days and working hours calculate themselves from the calendar and the "
                "public holidays on Lists. Type the hours in the yellow cells.")
    ws["A2"].font = Font(name=FONT, size=9, italic=True, color="595959")

    for i, (y, m) in enumerate(MONTHS):
        c = ws.cell(row=4, column=4 + i, value=date(y, m, 1))
        c.number_format = "mmm-yy"
        c.font = HEAD
        c.fill = HEAD_FILL
        c.alignment = Alignment(horizontal="center")
    ws["A4"] = "Calendar"
    ws["A4"].font = HEAD
    ws["A4"].fill = HEAD_FILL

    ws["A5"] = "Working days"
    ws["A6"] = "Working hours"
    for c in ("A5", "A6"):
        ws[c].font = BOLD
    for i, col in enumerate(UCOLS):
        ws[f"{col}5"] = f"=NETWORKDAYS({col}$4,EOMONTH({col}$4,0),Holidays)"
        ws[f"{col}5"].number_format = "#,##0"
        ws[f"{col}6"] = f"={col}5*Setup!$B$11"
        ws[f"{col}6"].number_format = "#,##0"
        for r in (5, 6):
            ws[f"{col}{r}"].font = BOLD
            ws[f"{col}{r}"].fill = GREY_FILL

    heads = [("A", "FY (year it ends)", 18), ("B", "Department", 16), ("C", "Hour Type", 18)]
    for ref, text, w in heads:
        ws[f"{ref}8"] = text
        ws[f"{ref}8"].font = HEAD
        ws[f"{ref}8"].fill = HEAD_FILL
        ws.column_dimensions[ref].width = w
    for i, (y, m) in enumerate(MONTHS):
        c = ws.cell(row=8, column=4 + i, value=date(y, m, 1))
        c.number_format = "mmm-yy"
        c.font = HEAD
        c.fill = HEAD_FILL
        c.alignment = Alignment(horizontal="center")
        ws.column_dimensions[UCOLS[i]].width = 11
    ws["P8"] = "FY total"
    ws["P8"].font = HEAD
    ws["P8"].fill = HEAD_FILL
    ws.column_dimensions["P"].width = 12
    ws.freeze_panes = "D9"

    r = 9
    for dept in UTIL_DEPTS:
        first = r
        for ht in HOUR_TYPES:
            ws[f"A{r}"] = "=Setup!$B$7"
            ws[f"B{r}"] = dept
            ws[f"C{r}"] = ht
            for col in UCOLS:
                c = ws[f"{col}{r}"]
                c.fill = YELLOW_FILL
                c.font = INPUT_FONT
                c.number_format = "#,##0.0"
            ws[f"P{r}"] = f"=SUM(D{r}:O{r})"
            ws[f"P{r}"].number_format = "#,##0.0"
            for ref in ("A", "B", "C"):
                ws[f"{ref}{r}"].font = BODY
            r += 1
        charge, noncharge = first, first + 1
        ws[f"C{r}"] = "Utilisation %"
        ws[f"C{r}"].font = BOLD
        for col in UCOLS:
            ws[f"{col}{r}"] = f"=IFERROR({col}{charge}/({col}{charge}+{col}{noncharge}),0)"
            ws[f"{col}{r}"].number_format = PCT
            ws[f"{col}{r}"].font = BOLD
            ws[f"{col}{r}"].fill = BAND_FILL
        ws[f"P{r}"] = f"=IFERROR(P{charge}/(P{charge}+P{noncharge}),0)"
        ws[f"P{r}"].number_format = PCT
        for ref in ("A", "B", "C", "P"):
            ws[f"{ref}{r}"].fill = BAND_FILL
        r += 1
        ws[f"C{r}"] = "FTE"
        ws[f"C{r}"].font = BOLD
        for col in UCOLS:
            ws[f"{col}{r}"] = f"=IFERROR(({col}{charge}+{col}{noncharge})/{col}$6,0)"
            ws[f"{col}{r}"].number_format = "0.00"
            ws[f"{col}{r}"].fill = BAND_FILL
        for ref in ("A", "B", "C", "P"):
            ws[f"{ref}{r}"].fill = BAND_FILL
        r += 2
    return ws


def build_pl_check(wb):
    """Paste the Xero P&L totals here. Proves the GL paste agrees with the P&L."""
    ws = wb.create_sheet("PL_Check")
    style_title(ws, "P&L control  -  does the GL paste agree with Xero",
                "Run the profit and loss you already run and type the five category totals into "
                "the yellow cells. Every line should read OK before you report.")
    ws.column_dimensions["A"].width = 46
    for col in "BCDE":
        ws.column_dimensions[col].width = 18
    ws.sheet_view.showGridLines = False
    for col, text in zip("ABCDE", ["Category", "Per the GL paste", "Per the Xero P&L",
                                   "Difference", "Status"]):
        ws[f"{col}4"] = text
        ws[f"{col}4"].font = HEAD
        ws[f"{col}4"].fill = HEAD_FILL
    last = GL_ROWS + 4
    for i, cat in enumerate(CATEGORY_ORDER):
        r = 5 + i
        ws[f"A{r}"] = cat
        ws[f"A{r}"].font = BODY
        ws[f"B{r}"] = (f'=SUMIFS(Cleanup!$A$5:$A${last},Cleanup!$E$5:$E${last},$A{r},'
                       f'Cleanup!$B$5:$B${last},Setup!$B$5)')
        c = ws[f"C{r}"]
        c.fill = YELLOW_FILL
        c.font = INPUT_FONT
        ws[f"D{r}"] = f"=$B{r}-$C{r}"
        ws[f"E{r}"] = f'=IF(ABS($D{r})<1,"OK","CHECK")'
        ws[f"E{r}"].font = BOLD
        ws[f"E{r}"].alignment = Alignment(horizontal="center")
        for col in "BCD":
            ws[f"{col}{r}"].number_format = MONEY
    r = 5 + len(CATEGORY_ORDER)
    ws[f"A{r}"] = "Net profit"
    ws[f"A{r}"].font = BOLD
    ws[f"B{r}"] = f"=SUM(B5:B{r-1})"
    ws[f"C{r}"] = f"=SUM(C5:C{r-1})"
    ws[f"D{r}"] = f"=$B{r}-$C{r}"
    ws[f"E{r}"] = f'=IF(ABS($D{r})<1,"OK","CHECK")'
    for col in "ABCDE":
        ws[f"{col}{r}"].font = BOLD
        ws[f"{col}{r}"].fill = GREY_FILL
        if col in "BCD":
            ws[f"{col}{r}"].number_format = MONEY
    return ws


def build_readme(wb):
    ws = wb.create_sheet("README", 0)
    style_title(ws, "CTS Financial Controller Pack",
                "One paste a month. Everything else is formulas.")
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 96
    steps = [
        ("EACH MONTH", ""),
        ("1.  Setup B5", "Pick the reporting month. The financial year, the period and every comparative follow from it."),
        ("2.  GL_Paste", "Run the Account transactions for P&L analysis report in Xero, year to date from 1 July, and paste it at A2 in the column order shown in row 1. This is the only thing you paste each month."),
        ("3.  PL_Check", "Type the five Xero P&L category totals. Every line should read OK."),
        ("4.  Read", "P&L FY27, then Summary."),
        ("", ""),
        ("ONCE A YEAR", ""),
        ("Budget FY27", "This year's budget or rolling forecast. Revise when the forecast moves."),
        ("Actual FY26", "Last year's actuals. Drives the PY and YOY columns."),
        ("Budget FY26", "Last year's budget."),
        ("Utilisation", "Hours per department per month. Working days and working hours calculate themselves."),
        ("Lists", "Public holidays, and any new Xero account."),
        ("", ""),
        ("HOW THE DEPARTMENT IS DECIDED", ""),
        ("Per transaction", "From the Cost Centres column, which carries the department name outright. On the August export all 3,771 lines were tagged. If it is ever blank the pack falls back to the bracket tag on the Job Numbers column."),
        ("Why not the account", "Three quarters of the P&L sits in accounts with no department in the name, Contract Support Staff and Equipment Hires among them. Only the transaction knows which department earned or spent it."),
        ("", ""),
        ("THINGS THAT WILL CATCH YOU OUT", ""),
        ("Sign convention", "Amount is Credit less Debit. Income is positive, costs are negative, so gross profit is Income plus Cost of Sales rather than minus."),
        ("Forecast months", "On the P&L, months up to the reporting month come from the GL. Months after it come from Budget FY27, so you always see a full twelve and the shortfall ahead."),
        ("Risk column", "No Activity, Low, Medium or High, off the thresholds on Setup. Below the materiality floor it is always Low."),
        ("Capacity", "GL_Paste holds 20,000 lines, about a full year at your volume. Setup tells you how many you have used."),
    ]
    r = 4
    for label, text in steps:
        if text == "":
            ws.cell(row=r, column=2, value=label).font = Font(name=FONT, size=11, bold=True, color=NAVY)
        else:
            ws.cell(row=r, column=2, value=label).font = BOLD
            c = ws.cell(row=r, column=3, value=text)
            c.font = BODY
            c.alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[r].height = 32
        r += 1
    return ws


def main():
    accounts = load_accounts("/tmp/claude-0/-home-user-Claude/"
                             "a91f56e6-b7da-5dad-8e0b-6d5f516f9360/scratchpad/lists.json")
    subs = subcategories(accounts)
    print(f"accounts {len(accounts)}  subcategories {len(subs)}  departments {len(DEPTS)}")

    wb = Workbook()
    wb.remove(wb.active)

    build_setup(wb)
    _, defs = build_lists(wb, accounts)
    build_gl_paste(wb)
    build_pl_check(wb)
    build_cleanup(wb)
    _, engine_index, _ = build_engine(wb, subs)
    build_paste_sheet(wb, "Budget FY27", subs,
                      "This year's budget or rolling forecast. Paste into the yellow months.")
    build_paste_sheet(wb, "Actual FY26", subs,
                      "Last year's actuals. Paste once, at the start of the year.")
    build_paste_sheet(wb, "Budget FY26", subs,
                      "Last year's budget. Paste once, at the start of the year.")
    build_pl(wb, subs, engine_index)
    build_summary(wb, subs)
    build_utilisation(wb)
    build_readme(wb)

    from openpyxl.workbook.defined_name import DefinedName
    for name, ref in defs.items():
        wb.defined_names.add(DefinedName(name, attr_text=ref))

    yellow = {"Setup", "GL_Paste", "PL_Check", "Budget FY27", "Actual FY26",
              "Budget FY26", "Utilisation"}
    for ws in wb.worksheets:
        ws.sheet_properties.tabColor = YELLOW if ws.title in yellow else NAVY

    out = "/home/user/Claude/pack/CTS Financial Controller Pack.xlsx"
    wb.save(out)
    print("saved", out)


if __name__ == "__main__":
    main()
