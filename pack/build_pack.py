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


def build_budget_paste(wb, subs):
    """Same shape as the Engine so a P&L row can read either with one offset."""
    ws = wb.create_sheet("Budget_Paste")
    style_title(ws, "Budget and rolling forecast  -  type here",
                "One row per department and subcategory, twelve months across. Same signs as the P&L: "
                "income positive, costs negative. Enter it once a year and revise the forecast when it changes.")
    for i, h in enumerate(["Key", "Department", "Category", "Subcategory"]):
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

    r = 5
    for dept in DEPTS:
        for cat, sub in subs:
            ws.cell(row=r, column=1, value=f"{dept}|{sub}").font = BODY
            ws.cell(row=r, column=2, value=dept).font = BODY
            ws.cell(row=r, column=3, value=cat).font = BODY
            ws.cell(row=r, column=4, value=sub).font = BODY
            for i in range(len(MONTHS)):
                c = ws.cell(row=r, column=5 + i)
                c.number_format = MONEY
                c.fill = YELLOW_FILL
                c.font = INPUT_FONT
            c = ws.cell(row=r, column=17,
                        value=f"=SUM(E{r}:{get_column_letter(4 + len(MONTHS))}{r})")
            c.number_format = MONEY
            c.font = BOLD
            r += 1
    return ws


def build_pl(wb, subs, engine_index):
    """Department, then category, then subcategory. Actual to the reporting month,
    budget or rolling forecast after it."""
    ws = wb.create_sheet("P&L FY27")
    style_title(ws, "Profit and loss  -  FY27 by department",
                "Click the plus and minus buttons on the left to open a department down to subcategory detail. "
                "Months up to the reporting month are actual. Months after it are budget or rolling forecast.")
    ws["A3"] = "=\"Reporting month: \"&TEXT(Setup!$B$5,\"mmm yyyy\")&\"   |   shaded columns are forecast\""
    ws["A3"].font = Font(name=FONT, size=10, bold=True, color="C00000")

    last_month_col = 1 + len(MONTHS)
    ws.cell(row=5, column=1, value="Line").font = HEAD
    ws.cell(row=5, column=1).fill = HEAD_FILL
    for i, (y, m) in enumerate(MONTHS):
        c = ws.cell(row=5, column=2 + i, value=date(y, m, 1))
        c.font = HEAD
        c.fill = HEAD_FILL
        c.number_format = "mmm-yy"
    tot_col = 2 + len(MONTHS)
    c = ws.cell(row=5, column=tot_col, value="FY total")
    c.font = HEAD
    c.fill = HEAD_FILL

    ws.column_dimensions["A"].width = 44
    for i in range(len(MONTHS) + 1):
        ws.column_dimensions[get_column_letter(2 + i)].width = 13
    ws.freeze_panes = "B6"
    ws.sheet_view.showGridLines = False

    def money_row(r, bold=False, fill=None):
        for i in range(len(MONTHS) + 1):
            c = ws.cell(row=r, column=2 + i)
            c.number_format = MONEY
            c.font = BOLD if bold else BODY
            if fill:
                c.fill = fill

    def value_formula(i, eng_row):
        """P&L month i sits in column B+i; the same month on the Engine is column E+i."""
        pl_col = get_column_letter(2 + i)
        eng_col = get_column_letter(5 + i)
        return (f"=IF({pl_col}$5<=Setup!$B$5,Engine!${eng_col}{eng_row},"
                f"Budget_Paste!${eng_col}{eng_row})")

    r = 6
    for dept in DEPTS:
        dept_row = r
        ws.cell(row=r, column=1, value=dept).font = Font(name=FONT, size=11, bold=True, color="FFFFFF")
        for i in range(len(MONTHS) + 2):
            ws.cell(row=r, column=1 + i).fill = HEAD_FILL
        money_row(r, bold=True)
        r += 1
        cat_rows = {}
        for cat in CATEGORY_ORDER:
            cat_subs = [s for c, s in subs if c == cat]
            if not cat_subs:
                continue
            cat_row = r
            ws.cell(row=r, column=1, value="    " + cat).font = BOLD
            for i in range(len(MONTHS) + 2):
                ws.cell(row=r, column=1 + i).fill = BAND_FILL
            money_row(r, bold=True, fill=BAND_FILL)
            r += 1
            first_sub = r
            for sub in cat_subs:
                ws.cell(row=r, column=1, value="        " + sub).font = BODY
                eng_row = engine_index[(dept, sub)]
                for i in range(len(MONTHS)):
                    ws.cell(row=r, column=2 + i, value=value_formula(i, eng_row))
                ws.cell(row=r, column=tot_col,
                        value=f"=SUM(B{r}:{get_column_letter(1 + len(MONTHS))}{r})")
                money_row(r)
                ws.row_dimensions[r].outlineLevel = 2
                ws.row_dimensions[r].hidden = True
                r += 1
            last_sub = r - 1
            for i in range(len(MONTHS) + 1):
                col = get_column_letter(2 + i)
                ws.cell(row=cat_row, column=2 + i,
                        value=f"=SUM({col}{first_sub}:{col}{last_sub})")
            money_row(cat_row, bold=True, fill=BAND_FILL)
            ws.row_dimensions[cat_row].outlineLevel = 1
            cat_rows[cat] = cat_row

        gp_row = r
        ws.cell(row=r, column=1, value="    Gross profit").font = BOLD
        for i in range(len(MONTHS) + 1):
            col = get_column_letter(2 + i)
            ws.cell(row=r, column=2 + i,
                    value=f"={col}{cat_rows['Income']}+{col}{cat_rows['Cost of Sales']}")
        money_row(r, bold=True, fill=GREY_FILL)
        ws.cell(row=r, column=1).fill = GREY_FILL
        ws.row_dimensions[r].outlineLevel = 1
        r += 1

        ws.cell(row=r, column=1, value="    Gross margin %").font = BODY
        for i in range(len(MONTHS) + 1):
            col = get_column_letter(2 + i)
            ws.cell(row=r, column=2 + i,
                    value=f"=IFERROR({col}{gp_row}/{col}{cat_rows['Income']},0)").number_format = PCT
        ws.row_dimensions[r].outlineLevel = 1
        r += 1

        np_row = r
        ws.cell(row=r, column=1, value="    Net profit").font = BOLD
        parts = "+".join(f"{{c}}{cat_rows[c2]}" for c2 in CATEGORY_ORDER if c2 in cat_rows)
        for i in range(len(MONTHS) + 1):
            col = get_column_letter(2 + i)
            ws.cell(row=r, column=2 + i, value="=" + parts.replace("{c}", col))
        money_row(r, bold=True, fill=GREY_FILL)
        ws.cell(row=r, column=1).fill = GREY_FILL
        ws.row_dimensions[r].outlineLevel = 1
        r += 1

        # The department header row carries its own net profit.
        for i in range(len(MONTHS) + 1):
            col = get_column_letter(2 + i)
            ws.cell(row=dept_row, column=2 + i, value=f"={col}{np_row}")
        money_row(dept_row, bold=True)
        for i in range(len(MONTHS) + 2):
            ws.cell(row=dept_row, column=1 + i).fill = HEAD_FILL
            if i:
                ws.cell(row=dept_row, column=1 + i).font = Font(name=FONT, size=10, bold=True, color="FFFFFF")
        r += 1
    ws.sheet_properties.outlinePr.summaryBelow = False
    return ws


def build_prior_year(wb, subs):
    """Last year, same shape as the Engine, so the year on year block works.
    Pasted once a year from the closed FY."""
    ws = wb.create_sheet("Prior_Year")
    style_title(ws, "Prior year  -  type or paste here once a year",
                "Last financial year by department and subcategory. Same signs as the P&L. "
                "Feeds the year on year block on the Summary.")
    for i, h in enumerate(["Key", "Department", "Category", "Subcategory"]):
        c = ws.cell(row=4, column=1 + i, value=h)
        c.font = HEAD
        c.fill = HEAD_FILL
    for i, (y, m) in enumerate(MONTHS):
        c = ws.cell(row=4, column=5 + i, value=date(y - 1, m, 1))
        c.font = HEAD
        c.fill = HEAD_FILL
        c.number_format = "mmm-yy"
    ws.cell(row=4, column=17, value="FY total").font = HEAD
    ws.cell(row=4, column=17).fill = HEAD_FILL
    ws.column_dimensions["A"].width = 34
    for col, w in (("B", 15), ("C", 15), ("D", 32)):
        ws.column_dimensions[col].width = w
    for i in range(len(MONTHS) + 1):
        ws.column_dimensions[get_column_letter(5 + i)].width = 12
    ws.freeze_panes = "E5"
    ws.sheet_view.showGridLines = False
    r = 5
    for dept in DEPTS:
        for cat, sub in subs:
            ws.cell(row=r, column=1, value=f"{dept}|{sub}").font = BODY
            ws.cell(row=r, column=2, value=dept).font = BODY
            ws.cell(row=r, column=3, value=cat).font = BODY
            ws.cell(row=r, column=4, value=sub).font = BODY
            for i in range(len(MONTHS)):
                c = ws.cell(row=r, column=5 + i)
                c.number_format = MONEY
                c.fill = YELLOW_FILL
                c.font = INPUT_FONT
            c = ws.cell(row=r, column=17,
                        value=f"=SUM(E{r}:{get_column_letter(4 + len(MONTHS))}{r})")
            c.number_format = MONEY
            c.font = BOLD
            r += 1
    return ws


def build_summary(wb, subs, engine_index):
    """Three rolling months, then year to date, quarter and year on year,
    laid out the way she sketched it and styled like the P&L."""
    ws = wb.create_sheet("Summary")
    style_title(ws, "Summary",
                "Three rolling months, then year to date, the quarter and year on year. "
                "Click the plus and minus buttons to open a department into its categories.")
    ws["A3"] = ('=\"Reporting month: \"&TEXT(Setup!$B$5,\"mmm yyyy\")'
                '&\"   |   period \"&Setup!$B$8&\" of 12\"')
    ws["A3"].font = Font(name=FONT, size=10, bold=True, color="C00000")

    # Column A, then her blocks. Blank spacer columns keep the blocks apart.
    heads = [
        ("A", "Line", 42), ("B", None, 13), ("C", None, 13), ("D", None, 13),
        ("E", "YTD", 13), ("F", "YTD vs forecast", 14), ("G", "YTD vs budget", 14),
        ("H", "Var $", 13), ("I", "Var %", 10), ("J", "", 3),
        ("K", "Quarter", 13), ("L", "Quarter budget", 14), ("M", "Var $", 13), ("N", "Var %", 10),
        ("O", "", 3),
        ("P", "Last year YTD", 14), ("Q", "This year YTD", 14), ("R", "Var $", 13), ("S", "Var %", 10),
    ]
    for col, label, width in heads:
        ws.column_dimensions[col].width = width
        if label is not None:
            c = ws[f"{col}5"]
            c.value = label
            c.font = HEAD
            c.fill = HEAD_FILL
            c.alignment = Alignment(horizontal="center", wrap_text=True)
    # The three rolling months are formulas off the reporting month.
    for col, offset in (("B", -2), ("C", -1), ("D", 0)):
        c = ws[f"{col}5"]
        c.value = f"=EDATE(Setup!$B$5,{offset})"
        c.number_format = "mmm yyyy"
        c.font = HEAD
        c.fill = HEAD_FILL
        c.alignment = Alignment(horizontal="center")
    ws["B4"] = "Rolling three months"
    ws["E4"] = "Year to date"
    ws["K4"] = "Quarter"
    ws["P4"] = "Year on year"
    for col in ("B", "E", "K", "P"):
        ws[f"{col}4"].font = Font(name=FONT, size=10, bold=True, color=NAVY)
    ws.freeze_panes = "B6"
    ws.sheet_view.showGridLines = False

    first, last = 5, 4 + len(DEPTS) * len(subs)
    hdr = f"{{s}}!$E$4:$P$4"
    body = f"{{s}}!$E${first}:$P${last}"
    deptcol = f"{{s}}!$B${first}:$B${last}"
    catcol = f"{{s}}!$C${first}:$C${last}"
    qstart = "DATE(YEAR(Setup!$B$5),FLOOR(MONTH(Setup!$B$5)-1,3)+1,1)"
    fystart = "DATE(Setup!$B$7-1,7,1)"

    def measure(sheet, dept, cat, which, month_expr=None):
        m = []
        if dept:
            m.append(f"({deptcol.format(s=sheet)}=\"{dept}\")")
        if cat:
            m.append(f"({catcol.format(s=sheet)}=\"{cat}\")")
        h = hdr.format(s=sheet)
        if which == "month":
            m.append(f"({h}={month_expr})")
        elif which == "ytd":
            m.append(f"({h}>={fystart})")
            m.append(f"({h}<=Setup!$B$5)")
        elif which == "qtr":
            m.append(f"({h}>={qstart})")
            m.append(f"({h}<=Setup!$B$5)")
        elif which == "ytd_ly":
            m.append(f"({h}>=DATE(Setup!$B$7-2,7,1))")
            m.append(f"({h}<=EDATE(Setup!$B$5,-12))")
        m.append(body.format(s=sheet))
        return "=SUMPRODUCT(" + "*".join(m) + ")"

    def variance(r, a, b, var, pct):
        ws[f"{var}{r}"] = f"={a}{r}-{b}{r}"
        ws[f"{pct}{r}"] = f"=IFERROR(({a}{r}-{b}{r})/ABS({b}{r}),0)"

    def write_row(r, label, dept, cat, bold=False, fill=None, white=False):
        c = ws.cell(row=r, column=1, value=("        " if cat else "") + label)
        c.font = Font(name=FONT, size=10, bold=bold, color="FFFFFF" if white else "000000")
        cells = {}
        # rolling months
        for col, off in (("B", -2), ("C", -1), ("D", 0)):
            cells[col] = measure("Engine", dept, cat, "month", f"EDATE(Setup!$B$5,{off})")
        cells["E"] = measure("Engine", dept, cat, "ytd")
        cells["F"] = measure("Budget_Paste", dept, cat, "ytd")
        cells["G"] = measure("Budget_Paste", dept, cat, "ytd")
        cells["K"] = measure("Engine", dept, cat, "qtr")
        cells["L"] = measure("Budget_Paste", dept, cat, "qtr")
        cells["P"] = measure("Prior_Year", dept, cat, "ytd_ly")
        cells["Q"] = f"=$E{r}"
        for col, f in cells.items():
            ws[f"{col}{r}"] = f
        variance(r, "E", "G", "H", "I")
        variance(r, "K", "L", "M", "N")
        variance(r, "Q", "P", "R", "S")
        for col, _, _ in heads:
            if col == "A":
                continue
            cc = ws[f"{col}{r}"]
            cc.number_format = PCT if col in ("I", "N", "S") else MONEY
            cc.font = Font(name=FONT, size=10, bold=bold,
                           color="FFFFFF" if white else "000000")
            if fill:
                cc.fill = fill

    r = 6
    for dept in DEPTS:
        write_row(r, dept, dept, None, bold=True, fill=HEAD_FILL, white=True)
        r += 1
        for cat in CATEGORY_ORDER:
            if not any(c == cat for c, _ in subs):
                continue
            write_row(r, cat, dept, cat, fill=BAND_FILL)
            ws.row_dimensions[r].outlineLevel = 1
            ws.row_dimensions[r].hidden = True
            r += 1
        r += 1
    write_row(r, "COMPANY TOTAL", None, None, bold=True, fill=GREY_FILL)
    ws.sheet_properties.outlinePr.summaryBelow = False
    return ws


def build_utilisation(wb):
    ws = wb.create_sheet("Utilisation")
    style_title(ws, "Utilisation and working hours",
                "Working hours are calculated from the calendar. Type the chargeable and non-chargeable hours per department.")
    ws.column_dimensions["A"].width = 18
    for col in "BCDEFGHI":
        ws.column_dimensions[col].width = 16

    heads = ["Month", "Working days", "Working hours", "Department",
             "Chargeable hrs", "Non-chargeable hrs", "Leave hrs",
             "Utilisation %", "FTE"]
    for i, h in enumerate(heads):
        c = ws.cell(row=5, column=1 + i, value=h)
        c.font = HEAD
        c.fill = HEAD_FILL
    ws.freeze_panes = "A6"
    ws.sheet_view.showGridLines = False

    r = 6
    for (y, m) in MONTHS:
        for dept in DEPTS:
            ws.cell(row=r, column=1, value=date(y, m, 1)).number_format = "mmm yyyy"
            ws.cell(row=r, column=2,
                    value=f"=NETWORKDAYS($A{r},EOMONTH($A{r},0),Holidays)").number_format = "#,##0"
            ws.cell(row=r, column=3, value=f"=$B{r}*Setup!$B$11").number_format = "#,##0"
            ws.cell(row=r, column=4, value=dept).font = BODY
            for col in (5, 6, 7):
                c = ws.cell(row=r, column=col)
                c.fill = YELLOW_FILL
                c.font = INPUT_FONT
                c.number_format = "#,##0.0"
            ws.cell(row=r, column=8,
                    value=f"=IFERROR($E{r}/($E{r}+$F{r}),0)").number_format = PCT
            ws.cell(row=r, column=9,
                    value=f"=IFERROR(($E{r}+$F{r})/$C{r},0)").number_format = "0.00"
            for col in (1, 2, 3, 8, 9):
                ws.cell(row=r, column=col).font = BODY
            r += 1
    return ws


def build_pl_check(wb):
    """Paste the Xero P&L summary here. It reconciles the GL paste to the P&L."""
    ws = wb.create_sheet("PL_Check")
    style_title(ws, "P&L control  -  paste the Xero profit and loss here",
                "Run the same profit and loss you already run and paste it at A8. "
                "The check below proves the GL paste agrees with the P&L, per category and in total.")
    ws.column_dimensions["A"].width = 46
    for col in "BCDE":
        ws.column_dimensions[col].width = 16
    ws.sheet_view.showGridLines = False

    ws["A4"] = "Category"
    ws["B4"] = "Per the GL paste"
    ws["C4"] = "Per the Xero P&L"
    ws["D4"] = "Difference"
    ws["E4"] = "Status"
    for col in "ABCDE":
        ws[f"{col}4"].font = HEAD
        ws[f"{col}4"].fill = HEAD_FILL

    last = 4 + GL_ROWS
    for i, cat in enumerate(CATEGORY_ORDER):
        r = 5 + i
        ws.cell(row=r, column=1, value=cat).font = BODY
        ws.cell(row=r, column=2,
                value=(f'=SUMIFS(Cleanup!$A$5:$A${last},Cleanup!$E$5:$E${last},$A{r},'
                       f'Cleanup!$B$5:$B${last},Setup!$B$5)')).number_format = MONEY
        c = ws.cell(row=r, column=3)
        c.fill = YELLOW_FILL
        c.font = INPUT_FONT
        c.number_format = MONEY
        ws.cell(row=r, column=4, value=f"=$B{r}-$C{r}").number_format = MONEY
        ws.cell(row=r, column=5,
                value=f'=IF(ABS($D{r})<1,"OK","CHECK")').font = BOLD

    ws["A11"] = "Paste the Xero profit and loss below this line, exactly as it comes."
    ws["A11"].font = Font(name=FONT, size=9, italic=True, color="595959")
    return ws


def build_readme(wb):
    ws = wb.create_sheet("README", 0)
    style_title(ws, "CTS Financial Controller Pack",
                "One workbook. Two pastes a month, then read.")
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 96

    steps = [
        ("EACH MONTH", ""),
        ("1.  Setup B5", "Pick the reporting month from the dropdown. The financial year, the period number and the comparative month all follow from it."),
        ("2.  GL_Paste", "Run the Account transactions for P&L analysis report in Xero, year to date from 1 July, and paste it at A2. Columns must stay in the order shown in row 1."),
        ("3.  PL_Check", "Paste the Xero profit and loss and enter the category totals in the yellow cells. Every line should read OK before you go further."),
        ("4.  Budget_Paste", "Once a year, and again whenever the rolling forecast changes."),
        ("5.  Utilisation", "Type the chargeable and non-chargeable hours per department. Working hours calculate themselves."),
        ("6.  Read", "P&L FY27, then Summary."),
        ("", ""),
        ("HOW THE DEPARTMENT IS DECIDED", ""),
        ("Per transaction", "From the Cost Centres column, which carries the department name outright. On the August export all 3,771 lines were tagged, so nothing fell through. If it is ever blank the pack falls back to the bracket tag on the Job Numbers column."),
        ("Why not the account", "Three quarters of the P&L sits in accounts with no department in the name, Contract Support Staff and Equipment Hires among them. Only the transaction knows which department earned or spent it."),
        ("Unmatched lines", "Anything with no tag lands in UNALLOCATED and is counted on Setup. Clear it to zero before you report."),
        ("", ""),
        ("THINGS THAT WILL CATCH YOU OUT", ""),
        ("Sign convention", "Amount is Credit less Debit. Income is positive, costs are negative, so gross profit is Income plus Cost of Sales rather than minus."),
        ("Forecast months", "On the P&L, months up to the reporting month come from the GL. Months after it come from Budget_Paste, so the year always shows a full twelve."),
        ("Adding an account", "A new Xero account must be added to Lists with its category and subcategory, or its lines show as UNKNOWN."),
        ("Capacity", "GL_Paste holds 20,000 lines, about a full financial year at your volume. Setup tells you how many you have used."),
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
            ws.row_dimensions[r].height = 30
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
    build_budget_paste(wb, subs)
    build_prior_year(wb, subs)
    build_cleanup(wb)
    _, engine_index, engine_last = build_engine(wb, subs)
    build_pl(wb, subs, engine_index)
    build_summary(wb, subs, engine_index)
    build_utilisation(wb)
    build_readme(wb)

    for name, ref in defs.items():
        wb.defined_names.add(__import__("openpyxl").workbook.defined_name.DefinedName(name, attr_text=ref))

    yellow_tabs = {"Setup", "GL_Paste", "PL_Check", "Budget_Paste", "Prior_Year", "Utilisation"}
    for ws in wb.worksheets:
        ws.sheet_properties.tabColor = YELLOW if ws.title in yellow_tabs else NAVY

    out = "/home/user/Claude/pack/CTS Financial Controller Pack.xlsx"
    wb.save(out)
    print("saved", out)
    print("engine rows", engine_last)


if __name__ == "__main__":
    main()
