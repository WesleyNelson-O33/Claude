"""The Employment Hero split file, rebuilt simply.

Three pastes, then every figure is a formula. No pivot tables, no VLOOKUP
chains across files, nothing to refresh.

It does the same jobs as the old split file: it holds the year's Employment
Hero data, it puts every employee in a state the way Gi does (residential
state), it reconciles back to the Employment Hero dashboard, and it tells you
what to type into the payroll tax workbook. The annual reconciliation for the
other states is on the last tab.
"""
import datetime, os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as cl
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule

OUT = "/home/user/Claude/ptax/out/EH Payroll Split FY2027.xlsx"
PR_ROWS, EA_ROWS, EM_ROWS = 5002, 25002, 1002
FY_START, FY_END = datetime.date(2026, 7, 1), datetime.date(2027, 6, 30)
STATES = ["ACT", "NSW", "QLD", "SA", "VIC", "WA"]
OTHER = ["ACT", "QLD", "SA", "VIC", "WA"]
MISSING = "** NOT FOUND **"

FONT = "Aptos Narrow"
NAVY, BLUE, BAND, YEL, GREY = "1F3864", "2E75B6", "D9E2F3", "FFF2CC", "F2F2F2"
H1 = Font(name=FONT, size=16, bold=True, color=NAVY)
H2 = Font(name=FONT, size=12, bold=True, color=NAVY)
BODY = Font(name=FONT, size=11)
BOLD = Font(name=FONT, size=11, bold=True)
WHITEB = Font(name=FONT, size=11, bold=True, color="FFFFFF")
SMALL = Font(name=FONT, size=9, italic=True, color="595959")
HEADFILL, BANDFILL = PatternFill("solid", fgColor=NAVY), PatternFill("solid", fgColor=BAND)
YELFILL, GREYFILL = PatternFill("solid", fgColor=YEL), PatternFill("solid", fgColor=GREY)
BLUEFILL = PatternFill("solid", fgColor=BLUE)
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
MONEY, MONEY0, PCT, DATE, HRS = '#,##0.00', '#,##0', '0.00%', 'dd/mm/yyyy', '#,##0.00'
MMM = 'mmm yyyy'

MONTHS = []
y, m = FY_START.year, FY_START.month
for _ in range(12):
    MONTHS.append(datetime.date(y, m, 1))
    m, y = (m + 1, y) if m < 12 else (1, y + 1)


def put(ws, ref, value, font=BODY, fmt=None, fill=None, wrap=False, border=False, align=None):
    c = ws[ref]
    c.value = value
    c.font = font
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill
    if wrap or align:
        c.alignment = Alignment(wrap_text=wrap, vertical="top" if wrap else None,
                                horizontal=align)
    if border:
        c.border = BOX
    return c


def head(ws, row, labels, start=1, fill=HEADFILL, h=30):
    for i, t in enumerate(labels):
        c = ws.cell(row=row, column=start + i, value=t)
        c.font = WHITEB
        c.fill = fill
        c.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[row].height = h


def widths(ws, spec):
    for k, v in spec.items():
        ws.column_dimensions[k].width = v


# ---------------------------------------------------------------- paste layouts
PR_COLS = ["Pay Process No", "Pay Date", "Period End Date", "Employee Id",
           "Employee First Name", "Employee Surname", "Total Hours", "Gross Earnings",
           "Pre-Tax Deduction", "Taxable Earnings", "Post-Tax Deduction", "PAYG",
           "PAYG %", "STSL", "SG Super", "Employer Contribution Super",
           "Member Voluntary Super", "Salary Sacrifice Super", "Net Earnings"]
EA_COLS = ["Pay Process No", "Pay Date", "Period End Date", "Employee Id",
           "Employee External Id", "Employee Name", "Pay Category Id",
           "Pay Category External Id", "Pay Category Name", "Units", "Unit Type",
           "Location Id", "Location External Id", "Location Name", "Notes", "Rate",
           "Rate Type", "Gross Earnings", "Taxable Earnings", "SG Super"]
EM_COLS = ["Employee Id", "First name", "Surname", "Job title", "Start date", "Gender",
           "Employment Type", "Residential State", "Residential Location",
           "Primary location"]

P = "'1 Pay run totals'!"
E = "'2 Earnings'!"
def pr(c):
    return "%s$%s$3:$%s$%d" % (P, c, c, PR_ROWS)
def ea(c):
    return "%s$%s$3:$%s$%d" % (E, c, c, EA_ROWS)

PR_STATE, PR_MONTH = pr("U"), pr("V")
PR_GROSS, PR_SG, PR_EC, PR_NET = pr("H"), pr("O"), pr("P"), pr("S")
PR_PAYG, PR_STSL, PR_PRE, PR_POST, PR_SS = pr("L"), pr("N"), pr("I"), pr("K"), pr("R")
EA_STATE, EA_COUNTS, EA_MONTH, EA_LINE = ea("V"), ea("W"), ea("X"), ea("Y")
EA_GROSS, EA_SG, EA_UNITS, EA_CAT = ea("R"), ea("T"), ea("J"), ea("I")
PICK = "'Start here'!$C$6"


CATEGORIES = [
    ("Annual Leave Taken", "Yes", "Salaries and wages", ""),
    ("Back Payment", "Yes", "Salaries and wages", ""),
    ("Bonus", "Yes", "Bonuses and commissions", ""),
    ("Bonus Leave Taken", "Yes", "Salaries and wages", ""),
    ("Casual - Overtime x 1.75", "Yes", "Salaries and wages", ""),
    ("Casual - Overtime x 2", "Yes", "Salaries and wages", ""),
    ("Casual - Public Holiday", "Yes", "Salaries and wages", ""),
    ("Casual - Sunday", "Yes", "Salaries and wages", ""),
    ("Casual No Meal Break", "Yes", "Salaries and wages", ""),
    ("Casual Ordinary Hours", "Yes", "Salaries and wages", ""),
    ("Casual working through meal break", "Yes", "Salaries and wages", ""),
    ("Compassionate Leave Taken", "Yes", "Salaries and wages", ""),
    ("Meal allowance while travelling - not local shows - less than 5 working days",
     "Yes", "Allowances",
     "Check this one. An allowance can be exempt up to the ATO rate. The payroll tax "
     "report from Employment Hero lists the exempt categories at the bottom."),
    ("Per Diems", "No", "Not taxable", "Exempt. This is the one Gi takes out every month."),
    ("Permanent - Overtime x 1.5", "Yes", "Salaries and wages", ""),
    ("Permanent - Overtime x1.25", "Yes", "Salaries and wages", ""),
    ("Permanent - Overtime x1.5", "Yes", "Salaries and wages", ""),
    ("Permanent - Overtime x2", "Yes", "Salaries and wages", ""),
    ("Permanent - Public Holiday not worked", "Yes", "Salaries and wages", ""),
    ("Permanent Ordinary Hours", "Yes", "Salaries and wages", ""),
    ("Permanent Travelling Time", "Yes", "Salaries and wages", ""),
    ("Personal/Carer's Leave Taken", "Yes", "Salaries and wages", ""),
    ("Time in lieu taken", "Yes", "Salaries and wages", ""),
    ("Unused Leave Payment (Type O)", "Yes", "Termination payments", ""),
    ("Unused leave payment (normal termination)", "Yes", "Termination payments", ""),
]
LINES = ["Salaries and wages", "Termination payments", "Bonuses and commissions",
         "Allowances", "Not taxable"]

RATES = {
    "ACT": (1750000, None, None, 0.0675,
            "6.75% while Australian wages stay under 20 million. No taper.",
            "payrolltax.gov.au and ACT Revenue Office"),
    "NSW": (1200000, None, None, 0.0545,
            "Lodged monthly, not reconciled here. It sits in the table so the "
            "Australian wages split is right.", "payrolltax.gov.au"),
    "QLD": (1300000, 1300000, 10400000, 0.0475,
            "Deduction falls by 1 dollar for every 7 dollars of Australian wages over "
            "1.3 million, nil at 10.4 million. Rate goes to 4.95% over 6.5 million.",
            "Queensland Revenue Office"),
    "SA": (600000, None, None, 0.0495,
           "The most you can deduct is 600,000, not the 1.5 million threshold. The rate "
           "shades in between 1.5 and 1.7 million of Australian wages, 4.95% above.",
           "payrolltax.gov.au"),
    "VIC": (1000000, 3000000, 5000000, 0.0485,
            "Deduction runs down between 3 million and 5 million of Australian wages.",
            "payrolltax.gov.au and SRO Victoria"),
    "WA": (1000000, 1000000, 7500000, 0.055,
           "Threshold falls by 2 dollars for every 13 dollars of Australian wages over "
           "1 million, nil at 7.5 million.", "payrolltax.gov.au and RevenueWA"),
}


def paste_header(ws, title, cols, helpers):
    put(ws, "A1", title, BOLD)
    ws.merge_cells("A1:%s1" % cl(min(len(cols), 20)))
    ws.row_dimensions[1].height = 34
    ws["A1"].alignment = Alignment(wrap_text=True, vertical="center")
    head(ws, 2, cols)
    for c in range(1, len(cols) + 1):
        ws.column_dimensions[cl(c)].width = 15
    if helpers:
        head(ws, 2, [lbl for _c, lbl, _f, _fm in helpers],
             start=helpers[0][0], fill=BLUEFILL)
    ws.freeze_panes = "A3"


def fill_helpers(ws, helpers, last):
    for r in range(3, last + 1):
        for col, _lbl, formula, fmt in helpers:
            put(ws, "%s%d" % (cl(col), r), formula % {"r": r}, fmt=fmt)


def sheet_payruns(wb):
    ws = wb.create_sheet("1 Pay run totals")
    helpers = [
        (21, "State", '=IF($D%(r)d="","",IFERROR(INDEX(\'3 Employees\'!$H$3:$H$'
                      + str(EM_ROWS) + ',MATCH(IFERROR(VALUE($D%(r)d),$D%(r)d),'
                      "'3 Employees'!$Y$3:$Y$" + str(EM_ROWS) + ',0)),"' + MISSING + '"))', None),
        (22, "Month", '=IF($B%(r)d="","",DATE(YEAR($B%(r)d),MONTH($B%(r)d),1))', MMM),
    ]
    paste_header(ws, "Paste the Employment Hero pay run totals report here, data only, "
                     "starting in cell A3. Do not paste its heading rows. This is the "
                     "one that carries net pay, PAYG, super and the deductions.",
                 PR_COLS, helpers)
    fill_helpers(ws, helpers, PR_ROWS)
    widths(ws, {"B": 12, "E": 18, "F": 18, "T": 3, "U": 14, "V": 12})
    ws.conditional_formatting.add("U3:U%d" % PR_ROWS, CellIsRule(
        operator="equal", formula=['"%s"' % MISSING],
        fill=PatternFill("solid", fgColor="FFC7CE"),
        font=Font(name=FONT, size=11, bold=True, color="9C0006")))
    return ws


def sheet_earnings(wb):
    ws = wb.create_sheet("2 Earnings")
    helpers = [
        (22, "State", '=IF($D%(r)d="","",IFERROR(INDEX(\'3 Employees\'!$H$3:$H$'
                      + str(EM_ROWS) + ',MATCH(IFERROR(VALUE($D%(r)d),$D%(r)d),'
                      "'3 Employees'!$Y$3:$Y$" + str(EM_ROWS) + ',0)),"' + MISSING + '"))', None),
        (23, "Counts?", '=IF($I%(r)d="","",IFERROR(VLOOKUP($I%(r)d,\'Pay categories\'!$A:$C,'
                        '2,FALSE),"** ADD TO LIST **"))', None),
        (24, "Month", '=IF($B%(r)d="","",DATE(YEAR($B%(r)d),MONTH($B%(r)d),1))', MMM),
        (25, "Return line", '=IF($I%(r)d="","",IFERROR(VLOOKUP($I%(r)d,\'Pay categories\'!'
                            '$A:$C,3,FALSE),"** ADD TO LIST **"))', None),
    ]
    paste_header(ws, "Paste the Employment Hero earnings report here, data only, starting "
                     "in cell A3. Do not paste its heading rows. This is the one with the "
                     "pay categories, the hours and the job numbers on it.",
                 EA_COLS, helpers)
    fill_helpers(ws, helpers, EA_ROWS)
    head(ws, 2, ["Employee Name", "Location External Id", "Location Name",
                 "Job Name and Job Number", "Units", "State"], start=27,
         fill=PatternFill("solid", fgColor="A9D08E"))
    put(ws, "AA1", "Filter column X to the month you want, then copy these six columns "
                   "into the timesheets tab of the payroll tax workbook.", BOLD, wrap=True)
    ws.merge_cells("AA1:AF1")
    for r in range(3, EA_ROWS + 1):
        put(ws, "AA%d" % r, '=IF($F%d="","",$F%d)' % (r, r))
        put(ws, "AB%d" % r, '=IF($F%d="","",$M%d)' % (r, r))
        put(ws, "AC%d" % r, '=IF($F%d="","",$N%d)' % (r, r))
        put(ws, "AD%d" % r, '=IF($F%d="","",$M%d&" - "&$N%d)' % (r, r, r))
        put(ws, "AE%d" % r, '=IF($F%d="","",$J%d)' % (r, r), fmt=HRS)
        put(ws, "AF%d" % r, '=IF($F%d="","",$V%d)' % (r, r))
    widths(ws, {"B": 12, "F": 24, "I": 30, "N": 32, "U": 3, "V": 14, "W": 11, "X": 12,
                "Y": 22, "Z": 3, "AA": 24, "AB": 18, "AC": 32, "AD": 40, "AE": 10, "AF": 10})
    red = CellIsRule(operator="equal", formula=['"** ADD TO LIST **"'],
                     fill=PatternFill("solid", fgColor="FFC7CE"),
                     font=Font(name=FONT, size=11, bold=True, color="9C0006"))
    ws.conditional_formatting.add("W3:W%d" % EA_ROWS, red)
    ws.conditional_formatting.add("Y3:Y%d" % EA_ROWS, red)
    ws.conditional_formatting.add("V3:V%d" % EA_ROWS, CellIsRule(
        operator="equal", formula=['"%s"' % MISSING],
        fill=PatternFill("solid", fgColor="FFC7CE"),
        font=Font(name=FONT, size=11, bold=True, color="9C0006")))
    ws.auto_filter.ref = "A2:Y%d" % EA_ROWS
    return ws


def sheet_employees(wb):
    ws = wb.create_sheet("3 Employees")
    paste_header(ws, "Paste the Employment Hero employee details report here, data only, "
                     "starting in cell A3. Do not paste its heading rows. This is what "
                     "puts every person in a state.", EM_COLS, [])
    head(ws, 2, ["Id as a number"], start=25, fill=BLUEFILL)
    for r in range(3, EM_ROWS + 1):
        put(ws, "Y%d" % r, '=IF($A%d="","",IFERROR(VALUE($A%d),$A%d))' % (r, r, r))
    widths(ws, {"A": 16, "B": 16, "C": 18, "D": 30, "H": 16, "J": 30, "X": 3, "Y": 16})
    return ws


def sheet_categories(wb):
    ws = wb.create_sheet("Pay categories")
    put(ws, "A1", "Does this pay category count for payroll tax?", H1)
    put(ws, "A2", "Anything in the earnings report that is not on this list turns red over "
                  "on 2 Earnings, and the count shows on Start here. Add it here and it "
                  "fixes itself.", SMALL, wrap=True)
    ws.merge_cells("A2:D2")
    ws.row_dimensions[2].height = 28
    head(ws, 4, ["Pay category", "Counts?", "Return line", "Note"])
    widths(ws, {"A": 62, "B": 12, "C": 26, "D": 68})
    for i, (name, counts, line, note) in enumerate(CATEGORIES):
        r = 5 + i
        put(ws, "A%d" % r, name, border=True)
        put(ws, "B%d" % r, counts, BOLD, fill=YELFILL, border=True, align="center")
        put(ws, "C%d" % r, line, fill=YELFILL, border=True)
        put(ws, "D%d" % r, note, SMALL, border=True, wrap=True)
    last = 4 + len(CATEGORIES) + 40
    d1 = DataValidation(type="list", formula1='"Yes,No"', allow_blank=True)
    d2 = DataValidation(type="list", formula1='"%s"' % ",".join(LINES), allow_blank=True)
    ws.add_data_validation(d1)
    ws.add_data_validation(d2)
    d1.add("B5:B%d" % last)
    d2.add("C5:C%d" % last)
    ws.freeze_panes = "A5"
    return ws


def sheet_rates(wb):
    ws = wb.create_sheet("Rates")
    put(ws, "A1", "Rates and thresholds", H1)
    put(ws, "A2", "These are the only numbers in this file that do not come from your own "
                  "data. I took them from a web search on 8 October 2026, not from a notice "
                  "or a letter. Have Pramila confirm them before you lodge, and check them "
                  "once a year.", SMALL, wrap=True)
    ws.merge_cells("A2:I2")
    ws.row_dimensions[2].height = 42
    head(ws, 4, ["State", "Most you can deduct for the year", "Taper starts at",
                 "Taper ends at", "Rate", "Taper factor", "Deduction for the year",
                 "What this means", "Source", "Checked on"])
    widths(ws, {"A": 8, "B": 16, "C": 14, "D": 14, "E": 10, "F": 12, "G": 16, "H": 62,
                "I": 34, "J": 14})
    for i, s in enumerate(STATES):
        r = 5 + i
        th, tf, tt, rate, note, src = RATES[s]
        put(ws, "A%d" % r, s, BOLD, border=True)
        put(ws, "B%d" % r, th, fmt=MONEY0, fill=YELFILL, border=True)
        put(ws, "C%d" % r, tf, fmt=MONEY0, fill=YELFILL, border=True)
        put(ws, "D%d" % r, tt, fmt=MONEY0, fill=YELFILL, border=True)
        put(ws, "E%d" % r, rate, fmt=PCT, fill=YELFILL, border=True)
        put(ws, "F%d" % r, '=IF(OR($C%d="",$D%d=""),1,MAX(0,MIN(1,($D%d-\'Annual other '
                           'states\'!$E$12)/($D%d-$C%d))))' % (r, r, r, r, r),
            fmt='0.0000', fill=GREYFILL, border=True)
        put(ws, "G%d" % r, "=$B%d*$F%d" % (r, r), BOLD, MONEY0, fill=GREYFILL, border=True)
        put(ws, "H%d" % r, note, SMALL, border=True, wrap=True)
        put(ws, "I%d" % r, src, SMALL, border=True, wrap=True)
        put(ws, "J%d" % r, datetime.date(2026, 10, 8), SMALL, DATE, fill=YELFILL, border=True)
        ws.row_dimensions[r].height = 46
    put(ws, "A12", "Yellow cells are yours to change. Grey cells work themselves out.", SMALL)
    ws.sheet_view.showGridLines = False
    return ws


YBM_BLOCKS = [("Wages that count for payroll tax", 4, "wages"),
              ("SG super", 20, "sg"),
              ("Employer contribution super", 36, "ec"),
              ("Total for payroll tax", 52, "total")]
YBM_YEAR = {"wages": 18, "sg": 34, "ec": 50, "total": 66}


def sheet_year(wb):
    ws = wb.create_sheet("Year by month")
    put(ws, "A1", "Every month, every state", H1)
    put(ws, "A2", "Nothing to fill in here. It all comes off the two pastes.", SMALL)
    widths(ws, {"A": 14, "B": 13, "C": 13, "K": 15})
    for c in range(4, 11):
        ws.column_dimensions[cl(c)].width = 14
    for title, top, kind in YBM_BLOCKS:
        put(ws, "A%d" % top, title, H2)
        head(ws, top + 1, ["Month", "From", "To"] + STATES + [MISSING, "Total"])
        for i, d1 in enumerate(MONTHS):
            r = top + 2 + i
            put(ws, "A%d" % r, d1, BOLD, MMM, border=True)
            put(ws, "B%d" % r, d1, SMALL, DATE, border=True)
            put(ws, "C%d" % r, "=EOMONTH($B%d,0)" % r, SMALL, DATE, border=True)
            for j in range(len(STATES) + 1):
                col = cl(4 + j)
                sref = "$%s$%d" % (col, top + 1)
                if kind == "wages":
                    f = ('=SUMIFS(%s,%s,%s,%s,$B%d,%s,"Yes")'
                         % (EA_GROSS, EA_STATE, sref, EA_MONTH, r, EA_COUNTS))
                elif kind == "sg":
                    f = "=SUMIFS(%s,%s,%s,%s,$B%d)" % (PR_SG, PR_STATE, sref, PR_MONTH, r)
                elif kind == "ec":
                    f = "=SUMIFS(%s,%s,%s,%s,$B%d)" % (PR_EC, PR_STATE, sref, PR_MONTH, r)
                else:
                    f = "=%s%d+%s%d+%s%d" % (col, 6 + i, col, 22 + i, col, 38 + i)
                put(ws, "%s%d" % (col, r), f, fmt=MONEY, border=True)
            put(ws, "L%d" % r, "=SUM(D%d:J%d)" % (r, r), BOLD, MONEY, border=True)
        tr = top + 14
        put(ws, "A%d" % tr, "Year", BOLD, fill=BANDFILL, border=True)
        for col in "BC":
            put(ws, "%s%d" % (col, tr), "", fill=BANDFILL, border=True)
        for j in range(len(STATES) + 2):
            col = cl(4 + j)
            put(ws, "%s%d" % (col, tr), "=SUM(%s%d:%s%d)" % (col, top + 2, col, tr - 1),
                BOLD, MONEY, fill=BANDFILL, border=True)
    ws.sheet_view.showGridLines = False
    return ws


def sheet_month(wb):
    ws = wb.create_sheet("Month summary")
    put(ws, "A1", "The month", H1)
    put(ws, "A3", "Month", BOLD)
    put(ws, "B3", "=%s" % PICK, H2, MMM)
    put(ws, "D3", "Change the month on the Start here tab.", SMALL)
    put(ws, "A4", "Goes in column", BOLD)
    put(ws, "B4", "=CHAR(64+IF(MONTH(" + PICK + ")>=7,MONTH(" + PICK
        + ")-5,MONTH(" + PICK + ")+7))", H2, align="left")
    put(ws, "D4", "of the payroll tax workbook, on the Main and NSW tab and on each state tab.", SMALL)

    put(ws, "A5", "Wages and super by state", H2)
    head(ws, 6, ["State", "Gross earnings", "Taken out, exempt",
                 "Wages for payroll tax", "SG super", "Employer contribution super",
                 "Total for payroll tax", "Hours"])
    widths(ws, {"A": 16, "B": 15, "C": 15, "D": 16, "E": 13, "F": 16, "G": 16, "H": 12})
    rowstates = STATES + [MISSING]
    for i, s in enumerate(rowstates):
        r = 7 + i
        put(ws, "A%d" % r, s, BOLD, border=True)
        put(ws, "B%d" % r, "=SUMIFS(%s,%s,$A%d,%s,%s)" % (EA_GROSS, EA_STATE, r, EA_MONTH, PICK),
            fmt=MONEY, border=True)
        put(ws, "C%d" % r, '=SUMIFS(%s,%s,$A%d,%s,%s,%s,"No")'
            % (EA_GROSS, EA_STATE, r, EA_MONTH, PICK, EA_COUNTS), fmt=MONEY, border=True)
        put(ws, "D%d" % r, "=B%d-C%d" % (r, r), BOLD, MONEY, border=True)
        put(ws, "E%d" % r, "=SUMIFS(%s,%s,$A%d,%s,%s)" % (PR_SG, PR_STATE, r, PR_MONTH, PICK),
            fmt=MONEY, border=True)
        put(ws, "F%d" % r, "=SUMIFS(%s,%s,$A%d,%s,%s)" % (PR_EC, PR_STATE, r, PR_MONTH, PICK),
            fmt=MONEY, border=True)
        put(ws, "G%d" % r, "=D%d+E%d+F%d" % (r, r, r), BOLD, MONEY, border=True)
        put(ws, "H%d" % r, "=SUMIFS(%s,%s,$A%d,%s,%s)" % (EA_UNITS, EA_STATE, r, EA_MONTH, PICK),
            fmt=HRS, border=True)
    tr = 7 + len(rowstates)
    put(ws, "A%d" % tr, "Total", BOLD, fill=BANDFILL, border=True)
    for col in "BCDEFGH":
        put(ws, "%s%d" % (col, tr), "=SUM(%s7:%s%d)" % (col, col, tr - 1), BOLD,
            MONEY if col != "H" else HRS, fill=BANDFILL, border=True)

    top = tr + 3
    put(ws, "A%d" % top, "What to type into the payroll tax workbook", H2)
    put(ws, "A%d" % (top + 1), "One block a state. These are the lines the return asks for. "
                               "The total on the right is the state's total wages, and it "
                               "matches the table above.", SMALL)
    ws.merge_cells("A%d:H%d" % (top + 1, top + 1))
    head(ws, top + 2, ["State", "Salaries and wages", "Superannuation",
                       "Termination payments", "Bonuses and commissions", "Allowances",
                       "Total", "Agrees?"])
    for i, s in enumerate(rowstates):
        r = top + 3 + i
        a = 7 + i
        put(ws, "A%d" % r, s, BOLD, border=True)
        put(ws, "B%d" % r, "=D%d-D%d-E%d-F%d" % (a, r, r, r), BOLD, MONEY, border=True)
        put(ws, "C%d" % r, "=E%d+F%d" % (a, a), BOLD, MONEY, border=True)
        for j, line in enumerate(["Termination payments", "Bonuses and commissions",
                                  "Allowances"]):
            put(ws, "%s%d" % (cl(4 + j), r),
                '=SUMIFS(%s,%s,$A%d,%s,%s,%s,"%s")'
                % (EA_GROSS, EA_STATE, r, EA_MONTH, PICK, EA_LINE, line),
                fmt=MONEY, border=True)
        put(ws, "G%d" % r, "=SUM(B%d:F%d)" % (r, r), BOLD, MONEY, border=True)
        put(ws, "H%d" % r, '=IF(ROUND(G%d-G%d,2)=0,"yes","** NO **")' % (r, a), border=True)
    tr2 = top + 3 + len(rowstates)
    put(ws, "A%d" % tr2, "Total", BOLD, fill=BANDFILL, border=True)
    for col in "BCDEFG":
        put(ws, "%s%d" % (col, tr2), "=SUM(%s%d:%s%d)" % (col, top + 3, col, tr2 - 1),
            BOLD, MONEY, fill=BANDFILL, border=True)
    put(ws, "H%d" % tr2, "", fill=BANDFILL, border=True)
    ws.sheet_view.showGridLines = False
    return ws


def sheet_reconcile(wb):
    ws = wb.create_sheet("Reconcile to EH")
    put(ws, "A1", "Does this file agree with Employment Hero?", H1)
    put(ws, "A2", "Run the Employment Hero dashboard for the year to date and type its "
                  "figures in the yellow column. Every difference has to be zero before "
                  "you use anything else in this file.", SMALL, wrap=True)
    ws.merge_cells("A2:D2")
    ws.row_dimensions[2].height = 30
    head(ws, 4, ["", "This file", "Employment Hero", "Difference"])
    widths(ws, {"A": 46, "B": 18, "C": 18, "D": 16})
    rows = [
        ("Net earnings", "=SUM(%s)" % PR_NET),
        ("PAYG plus STSL", "=SUM(%s)+SUM(%s)" % (PR_PAYG, PR_STSL)),
        ("SG super", "=SUM(%s)" % PR_SG),
        ("Pre-tax plus post-tax deductions", "=SUM(%s)+SUM(%s)" % (PR_PRE, PR_POST)),
        ("Employer contribution super, not shown on the dashboard", "=SUM(%s)" % PR_EC),
    ]
    for i, (label, f) in enumerate(rows):
        r = 5 + i
        put(ws, "A%d" % r, label, BOLD, border=True)
        put(ws, "B%d" % r, f, fmt=MONEY, border=True)
        if i < 4:
            put(ws, "C%d" % r, 0, fmt=MONEY, fill=YELFILL, border=True)
            put(ws, "D%d" % r, "=B%d-C%d" % (r, r), BOLD, MONEY, border=True)
        else:
            put(ws, "C%d" % r, "for information", SMALL, border=True)
            put(ws, "D%d" % r, "", border=True)
    put(ws, "A12", "The two pastes against each other", H2)
    put(ws, "A13", "The pay run totals report and the earnings report should carry the "
                   "same gross and the same SG super. A few cents apart is Employment "
                   "Hero rounding super line by line and does not matter.", SMALL, wrap=True)
    ws.merge_cells("A13:D13")
    ws.row_dimensions[13].height = 30
    head(ws, 14, ["", "Pay run totals", "Earnings report", "Difference"])
    for i, (label, a, b) in enumerate([
            ("Gross earnings", "=SUM(%s)" % PR_GROSS, "=SUM(%s)" % EA_GROSS),
            ("SG super", "=SUM(%s)" % PR_SG, "=SUM(%s)" % EA_SG)]):
        r = 15 + i
        put(ws, "A%d" % r, label, BOLD, border=True)
        put(ws, "B%d" % r, a, fmt=MONEY, border=True)
        put(ws, "C%d" % r, b, fmt=MONEY, border=True)
        put(ws, "D%d" % r, "=B%d-C%d" % (r, r), BOLD, MONEY, border=True)
    ws.sheet_view.showGridLines = False
    return ws


def sheet_annual(wb):
    ws = wb.create_sheet("Annual other states")
    put(ws, "A1", "Annual reconciliation, the other states", H1)
    put(ws, "A2", "NSW is lodged every month so NSW is not reconciled here. It is in the "
                  "table because the deduction each state gives you depends on your total "
                  "Australian wages.", SMALL, wrap=True)
    ws.merge_cells("A2:N2")
    head(ws, 5, ["State", "Wages that count", "SG super", "Employer contribution super",
                 "Total taxable wages", "Share of Australian wages", "Days liable",
                 "Deduction for the year", "Deduction entitlement", "Taxable amount",
                 "Rate", "Payroll tax for the year", "Already accrued", "Still to pay"])
    widths(ws, {"A": 8, "B": 15, "C": 13, "D": 16, "E": 16, "F": 13, "G": 10, "H": 15,
                "I": 15, "J": 15, "K": 9, "L": 15, "M": 14, "N": 14})
    for i, s in enumerate(STATES):
        r = 6 + i
        col = cl(4 + i)
        put(ws, "A%d" % r, s, BOLD, border=True)
        put(ws, "B%d" % r, "='Year by month'!%s%d" % (col, YBM_YEAR["wages"]), fmt=MONEY, border=True)
        put(ws, "C%d" % r, "='Year by month'!%s%d" % (col, YBM_YEAR["sg"]), fmt=MONEY, border=True)
        put(ws, "D%d" % r, "='Year by month'!%s%d" % (col, YBM_YEAR["ec"]), fmt=MONEY, border=True)
        put(ws, "E%d" % r, "=B%d+C%d+D%d" % (r, r, r), BOLD, MONEY, border=True)
        put(ws, "F%d" % r, "=IF($E$12=0,0,E%d/$E$12)" % r, fmt=PCT, border=True)
        put(ws, "G%d" % r, 365, fmt=MONEY0, fill=YELFILL, border=True)
        put(ws, "H%d" % r, "=Rates!G%d" % (5 + i), fmt=MONEY0, border=True)
        put(ws, "I%d" % r, "=H%d*F%d*G%d/365" % (r, r, r), fmt=MONEY, border=True)
        put(ws, "J%d" % r, "=MAX(0,E%d-I%d)" % (r, r), fmt=MONEY, border=True)
        put(ws, "K%d" % r, "=Rates!E%d" % (5 + i), fmt=PCT, border=True)
        put(ws, "L%d" % r, "=J%d*K%d" % (r, r), BOLD, MONEY, border=True)
        if s in OTHER:
            put(ws, "M%d" % r, 0, fmt=MONEY, fill=YELFILL, border=True)
            put(ws, "N%d" % r, "=L%d-M%d" % (r, r), BOLD, MONEY, border=True)
        else:
            put(ws, "M%d" % r, "lodged monthly", SMALL, border=True)
            put(ws, "N%d" % r, "", border=True)
    put(ws, "A12", "Total", BOLD, fill=BANDFILL, border=True)
    for col in list("BCDE") + ["L", "M", "N"]:
        put(ws, "%s12" % col, "=SUM(%s6:%s11)" % (col, col), BOLD, MONEY,
            fill=BANDFILL, border=True)
    for col in "FGHIJK":
        put(ws, "%s12" % col, "", fill=BANDFILL, border=True)
    put(ws, "A14", "Column E total is your total Australian wages for the year. Type what "
                   "you have already accrued for each state in column M. Days liable is "
                   "365 for a full year.", SMALL, wrap=True)
    ws.merge_cells("A14:N14")
    ws.sheet_view.showGridLines = False
    return ws


def sheet_start(wb):
    ws = wb.create_sheet("Start here", 0)
    widths(ws, {"A": 3, "B": 44, "C": 20, "D": 18, "E": 18, "F": 18, "G": 18, "H": 18})
    put(ws, "B1", "Employment Hero payroll split", H1)
    put(ws, "B2", "Corporate Technology Services  .  1 July 2026 to 30 June 2027", SMALL)

    put(ws, "B4", "The month you are working on", H2)
    put(ws, "C6", MONTHS[2], H2, MMM, fill=YELFILL)
    put(ws, "D6", "Pick the month here. The Month summary tab follows it.", SMALL)
    dv = DataValidation(type="list", formula1="='Year by month'!$B$6:$B$17", allow_blank=False)
    ws.add_data_validation(dv)
    dv.add("C6")

    put(ws, "B8", "What you do", H2)
    steps = [
        ("Step 1", "Run the pay run totals report in Employment Hero for the year so far. "
                   "Paste it on 1 Pay run totals."),
        ("Step 2", "Run the earnings report for the year so far. Paste it on 2 Earnings."),
        ("Step 3", "Run the employee details report. Paste it on 3 Employees."),
        ("Step 4", "Clear every check below."),
        ("Step 5", "Reconcile to EH. Type the dashboard figures in and get the differences "
                   "to zero."),
        ("Step 6", "Read Month summary. It tells you what to type into the payroll tax "
                   "workbook."),
        ("Step 7", "For the job number allocation, filter column X on 2 Earnings to your "
                   "month and copy the six green columns."),
    ]
    r = 9
    for label, text in steps:
        put(ws, "B%d" % r, label, BOLD)
        put(ws, "C%d" % r, text)
        ws.merge_cells("C%d:H%d" % (r, r))
        r += 1

    put(ws, "B18", "Checks", H2)
    checks = [
        ("Lines on 1 Pay run totals", "=COUNT(%s)" % pr("B"), MONEY0),
        ("Lines on 2 Earnings", "=COUNT(%s)" % ea("B"), MONEY0),
        ("People on 3 Employees", "=COUNTA('3 Employees'!$A$3:$A$%d)" % EM_ROWS, MONEY0),
        ("Lines where the state could not be found",
         '=COUNTIF(%s,"%s")+COUNTIF(%s,"%s")' % (PR_STATE, MISSING, EA_STATE, MISSING), MONEY0),
        ("Pay categories not on the list yet",
         '=COUNTIF(%s,"** ADD TO LIST **")' % EA_COUNTS, MONEY0),
        ("Lines dated outside this financial year",
         '=COUNTIFS(%s,"<"&$C$30)+COUNTIFS(%s,">"&$C$31)+COUNTIFS(%s,"<"&$C$30)'
         '+COUNTIFS(%s,">"&$C$31)' % (pr("B"), pr("B"), ea("B"), ea("B")), MONEY0),
        ("Gross earnings, the two reports apart",
         "=ROUND(SUM(%s)-SUM(%s),2)" % (PR_GROSS, EA_GROSS), MONEY),
        ("Pay date column holds dates, both pastes",
         '=IF(AND(COUNT(%s)=COUNTA(%s),COUNT(%s)=COUNTA(%s)),"ok","** CHECK THE COLUMNS **")'
         % (pr("B"), pr("B"), ea("B"), ea("B")), None),
    ]
    r = 19
    for label, f, fmt in checks:
        put(ws, "B%d" % r, label)
        put(ws, "C%d" % r, f, BOLD, fmt)
        r += 1
    put(ws, "B30", "Year starts", BOLD)
    put(ws, "C30", FY_START, BOLD, DATE, fill=YELFILL)
    put(ws, "B31", "Year ends", BOLD)
    put(ws, "C31", FY_END, BOLD, DATE, fill=YELFILL)

    put(ws, "B33", "This month, by state", H2)
    head(ws, 34, ["", "State", "Wages for payroll tax", "Super", "Total for payroll tax"])
    for i, s in enumerate(STATES):
        r = 35 + i
        put(ws, "B%d" % r, s, BOLD, border=True)
        put(ws, "C%d" % r, "='Month summary'!D%d" % (7 + i), fmt=MONEY, border=True)
        put(ws, "D%d" % r, "='Month summary'!E%d+'Month summary'!F%d" % (7 + i, 7 + i),
            fmt=MONEY, border=True)
        put(ws, "E%d" % r, "='Month summary'!G%d" % (7 + i), BOLD, MONEY, border=True)
    put(ws, "B41", "Total", BOLD, fill=BANDFILL, border=True)
    for col in "CDE":
        put(ws, "%s41" % col, "=SUM(%s35:%s40)" % (col, col), BOLD, MONEY,
            fill=BANDFILL, border=True)

    put(ws, "B43", "What changed from the old file", H2)
    for i, t in enumerate([
        "There are no pivot tables. Nothing to refresh, nothing to drag.",
        "The state still comes from the employee's residential state, the same as Gi.",
        "Per diems are still taken out, and still only in the summary, not in the data.",
        "Super still comes off the pay run totals report, so it matches what Gi uses.",
        "The old tabs for casuals, overtime and Queensland gross to net are not here. "
        "Say the word and I will add them.",
    ]):
        put(ws, "B%d" % (44 + i), t)
        ws.merge_cells("B%d:H%d" % (44 + i, 44 + i))
    ws.sheet_view.showGridLines = False
    return ws


def main():
    wb = Workbook()
    wb.remove(wb.active)
    sheet_payruns(wb)
    sheet_earnings(wb)
    sheet_employees(wb)
    sheet_categories(wb)
    sheet_rates(wb)
    sheet_year(wb)
    sheet_month(wb)
    sheet_reconcile(wb)
    sheet_annual(wb)
    sheet_start(wb)
    order = ["Start here", "1 Pay run totals", "2 Earnings", "3 Employees",
             "Pay categories", "Rates", "Month summary", "Year by month",
             "Reconcile to EH", "Annual other states"]
    wb._sheets = [wb[n] for n in order]
    colours = {"Start here": NAVY, "1 Pay run totals": "FFC000", "2 Earnings": "FFC000",
               "3 Employees": "FFC000", "Pay categories": BLUE, "Rates": BLUE,
               "Month summary": "FF6600", "Year by month": "A9D08E",
               "Reconcile to EH": "A9D08E", "Annual other states": "FF6600"}
    for n, c in colours.items():
        wb[n].sheet_properties.tabColor = c
    wb.active = 0
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    wb.save(OUT)
    print("wrote", OUT, "%.1f KB" % (os.path.getsize(OUT) / 1024))


if __name__ == "__main__":
    main()
