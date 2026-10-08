"""Build the payroll tax annual reconciliation workbook for the other states.

Two pastes, everything else by formula. No pivot tables.
"""
import datetime, os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as cl
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule

OUT = "/home/user/Claude/ptax/out/Payroll Tax Annual Reconciliation FY2027.xlsx"
ROWS = 30002          # last row of the earnings formula range
EROWS = 1002          # last row of the employees formula range
FY_START = datetime.date(2026, 7, 1)
FY_END = datetime.date(2027, 6, 30)
STATES = ["ACT", "NSW", "QLD", "SA", "VIC", "WA"]
OTHER = ["ACT", "QLD", "SA", "VIC", "WA"]

FONT = "Aptos Narrow"
NAVY = "1F3864"
BAND = "D9E2F3"
PASTE = "FFF2CC"
INPUT_ = "FFF2CC"
GREY = "F2F2F2"
ORANGE = "FF6600"

H1 = Font(name=FONT, size=16, bold=True, color=NAVY)
H2 = Font(name=FONT, size=12, bold=True, color=NAVY)
BODY = Font(name=FONT, size=11)
BOLD = Font(name=FONT, size=11, bold=True)
WHITEB = Font(name=FONT, size=11, bold=True, color="FFFFFF")
SMALL = Font(name=FONT, size=9, italic=True, color="595959")
HEADFILL = PatternFill("solid", fgColor=NAVY)
BANDFILL = PatternFill("solid", fgColor=BAND)
PASTEFILL = PatternFill("solid", fgColor=PASTE)
GREYFILL = PatternFill("solid", fgColor=GREY)
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
MONEY = '#,##0.00'
MONEY0 = '#,##0'
PCT = '0.00%'
DATE = 'dd/mm/yyyy'

MONTHS = []
y, m = FY_START.year, FY_START.month
for _ in range(12):
    first = datetime.date(y, m, 1)
    nm, ny = (m + 1, y) if m < 12 else (1, y + 1)
    last = datetime.date(ny, nm, 1) - datetime.timedelta(days=1)
    MONTHS.append((first, last))
    m, y = nm, ny


def put(ws, ref, value, font=BODY, fmt=None, fill=None, align=None, wrap=False, border=False):
    c = ws[ref]
    c.value = value
    c.font = font
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill
    if align or wrap:
        c.alignment = Alignment(horizontal=align, wrap_text=wrap, vertical="top" if wrap else None)
    if border:
        c.border = BOX
    return c


def header_row(ws, row, labels, start=1, fill=HEADFILL, font=WHITEB):
    for i, t in enumerate(labels):
        c = ws.cell(row=row, column=start + i, value=t)
        c.font = font
        c.fill = fill
        c.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[row].height = 30


def widths(ws, spec):
    for col, w in spec.items():
        ws.column_dimensions[col].width = w


EARN_COLS = ["Pay Process No", "Pay Date", "Period End Date", "Employee Id",
             "Employee External Id", "Employee Name", "Pay Category Id",
             "Pay Category External Id", "Pay Category Name", "Units", "Unit Type",
             "Location Id", "Location External Id", "Location Name", "Notes",
             "Rate", "Rate Type", "Gross Earnings", "Taxable Earnings", "SG Super"]
# A..T above. Helper columns:
C_STATE, C_COUNTS = "V", "W"
C_PAYDATE, C_EMPID, C_CAT, C_GROSS, C_SUPER = "B", "D", "I", "R", "T"

EMP_COLS = ["Employee Id", "First name", "Surname", "Job title", "Start date",
            "Gender", "Employment Type", "Residential State", "Residential Location",
            "Primary location"]
E_ID, E_STATE, E_KEY = "A", "H", "Y"

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
    ("Per Diems", "No", "Not taxable", "Exempt. Gi deducts per diems every month."),
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

# From payrolltax.gov.au and the state revenue offices, read on 8 October 2026.
# Threshold is the maximum yearly deduction. Taper from and taper to are the
# Australian wage levels between which that deduction runs down to nil.
RATES = {
    "ACT": (1750000, None, None, 0.0675,
            "6.75% while Australian wages stay under 20 million. No taper.",
            "payrolltax.gov.au and ACT Revenue Office"),
    "NSW": (1200000, None, None, 0.0545,
            "Lodged monthly, not here. It sits in this table so the Australian "
            "wages split is right.",
            "payrolltax.gov.au"),
    "QLD": (1300000, 1300000, 10400000, 0.0475,
            "Deduction falls by 1 dollar for every 7 dollars of Australian wages "
            "over 1.3 million, and is nil at 10.4 million. Rate goes to 4.95% "
            "over 6.5 million.",
            "Queensland Revenue Office"),
    "SA": (600000, None, None, 0.0495,
           "The most you can deduct is 600,000, not the 1.5 million threshold. "
           "The rate shades in between 1.5 and 1.7 million of Australian wages "
           "and is 4.95% above that.",
           "payrolltax.gov.au"),
    "VIC": (1000000, 3000000, 5000000, 0.0485,
            "Deduction runs down between 3 million and 5 million of Australian wages.",
            "payrolltax.gov.au and SRO Victoria"),
    "WA": (1000000, 1000000, 7500000, 0.055,
           "Threshold falls by 2 dollars for every 13 dollars of Australian wages "
           "over 1 million, and is nil at 7.5 million.",
           "payrolltax.gov.au and RevenueWA"),
}


E = "'1 Earnings'!"
def er(c):
    return "%s$%s$3:$%s$%d" % (E, c, c, ROWS)
R_STATE, R_COUNTS, R_LINE = er(C_STATE), er(C_COUNTS), er("X")
R_DATE, R_GROSS, R_SUPER = er(C_PAYDATE), er(C_GROSS), er(C_SUPER)
C_LINE = "X"


def sheet_start(wb):
    ws = wb.create_sheet("Start here")
    widths(ws, {"A": 4, "B": 46, "C": 20, "D": 20, "E": 20, "F": 20, "G": 20, "H": 20})
    put(ws, "B1", "Payroll tax: annual reconciliation for the other states", H1)
    put(ws, "B2", "Corporate Technology Services  .  1 July 2026 to 30 June 2027", SMALL)

    put(ws, "B4", "What this is for", H2)
    for i, t in enumerate([
        "NSW is lodged every month, so NSW is not reconciled here.",
        "This works out the year's payroll tax for ACT, QLD, SA, VIC and WA,",
        "and tells you what is still to pay after what you have already accrued.",
        "There are no pivot tables. Two pastes and the rest is formulas.",
    ]):
        put(ws, "B%d" % (5 + i), t)

    put(ws, "B10", "What you do", H2)
    steps = [
        ("Step 1", "Run the earnings report in Employment Hero for the whole year. "
                   "Paste it on the tab called 1 Earnings."),
        ("Step 2", "Run the employee details report. Paste it on the tab called 2 Employees."),
        ("Step 3", "Look at Pay categories. Anything new shows up in red. "
                   "Say whether it counts for payroll tax."),
        ("Step 4", "Look at Rates. Check the five numbers against the revenue offices once a year."),
        ("Step 5", "Type what you have already accrued on the Accrued tab."),
        ("Step 6", "Read the answer on the Annual tab."),
    ]
    r = 11
    for label, text in steps:
        put(ws, "B%d" % r, label, BOLD)
        put(ws, "C%d" % r, text, BODY)
        ws.merge_cells("C%d:H%d" % (r, r))
        r += 1

    put(ws, "B19", "Checks", H2)
    put(ws, "C19", "These all have to be clear before you believe the answer.", SMALL)
    checks = [
        ("Lines pasted on 1 Earnings", "=COUNT(%s)" % R_DATE, MONEY0),
        ("Employees pasted on 2 Employees", "=COUNTA('2 Employees'!$A$3:$A$%d)" % EROWS, MONEY0),
        ("Lines where the state could not be found", '=COUNTIF(%s,"** NOT FOUND **")' % R_STATE, MONEY0),
        ("Pay categories not on the list yet", '=COUNTIF(%s,"** ADD TO LIST **")' % R_COUNTS, MONEY0),
        ("Lines dated outside this financial year",
         '=COUNTIFS(%s,"<"&$C$33)+COUNTIFS(%s,">"&$C$34)' % (R_DATE, R_DATE), MONEY0),
        ("Gross earnings pasted, all lines", "=SUM(%s)" % R_GROSS, MONEY),
        ("Gross earnings that count for payroll tax",
         '=SUMIFS(%s,%s,"Yes")' % (R_GROSS, R_COUNTS), MONEY),
        ("Super pasted, all lines", "=SUM(%s)" % R_SUPER, MONEY),
        ("Earliest pay date in the paste", "=MIN(%s)" % R_DATE, DATE),
        ("Latest pay date in the paste", "=MAX(%s)" % R_DATE, DATE),
        ("Pay date column holds dates",
         '=IF(COUNT(%s)=COUNTA(%s),"ok","** CHECK THE COLUMNS **")' % (R_DATE, R_DATE), None),
        ("Gross earnings column holds numbers",
         '=IF(COUNT(%s)=COUNTA(%s),"ok","** CHECK THE COLUMNS **")' % (R_GROSS, R_GROSS), None),
    ]
    r = 20
    for label, f, fmt in checks:
        put(ws, "B%d" % r, label)
        put(ws, "C%d" % r, f, BOLD, fmt)
        r += 1
    put(ws, "B33", "Year starts", BOLD)
    put(ws, "C33", FY_START, BOLD, DATE, fill=PASTEFILL)
    put(ws, "B34", "Year ends", BOLD)
    put(ws, "C34", FY_END, BOLD, DATE, fill=PASTEFILL)

    put(ws, "B36", "The answer", H2)
    put(ws, "C35", "This file comes to you with July to October already pasted in, so you "
                   "can see it working. Until the whole year is in, a full year's deduction "
                   "is being set against part of a year's wages, so the tax reads low or nil. "
                   "That is expected. Repaste both tabs after 30 June and it is right.",
        SMALL)
    ws.merge_cells("C35:H35")
    ws.row_dimensions[35].height = 30
    ws["C35"].alignment = Alignment(wrap_text=True, vertical="center")
    header_row(ws, 37, ["", "State", "Payroll tax for the year", "Already accrued",
                        "Still to pay"], start=1)
    for i, s in enumerate(OTHER):
        r = 38 + i
        put(ws, "B%d" % r, s, BOLD, border=True)
        put(ws, "C%d" % r, "=Annual!L%d" % (6 + STATES.index(s)), None or BODY, MONEY, border=True)
        put(ws, "D%d" % r, "=Annual!M%d" % (6 + STATES.index(s)), BODY, MONEY, border=True)
        put(ws, "E%d" % r, "=Annual!N%d" % (6 + STATES.index(s)), BOLD, MONEY, border=True)
    r = 38 + len(OTHER)
    put(ws, "B%d" % r, "Total", BOLD, fill=BANDFILL, border=True)
    for col in "CDE":
        put(ws, "%s%d" % (col, r), "=SUM(%s38:%s%d)" % (col, col, r - 1), BOLD, MONEY,
            fill=BANDFILL, border=True)

    put(ws, "B%d" % (r + 2), "Where the figures come from", H2)
    for i, t in enumerate([
        "Wages and super come from the earnings report you paste, nothing else.",
        "The state is the employee's residential state, off the employee details report.",
        "That is the basis Gi uses, and it is the basis the Employment Hero payroll tax "
        "report uses when you group it by residential location.",
        "Per diems are the only thing taken out, because they are exempt.",
        "Employer contributed super is not in the earnings report. Type it on the Annual tab "
        "if there is any. For the other states it has been nil.",
    ]):
        put(ws, "B%d" % (r + 3 + i), t)
        ws.merge_cells("B%d:H%d" % (r + 3 + i, r + 3 + i))
    ws.sheet_view.showGridLines = False
    return ws


def sheet_earnings(wb):
    ws = wb.create_sheet("1 Earnings")
    put(ws, "A1", "Paste the Employment Hero earnings report here, data only, starting "
                  "in cell A3. Do not paste its heading rows. The three columns in blue "
                  "fill themselves in.", BOLD)
    ws.merge_cells("A1:T1")
    ws.row_dimensions[1].height = 30
    ws["A1"].alignment = Alignment(wrap_text=True, vertical="center")
    header_row(ws, 2, EARN_COLS)
    header_row(ws, 2, ["State", "Counts?", "Return line"], start=22,
               fill=PatternFill("solid", fgColor="2E75B6"))
    for c in range(1, 21):
        ws.column_dimensions[cl(c)].width = 15
    widths(ws, {"B": 12, "F": 22, "I": 30, "N": 34, "U": 3, "V": 14, "W": 12, "X": 24})
    for r in range(3, ROWS + 1):
        put(ws, "V%d" % r,
            '=IF($D%d="","",IFERROR(INDEX(\'2 Employees\'!$H$3:$H$%d,'
            'MATCH(IFERROR(VALUE($D%d),$D%d),\'2 Employees\'!$Y$3:$Y$%d,0)),"** NOT FOUND **"))'
            % (r, EROWS, r, r, EROWS))
        put(ws, "W%d" % r,
            '=IF($I%d="","",IFERROR(VLOOKUP($I%d,\'Pay categories\'!$A:$C,2,FALSE),'
            '"** ADD TO LIST **"))' % (r, r))
        put(ws, "X%d" % r,
            '=IF($I%d="","",IFERROR(VLOOKUP($I%d,\'Pay categories\'!$A:$C,3,FALSE),'
            '"** ADD TO LIST **"))' % (r, r))
    red = CellIsRule(operator="equal", formula=['"** NOT FOUND **"'],
                     fill=PatternFill("solid", fgColor="FFC7CE"),
                     font=Font(name=FONT, size=11, bold=True, color="9C0006"))
    red2 = CellIsRule(operator="equal", formula=['"** ADD TO LIST **"'],
                      fill=PatternFill("solid", fgColor="FFC7CE"),
                      font=Font(name=FONT, size=11, bold=True, color="9C0006"))
    ws.conditional_formatting.add("V3:V%d" % ROWS, red)
    ws.conditional_formatting.add("W3:X%d" % ROWS, red2)
    ws.freeze_panes = "A3"
    return ws


def sheet_employees(wb):
    ws = wb.create_sheet("2 Employees")
    put(ws, "A1", "Paste the Employment Hero employee details report here, data only, "
                  "starting in cell A3. Do not paste its heading rows.", BOLD)
    ws.merge_cells("A1:J1")
    ws.row_dimensions[1].height = 30
    ws["A1"].alignment = Alignment(wrap_text=True, vertical="center")
    header_row(ws, 2, EMP_COLS)
    header_row(ws, 2, ["Id as a number"], start=25,
               fill=PatternFill("solid", fgColor="2E75B6"))
    for c in range(1, 11):
        ws.column_dimensions[cl(c)].width = 18
    widths(ws, {"D": 30, "H": 16, "J": 30, "X": 3, "Y": 16})
    for r in range(3, EROWS + 1):
        put(ws, "Y%d" % r, '=IF($A%d="","",IFERROR(VALUE($A%d),$A%d))' % (r, r, r))
    ws.freeze_panes = "A3"
    return ws


def sheet_categories(wb):
    ws = wb.create_sheet("Pay categories")
    put(ws, "A1", "Does this pay category count for payroll tax?", H1)
    put(ws, "A2", "Anything the earnings report has that is not on this list shows as "
                  "** ADD TO LIST ** over on 1 Earnings, and the count shows on Start here. "
                  "Add it here and it fixes itself.", SMALL)
    ws.merge_cells("A2:D2")
    ws.row_dimensions[2].height = 28
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="center")
    header_row(ws, 4, ["Pay category", "Counts?", "Return line", "Note"])
    widths(ws, {"A": 62, "B": 12, "C": 26, "D": 70})
    for i, (name, counts, line, note) in enumerate(CATEGORIES):
        r = 5 + i
        put(ws, "A%d" % r, name, border=True)
        put(ws, "B%d" % r, counts, BOLD, fill=INPUT_ and PASTEFILL, align="center", border=True)
        put(ws, "C%d" % r, line, fill=PASTEFILL, border=True)
        put(ws, "D%d" % r, note, SMALL, border=True)
        ws["D%d" % r].alignment = Alignment(wrap_text=True, vertical="top")
    last = 4 + len(CATEGORIES) + 40
    dv1 = DataValidation(type="list", formula1='"Yes,No"', allow_blank=True)
    dv2 = DataValidation(type="list", formula1='"%s"' % ",".join(LINES), allow_blank=True)
    ws.add_data_validation(dv1)
    ws.add_data_validation(dv2)
    dv1.add("B5:B%d" % last)
    dv2.add("C5:C%d" % last)
    ws.freeze_panes = "A5"
    return ws


def sheet_rates(wb):
    ws = wb.create_sheet("Rates")
    put(ws, "A1", "Rates and thresholds", H1)
    put(ws, "A2", "Check these once a year against each revenue office. They are the only "
                  "numbers in this file that come from outside your own data. "
                  "I took them from a web search on 8 October 2026, not from a letter or "
                  "a notice, so have Pramila confirm them before you lodge.", SMALL)
    ws.merge_cells("A2:I2")
    ws.row_dimensions[2].height = 42
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="center")
    header_row(ws, 4, ["State", "Most you can deduct for the year", "Taper starts at",
                       "Taper ends at", "Rate", "Taper factor", "Deduction for the year",
                       "What this means", "Source", "Checked on"])
    widths(ws, {"A": 8, "B": 16, "C": 14, "D": 14, "E": 10, "F": 12, "G": 16,
                "H": 62, "I": 34, "J": 14})
    for i, s in enumerate(STATES):
        r = 5 + i
        th, tf, tt, rate, note, src = RATES[s]
        put(ws, "A%d" % r, s, BOLD, border=True)
        put(ws, "B%d" % r, th, fmt=MONEY0, fill=PASTEFILL, border=True)
        put(ws, "C%d" % r, tf, fmt=MONEY0, fill=PASTEFILL, border=True)
        put(ws, "D%d" % r, tt, fmt=MONEY0, fill=PASTEFILL, border=True)
        put(ws, "E%d" % r, rate, fmt=PCT, fill=PASTEFILL, border=True)
        put(ws, "F%d" % r,
            '=IF(OR($C%d="",$D%d=""),1,MAX(0,MIN(1,($D%d-Annual!$E$12)/($D%d-$C%d))))'
            % (r, r, r, r, r), fmt='0.0000', fill=GREYFILL, border=True)
        put(ws, "G%d" % r, "=$B%d*$F%d" % (r, r), BOLD, MONEY0, fill=GREYFILL, border=True)
        put(ws, "H%d" % r, note, SMALL, border=True)
        ws["H%d" % r].alignment = Alignment(wrap_text=True, vertical="top")
        put(ws, "I%d" % r, src, SMALL, border=True)
        ws["I%d" % r].alignment = Alignment(wrap_text=True, vertical="top")
        put(ws, "J%d" % r, datetime.date(2026, 10, 8), SMALL, DATE, fill=PASTEFILL, border=True)
        ws.row_dimensions[r].height = 46
    put(ws, "A12", "Yellow cells are the ones you change. Grey cells work themselves out.", SMALL)
    ws.sheet_view.showGridLines = False
    return ws


def sheet_monthly(wb):
    ws = wb.create_sheet("Monthly")
    put(ws, "A1", "Month by month", H1)
    put(ws, "A2", "Wages that count, then super, then the two added together. "
                  "Nothing to fill in here.", SMALL)
    blocks = [("Wages that count for payroll tax", 4, "wages"),
              ("Super", 20, "super"),
              ("Total taxable wages", 36, "total")]
    widths(ws, {"A": 14, "B": 13, "C": 13})
    for c in range(4, 11):
        ws.column_dimensions[cl(c)].width = 14
    for title, top, kind in blocks:
        put(ws, "A%d" % top, title, H2)
        header_row(ws, top + 1, ["Month", "From", "To"] + STATES + ["Total"])
        for i, (d1, d2) in enumerate(MONTHS):
            r = top + 2 + i
            put(ws, "A%d" % r, d1.strftime("%b %Y"), BOLD, border=True)
            put(ws, "B%d" % r, d1, SMALL, DATE, border=True)
            put(ws, "C%d" % r, d2, SMALL, DATE, border=True)
            for j, s in enumerate(STATES):
                col = cl(4 + j)
                if kind == "wages":
                    f = ('=SUMIFS(%s,%s,$%s$%d,%s,"Yes",%s,">="&$B%d,%s,"<="&$C%d)'
                         % (R_GROSS, R_STATE, col, top + 1, R_COUNTS, R_DATE, r, R_DATE, r))
                elif kind == "super":
                    f = ('=SUMIFS(%s,%s,$%s$%d,%s,">="&$B%d,%s,"<="&$C%d)'
                         % (R_SUPER, R_STATE, col, top + 1, R_DATE, r, R_DATE, r))
                else:
                    f = "=%s%d+%s%d" % (col, 6 + i, col, 22 + i)
                put(ws, "%s%d" % (col, r), f, fmt=MONEY, border=True)
            put(ws, "K%d" % r, "=SUM(D%d:I%d)" % (r, r), BOLD, MONEY, border=True)
        tr = top + 14
        put(ws, "A%d" % tr, "Year", BOLD, fill=BANDFILL, border=True)
        put(ws, "B%d" % tr, "", fill=BANDFILL, border=True)
        put(ws, "C%d" % tr, "", fill=BANDFILL, border=True)
        for j in range(len(STATES) + 1):
            col = cl(4 + j)
            put(ws, "%s%d" % (col, tr), "=SUM(%s%d:%s%d)" % (col, top + 2, col, tr - 1),
                BOLD, MONEY, fill=BANDFILL, border=True)
    ws.sheet_view.showGridLines = False
    return ws


def sheet_accrued(wb):
    ws = wb.create_sheet("Accrued")
    put(ws, "A1", "What you have already accrued or paid", H1)
    put(ws, "A2", "Type what you actually posted each month. It is filled in with the "
                  "standard monthly accrual as a starting point. Overtype anything that "
                  "is different.", SMALL)
    ws.merge_cells("A2:H2")
    ws.row_dimensions[2].height = 28
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="center")
    header_row(ws, 4, ["Month"] + OTHER + ["Total"])
    widths(ws, {"A": 14, "B": 13, "C": 13, "D": 13, "E": 13, "F": 13, "G": 13})
    standard = {"ACT": 0, "QLD": 500, "SA": 200, "VIC": 1300, "WA": 700}
    for i, (d1, _d2) in enumerate(MONTHS):
        r = 5 + i
        put(ws, "A%d" % r, d1.strftime("%b %Y"), BOLD, border=True)
        for j, s in enumerate(OTHER):
            put(ws, "%s%d" % (cl(2 + j), r), standard[s], fmt=MONEY,
                fill=PASTEFILL, border=True)
        put(ws, "G%d" % r, "=SUM(B%d:F%d)" % (r, r), BOLD, MONEY, border=True)
    put(ws, "A17", "Year", BOLD, fill=BANDFILL, border=True)
    for j in range(len(OTHER) + 1):
        col = cl(2 + j)
        put(ws, "%s17" % col, "=SUM(%s5:%s16)" % (col, col), BOLD, MONEY,
            fill=BANDFILL, border=True)
    ws.sheet_view.showGridLines = False
    return ws


D1, D2 = "'Start here'!$C$33", "'Start here'!$C$34"
MW, MS = 18, 34          # the year rows of the wages and super blocks on Monthly


def sheet_annual(wb):
    ws = wb.create_sheet("Annual")
    put(ws, "A1", "The annual reconciliation", H1)
    put(ws, "A2", "Yellow cells are the only ones you type in. Everything else is a formula.",
        SMALL)
    header_row(ws, 5, [
        "State", "Wages that count", "Super", "Employer contributed super",
        "Total taxable wages", "Share of Australian wages", "Days liable",
        "Deduction for the year", "Deduction entitlement", "Taxable amount", "Rate",
        "Payroll tax for the year", "Already accrued", "Still to pay"])
    widths(ws, {"A": 8, "B": 15, "C": 13, "D": 15, "E": 16, "F": 13, "G": 10, "H": 15,
                "I": 15, "J": 15, "K": 9, "L": 15, "M": 14, "N": 14})
    for i, s in enumerate(STATES):
        r = 6 + i
        mcol = cl(4 + i)
        put(ws, "A%d" % r, s, BOLD, border=True)
        put(ws, "B%d" % r, "=Monthly!%s%d" % (mcol, MW), fmt=MONEY, border=True)
        put(ws, "C%d" % r, "=Monthly!%s%d" % (mcol, MS), fmt=MONEY, border=True)
        put(ws, "D%d" % r, 0, fmt=MONEY, fill=PASTEFILL, border=True)
        put(ws, "E%d" % r, "=B%d+C%d+D%d" % (r, r, r), BOLD, MONEY, border=True)
        put(ws, "F%d" % r, '=IF($E$12=0,0,E%d/$E$12)' % r, fmt=PCT, border=True)
        put(ws, "G%d" % r, 365, fmt=MONEY0, fill=PASTEFILL, border=True)
        put(ws, "H%d" % r, "=Rates!G%d" % (5 + i), fmt=MONEY0, border=True)
        put(ws, "I%d" % r, "=H%d*F%d*G%d/365" % (r, r, r), fmt=MONEY, border=True)
        put(ws, "J%d" % r, "=MAX(0,E%d-I%d)" % (r, r), fmt=MONEY, border=True)
        put(ws, "K%d" % r, "=Rates!E%d" % (5 + i), fmt=PCT, border=True)
        put(ws, "L%d" % r, "=J%d*K%d" % (r, r), BOLD, MONEY, border=True)
        if s in OTHER:
            put(ws, "M%d" % r, "=Accrued!%s17" % cl(2 + OTHER.index(s)), fmt=MONEY, border=True)
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
    put(ws, "A13", "Column E, total, is your total Australian wages for the year. "
                   "The deduction each state gives you is worked out off it.", SMALL)
    ws.merge_cells("A13:N13")
    put(ws, "A14", "Days liable is 365 for a full year. Change it only if the business "
                   "was not employing for the whole year. The pay dates you have pasted "
                   "run from", SMALL)
    ws.merge_cells("A14:I14")
    put(ws, "J14", "=MIN(%s)" % R_DATE, SMALL, DATE)
    put(ws, "K14", "to", SMALL)
    put(ws, "L14", "=MAX(%s)" % R_DATE, SMALL, DATE)

    put(ws, "A16", "Breakdown for the return forms", H2)
    put(ws, "A17", "The forms ask for these separately. Super is the same figure as "
                   "column C above.", SMALL)
    ws.merge_cells("A17:H17")
    cols = ["Salaries and wages", "Termination payments", "Bonuses and commissions",
            "Allowances"]
    header_row(ws, 19, ["State"] + cols + ["Super", "Total", "Not taxable, for information"])
    for i, s in enumerate(STATES):
        r = 20 + i
        put(ws, "A%d" % r, s, BOLD, border=True)
        for j, line in enumerate(cols):
            col = cl(2 + j)
            put(ws, "%s%d" % (col, r),
                '=SUMIFS(%s,%s,$A%d,%s,%s$19,%s,">="&%s,%s,"<="&%s)'
                % (R_GROSS, R_STATE, r, R_LINE, col, R_DATE, D1, R_DATE, D2),
                fmt=MONEY, border=True)
        put(ws, "F%d" % r, "=C%d" % (6 + i), fmt=MONEY, border=True)
        put(ws, "G%d" % r, "=SUM(B%d:F%d)" % (r, r), BOLD, MONEY, border=True)
        put(ws, "H%d" % r,
            '=SUMIFS(%s,%s,$A%d,%s,"Not taxable",%s,">="&%s,%s,"<="&%s)'
            % (R_GROSS, R_STATE, r, R_LINE, R_DATE, D1, R_DATE, D2),
            SMALL, MONEY, border=True)
    put(ws, "A26", "Total", BOLD, fill=BANDFILL, border=True)
    for j in range(7):
        col = cl(2 + j)
        put(ws, "%s26" % col, "=SUM(%s20:%s25)" % (col, col), BOLD, MONEY,
            fill=BANDFILL, border=True)
    put(ws, "A28", "Column G total here should equal column E total above. "
                   "If it does not, a pay category has no return line against it.", SMALL)
    ws.merge_cells("A28:H28")
    put(ws, "A29", "Difference", BOLD)
    put(ws, "B29", "=G26-E12", BOLD, MONEY)
    ws.sheet_view.showGridLines = False
    return ws


def main():
    wb = Workbook()
    wb.remove(wb.active)
    sheet_start(wb)
    sheet_earnings(wb)
    sheet_employees(wb)
    sheet_categories(wb)
    sheet_rates(wb)
    sheet_monthly(wb)
    sheet_accrued(wb)
    sheet_annual(wb)
    colours = {"Start here": "1F3864", "1 Earnings": "FFC000", "2 Employees": "FFC000",
               "Pay categories": "2E75B6", "Rates": "2E75B6", "Monthly": "A9D08E",
               "Accrued": "FFC000", "Annual": "FF6600"}
    for name, col in colours.items():
        wb[name].sheet_properties.tabColor = col
    wb.active = 0
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    wb.save(OUT)
    print("wrote", OUT, "%.1f KB" % (os.path.getsize(OUT) / 1024))


if __name__ == "__main__":
    main()
