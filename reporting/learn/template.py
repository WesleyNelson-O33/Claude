"""
A COMPLETE, TINY VERSION OF THE PATTERN.  Read top to bottom - it is short.

Everything I built for you is this same shape, just bigger. Run it, open the
file, then change something and run it again. That loop is how you learn it.

    python3 template.py

The rule the whole thing rests on:
    ONE sheet you paste into.  EVERY other sheet is formulas pointing at it.
Break that rule and you are back to copying numbers between files by hand.

Build order matters, and the order is always:
    1. run this script            -> writes the file with formulas, no answers
    2. run recalc                 -> LibreOffice opens it, works out the answers
    3. run tools/fix_outline.py   -> puts back the +/- button settings that
                                     step 2 silently throws away
Do step 3 second-last or the collapse buttons in the Report sit on the wrong
rows. I lost three workbooks to that before I worked out what was happening.
"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.properties import Outline

# ---------------------------------------------------------------- SETTINGS --
# Change these and re-run. This is the first thing to experiment with.
ACCOUNTS = ["Sales", "Materials", "Wages", "Rent", "Insurance"]
CATEGORY = {"Sales": "Income", "Materials": "Cost of Sales",
            "Wages": "Cost of Sales", "Rent": "Overheads", "Insurance": "Overheads"}
PASTE_ROWS = 2000          # how many rows of pasted data to allow for
OUT = "Example Pack.xlsx"

# ------------------------------------------------------------------ STYLES --
ARIAL = "Arial"
HEAD = PatternFill("solid", fgColor="1F3864")          # dark blue header
INPUT = PatternFill("solid", fgColor="FFF2CC")         # yellow = you type here
TOTAL = PatternFill("solid", fgColor="EDEDED")
BOX = Border(*[Side(style="thin", color="BFBFBF")] * 4)
MONEY = '$#,##0.00;($#,##0.00);"-"'                    # third bit = how zero shows
wb = Workbook()

def header(ws, row, labels):
    for i, t in enumerate(labels, start=1):
        c = ws.cell(row=row, column=i, value=t)
        c.font = Font(name=ARIAL, size=10, bold=True, color="FFFFFF")
        c.fill = HEAD; c.border = BOX
        c.alignment = Alignment(horizontal="center", wrap_text=True)

# ============================================================== 1. THE INPUT =
# Bare on purpose. Excel refuses to paste a whole-sheet selection anywhere
# except A1, so never put a title or a control panel above your paste zone.
data = wb.active
data.title = "Data"
header(data, 1, ["Date", "Account", "Description", "Amount"])
for col, w in zip("ABCD", (12, 20, 40, 14)):
    data.column_dimensions[col].width = w
for r in range(2, PASTE_ROWS + 2):                      # style the empty rows
    for col in range(1, 5):
        c = data.cell(row=r, column=col)
        c.font = Font(name=ARIAL, size=9, color="0000FF")   # blue = an input
        c.fill = INPUT; c.border = BOX
        if col == 1: c.number_format = 'dd/mm/yyyy'
        if col == 4: c.number_format = MONEY
# a few example rows so the pack is not empty on first open
for i, (d, a, desc, amt) in enumerate([
        ("2026-07-05", "Sales", "Invoice 1001", 12000), ("2026-07-06", "Materials", "Timber", -3200),
        ("2026-07-31", "Wages", "July payroll", -4100), ("2026-08-03", "Sales", "Invoice 1002", 15500),
        ("2026-08-09", "Materials", "Fixings", -1850), ("2026-08-31", "Wages", "August payroll", -4100),
        ("2026-08-31", "Rent", "August rent", -2500), ("2026-08-31", "Insurance", "Policy renewal", -900)]):
    import datetime as dt
    data.cell(2 + i, 1, dt.date.fromisoformat(d))
    data.cell(2 + i, 2, a); data.cell(2 + i, 3, desc); data.cell(2 + i, 4, amt)

# ============================================================= 2. THE ENGINE =
# Aggregate ONCE here; every report then reads these cells instead of
# re-scanning the data. With one sheet it barely matters. With eight it is the
# difference between a pack that recalculates in seconds and one that crawls.
DP = f"Data!$B$2:$B${PASTE_ROWS + 1}"       # account column
DD = f"Data!$A$2:$A${PASTE_ROWS + 1}"       # date column
DV = f"Data!$D$2:$D${PASTE_ROWS + 1}"       # amount column
eng = wb.create_sheet("Engine")
header(eng, 1, ["Account", "Category", "Jul", "Aug", "Total"])
for i, a in enumerate(ACCOUNTS):
    r = 2 + i
    eng.cell(r, 1, a).font = Font(name=ARIAL, size=10)
    eng.cell(r, 2, CATEGORY[a]).font = Font(name=ARIAL, size=10)
    for k, (col, lo, hi) in enumerate(((3, "2026-07-01", "2026-07-31"),
                                       (4, "2026-08-01", "2026-08-31"))):
        # SUMIFS = add up column DV where the account matches AND the date is in range
        c = eng.cell(r, col, f'=SUMIFS({DV},{DP},$A{r},{DD},">={lo}",{DD},"<={hi}")')
        c.number_format = MONEY; c.font = Font(name=ARIAL, size=10)
    eng.cell(r, 5, f"=SUM(C{r}:D{r})").number_format = MONEY
for col, w in zip("ABCDE", (20, 16, 14, 14, 14)):
    eng.column_dimensions[col].width = w

# ============================================================= 3. THE REPORT =
# Category rows that collapse to hide the accounts underneath. The trick is
# outlineLevel on the detail rows plus summaryBelow=False, because our summary
# row sits ABOVE its detail. Miss that and the +/- buttons land on wrong rows.
rep = wb.create_sheet("Report")
header(rep, 1, ["Line", "Jul", "Aug", "Total"])
r = 2
for cat in ["Income", "Cost of Sales", "Overheads"]:
    members = [i for i, a in enumerate(ACCOUNTS) if CATEGORY[a] == cat]
    cat_row = r
    rep.cell(r, 1, cat).font = Font(name=ARIAL, size=11, bold=True)
    for col in range(1, 5): rep.cell(r, col).fill = TOTAL
    r += 1
    for i in members:                                    # one row per account
        rep.cell(r, 1, f"    {ACCOUNTS[i]}").font = Font(name=ARIAL, size=10)
        for col in (2, 3, 4):
            c = rep.cell(r, col, f"=Engine!{chr(65 + col)}{2 + i}")
            c.number_format = MONEY; c.font = Font(name=ARIAL, size=10)
        rep.row_dimensions[r].outlineLevel = 1           # <- makes it collapsible
        rep.row_dimensions[r].hidden = True              # <- starts collapsed
        r += 1
    for col in (2, 3, 4):                                # category total
        c = rep.cell(cat_row, col, f"=SUM({chr(64 + col)}{cat_row + 1}:{chr(64 + col)}{r - 1})")
        c.number_format = MONEY; c.font = Font(name=ARIAL, size=11, bold=True)
        c.fill = TOTAL
rep.cell(r, 1, "NET").font = Font(name=ARIAL, size=11, bold=True)
for col in (2, 3, 4):
    c = rep.cell(r, col, f"=SUMIF($A$2:$A${r-1},\"Income\",{chr(64+col)}$2:{chr(64+col)}${r-1})"
                         f"+SUMIF($A$2:$A${r-1},\"Cost of Sales\",{chr(64+col)}$2:{chr(64+col)}${r-1})"
                         f"+SUMIF($A$2:$A${r-1},\"Overheads\",{chr(64+col)}$2:{chr(64+col)}${r-1})")
    c.number_format = MONEY; c.font = Font(name=ARIAL, size=11, bold=True)
rep.column_dimensions["A"].width = 30
for col in "BCD": rep.column_dimensions[col].width = 14

# ============================================================ 4. THE CONTROL =
# A pack nobody checks is a pack nobody should trust. One control that proves
# nothing fell between the paste and the report is worth ten pretty charts.
ctl = wb.create_sheet("Check")
header(ctl, 1, ["Check", "Result", "Should be", "OK?"])
ctl.cell(2, 1, "Total pasted less total on the Engine").font = Font(name=ARIAL, size=10)
ctl.cell(2, 2, f"=ROUND(SUM({DV})-SUM(Engine!$E$2:$E${1+len(ACCOUNTS)}),2)").number_format = MONEY
ctl.cell(2, 3, "nil").font = Font(name=ARIAL, size=10)
ctl.cell(2, 4, '=IF(ROUND(B2,2)=0,"yes","NO - an account is missing from the Engine")')
ctl.cell(2, 4).font = Font(name=ARIAL, size=10, bold=True)
for col, w in zip("ABCD", (40, 16, 12, 44)): ctl.column_dimensions[col].width = w

# ----------------------------------------------------------------- FINISH ---
for ws in wb.worksheets:
    ws.sheet_view.showGridLines = False
# set LAST: our summary rows sit above their detail
rep.sheet_properties.outlinePr = Outline(summaryBelow=False, summaryRight=False, applyStyles=False)
wb.save(OUT)
print(f"wrote {OUT} - open it, then come back and change something above")
