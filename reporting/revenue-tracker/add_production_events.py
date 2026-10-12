"""Production changes agreed with the Production manager (Oct 2026), applied to a live filled tracker.

  * Lists AD: Event Type list (AGM, Investor Day, Town Hall, Results Announcements) - add more under it.
  * Production: new 'Event Type' column (dropdown) at the end of the table, so nothing that feeds Finance moves.
  * New sheet 'Production by Event Date': every Production job in event-date order (oldest or newest first),
    read-only, filterable by month / event type / client, with totals that follow the filter (total margin
    included) and averages by event type. The Production sheet itself never changes order, so Finance
    always pulls through correctly.

Usage: python add_production_events.py <tracker.xlsx> <out.xlsx>
"""
import sys

import openpyxl
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import TableColumn

SRC, OUT = sys.argv[1:3]
FONT = "Arial"
NAVY, SLATE = "1F3864", "1F4E79"
FILL = lambda c: PatternFill("solid", fgColor=c)
HAIR = Side(style="hair", color="C9D3E3")
THIN = Side(style="thin", color="C9D3E3")
MONEY = '$#,##0.00;[Red]-$#,##0.00;"-"'
N = 1500                      # Production rows 5..1504
EVENT_TYPES = ["AGM", "Investor Day", "Town Hall", "Results Announcements"]

wb = openpyxl.load_workbook(SRC)


def protect(ws):
    p = ws.protection
    p.sheet, p.password = True, "CTS1234"
    p.autoFilter = p.formatCells = p.formatColumns = p.formatRows = False
    p.sort = p.insertRows = p.deleteRows = True


def hdr(c, text, fill=SLATE):
    c.value = text
    c.font = Font(name=FONT, size=10, bold=True, color="FFFFFF")
    c.fill = FILL(fill)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


# ------------------------------------------------------------- Lists: event types
ls = wb["Lists"]
hdr(ls["AD4"], "Event Type", NAVY)
for i in range(20):
    c = ls.cell(5 + i, 30, EVENT_TYPES[i] if i < len(EVENT_TYPES) else None)
    c.font = Font(name=FONT, size=10)
    c.fill = FILL("FFF2CC")
    c.protection = Protection(locked=False)
ls["AD26"] = "Production's event types. Add a new one in the next empty cell above - the dropdown picks it up."
ls["AD26"].font = Font(name=FONT, size=9, italic=True, color="595959")
ls.column_dimensions["AD"].width = 24
wb.defined_names["lst_EventType"] = DefinedName(
    "lst_EventType", attr_text="OFFSET(Lists!$AD$5,0,0,MAX(1,COUNTA(Lists!$AD$5:$AD$24)),1)")

# ------------------------------------------------------- Production: Event Type
pr = wb["Production"]
head = {pr.cell(4, c).value: c for c in range(1, pr.max_column + 1) if pr.cell(4, c).value}
assert "Event Type" not in head, "already applied"
col = 49                                               # AW - straight after Finance Row, table ends at AV
L = "AW"
hdr(pr.cell(4, col), "Event Type", NAVY)
src_in = pr["U5"]                                      # Event Grouping - a white typing cell
for r in range(5, 5 + N):
    c = pr.cell(r, col)
    c.font = Font(name=FONT, size=10)
    c.fill = FILL("FFFFFF")
    c.border = Border(bottom=HAIR)
    c.protection = Protection(locked=False)
pr.column_dimensions[L].width = 22
dv = DataValidation(type="list", formula1="lst_EventType", allow_blank=True, showErrorMessage=True,
                    errorStyle="warning", errorTitle="Event Type",
                    error="That is not on the Event Type list (Lists AD). Click Yes to keep it, or add it to the list.")
dv.add(f"{L}5:{L}{4 + N}")
pr.add_data_validation(dv)
t = pr.tables["tbl_Production"]
t.ref = f"A4:{L}{4 + N}"
t.tableColumns.append(TableColumn(id=len(t.tableColumns) + 1, name="Event Type"))
if t.autoFilter is not None:
    t.autoFilter.ref = t.ref

# --------------------------------------------- Production by Event Date (view)
name = "Production by Event Date"
if name in wb.sheetnames:
    del wb[name]
vw = wb.create_sheet(name, wb.sheetnames.index("Production") + 1)
vw.sheet_properties.tabColor = "2E75B6"
vw.sheet_view.showGridLines = False
for c in range(1, 23):
    for r in (1, 2):
        vw.cell(r, c).fill = FILL(NAVY)
vw["A1"] = "Production by Event Date"
vw["A1"].font = Font(name=FONT, size=18, bold=True, color="FFFFFF")
vw["A2"] = ("Every Production job in event-date order, straight from the Production sheet - nothing is typed here. "
            "Filter any column (month, event type, client). The totals follow the filter. Change jobs on the Production sheet.")
vw["A2"].font = Font(name=FONT, size=10, italic=True, color="D9E2F3")
vw.row_dimensions[1].height = 30

# order picker
vw["A4"] = "Order"
vw["A4"].font = Font(name=FONT, size=10, bold=True)
vw["B4"] = "Oldest first"
vw["B4"].fill = FILL("FFFFFF")
vw["B4"].font = Font(name=FONT, size=10, bold=True, color=NAVY)
vw["B4"].border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
vw["B4"].protection = Protection(locked=False)
odv = DataValidation(type="list", formula1='"Oldest first,Newest first"', allow_blank=False)
odv.add("B4")
vw.add_data_validation(odv)
vw["C4"] = "<- pick Oldest first or Newest first. Jobs with no event date go to the bottom."
vw["C4"].font = Font(name=FONT, size=9, italic=True, color="595959")

P = "Production!"
pcol = lambda h: f"{P}${openpyxl.utils.get_column_letter(head[h])}$5:${openpyxl.utils.get_column_letter(head[h])}${4 + N}"
cols = [  # (header, source header on Production or special, format, width)
    ("Event Date", "Event Date", "dd-mmm-yy", 11),
    ("Event Month", "_month", "mmm-yy", 9),
    ("Job Number", "Job Number", "@", 12),
    ("Client", "Client", None, 18),
    ("Job Description", "Job Description", None, 36),
    ("Event Type", "_etype", None, 18),
    ("Event Grouping", "Event Grouping", None, 14),
    ("Cost Centre", "Cost Centre", None, 12),
    ("Current RMS No", "Current RMS No", "@", 10),
    ("Expected Revenue Ex GST", "Expected Revenue Ex GST", MONEY, 13),
    ("Invoiced Ex GST", "Revenue Ex GST", MONEY, 13),
    ("Invoice No", "Invoice No", None, 11),
    ("Invoiced", "Invoiced", None, 16),
    ("Discounts Given", "Discounts Given", MONEY, 12),
    ("Cross Hire", "Cross Hire Expense", MONEY, 11),
    ("Conf. Call", "Conf. Call Costs", MONEY, 11),
    ("Transcription", "Transcription Costs", MONEY, 11),
    ("Labour (Internal)", "Labour Expense (Internal)", MONEY, 12),
    ("Total Expense", "Total Expense", MONEY, 12),
    ("Total Margin", "Margin", MONEY, 13),
    ("Margin %", "_mpct", "0.0%", 9),
    ("Notes", "Notes", None, 30),
]
HR, FIRST = 7, 8                       # header row, first data row
LAST = FIRST + N - 1
TOT = 6                                # totals row (follows filters)
for i, (h, s, fmt, w) in enumerate(cols):
    c = i + 1
    CLc = openpyxl.utils.get_column_letter(c)
    hdr(vw.cell(HR, c), h)
    vw.column_dimensions[CLc].width = w
vw.row_dimensions[HR].height = 32
vw.cell(TOT, 1, "TOTAL (filtered)").font = Font(name=FONT, size=10, bold=True, color=NAVY)
for c in range(1, len(cols) + 1):
    vw.cell(TOT, c).fill = FILL("D9E2F3")

# hidden sort key per Production row (column Z), position of each view row (column Y)
KEYC, POSC = "Z", "Y"
hdr(vw[f"{KEYC}{HR}"], "Sort key")
hdr(vw[f"{POSC}{HR}"], "Prod row")
jobs, dates = pcol("Job Number"), pcol("Event Date")
for i in range(N):
    r = FIRST + i
    k = i + 1
    vw[f"{KEYC}{r}"] = (f'=IF(INDEX({jobs},{k})&""="","",IF(ISNUMBER(INDEX({dates},{k})),'
                        f'IF($B$4="Newest first",-INDEX({dates},{k}),INDEX({dates},{k}))+{k}/100000,1000000+{k}/100000))')
    vw[f"{POSC}{r}"] = f'=IFERROR(MATCH(SMALL(${KEYC}${FIRST}:${KEYC}${LAST},{k}),${KEYC}${FIRST}:${KEYC}${LAST},0),"")'
    for i2, (h, s, fmt, w) in enumerate(cols):
        c = vw.cell(r, i2 + 1)
        g = f"${POSC}{r}"
        if s == "_month":
            fx = f'=IF({g}="","",IF(ISNUMBER(INDEX({dates},{g})),EOMONTH(INDEX({dates},{g}),0),""))'
        elif s == "_etype":
            fx = f'=IF({g}="","",INDEX({P}$AW$5:$AW${4 + N},{g})&"")'
        elif s == "_mpct":
            fx = f'=IF({g}="","",IFERROR(IF(MAX(N(J{r}),N(K{r}))=0,"",T{r}/MAX(N(J{r}),N(K{r}))),""))'
        elif fmt in (MONEY,) or s == "Event Date":
            fx = f'=IF({g}="","",IF(INDEX({pcol(s)},{g})&""="","",INDEX({pcol(s)},{g})))'
        else:
            fx = f'=IF({g}="","",INDEX({pcol(s)},{g})&"")'
        c.value = fx
        c.font = Font(name=FONT, size=10, color="262626")
        c.border = Border(bottom=HAIR)
        if fmt:
            c.number_format = fmt
# hidden revenue base (invoiced, or expected if higher) - the same base Margin uses on the Production sheet
hdr(vw[f"X{HR}"], "Revenue base")
for r in range(FIRST, LAST + 1):
    vw[f"X{r}"] = f'=IF($Y{r}="","",MAX(N(J{r}),N(K{r})))'
    vw[f"X{r}"].number_format = MONEY
for c in ("X", KEYC, POSC):
    vw.column_dimensions[c].hidden = True

# totals that follow the filter
for i, (h, s, fmt, w) in enumerate(cols):
    CLc = openpyxl.utils.get_column_letter(i + 1)
    x = vw.cell(TOT, i + 1)
    if fmt == MONEY:
        x.value = f"=SUBTOTAL(109,{CLc}{FIRST}:{CLc}{LAST})"
    elif h == "Job Number":
        x.value = (f'=SUMPRODUCT(SUBTOTAL(103,OFFSET({CLc}{FIRST},ROW({CLc}{FIRST}:{CLc}{LAST})-ROW({CLc}{FIRST}),0)),'
                   f'--({CLc}{FIRST}:{CLc}{LAST}<>""))&" jobs"')
    elif h == "Margin %":
        x.value = f'=IFERROR(T{TOT}/SUBTOTAL(109,X{FIRST}:X{LAST}),"")'
    if x.value:
        x.font = Font(name=FONT, size=10, bold=True, color=NAVY)
        if fmt:
            x.number_format = fmt
vw.auto_filter.ref = f"A{HR}:{openpyxl.utils.get_column_letter(len(cols))}{LAST}"
vw.freeze_panes = f"D{FIRST}"
red = FILL("F8D7DA")
vw.conditional_formatting.add(f"T{FIRST}:T{LAST}", FormulaRule(formula=[f'AND(ISNUMBER(T{FIRST}),T{FIRST}<0)'],
                                                                font=Font(color="9C0006", bold=True)))

# averages by event type (to the right)
S0 = 28                       # column AB onwards
labels = ["Event Type", "Jobs", "Revenue", "Average Revenue", "Average Total Expense", "Average Total Margin", "Margin %",
          "Average Discount"]
for i, h in enumerate(labels):
    hdr(vw.cell(HR, S0 + i), h, NAVY)
    vw.column_dimensions[openpyxl.utils.get_column_letter(S0 + i)].width = 14 if i else 22
vw.cell(HR - 1, S0, "AVERAGE BY EVENT TYPE  -  all Production jobs with that event type").font = Font(name=FONT, size=10, bold=True, color=NAVY)
et = f"{P}$AW$5:$AW${4 + N}"
mg, ex, ds = pcol("Margin"), pcol("Total Expense"), pcol("Discounts Given")
for i in range(21):
    r = FIRST + i
    a = vw.cell(r, S0, f'=IF(INDEX(Lists!$AD$5:$AD$24,{i + 1})="","",INDEX(Lists!$AD$5:$AD$24,{i + 1}))' if i < 20 else "No event type set")
    a.font = Font(name=FONT, size=10, bold=True)
    A = openpyxl.utils.get_column_letter(S0)
    crit = f'{et},${A}{r}' if i < 20 else f'{et},"",{jobs},"<>"'
    f = [
        f'=IF(${A}{r}="","",COUNTIFS({crit}))',
        f'=IF(${A}{r}="","",SUMIFS({mg},{crit})+SUMIFS({ex},{crit}))',
        f'=IF(OR(${A}{r}="",N({openpyxl.utils.get_column_letter(S0 + 1)}{r})=0),"",{openpyxl.utils.get_column_letter(S0 + 2)}{r}/{openpyxl.utils.get_column_letter(S0 + 1)}{r})',
        f'=IF(OR(${A}{r}="",N({openpyxl.utils.get_column_letter(S0 + 1)}{r})=0),"",SUMIFS({ex},{crit})/{openpyxl.utils.get_column_letter(S0 + 1)}{r})',
        f'=IF(OR(${A}{r}="",N({openpyxl.utils.get_column_letter(S0 + 1)}{r})=0),"",SUMIFS({mg},{crit})/{openpyxl.utils.get_column_letter(S0 + 1)}{r})',
        f'=IF(OR(${A}{r}="",N({openpyxl.utils.get_column_letter(S0 + 2)}{r})=0),"",SUMIFS({mg},{crit})/{openpyxl.utils.get_column_letter(S0 + 2)}{r})',
        f'=IF(OR(${A}{r}="",N({openpyxl.utils.get_column_letter(S0 + 1)}{r})=0),"",SUMIFS({ds},{crit})/{openpyxl.utils.get_column_letter(S0 + 1)}{r})',
    ]
    for j, fx in enumerate(f):
        x = vw.cell(r, S0 + 1 + j, fx)
        x.font = Font(name=FONT, size=10)
        x.border = Border(bottom=HAIR)
        x.number_format = "0" if j == 0 else ("0.0%" if j == 5 else MONEY)
vw.cell(FIRST + 22, S0, ("Revenue = invoiced, or expected if not yet invoiced (the same base as Margin on the Production sheet). "
                         "Total Margin = revenue less cross hire, conference call, transcription and internal labour. "
                         "Company overheads are not allocated to jobs, so this is the job margin before overheads.")).font = Font(name=FONT, size=9, italic=True, color="595959")
protect(vw)

# --------------------------------------- Deferrals highlight (Excel keeps it as an extension openpyxl drops)
wb["Deferrals"].conditional_formatting.add(
    "AG7:CZ506", FormulaRule(formula=["AG$5='Deferral Journal'!$C$4"], fill=FILL("FFF3CD"), font=Font(name=FONT, color="000000")))

# ---------------------------------------------------------------- Read Me
rm = wb["Read Me"]
r0 = rm.max_row + 2
lines = [
    ("PRODUCTION - EVENT DATE VIEW AND EVENT TYPE", True),
    ("Production keeps typing on the Production sheet exactly as before. The new Event Type column (far right, after Finance Row) has a dropdown: AGM, Investor Day, Town Hall, Results Announcements. Add new types on Lists AD.", False),
    ("The 'Production by Event Date' sheet lists every Production job in event-date order - pick Oldest first or Newest first in B4. Filter by Event Month, Event Type, Client or anything else; the TOTAL row follows the filter, including Total Margin and Margin %.", False),
    ("It is read-only on purpose: the Production sheet never changes order, so every line keeps pulling through to Finance. To change a job, change it on the Production sheet.", False),
    ("Average by Event Type (right of the view) shows jobs, revenue, average revenue, average total expense, average total margin and average discount for each event type.", False),
]
for i, (t_, h) in enumerate(lines):
    c = rm.cell(r0 + i, 2, t_)
    if h:
        c.font = Font(name=FONT, size=11, bold=True, color="FFFFFF")
        c.fill = FILL(NAVY)
    else:
        c.font = Font(name=FONT, size=11, color="262626")
        c.alignment = Alignment(wrap_text=True, vertical="top")
        rm.row_dimensions[r0 + i].height = 15 * (len(t_) // 105 + 1) + 2

wb.save(OUT)
print("saved", OUT)
