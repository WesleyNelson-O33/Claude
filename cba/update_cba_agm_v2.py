"""Second pass on Danica's edited CBA AGM 2026 Costs and Hours workbook. Keeps her edits as they are and adds:
a GST and totals block on Summary, and a Staff Hours by Date tab built from the Roster vs Timesheet Comp2026 1 tab
(timesheet lines at location CBA AGM 2026 [PRD] or [VID])."""
import sys, datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

SRC, OUT = sys.argv[1:3]
F = "Arial"
HEAD = PatternFill("solid", fgColor="1F3864")
AMBER = PatternFill("solid", fgColor="FFE699")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
MONEY = '#,##0.00;(#,##0.00);"-"'
HRS = '0.00;(0.00);"-"'
RTAB = "Roster vs Timesheet Comp2026 1"

wb = openpyxl.load_workbook(SRC)


def put(ws, ref, v, bold=False, fmt=None, fill=None, wrap=False, head=False, size=10):
    c = ws[ref]
    c.value = v
    c.font = Font(name=F, bold=bold or head, size=size, color="FFFFFF" if head else "000000")
    c.alignment = Alignment(wrap_text=wrap, vertical="top")
    if head:
        c.fill = HEAD
    elif fill:
        c.fill = fill
    if head or fmt:
        c.border = BOX
    if fmt:
        c.number_format = fmt
    return c


# ---------------------------------------------------------------- Summary: GST and totals
sm = wb["Summary"]
assert sm["B3"].value == "='Supplier Costs'!F16" and sm["B4"].value == "='Supplier Costs'!F23"
assert sm["C6"].value == "Hours. No dates by day. October missing."
sm["C6"].value = "Hours, monthly totals from earlier reports. Dates by day are on the Staff Hours by Date tab."
r = 16
put(sm, "A%d" % r, "Costs with GST", bold=True, size=12)
for col, h in zip("ABCD", ["Item", "Ex GST", "GST", "Inc GST"]):
    put(sm, "%s%d" % (col, r + 1), h, head=True)
lines = [("Supplier bills in Xero", 16), ("Encore, PO amount (PO-00016949)", 23)]
for i, (label, src) in enumerate(lines, r + 2):
    put(sm, "A%d" % i, label, fmt="@")
    for col, sc in zip("BCD", "FGH"):
        put(sm, "%s%d" % (col, i), "='Supplier Costs'!%s%d" % (sc, src), fmt=MONEY)
t = r + 2 + len(lines)
put(sm, "A%d" % t, "Total supplier costs", bold=True, fmt="@")
for col in "BCD":
    put(sm, "%s%d" % (col, t), "=SUM(%s%d:%s%d)" % (col, r + 2, col, t - 1), bold=True, fmt=MONEY)
put(sm, "A%d" % (t + 1), "The Encore line is the PO amount, not a bill. Start Webcast is included as billed, still waiting on "
                         "DL/TW approval.", wrap=True)
sm.merge_cells("A%d:D%d" % (t + 1, t + 1)); sm.row_dimensions[t + 1].height = 30
sm.column_dimensions["D"].width = 14

# ---------------------------------------------------------------- Staff Hours by Date
rs = wb[RTAB]
hdr = [c.value for c in rs[1]]
ix = {k: i for i, k in enumerate(hdr) if k}


def hours(td):
    return round(td.total_seconds() / 3600, 2) if isinstance(td, datetime.timedelta) else None


def hhmm(td):
    if not isinstance(td, datetime.timedelta):
        return None
    m = int(td.total_seconds() // 60)
    return "%d:%02d" % (m // 60, m % 60)


lines = []
for row in rs.iter_rows(min_row=2):
    g = lambda k: row[ix[k]].value
    tl = str(g("Timesheet Location") or "")
    if "CBA AGM 2026" in tl:
        rl = str(g("Roster Location") or "")
        lines.append([g("Timesheet Start Date"), "%s %s" % (g("First Name"), str(g("Surname")).title()), tl,
                      hhmm(g("Timesheet Start Time")), hhmm(g("Timesheet End Time")), hhmm(g("Timesheet Breaks")),
                      hours(g("Timesheet Duration")), g("Timesheet Status"), hours(g("Roster Duration")), rl, row[0].row])
lines.sort(key=lambda x: (x[0], x[1]))

sd = wb.create_sheet("Staff Hours by Date", index=wb.sheetnames.index("Staff Hours") + 1)
sd.sheet_view.showGridLines = False
put(sd, "A1", "CBA AGM 2026, job 2506208. Staff hours by date", bold=True, size=13)
put(sd, "A2", "From the Roster vs Timesheet Comp2026 1 tab, 20 March to 30 October 2026. Only timesheets at location "
              "CBA AGM 2026 are counted. This report only shows timesheets that matched a rostered shift, so hours "
              "worked without a roster are missing. Duncan Lugstein and Daniel Sobkowski have none here.", wrap=True, fill=AMBER)
sd.merge_cells("A2:K2"); sd.row_dimensions[2].height = 45
H = ["Date", "Employee", "Timesheet location", "Start", "Finish", "Break", "Hours worked", "Timesheet status",
     "Rostered hours", "Roster location", "Source row"]
for i, h in enumerate(H, 1):
    put(sd, "%s4" % chr(64 + i), h, head=True, wrap=True)
for r, vals in enumerate(lines, 5):
    for i, v in enumerate(vals, 1):
        c = put(sd, "%s%d" % (chr(64 + i), r), v, fmt=["DD/MM/YYYY", "@", "@", "@", "@", "@", HRS, "@", HRS, "@", "0"][i - 1])
last = 4 + len(lines)
put(sd, "F%d" % (last + 1), "Total", bold=True)
put(sd, "G%d" % (last + 1), "=SUM(G5:G%d)" % last, bold=True, fmt=HRS)

r = last + 4
put(sd, "A%d" % r, "Hours by employee", bold=True, size=12)
for col, h in zip("ABC", ["Employee", "Hours worked", "Days"]):
    put(sd, "%s%d" % (col, r + 1), h, head=True)
names = sorted({x[1] for x in lines})
for i, n in enumerate(names, r + 2):
    put(sd, "A%d" % i, n, fmt="@")
    put(sd, "B%d" % i, '=SUMIFS($G$5:$G$%d,$B$5:$B$%d,A%d)' % (last, last, i), fmt=HRS)
    put(sd, "C%d" % i, '=COUNTIFS($B$5:$B$%d,A%d)' % (last, i), fmt="0")
e = r + 2 + len(names)
put(sd, "A%d" % e, "Total", bold=True)
put(sd, "B%d" % e, "=SUM(B%d:B%d)" % (r + 2, e - 1), bold=True, fmt=HRS)

r = e + 3
put(sd, "A%d" % r, "By month", bold=True, size=12)
for col, h in zip("ABC", ["Month", "Hours worked", "Days"]):
    put(sd, "%s%d" % (col, r + 1), h, head=True)
months = sorted({datetime.date(x[0].year, x[0].month, 1) for x in lines})
for i, m in enumerate(months, r + 2):
    nxt = datetime.date(m.year + (m.month == 12), m.month % 12 + 1, 1)
    put(sd, "A%d" % i, m, fmt="MMM YYYY")
    put(sd, "B%d" % i, '=SUMIFS($G$5:$G$%d,$A$5:$A$%d,">="&A%d,$A$5:$A$%d,"<"&DATE(%d,%d,1))'
        % (last, last, i, last, nxt.year, nxt.month), fmt=HRS)
    put(sd, "C%d" % i, '=COUNTIFS($A$5:$A$%d,">="&A%d,$A$5:$A$%d,"<"&DATE(%d,%d,1))' % (last, i, last, nxt.year, nxt.month), fmt="0")
for col, w in zip("ABCDEFGHIJK", [12, 30, 22, 8, 8, 8, 10, 12, 10, 28, 9]):
    sd.column_dimensions[col].width = w
sd.freeze_panes = "A5"

wb.save(OUT)
print(len(lines), "lines", sum(x[6] for x in lines), "hours")
for n in names:
    print(n, sum(x[6] for x in lines if x[1] == n))
