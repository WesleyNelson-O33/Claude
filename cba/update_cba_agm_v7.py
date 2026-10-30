"""Seventh pass on Danica's CBA AGM 2026 Costs and Hours 29102026 file. Staff Hours by Date now finds the CBA AGM 2026
lines itself. One section per Timesheet tab, oldest first, each with spare rows. In each row, column L uses
SMALL(IF(...)) to find the next timesheet row at location CBA AGM 2026 that is not rejected. The other columns read that
row with INDEX. Column M marks a line already listed from an earlier tab as Duplicate, and every total only counts Yes.
The Summary hours block reads the same columns. Her other edits and the WIP lines are left alone."""
import sys, datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.formula import ArrayFormula

SRC, OUT = sys.argv[1:3]
F = "Arial"
HEAD = PatternFill("solid", fgColor="1F3864")
SECT = PatternFill("solid", fgColor="D9E1F2")
AMBER = PatternFill("solid", fgColor="FFE699")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HRS = '0.00;(0.00);"-"'
TAB = "Staff Hours by Date"
SPARE = 10
LAST_SRC = 6000
MATCH = "CBA AGM 2026"

wb = openpyxl.load_workbook(SRC)


def put(ws, ref, v, bold=False, fmt=None, fill=None, wrap=False, head=False, size=10):
    c = ws[ref]
    c.value = v
    c.font = Font(name=F, bold=bold or head, size=size, color="FFFFFF" if head else "000000")
    c.alignment = Alignment(wrap_text=wrap, vertical="top")
    c.fill = HEAD if head else (fill or PatternFill())
    c.border = BOX if (head or fmt) else Border()
    if fmt:
        c.number_format = fmt


# ---------------------------------------------------------------- the Timesheet tabs, oldest first
src = []
for ws in wb.worksheets:
    if not ws.title.startswith("Timesheet "):
        continue
    hc = {c.value: c.column_letter for c in ws[1] if c.value}
    assert ws.max_row <= LAST_SRC, ws.title
    first_date, found, names = None, 0, {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        g = lambda k: row[openpyxl.utils.column_index_from_string(hc[k]) - 1]
        d = g("Start Date")
        if d and (first_date is None or d < first_date):
            first_date = d
        if MATCH in str(g("Location") or "") and g("Status") != "Rejected":
            found += 1
            n = "%s %s" % (g("First Name"), str(g("Surname")).title())
            names[n] = names.get(n, 0) + g("Duration").total_seconds() / 3600
    src.append((first_date, ws.title, hc, found, names))
src.sort()
assert len(src) == 8

# ---------------------------------------------------------------- rebuild Staff Hours by Date
idx = wb.sheetnames.index(TAB)
del wb[TAB]
sd = wb.create_sheet(TAB, idx)
sd.sheet_view.showGridLines = False
put(sd, "A1", "CBA AGM 2026, job 2506208. Staff hours by date", bold=True, size=13)
put(sd, "A2", "This tab finds the lines itself. For each Timesheet tab, column L looks for the next row whose Location "
              "contains CBA AGM 2026 and whose Status is not Rejected. The other columns read that row. Column M says "
              "Duplicate when the same Timesheet Id is already listed from an earlier tab, and the totals only count Yes. Column N "
              "marks the first line for that person on that day, so Days counts each day once. "
              "Each section has %d spare rows for new timesheets. The check at the bottom says if a section runs out." % SPARE,
    wrap=True, fill=AMBER)
sd.merge_cells("A2:N2"); sd.row_dimensions[2].height = 58
H = ["Date", "Employee", "Location", "Start", "Finish", "Unpaid break", "Hours worked", "Status", "Work type",
     "Timesheet Id", "Source tab", "Source row", "Counted", "New day"]
for i, h in enumerate(H):
    put(sd, "%s3" % chr(65 + i), h, head=True, wrap=True)

r = 4
sections = []
for first_date, title, hc, found, _ in src:
    q = "'%s'!" % title
    rng = lambda k: "%s$%s$2:$%s$%d" % (q, hc[k], hc[k], LAST_SRC)
    put(sd, "A%d" % r, title, bold=True, fill=SECT)
    for col in "BCDEFGHIJKLMN":
        sd["%s%d" % (col, r)].fill = SECT
    r += 1
    s_first = r
    slots = found + SPARE
    for k in range(1, slots + 1):
        L = "$L%d" % r
        cond = 'ISNUMBER(SEARCH("%s",%s))*(%s<>"Rejected")' % (MATCH, rng("Location"), rng("Status"))
        sd["L%d" % r] = ArrayFormula("L%d" % r, '=IFERROR(SMALL(IF(%s,ROW(%s)),ROWS(L$%d:L%d)),"")' % (cond, rng("Location"), s_first, r))
        sd["L%d" % r].number_format = "0"
        get = lambda key: "INDEX(%s$%s:$%s,%s)" % (q, hc[key], hc[key], L)
        cells = {
            "A": ("=IF(%s=\"\",\"\",%s)" % (L, get("Start Date")), "DD/MM/YYYY"),
            "B": ("=IF(%s=\"\",\"\",%s&\" \"&PROPER(%s))" % (L, get("First Name"), get("Surname")), "@"),
            "C": ("=IF(%s=\"\",\"\",%s)" % (L, get("Location")), "@"),
            "D": ("=IF(%s=\"\",\"\",%s)" % (L, get("Start Time")), "h:mm"),
            "E": ("=IF(%s=\"\",\"\",%s)" % (L, get("End Time")), "h:mm"),
            "F": ("=IF(%s=\"\",\"\",%s)" % (L, get("Unpaid Break Duration")), "[h]:mm"),
            "G": ("=IF(%s=\"\",\"\",ROUND(%s*24,2))" % (L, get("Duration")), HRS),
            "H": ("=IF(%s=\"\",\"\",%s)" % (L, get("Status")), "@"),
            "I": ("=IF(%s=\"\",\"\",IF(%s=\"\",\"\",%s))" % (L, get("Work Type"), get("Work Type")), "@"),
            "J": ("=IF(%s=\"\",\"\",%s)" % (L, get("Timesheet Id")), "0"),
            "K": ("=IF(%s=\"\",\"\",\"%s\")" % (L, title), "@"),
            "M": ("=IF(%s=\"\",\"\",IF(COUNTIF($J$3:J%d,J%d)>0,\"Duplicate\",\"Yes\"))" % (L, r - 1, r), "@"),
            "N": ('=IF(M%d<>"Yes",0,IF(COUNTIFS($B$4:B%d,B%d,$A$4:A%d,A%d,$M$4:M%d,"Yes")>0,0,1))' % (r, r - 1, r, r - 1, r, r - 1), "0"),
        }
        for col, (f, fmt) in cells.items():
            put(sd, "%s%d" % (col, r), f, fmt=fmt)
        sd["L%d" % r].font = Font(name=F, size=10)
        sd["L%d" % r].border = BOX
        r += 1
    sections.append((title, hc, s_first, r - 1, slots))
first, last = 5, r - 1
A, B, C, G, M, N = ("$%s$%d:$%s$%d" % (c, first, c, last) for c in "ABCGMN")
r += 1
put(sd, "F%d" % r, "Total", bold=True)
put(sd, "G%d" % r, '=SUMIFS(%s,%s,"Yes")' % (G, M), bold=True, fmt=HRS)
total_row = r
TOTAL = "'%s'!G%d" % (TAB, total_row)

# ---------------------------------------------------------------- by employee
all_names = {}
for *_, names in src:
    for n, h in names.items():
        all_names[n] = all_names.get(n, 0) + h
names = sorted(all_names, key=lambda n: -all_names[n])
r += 3
put(sd, "A%d" % r, "Hours by employee", bold=True, size=12)
for col, h in zip("ABC", ["Employee", "Hours worked", "Days"]):
    put(sd, "%s%d" % (col, r + 1), h, head=True)
emp_first = r + 2
for i, n in enumerate(names, emp_first):
    put(sd, "A%d" % i, n, fmt="@")
    put(sd, "B%d" % i, '=SUMIFS(%s,%s,A%d,%s,"Yes")' % (G, B, i, M), fmt=HRS)
    put(sd, "C%d" % i, "=SUMIFS(%s,%s,A%d)" % (N, B, i), fmt="0")
e = emp_first + len(names)
put(sd, "A%d" % e, "Total", bold=True)
put(sd, "B%d" % e, "=SUM(B%d:B%d)" % (emp_first, e - 1), bold=True, fmt=HRS)
emp_total = e

# ---------------------------------------------------------------- by month
r = e + 3
put(sd, "A%d" % r, "By month", bold=True, size=12)
for col, h in zip("ABCD", ["Month", "Production [PRD]", "Video [VID]", "Total"]):
    put(sd, "%s%d" % (col, r + 1), h, head=True)
m_first = r + 2
months, m = [], datetime.date(2025, 1, 1)
while m <= datetime.date(2026, 11, 1):
    months.append(m)
    m = datetime.date(m.year + (m.month == 12), m.month % 12 + 1, 1)
for i, m in enumerate(months, m_first):
    crit = '%s,">="&A%d,%s,"<"&EDATE(A%d,1),%s,"Yes"' % (A, i, A, i, M)
    put(sd, "A%d" % i, m, fmt="MMM YYYY")
    put(sd, "B%d" % i, '=SUMIFS(%s,%s,%s,"*[PRD]*")' % (G, crit, C), fmt=HRS)
    put(sd, "C%d" % i, '=SUMIFS(%s,%s,%s,"*[VID]*")' % (G, crit, C), fmt=HRS)
    put(sd, "D%d" % i, "=B%d+C%d" % (i, i), fmt=HRS)
me = m_first + len(months)
put(sd, "A%d" % me, "Total", bold=True)
for col in "BCD":
    put(sd, "%s%d" % (col, me), "=SUM(%s%d:%s%d)" % (col, m_first, col, me - 1), bold=True, fmt=HRS)

# ---------------------------------------------------------------- checks
r = me + 3
put(sd, "A%d" % r, "Checks", bold=True, size=12)
for col, h in zip("ABCDE", ["Timesheet tab", "Lines in the tab", "Rows in its section", "Hours in the tab", "Result"]):
    put(sd, "%s%d" % (col, r + 1), h, head=True, wrap=True)
c_first = r + 2
for i, (title, hc, s1, s2, slots) in enumerate(sections, c_first):
    q = "'%s'!" % title
    rng = lambda k: "%s$%s$2:$%s$%d" % (q, hc[k], hc[k], LAST_SRC)
    put(sd, "A%d" % i, title, fmt="@")
    put(sd, "B%d" % i, '=SUMPRODUCT(ISNUMBER(SEARCH("%s",%s))*(%s<>"Rejected"))' % (MATCH, rng("Location"), rng("Status")), fmt="0")
    put(sd, "C%d" % i, slots, fmt="0")
    put(sd, "D%d" % i, '=ROUND(SUMIFS(%s,%s,"*%s*",%s,"<>Rejected")*24,2)' % (rng("Duration"), rng("Location"), MATCH, rng("Status")), fmt=HRS)
    put(sd, "E%d" % i, '=IF(B%d<=C%d,"OK","Add rows to this section")' % (i, i), fmt="@")
c_last = c_first + len(sections) - 1
k = c_last + 1
put(sd, "A%d" % k, "All tabs", bold=True)
put(sd, "D%d" % k, "=SUM(D%d:D%d)" % (c_first, c_last), bold=True, fmt=HRS)
put(sd, "A%d" % (k + 1), "Less lines marked Duplicate")
put(sd, "D%d" % (k + 1), '=SUMIFS(%s,%s,"Duplicate")' % (G, M), fmt=HRS)
put(sd, "A%d" % (k + 2), "Should equal the Total", bold=True)
put(sd, "D%d" % (k + 2), "=ROUND(D%d-D%d,2)" % (k, k + 1), bold=True, fmt=HRS)
put(sd, "E%d" % (k + 2), '=IF(ROUND(D%d-G%d,2)=0,"OK","Check")' % (k + 2, total_row), bold=True, fmt="@")
put(sd, "A%d" % (k + 3), "Hours by employee add up to the Total", bold=True)
put(sd, "E%d" % (k + 3), '=IF(ROUND(B%d-G%d,2)=0,"OK","Someone new. Add their name to Hours by employee")' % (emp_total, total_row), bold=True, fmt="@")
put(sd, "A%d" % (k + 4), "By month adds up to the Total", bold=True)
put(sd, "E%d" % (k + 4), '=IF(ROUND(D%d-G%d,2)=0,"OK","Check")' % (me, total_row), bold=True, fmt="@")

for col, w in zip("ABCDEFGHIJKLMN", [13, 30, 34, 8, 8, 9, 10, 11, 22, 12, 26, 8, 10, 8]):
    sd.column_dimensions[col].width = w
sd.freeze_panes = "A4"

# ---------------------------------------------------------------- Summary
sm = wb["Summary"]
assert sm["A7"].value == "Staff hours, earlier monthly reports" and sm["A25"].value == "Staff hours by employee"
assert sm["B21"].value == "='Staff Hours'!D22"
sm["A7"].value = "Staff hours, EH timesheets"
sm["B7"].value = "=" + TOTAL
sm["C7"].value = "Hours. Every EH timesheet at CBA AGM 2026, PRD and VID, January 2025 to 30 October 2026."
assert sm["A13"].value.startswith("3. Staff hours by person and date are still to come")
sm["A13"].value = ("3. Staff hours by person and date are from the Employment Hero timesheet exports, 1 January 2025 "
                   "to 30 October 2026. See the Staff Hours by Date tab.")
for rg in list(sm.merged_cells.ranges):
    if rg.min_row >= 25:
        sm.unmerge_cells(str(rg))
for row in sm.iter_rows(min_row=25, max_row=max(sm.max_row, 60)):
    for c in row:
        c.value = None
        c.fill, c.border = PatternFill(), Border()
D = "'%s'!" % TAB
r = 25
put(sm, "A%d" % r, "Staff hours by employee", bold=True, size=12)
for col, h in zip("ABCDE", ["Employee", "Hours worked", "Days", "First date", "Last date"]):
    put(sm, "%s%d" % (col, r + 1), h, head=True, wrap=True)
for k, i in enumerate(range(r + 2, r + 2 + len(names))):
    put(sm, "A%d" % i, "=%sA%d" % (D, emp_first + k), fmt="@")
    put(sm, "B%d" % i, "=%sB%d" % (D, emp_first + k), fmt=HRS)
    put(sm, "C%d" % i, "=%sC%d" % (D, emp_first + k), fmt="0")
    put(sm, "D%d" % i, '=_xlfn.MINIFS(%s%s,%s%s,A%d,%s%s,"Yes")' % (D, A, D, B, i, D, M), fmt="DD/MM/YYYY")
    put(sm, "E%d" % i, '=_xlfn.MAXIFS(%s%s,%s%s,A%d,%s%s,"Yes")' % (D, A, D, B, i, D, M), fmt="DD/MM/YYYY")
t = r + 2 + len(names)
put(sm, "A%d" % t, "Total", bold=True, fmt="@")
put(sm, "B%d" % t, "=SUM(B%d:B%d)" % (r + 2, t - 1), bold=True, fmt=HRS)
put(sm, "A%d" % (t + 1), "From the Employment Hero timesheet exports. Hours by date are on the Staff Hours by Date tab.", wrap=True)
sm.merge_cells("A%d:E%d" % (t + 1, t + 1))

wb.save(OUT)
print("rows", first, last, "total row", total_row, [(s[0], s[4]) for s in sections])
