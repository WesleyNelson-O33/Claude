"""WIP 11300 Journal Edit Schedule for September 2026, laid out like Danica's August one
(WIP_Adjustment__Schedule_10_09_2026.xlsx): starting position, lines to change by journal,
where to find each item, notes, and raw copies of the WIP Schedule and the Xero 11300 export."""
import datetime, sys
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import column_index_from_string as CI

WIPF, XEROF, OUT = sys.argv[1], sys.argv[2], sys.argv[3]
RED, AMB = PatternFill("solid", fgColor="F4B6B6"), PatternFill("solid", fgColor="FCE4B6")
BOLD = Font(bold=True)

src_w = openpyxl.load_workbook(WIPF, data_only=True)["Schedule Updated"]
src_x = openpyxl.load_workbook(XEROF, data_only=True).worksheets[0]

# Xero 11300 rows by reference and amount
xrows = []
for r in range(1, src_x.max_row + 1):
    d = src_x.cell(r, 1).value
    if isinstance(d, datetime.datetime):
        amt = (src_x.cell(r, 5).value or 0) - (src_x.cell(r, 6).value or 0)
        xrows.append((r, d, str(src_x.cell(r, 4).value), str(src_x.cell(r, 3).value), round(amt, 2)))
x_close = round(sum(x[4] for x in xrows) - 380517.76, 2)
w_close = round(sum(src_w.cell(r, CI("BZ")).value or 0 for r in range(6, src_w.max_row + 1)
                    if isinstance(src_w.cell(r, CI("BZ")).value, (int, float))), 2)

def xero_row(ref, text, amt):
    for r, d, rf, ds, a in xrows:
        if rf == ref and text in ds and abs(a - amt) < 0.006:
            return r
    raise KeyError((ref, text, amt))

ITEMS = [  # item, WIP schedule row, Xero journal, text in Xero description, amount as posted, line
    ("ASX AGM 22.10.26 pre-prod labour", 480, "#71186", "ASX - AGM - 22.10.26", -1625),
    ("Cochlear AGM 26.10.26 pre-prod labour", 481, "#71186", "Cochlear - AGM - 26.10.26", -2697.5),
    ("TLC AGM 14.10.26 pre-prod labour", 482, "#71186", "TLC - 2026 AGM - 14.10.26", -715),
    ("RSNSW OGM 7.10.26 pre-prod labour", 483, "#71186", "RSNSW Ordinary General Meeting 7.10.2026", -195),
    ("Garvan Breast Cancer Lecture 1.10.26 pre-prod labour", 484, "#71186", "Garvan - Breast Cancer", -130),
    ("Microsoft SME&C Town Hall 14.10.26 pre-prod labour", 485, "#71186", "Microsoft - SME&C", -487.5),
    ("Garvan Suttons Event 15.10.26 pre-prod labour", 486, "#71186", "Garvan - Suttons", -97.5),
    ("ASX CEO Connect 27.10.26 pre-prod labour", 487, "#71186", "ASX CEO Connect 27.10.26", -260),
    ("ART AMM Captioning 25.11.26 pre-prod labour", 488, "#71188", "ART - AMM Captioning", -97.5),
    ("Automic Noumi AGM 5.11.26 pre-prod labour", 489, "#71188", "Automic - Noumi", -162.5),
]
WHY = ("Not changed yet, needs your OK. Pre-production work done in September for an event in {when}, and not invoiced at 30/09 "
       "(no invoice in Xero to 14/10). By the same basis as August's #70428 lines, accruing the unbilled revenue in "
       "September means a debit to WIP. Xero and the schedule both have it as a credit. To fix: in {j} move 11300 to Debit "
       "and 42100 Production Labour to Credit, and change schedule row {row} to a positive figure.")

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Journal Edit Schedule"
ws["A1"] = "WIP 11300 — Journal Edit Schedule, September 2026"; ws["A1"].font = Font(bold=True, size=13)
ws["B3"], ws["G3"] = "STARTING POSITION", "Ending  POSITION (if every line below is changed)"
ws["B4"], ws["E4"] = "Raw Copy of WIP Schedule ", w_close
ws["B5"], ws["E5"] = "Raw Xero WIP 11300", x_close
ws["B6"], ws["E6"] = "Variance to clear", "=E4-E5"
ws["G4"], ws["G5"], ws["G6"] = "Corrected WIP Schedule", "Corrected Xero WIP 11300", "Variance "
for c in ("B3", "G3", "B4", "B5", "B6", "G4", "G5", "G6"): ws[c].font = BOLD
hdr = ["Item", "Journal", "Line to change", "Xero now", "Schedule says", "Effect on\nvariance", "Variance\nafter",
       "What to do in Xero", "Done", "Should be", "Effect on 11300\nif changed"]
for i, h in enumerate(hdr):
    c = ws.cell(8, 2 + i, h); c.font = BOLD; c.alignment = Alignment(wrap_text=True)
r = 10
ws.cell(r, 2, "JOURNAL #71186 — Pre-PROD & INT Work & Expense Accrual for October 2026, and #71188 — WIP Journal September 2026").font = BOLD
first = r + 1
for item, srow, j, txt, amt in ITEMS:
    r += 1
    sched = round(src_w.cell(srow, CI("BY")).value, 2)
    assert sched == amt, (srow, sched, amt)
    when = "November" if "11.26" in item else "October"
    vals = [item, j, "42100 Production Labour / 11300 WIP", amt, sched, "=F%d-E%d" % (r, r), "=F%d-E%d" % (r, r),
            WHY.format(when=when, j=j, row=srow), None, -amt, "=K%d-E%d" % (r, r)]
    for i, v in enumerate(vals): ws.cell(r, 2 + i, v)
    for col in range(2, 13): ws.cell(r, col).fill = RED
r += 1
ws.cell(r, 2, "QUESTION — left as it is, same as August")
ws.cell(r, 2).font = BOLD
r += 1
aw_x = xero_row("#71188", "SYD 28.06 Services", -148)
for i, v in enumerate(["Aware SYD 28.06 Services labour revenue pre work", "#71188", "42800 Tech Consulting / 11300 WIP",
                       -148, round(src_w.cell(458, CI("BY")).value, 2), "=F%d-E%d" % (r, r), "=F%d-E%d" % (r, r),
                       "Not changed. Xero and the schedule agree as a credit, the same way August's 1,450 was left. "
                       "There is no services invoice for job 26062901 in Xero, so if the labour has not been invoiced it "
                       "would be a debit like the lines above. Please confirm.", None, None, None]):
    ws.cell(r, 2 + i, v)
for col in range(2, 13): ws.cell(r, col).fill = AMB
last = r
r += 1
ws.cell(r, 4, "TOTAL"); ws.cell(r, 4).font = BOLD
for col in "EFGL":
    ws["%s%d" % (col, r)] = "=SUM(%s%d:%s%d)" % (col, first, col, last)
tot = r
ws["J4"] = "=E4+L%d" % tot
ws["J5"] = "=E5+L%d" % tot
ws["J6"] = "=J4-J5"
r += 2
ws.cell(r, 2, "The WIP Schedule and Xero 11300 already agree at 30/09/2026, so there is no variance to clear this month. "
              "The lines above do not change the variance. They change both sides by the same amount, so the schedule "
              "and Xero must be changed together.")
r += 2
ws.cell(r, 2, "WHERE TO FIND EACH ITEM — rows are highlighted on both raw tabs").font = BOLD
r += 2
for i, h in enumerate(["Item", "Schedule row", "Xero tab row", "Colour", "Meaning"]):
    ws.cell(r, 2 + i, h).font = BOLD
for item, srow, j, txt, amt in ITEMS:
    r += 1
    xr = xero_row(j, txt, amt)
    for i, v in enumerate([item, srow, xr, "RED", "Change in Xero %s and on the schedule — confirm first" % j]):
        ws.cell(r, 2 + i, v)
r += 1
for i, v in enumerate(["Aware SYD 28.06 Services $148", 458, aw_x, "AMB", "Left as is — confirm"]):
    ws.cell(r, 2 + i, v)
r += 2
ws.cell(r, 2, "RED = change the journal in Xero (and here, the schedule as well).")
ws.cell(r + 1, 2, "AMBER = left as it is, needs a decision.")
for col, w in zip("ABCDEFGHIJKL", (3, 44, 10, 34, 12, 13, 12, 12, 70, 8, 12, 15)):
    ws.column_dimensions[col].width = w
for rr in range(first, last + 1):
    ws.cell(rr, 9).alignment = Alignment(wrap_text=True, vertical="top")
for rr in range(4, tot + 1):
    for col in "EFGHJKL":
        ws["%s%d" % (col, rr)].number_format = "#,##0.00;(#,##0.00);-"

# What I did
n = wb.create_sheet("What i did ")
notes = [
    ("How the September WIP was checked", None),
    ("Notes on how September's WIP 11300 was reconciled and what each decision is based on.", None),
    (None, "STARTING POSITION"),
    (None, "The WIP Schedule (2026-09_WIP.xlsx) closes at %s. The Xero 11300 export (1 July to 30 September) closes at %s." % (f"{w_close:,.2f}", f"{x_close:,.2f}")),
    (None, "July and August also agree month by month, so the opening balance is right."),
    (None, "LINE BY LINE"),
    (None, "Every September amount in the schedule was matched to a line in Xero. Xero combines a few lines (UBank's three, "),
    (None, "Aware's labour and freight, the APA and Aware Appspace licence lines), but the totals are the same."),
    (None, "Journals checked: #70539 (1 September reversal of August's auto-reversing lines), #71025 to #71044, #71175 to #71180, "),
    (None, "#71192 (licence releases and reclasses), #71186 (pre-production and expense accrual for October) and #71188 (WIP Journal September)."),
    (None, "BASIS FOR WHICH SIDE EACH LINE BELONGS ON"),
    (None, "Same test as August. Work done in the month but not yet invoiced is accrued revenue, so WIP goes up on the debit side."),
    (None, "Revenue invoiced before the work is deferred, so WIP goes on the credit side."),
    (None, "WHAT IS DIFFERENT FROM AUGUST"),
    (None, "In August, Xero had the pre-production lines as debits and the schedule had them as credits, so the schedule was changed."),
    (None, "In September, the schedule and Xero both have the ten pre-production lines as credits (rows 480 to 489, journals #71186 and #71188)."),
    (None, "None of the ten events was invoiced by 14 October, so by August's basis they should be debits in both places."),
    (None, "Nothing has been changed in Xero or the schedule. The lines are listed on the Journal Edit Schedule tab for your decision."),
    (None, "If they are changed, WIP 11300 and the schedule both move from a credit of 146,188.18 to a credit of 133,253.18."),
    (None, "ITEMS LEFT AS THEY ARE"),
    (None, "Aware SYD 28.06 Services (row 458, 148): left as a credit, the same way August's 1,450 was left. Flagged to confirm."),
    (None, "REVENUE TRACKER"),
    (None, "The tracker's WIP Movements and Deferrals now carry the same September lines as the schedule and Xero, and agree job by job."),
    (None, "If the ten lines are changed, change the same rows in the tracker (WIP Movements rows 191 to 200) to positive figures."),
]
for t in notes: n.append(list(t))
n["A1"].font = Font(bold=True, size=13)
for rr in range(1, n.max_row + 1):
    if n.cell(rr, 2).value and n.cell(rr, 2).value.isupper(): n.cell(rr, 2).font = BOLD
n.column_dimensions["A"].width = 4; n.column_dimensions["B"].width = 120

# Raw copies
rw = wb.create_sheet("Raw Copy of WIP Schedule ")
for row in src_w.iter_rows(min_row=1, max_row=src_w.max_row, max_col=src_w.max_column, values_only=True):
    rw.append(list(row))
for item, srow, *_ in ITEMS:
    for c in range(1, src_w.max_column + 1): rw.cell(srow, c).fill = RED
for c in range(1, src_w.max_column + 1): rw.cell(458, c).fill = AMB
rx = wb.create_sheet("Raw Xero WIP 11300")
for row in src_x.iter_rows(values_only=True):
    rx.append(list(row))
for item, srow, j, txt, amt in ITEMS:
    for c in range(1, 10): rx.cell(xero_row(j, txt, amt), c).fill = RED
for c in range(1, 10): rx.cell(aw_x, c).fill = AMB
rx.column_dimensions["C"].width = 90
wb.save(OUT)
print("WIP", w_close, "Xero", x_close, "if changed", round(w_close + 2 * 6467.5, 2))
