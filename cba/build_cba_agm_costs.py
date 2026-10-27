"""CBA AGM 2026 (job 2506208 / 2506208V) supplier costs and staff hours for Duncan's invoice request.
Sources: Xero bills and purchase orders read 27 Oct 2026, the FY27 Revenue Tracker 29092026 WIP Movements tab,
and the September and August PRD Labour Hour Reports and EH raw data as read in earlier chats."""
import sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

OUT = sys.argv[1]
F = "Arial"
HEAD = PatternFill("solid", fgColor="1F3864")
AMBER = PatternFill("solid", fgColor="FFE699")
RED = PatternFill("solid", fgColor="FFC7CE")
GREY = PatternFill("solid", fgColor="F2F2F2")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
MONEY = '#,##0.00;(#,##0.00);"-"'
HRS = '0.00;(0.00);"-"'

wb = openpyxl.Workbook()


def sheet(title, first=False):
    ws = wb.active if first else wb.create_sheet()
    ws.title = title
    ws.sheet_view.showGridLines = False
    return ws


def text(ws, ref, v, bold=False, size=10, wrap=False, fill=None, color="000000"):
    c = ws[ref]
    c.value = v
    c.font = Font(name=F, bold=bold, size=size, color=color)
    c.alignment = Alignment(wrap_text=wrap, vertical="top")
    if fill:
        c.fill = fill
    return c


def table(ws, row, headers, rows, widths, fmts):
    for i, h in enumerate(headers, 1):
        c = ws.cell(row, i, h)
        c.font = Font(name=F, bold=True, color="FFFFFF", size=10)
        c.fill = HEAD
        c.alignment = Alignment(wrap_text=True, vertical="center")
        c.border = BOX
    for r, vals in enumerate(rows, row + 1):
        fill = vals[-1] if isinstance(vals[-1], PatternFill) else None
        vals = vals[:-1] if fill else vals
        for i, v in enumerate(vals, 1):
            c = ws.cell(r, i, v)
            c.font = Font(name=F, size=10)
            c.alignment = Alignment(wrap_text=True, vertical="top")
            c.border = BOX
            if fmts.get(i):
                c.number_format = fmts[i]
            if fill:
                c.fill = fill
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = ws.cell(row + 1, 1)
    return row + len(rows)


def total(ws, row, label_col, cols, first, last, label="Total"):
    c = ws.cell(row, label_col, label)
    c.font = Font(name=F, bold=True, size=10)
    for col in cols:
        L = get_column_letter(col)
        c = ws.cell(row, col, "=SUM(%s%d:%s%d)" % (L, first, L, last))
        c.font = Font(name=F, bold=True, size=10)
        c.number_format = MONEY
        c.border = Border(top=Side(style="thin"), bottom=Side(style="double"))


# ---------------------------------------------------------------- Supplier Costs
sc = sheet("Supplier Costs", first=True)
text(sc, "A1", "CBA AGM 2026, job 2506208. Supplier bills in Xero", bold=True, size=13)
text(sc, "A2", "Found by searching every Xero bill from 1 June 2025 for 2506208 or CBA AGM 2026 in the bill number "
               "or line description. Read from Xero on 27 October 2026. Amounts are what Xero holds.", wrap=True)
sc.merge_cells("A2:K2"); sc.row_dimensions[2].height = 30
B = [
    ("2026-07-03", "Red Bee Media Pty Ltd", "INV7060003977", "Testing of live remote captioning for Viostream", "51400", 365.00, 36.50, "Paid", ""),
    ("2026-09-30", "Goget Carshare", "INV10498230926", "Tyler Wood, CBA AGM, Chair Review, 21/09/2026", "58115", 72.00, 7.20, "Paid",
     "One line of a 572.29 bill. The other lines are other jobs."),
    ("2026-10-01", "Elan Projects t/as PopUp WiFi", "IAU35330", "Pop-up wifi hire", "51200 5,720.00 / 51400 780.00", 6500.00, 650.00, "Paid", ""),
    ("2026-10-01", "Pop-Up Gardens", "INV2070", "Pop-up plant hire", "51400 2,125.00 / 58215 780.00", 2905.00, 290.50, "Paid", ""),
    ("2026-10-01", "Valiant Corporate", "INV87154", "Indoor chair hire", "51100 1,424.00 / 52500 99.68 / 58215 655.04", 2178.72, 217.87, "Paid", ""),
    ("2026-10-07", "Tony Taouk (T/A New Office)", "CTS-20261007", "Acoustic screen hire", "51200", 3072.73, 307.27, "Paid", ""),
    ("2026-10-07", "Officeworks Superstores Pty Ltd", "INV632570641", "HP keyboard and mouse combo", "51100", 98.18, 9.82, "Unpaid", ""),
    ("2026-10-14", "Start Webcast Pty Ltd", "INV-202222", "Equipment and services hire", "51315 7,315.00 / 51200 4,309.60", 11624.60, 1162.46, "Unpaid",
     "Bill says DL/TW yet to approve.", AMBER),
    ("2026-10-19", "Access Innovation Media Pty Ltd", "INV22272", "Services hire", "51400", 240.00, 24.00, "Unpaid", ""),
    ("2026-10-19", "Access Innovation Media Pty Ltd", "INV22273", "Services hire", "51400", 840.00, 84.00, "Unpaid", ""),
    ("2026-10-19", "Access Innovation Media Pty Ltd", "INV22274", "Services hire", "51400", 620.00, 62.00, "Unpaid", ""),
    ("2026-10-19", "Access Innovation Media Pty Ltd", "INV22275", "Services hire", "51400", 150.00, 15.00, "Unpaid", ""),
]
hdr = ["Bill date", "Supplier", "Bill number", "Description", "Account", "Ex GST", "GST", "Inc GST", "Status", "Notes"]
rows = []
for r, b in enumerate(B, 5):
    fill = b[9:] if len(b) > 9 else ()
    rows.append(list(b[:7]) + ["=F%d+G%d" % (r, r)] + list(b[7:9]) + list(fill))
last = table(sc, 4, hdr, rows, [11, 30, 16, 40, 30, 12, 11, 12, 9, 40], {6: MONEY, 7: MONEY, 8: MONEY})
total(sc, last + 1, 5, [6, 7, 8], 5, last, "Total bills")
r = last + 3
text(sc, "E%d" % r, "Paid", bold=True)
sc["F%d" % r] = '=SUMIFS(F5:F%d,I5:I%d,"Paid")' % (last, last)
text(sc, "E%d" % (r + 1), "Unpaid", bold=True)
sc["F%d" % (r + 1)] = '=SUMIFS(F5:F%d,I5:I%d,"Unpaid")' % (last, last)
for x in (r, r + 1):
    sc["F%d" % x].number_format = MONEY; sc["F%d" % x].font = Font(name=F, size=10)

r += 3
text(sc, "A%d" % r, "Purchase order not yet billed", bold=True, size=12)
po_row = r + 2
table(sc, r + 1, ["PO date", "Supplier", "PO number", "Reference", "Account", "Ex GST", "GST", "Inc GST", "Status", "Notes"],
      [["2026-10-09", "Encore Event Technologies Pty Ltd", "PO-00016949", "Q-000000985630 - CBA 2026 AGM", "Not shown by Xero",
        18318.43, 1831.93, "=F%d+G%d" % (po_row, po_row), "Authorised",
        "No Encore bill in Xero yet. Final cost could differ from the PO.", AMBER]],
      [11, 30, 16, 40, 30, 12, 11, 12, 9, 40], {6: MONEY, 7: MONEY, 8: MONEY})
sc.freeze_panes = "A5"
r = po_row + 2
text(sc, "E%d" % r, "Bills plus PO", bold=True)
for col in "FGH":
    c = sc["%s%d" % (col, r)]
    c.value = "=%s%d+%s%d" % (col, last + 1, col, po_row)
    c.number_format = MONEY; c.font = Font(name=F, bold=True, size=10)
text(sc, "A%d" % (r + 2), "Not checked: bills coded to job 2506208 by tracking only, with no job number or AGM in the text. "
                          "Reconciled spend money and expense claims could not be searched. Nothing in Xero mentions the video job 2506208V.",
     wrap=True, fill=AMBER)
sc.merge_cells("A%d:J%d" % (r + 2, r + 2)); sc.row_dimensions[r + 2].height = 30

# ---------------------------------------------------------------- Staff Hours
sh = sheet("Staff Hours")
text(sh, "A1", "CBA AGM 2026, job 2506208. Staff hours", bold=True, size=13)
text(sh, "A2", "Monthly totals only. No day-by-day dates were in any source I have. The dates by person must come from the "
               "Employment Hero timesheet export for job 2506208 and 2506208V.", wrap=True, fill=AMBER)
sh.merge_cells("A2:G2"); sh.row_dimensions[2].height = 30
H = [
    ("Sep 2026", "Tyler Wood", "2506208 [PRD]", 41.00, "September 2026 PRD Labour Hour Report, PROD tab", "Only rows 1 to 18 of 71 were seen."),
    ("Sep 2026", "Duncan Lugstein", "2506208 [PRD]", 8.25, "September 2026 PRD Labour Hour Report, PROD tab", "Only rows 1 to 18 of 71 were seen."),
    ("Aug 2026", "Duncan Lugstein", "2506208 [PRD]", 2.50, "August 2026 PRD Labour Hour Report", "Only rows 6 to 9 were seen."),
    ("Aug 2026", "Not split by person", "2506208 [PRD]", 7.50, "Employment Hero raw data (10.00 hours in August)",
     "10.00 August hours in EH, less Duncan's 2.50. Who worked them is unknown.", AMBER),
    ("Month unknown", "Daniel Sobkowski", "2506208 [PRD]", 12.25, "ALL DEPT Labour Hours Report, CONS Other CC tab",
     "Month filter was (Multiple Items). May overlap the months above.", AMBER),
    ("Oct 2026", "Not available", "2506208 [PRD]", None, "None", "No October hours seen. This covers the event itself.", RED),
]
hdr = ["Month", "Employee", "Job", "Hours", "Source", "Notes"]
last = table(sh, 4, hdr, [list(h) for h in H], [14, 22, 15, 10, 46, 52], {4: HRS})
c = sh.cell(last + 1, 3, "Total hours seen"); c.font = Font(name=F, bold=True, size=10)
c = sh.cell(last + 1, 4, "=SUM(D5:D%d)" % last); c.font = Font(name=F, bold=True, size=10); c.number_format = HRS

r = last + 4
text(sh, "A%d" % r, "Pre-production labour accrued to WIP on the revenue tracker", bold=True, size=12)
text(sh, "A%d" % (r + 1), "FY27 Revenue Tracker 29092026, WIP Movements tab. Production at $130 per hour, video at $95 per hour.", wrap=True)
sh.merge_cells("A%d:F%d" % (r + 1, r + 1))
W = [
    ("Jun 2026", "2506208 [PRD]", "Migrated opening balance", 6390.00, "WIP Movements row 18", ""),
    ("Jun 2026", "2506208V [VID]", "Migrated opening balance", 142.50, "WIP Movements row 17", ""),
    ("Jul 2026", "2506208 [PRD]", "Accrual, unbilled work", 2535.00, "WIP Movements row 19", ""),
    ("Aug 2026", "2506208 [PRD]", "Accrual, unbilled work", 1787.50, "WIP Movements row 20",
     "Does not match the 10.00 August hours in EH. Check before billing.", AMBER),
    ("Sep 2026", "2506208 [PRD]", "Accrual, unbilled work", 6402.50, "WIP Movements row 21",
     "Tracker shows 6,042.50, a typo already on the Fix on tracker sheet. 6,402.50 matches the 49.25 hours above.", AMBER),
]
first = r + 3
last = table(sh, r + 2, ["Month", "Job", "Type", "Amount ex GST", "Source", "Notes"], [list(w) for w in W],
             [14, 22, 15, 14, 46, 52], {4: MONEY})
total(sh, last + 1, 3, [4], first, last, "Total accrued")
sh.freeze_panes = "A5"

# ---------------------------------------------------------------- Unclear
uc = sheet("Unclear Items")
text(uc, "A1", "CBA AGM items not counted above", bold=True, size=13)
text(uc, "A2", "These say CBA AGM with no year or job number, or the labels conflict. By date they look like the 2025 AGM "
               "(job 3059), but that is not proof. They are not in the totals.", wrap=True)
uc.merge_cells("A2:H2"); uc.row_dimensions[2].height = 30
U = [
    ("2025-07-02", "GM Taxi Pay", "CBA AGM", "Cab from hotel to Woolloongabba", "58115", 16.23, 1.62, "Paid"),
    ("2025-10-01", "NW Group Australasia", "019453", "CBA AGM Rehearsal (PO-00016795)", "51200", 1446.38, 144.64, "Paid"),
    ("2025-10-09", "Hire Intelligence", "PO-00016800", "CS81531, CBA AGM Additional. Bill CS162740 says CBA AGM 2025", "", 470.00, 47.00, "Billed"),
    ("2025-10-31", "Goget Carshare", "10498231025", "Eight CBA AGM lines, 7 to 27 October 2025", "58115", 639.92, 64.00, "Paid"),
    ("2026-06-13", "Swan Taxis Pty Ltd", "13062026", "Bill number says CBA AGM 2027, line says CBA AGM 2026", "58115 / 59015", 57.28, 5.72, "Paid"),
]
last = table(uc, 4, ["Date", "Supplier", "Bill / PO", "Description", "Account", "Ex GST", "GST", "Status"],
             [list(u) for u in U], [11, 24, 14, 50, 14, 12, 11, 10], {6: MONEY, 7: MONEY})
uc.freeze_panes = "A5"

# ---------------------------------------------------------------- What I need
nd = sheet("What I need", first=False)
wb.move_sheet(nd, offset=-3)
text(nd, "A1", "CBA AGM 2026 invoice. Status and what is still needed", bold=True, size=13)
L = [
    ("Supplier bills in Xero", "='Supplier Costs'!F%d" % (len(B) + 5), "Ex GST, 12 bills"),
    ("Encore PO not yet billed", "='Supplier Costs'!F%d" % po_row, "Ex GST, PO-00016949"),
    ("Bills plus Encore PO", "='Supplier Costs'!F%d" % (po_row + 2), "Ex GST"),
    ("Staff hours seen so far", "='Staff Hours'!D%d" % (len(H) + 5), "Hours. No dates by day. October missing."),
]
for i, (a, f, n) in enumerate(L, 3):
    text(nd, "A%d" % i, a, bold=True)
    c = nd["B%d" % i]; c.value = f; c.font = Font(name=F, size=10)
    c.number_format = HRS if "hours" in a else MONEY
    text(nd, "C%d" % i, n)
text(nd, "A9", "To finish this for Duncan", bold=True, size=12)
N = [
    "1. Employment Hero timesheet export for jobs 2506208 and 2506208V, every date from the start of the job to today. This gives the hours by person and date. October is the big gap.",
    "2. Encore's final bill for PO-00016949, or confirm the PO amount is the figure to use.",
    "3. Start Webcast INV-202222 is waiting on DL/TW approval. Confirm the 11,624.60 ex GST is right.",
    "4. Any October GoGet, taxi, travel or expense claims for the event week. None are in Xero yet.",
    "5. The charge-out rates Duncan wants on the invoice. The tracker uses $130 production and $95 video.",
    "6. Confirm the Unclear Items tab can be left out. They look like the 2025 AGM.",
]
for i, n in enumerate(N, 10):
    text(nd, "A%d" % i, n, wrap=True)
    nd.merge_cells("A%d:C%d" % (i, i)); nd.row_dimensions[i].height = 30
nd.column_dimensions["A"].width = 34; nd.column_dimensions["B"].width = 16; nd.column_dimensions["C"].width = 60

wb.save(OUT)
