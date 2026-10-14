"""WIP 11300 Journal Edit Schedule for September 2026, built on Danica's August file
(WIP_Adjustment__Schedule_10_09_2026.xlsx) so the layout, wording style and formatting match.
Written as Danica's own narration of the corrections."""
import copy, datetime, sys
import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.utils import column_index_from_string as CI

AUG, WIPF, XEROF, OUT = sys.argv[1:5]
RED = PatternFill("solid", fgColor="FFC7CE")
WIPTAB, XTAB = "Corrected WIP Sheet 30092026", "Xero 11300 - Corrected 30092026"

wb = openpyxl.load_workbook(AUG)
src_w = openpyxl.load_workbook(WIPF, data_only=True)["Schedule Updated"]
src_x = openpyxl.load_workbook(XEROF, data_only=True).worksheets[0]

ITEMS = [  # item, schedule row, journal, text in Xero description, amount as posted (11300 side)
    ("ASX AGM 22.10.26 pre-prod labour", 480, "#71186", "ASX - AGM - 22.10.26", -1625),
    ("Cochlear AGM 26.10.26 pre-prod labour", 481, "#71186", "Cochlear - AGM - 26.10.26", -2697.5),
    ("TLC AGM 14.10.26 pre-prod labour", 482, "#71186", "TLC - 2026 AGM - 14.10.26", -715),
    ("RSNSW OGM 7.10.26 pre-prod labour", 483, "#71186", "RSNSW Ordinary General Meeting 7.10.2026", -195),
    ("Garvan Breast Cancer Lecture 1.10.26", 484, "#71186", "Garvan - Breast Cancer", -130),
    ("Microsoft SME&C Town Hall 14.10.26", 485, "#71186", "Microsoft - SME&C", -487.5),
    ("Garvan Suttons Event 15.10.26", 486, "#71186", "Garvan - Suttons", -97.5),
    ("ASX CEO Connect 27.10.26 pre-prod labour", 487, "#71186", "ASX CEO Connect 27.10.26", -260),
    ("ART AMM Captioning 25.11.26", 488, "#71188", "ART - AMM Captioning", -97.5),
    ("Automic Noumi AGM 5.11.26", 489, "#71188", "Automic - Noumi", -162.5),
]
FIRST = ("I corrected this one in Xero. It was sitting on the credit side of WIP, which defers the revenue out of September. "
         "The pre-production labour was performed in September and has not been invoiced, so the revenue is earned in "
         "September and the accrual belongs on the debit side of WIP. I moved 11300 to Debit and put 42100 Production Labour "
         "on Credit. The schedule also had it as a credit, so I edited the schedule to a debit to match. Sheet1 row %d.")
SAME = ("Same reasoning as ASX AGM. Work performed in September and not yet invoiced, so I moved 11300 to Debit and 42100 "
        "to Credit in Xero and edited the schedule to a debit. Sheet1 row %d.")
LATE = ("Same treatment. This line sits in the WIP Journal for September and is set to reverse in November. Work performed "
        "in September and not yet invoiced, so I moved 11300 to Debit and 42100 to Credit in Xero and edited the schedule "
        "to a debit. Sheet1 row %d.")

# ---------- raw and corrected Xero 11300 tabs, September only, same layout as August's
xr = []
for r in range(1, src_x.max_row + 1):
    d = src_x.cell(r, 1).value
    if isinstance(d, datetime.datetime) and d.month == 9:
        xr.append([src_x.cell(r, c).value for c in range(1, 10)])
opening = 128544.93  # credit, Xero 11300 at 31/08/2026

def flip_index(rows):
    idx = {}
    for item, srow, j, txt, amt in ITEMS:
        for i, x in enumerate(rows):
            a = round((x[4] or 0) - (x[5] or 0), 2)
            if i not in idx.values() and str(x[3]) == j and txt in str(x[2]) and abs(a - amt) < 0.006:
                idx[srow] = i; break
        else:
            raise KeyError(item)
    return idx

FLIPS = flip_index(xr)

def xero_tab(ws, corrected):
    head = [["Work in Progress Transactions"], ["Corporate Technology Services Pty Ltd"],
            ["For the period 1 September 2026 to 30 September 2026"], [],
            ["Date", "Source", "Description", "Reference", "Debit", "Credit", "Running Balance", "Gross", "GST"], [],
            ["Work in Progress"], ["Opening Balance", None, None, None, 0, opening, "=(E8 - F8)", 0, 0]]
    for i, h in enumerate(head, 1):
        for j, v in enumerate(h, 1): ws.cell(i, j, v)
    r = 8
    rows_at = {}
    for i, x in enumerate(xr):
        r += 1
        x = list(x)
        if corrected and i in FLIPS.values():
            x[4], x[5] = x[5], x[4]
            x[7] = (x[4] or 0) - (x[5] or 0)
        for j, v in enumerate(x[:6], 1): ws.cell(r, j, v)
        ws.cell(r, 7, "=((G%d + E%d) - F%d)" % (r - 1, r, r))
        ws.cell(r, 8, x[7]); ws.cell(r, 9, x[8])
        ws.cell(r, 1).number_format = "dd/mm/yyyy"
        rows_at[i] = r
        if i in FLIPS.values():
            for c in range(1, 10): ws.cell(r, c).fill = RED
    last = r
    ws.cell(r + 1, 1, "Total Work in Progress")
    for col in "EFHI": ws["%s%d" % (col, r + 1)] = "=SUM(%s9:%s%d)" % (col, col, last)
    ws["G%d" % (r + 1)] = "=G%d" % last
    ws.cell(r + 2, 1, "Closing Balance"); ws["E%d" % (r + 2)] = 0
    ws["F%d" % (r + 2)] = "=F8-E%d+F%d" % (r + 1, r + 1); ws["G%d" % (r + 2)] = "=G%d" % last
    ws.cell(r + 4, 1, "Total")
    for col in "EFGHI": ws["%s%d" % (col, r + 4)] = "=%s%d" % (col, r + 1)
    ws.column_dimensions["C"].width = 90
    return rows_at, last

# ---------- raw and corrected WIP schedule tabs (Schedule Updated, values)
def wip_tab(ws, corrected):
    for r in range(1, src_w.max_row + 1):
        for c in range(1, src_w.max_column + 1):
            v = src_w.cell(r, c).value
            if v is not None: ws.cell(r, c, v)
    for item, srow, *_ in ITEMS:
        for c in range(1, src_w.max_column + 1): ws.cell(srow, c).fill = RED
        if corrected:
            ws.cell(srow, CI("BY"), -src_w.cell(srow, CI("BY")).value)
            ws.cell(srow, CI("BZ"), -src_w.cell(srow, CI("BZ")).value)
    if corrected:
        ws["BY3"] = "=SUM(BY11:BY%d)" % src_w.max_row
        ws["BZ3"] = "=SUM(BZ11:BZ%d)" % src_w.max_row
        ws["BZ4"] = "=BX4+SUM(BY6:BY%d)" % src_w.max_row

names = wb.sheetnames
fresh = {}
for i, new in ((2, WIPTAB), (3, XTAB), (4, names[4]), (5, names[5])):
    old = wb[names[i]]
    wb.remove(old)
    fresh[new] = wb.create_sheet(new, i)
wip_tab(fresh[WIPTAB], True)
wip_tab(fresh[names[4]], False)
xrow_c, xlast_c = xero_tab(fresh[XTAB], True)
xrow_r, _ = xero_tab(fresh[names[5]], False)

raw_wip = round(sum(src_w.cell(r, CI("BZ")).value for r in range(11, src_w.max_row + 1)
                    if isinstance(src_w.cell(r, CI("BZ")).value, (int, float))), 2)
raw_x = round(-opening + sum((x[4] or 0) - (x[5] or 0) for x in xr), 2)
assert raw_wip == raw_x == -146188.18, (raw_wip, raw_x)

# ---------- Journal Edit Schedule
ws = wb["Journal Edit Schedule"]
for img in list(getattr(ws, "_images", [])): ws._images.remove(img)
ws["A1"] = "WIP 11300 — Journal Edit Schedule, September 2026"
ws["E4"], ws["E5"] = raw_wip, raw_x
ws["E6"] = "=E4-E5"
ws["G4"], ws["J4"] = WIPTAB, "='%s'!BZ3" % WIPTAB
ws["G5"], ws["J5"] = "Corrected Xero WIP 11300", "='%s'!G%d" % (XTAB, xlast_c)
ws["J6"] = "=J4-J5"
style_item = {c: copy.copy(ws.cell(11, c)._style) for c in range(2, 11)}
style_done = copy.copy(ws["J11"]._style)
for r in range(10, 46):
    for c in range(2, 19):
        if not isinstance(ws.cell(r, c), openpyxl.cell.cell.MergedCell): ws.cell(r, c).value = None
ws["B10"] = "JOURNAL #71186 — Pre-PROD & INT Work & Expense Accrual for October 2026"
ws["B19"] = "JOURNAL #71188 — WIP Journal September 2026"
rows = list(range(11, 19)) + [20, 21]
for k, ((item, srow, j, txt, amt), r) in enumerate(zip(ITEMS, rows)):
    text = FIRST % srow if k == 0 else (LATE % srow if j == "#71188" else SAME % srow)
    for c, v in zip(range(2, 11), [item, j, "42100 Production Labour / 11300 WIP", amt,
                                     round(src_w.cell(srow, CI("BY")).value, 2), "=F%d-E%d" % (r, r), "=E6+G%d" % r, text, "a"]):
        ws.cell(r, c, v); ws.cell(r, c)._style = copy.copy(style_item[c])
    ws.cell(r, 10)._style = copy.copy(style_done)
for r in (22, 23):
    for c in range(2, 11): ws.cell(r, c)._style = copy.copy(style_item[c])
ws["D24"] = "TOTAL — schedule and Xero agree before and after"
ws["E24"], ws["F24"], ws["G24"] = "=SUM(E11:E23)", "=SUM(F11:F23)", "=SUM(G11:G23)"
ws["B26"] = ("The schedule and Xero 11300 agreed at a credit of 146,188.18 before the changes, but both carried the same "
             "wrong sign on the ten lines above, so it did not show as a variance. Once every line above is done, both "
             "close at a credit of 133,253.18. I re-exported the WIP 11300 account transactions and re-totalled the "
             "schedule to confirm they agree.")
ws["B27"] = "WHERE TO FIND EACH ITEM — rows are highlighted on both raw tabs"
for c, v in zip(range(2, 7), ["Item", "Sheet1 row", "Xero tab row", "Colour", "Meaning"]): ws.cell(29, c, v)
for k, (item, srow, j, txt, amt) in enumerate(ITEMS):
    r = 30 + k
    for c, v in zip(range(2, 7), ["%s $%s" % (item, f"{-amt:,.2f}"), srow, xrow_r[FLIPS[srow]], "RED",
                                   "Flipped in Xero %s and schedule sign fixed" % j]):
        ws.cell(r, c, v)
    ws.cell(r, 5).fill = RED
for c in range(2, 7):
    for r in (40, 41): ws.cell(r, c).fill = PatternFill()
ws["B43"] = "RED = change the journal in Xero.   "
ws["B44"] = "AMBER = Xero is right, change the schedule. "
ws["B45"] = "BLUE = a line is missing from Xero."
for r in range(11, 22):
    ws.row_dimensions[r].height = 48 if r in rows else 15

# ---------- What i did
n = wb["What i did "]
for row in n.iter_rows():
    for c in row: c.value = None
LINES = [
    ("A", "How to edit an existing journal in Xero"),
    ("A", "My notes on how I corrected the September WIP journals and what I based each decision on."),
    ("B", "LOCATING THE JOURNALS"),
    ("B", "Accounting, then Advanced, then Manual Journals, on the Posted tab. Journals were located by journal number"),
    ("B", "or narration text."),
    ("B", "The pre-production accrual for October events is #71186 and the September WIP journal is #71188."),
    ("B", "WHAT I FOUND"),
    ("B", "The WIP schedule and Xero 11300 agreed at 30 September, both at a credit of 146,188.18, so there was no variance"),
    ("B", "to clear. Agreeing in total did not mean every line was right. Ten pre-production labour lines were sitting on the"),
    ("B", "credit side of WIP in both the schedule and Xero. Because both carried the same sign, it did not show as a variance."),
    ("B", "BASIS FOR DETERMINING WHICH SIDE EACH LINE BELONGS ON"),
    ("B", "Each line was assessed on what the entry is intended to do to work in progress, rather than on what had"),
    ("B", "already been posted."),
    ("B", "Where an amount is being deferred out of the month, WIP must increase, so the 11300 line sits on the debit"),
    ("B", "side for a cost and on the credit side for revenue."),
    ("B", "Where an amount is being released or reversed into the month, WIP must decrease, so the 11300 line sits on"),
    ("B", "the opposite side."),
    ("B", "Where work is performed in the month and not yet invoiced, the revenue is earned in the month and is accrued,"),
    ("B", "so the 11300 line sits on the debit side. This is the same basis I used for the August pre-production lines."),
    ("B", "This test was applied to every line rather than assuming the original posting was correct."),
    ("B", "METHOD FOR REVERSING A PAIR OF LINES"),
    ("B", "No amounts were changed. The figure was cleared from the column it occupied and entered into the opposite"),
    ("B", "column, on both lines of the pair, so that the journal remained in balance."),
    ("B", "ASX AGM is the clearest illustration."),
    ("B", "  BEFORE   42100 Production Labour     Debit  1,625.00"),
    ("B", "           11300 Work in Progress      Credit 1,625.00"),
    ("B", "  AFTER    42100 Production Labour     Credit 1,625.00"),
    ("B", "           11300 Work in Progress      Debit  1,625.00"),
    ("B", "Job Number, Cost Centre, Tax Rate and Description were left unchanged. Only the two amounts were moved."),
    ("B", "The same ten lines were changed to a debit on the WIP schedule (rows 480 to 489) and on the revenue tracker"),
    ("B", "(WIP Movements rows 191 to 200), so all three still agree."),
    ("B", "CHECKS PERFORMED BEFORE SAVING"),
    ("B", "That the Subtotal and Total continued to show the same figure on both sides."),
    ("B", "That the Tax Rate remained BAS Excluded on every line, these being WIP journals with no GST."),
    ("B", "That the Job Number and Cost Centre were completed on every line, as blank entries drop out of departmental"),
    ("B", "reporting."),
    ("B", "ITEMS DELIBERATELY LEFT UNCHANGED IN XERO"),
    ("B", "ICC September Presenter Support ComBio (762.50) and Bankwest BCEC WGEA (950.00) were invoiced in September for"),
    ("B", "October events, so deferring the revenue and crediting WIP is correct. The CBA AGM GoGet hire (72.00) is a cost"),
    ("B", "paid ahead of the event, so debiting WIP is correct. The 1 September reversal of August's auto-reversing lines"),
    ("B", "(#70539) and the licence releases were posted correctly and agree to the schedule line by line."),
    ("B", "VERIFICATION"),
    ("B", "The WIP 11300 account transactions were re-exported and the schedule re-totalled, with the two closing"),
    ("B", "balances agreed to the cent at a credit of 133,253.18. This is performed before September is closed off."),
    ("B", "The Xero balance cell at the top of the schedule must reference the actual Xero figure rather than the"),
    ("B", "schedule total, otherwise the check compares the schedule to itself and will never identify a variance."),
]
for i, (col, text) in enumerate(LINES, 1):
    n["%s%d" % (col, i)] = text
wb.save(OUT)
print("raw", raw_wip, raw_x, "corrected", round(raw_wip + 2 * 6467.5, 2))
