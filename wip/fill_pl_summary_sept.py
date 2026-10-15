"""Fill the Sep summary sheet of Danica's PL and Revenue summary, splitting each income account by cost
centre the same way August's "Aug-26 P&L by Cost Centre" was built: each invoice line goes to the cost
centre of its tracker Finance row (shared Production/Video invoices split using the video lines first),
plus the WIP and deferral journals at the cost centre of their tracker line."""
import pickle, sys
import openpyxl
from openpyxl.styles import Font

SRC, RES, OUT = sys.argv[1], sys.argv[2], sys.argv[3]
res, tabtot = pickle.load(open(RES, "rb"))
wb = openpyxl.load_workbook(SRC)
ws = wb["Sep summary"]
COLS = {"ONSITE": "C", "PRODUCTION": "D", "VIDEO": "E", "INTEGRATION": "F", "CONSULTING": "G", "CTS": "H"}
TAB = {"45015": "4515"}
ws["K1"], ws["L1"] = "Xero (account tab)", "Check"
for c in ("K1", "L1"): ws[c].font = Font(bold=True)
for r in range(2, 19):
    code = str(ws["A%d" % r].value).strip()
    for col in "CDEFGHI": ws["%s%d" % (col, r)] = 0
    ws["J%d" % r] = "=SUM(C%d:I%d)" % (r, r)
    split = res.get(code, {})
    for cc, v in split.items():
        ws["%s%d" % (COLS[cc], r)] = v
    tab = TAB.get(code, code)
    if tab in wb.sheetnames:
        t = wb[tab]
        tr = max(i for i in range(1, t.max_row + 1) if t.cell(i, 1).value == "Total")
        ws["K%d" % r] = "='%s'!F%d-'%s'!E%d" % (tab, tr, tab, tr)
    else:
        ws["K%d" % r] = 0
    ws["L%d" % r] = '=IF(ROUND(J%d-K%d,2)=0,"Ties","CHECK "&TEXT(J%d-K%d,"#,##0.00"))' % (r, r, r, r)
for col in "CDEFGHIJK":
    ws["%s19" % col] = "=SUM(%s2:%s18)" % (col, col)
ws["L19"] = '=IF(ROUND(J19-K19,2)=0,"Ties","CHECK")'
for r in range(2, 20):
    for col in "CDEFGHIJK": ws["%s%d" % (col, r)].number_format = '#,##0.00;[Red]-#,##0.00;-'

ws["A21"] = "Type these into Month-End column H (month = Sep-26):"; ws["A21"].font = Font(bold=True)
for i, (cc, col) in enumerate(list(COLS.items()) + [("No cost centre set", None)]):
    ws["B%d" % (22 + i)] = cc
    ws["C%d" % (22 + i)] = "=%s19" % col if col else 0
    ws["C%d" % (22 + i)].number_format = '#,##0.00;[Red]-#,##0.00;-'
NOTES = [
    "Notes",
    "Built the same way as August: each September invoice line in Xero (by account) goes to the cost centre of its Finance row in the tracker, plus the WIP journals (#70539, #71222, #71230) and the licence and deferral journals at the cost centre of their tracker line. Each account ties to the account transactions on its tab.",
    "Invoices shared between Production and Video (INV-10623, 10624, 10627, 10632, 10663, 10670) are split using the video lines first: Video Labour lines and the Equipment Hire line called Video go to VIDEO. Each split agrees to the tracker's Video row for that invoice.",
    "42150 Video Labour has 2,125.00 on PRODUCTION. These are video labour lines on invoices whose only Finance row is Production, the same as August's 3,610.00.",
    "INV-10652 (Aware Super maintenance, 11,000.00, ONSITE) and INV-10669 (Global Philanthropic B-Roll, 1,249.00, VIDEO) are in Xero but not typed on the tracker's Finance tab yet. They are allocated by their job's cost centre.",
    "CN-10666 (DTTL, -1,186.15) is on 42300 Installation Labour and sits on ONSITE because its Finance row (ONS-0051, job 2347) is Onsite. The credit note description is DTTL Melbourne Bluetooth Audio, which is an Integration job. Check the cost centre on the credit note in Xero.",
    "42100 Production Labour is 42,285.00 on the account tab. This includes the ten pre-production lines after they were corrected to a credit to revenue (#71222 and #71230). The earlier P&L run showed 29,350.00 before the correction.",
    "This split uses the tracker's cost centre for each invoice. To test that Xero is tagged the same way, run the P&L in Xero by the Cost Centres tracking category and compare.",
]
for i, n in enumerate(NOTES):
    ws["A%d" % (31 + i)] = n
ws["A31"].font = Font(bold=True)
ws.column_dimensions["K"].width = 16; ws.column_dimensions["L"].width = 14
wb.save(OUT)
tot = {cc: round(sum(v.get(cc, 0) for v in res.values()), 2) for cc in COLS}
print(tot, round(sum(tot.values()), 2), round(sum(tabtot.values()), 2))
