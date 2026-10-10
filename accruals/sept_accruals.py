"""September 2026 expense accruals: update the 2-1210 reconciliation and write the Xero journal.

Amounts are ex GST, as the bills are recorded in Xero.
"""
import csv, re, sys, zipfile

SRC, OUTX, OUTCSV = sys.argv[1], sys.argv[2], sys.argv[3]
EDITS = {  # MAIN sheet
    "C125": "3417",                    # August accounting fees as posted in Xero (file had 1,800)
    "C127": "3417-342",                # accrual less ASIC annual review fee (GST free)
    "F127": "300-F126",                # electricity accrual balance held at 300
    "I127": "10000-184.06-582.73",     # 10,000 accrual less Swamp and Tec Art bills
}
z = zipfile.ZipFile(SRC)
sheet = z.read("xl/worksheets/sheet2.xml").decode("utf-8")
for ref, f in EDITS.items():
    sheet, n = re.subn(r'(<c r="%s"[^>]*>)<f>[^<]*</f><v>[^<]*</v>' % ref, r"\g<1><f>%s</f>" % f, sheet)
    assert n == 1, ref
wb = z.read("xl/workbook.xml").decode("utf-8")
if "fullCalcOnLoad" not in wb:
    wb = re.sub(r"<calcPr([^>]*?)/>", r'<calcPr\1 fullCalcOnLoad="1"/>', wb)
with zipfile.ZipFile(OUTX, "w", zipfile.ZIP_DEFLATED) as o:
    for it in z.infolist():
        data = sheet.encode() if it.filename == "xl/worksheets/sheet2.xml" else \
               wb.encode() if it.filename == "xl/workbook.xml" else z.read(it.filename)
        o.writestr(it, data)

aug_elec_balance = 410.63
lines = [("61050", round(3417 - 342, 2)), ("61901", 1200.00),
         ("64450", round(300 - aug_elec_balance, 2)), ("61730", round(10000 - 184.06 - 582.73, 2))]
lines.append(("21210", -round(sum(a for _, a in lines), 2)))
NARR = "Accrual Expenses September 2026"
with open(OUTCSV, "w", newline="") as fh:
    w = csv.writer(fh)
    w.writerow(["*Narration", "*Date", "Description", "*AccountCode", "*TaxRate", "*Amount",
                "TrackingName1", "TrackingOption1", "TrackingName2", "TrackingOption2"])
    for acct, amt in lines:
        w.writerow([NARR, "30/09/2026", NARR, acct, "BAS Excluded", "%.2f" % amt,
                    "Cost Centres", "CTS", "Job Numbers", "9000 - OFFICE / ADMIN [CTS]"])
assert abs(sum(a for _, a in lines)) < 0.005
for l in lines: print(l)
