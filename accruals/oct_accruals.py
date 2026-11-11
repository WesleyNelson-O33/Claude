"""October 2026 expense accruals: add the October block to the 2-1210 reconciliation and write the Xero journal.
Edits the sheet XML directly; the file carries threaded comments and drawings."""
import csv, os, re, zipfile
from lxml import etree
from openpyxl.utils import column_index_from_string as CI, get_column_letter as CL
SRC = "/root/.claude/uploads/fc902e4e-1220-55ba-999f-a6d0cad6db14/192fc465-2026-09_Accrual_Expenses_2-1210_Reconciliation.xlsx"
OUT = "/home/user/Claude/accruals/out"
NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
def q(t): return "{%s}%s" % (NS, t)
def colname(i): return CL(i)
def colnum(s): return CI(s)
exec(open("/tmp/claude-0/-home-user-Claude/fc902e4e-1220-55ba-999f-a6d0cad6db14/scratchpad/xmlmachinery.py").read())

STORE_DJ, CABLE_CHICK = 1096.36, 2477.73      # October instant asset bills, ex GST, as posted to 61730
LINES = [("61050", "Accounting Fees", 3417.00), ("61901", "Business Advisory", 1200.00),
         ("61730", "Instant Asset Write Off", round(10000 - STORE_DJ - CABLE_CHICK, 2))]
XERO_OPENING = 24250.05

os.makedirs(OUT, exist_ok=True)
z = zipfile.ZipFile(SRC); parts = {n: z.read(n) for n in z.namelist()}; order = [i.filename for i in z.infolist()]; z.close()
ss = Strings(parts["xl/sharedStrings.xml"]); ws = Sheet(parts["xl/worksheets/sheet2.xml"], ss)
def styled(r, spec, like):
    row = ws.set_row(r, spec); src = Sheet.cells(ws.rowel(like))
    for col, c in Sheet.cells(row).items():
        if src.get(col) is not None and src[col].get("s"): c.set("s", src[col].get("s"))
ws.set_row(129, [(1, "s", "Accruals (October)"), (3, "n", 3417), (5, "n", 1200),
                 (9, "f", "10000-%s-%s" % (STORE_DJ, CABLE_CHICK))])
styled(129, [(1, "s", "Accruals (October)"), (3, "n", 3417), (5, "n", 1200), (9, "f", "10000-%s-%s" % (STORE_DJ, CABLE_CHICK))], 127)
spec = [(1, "s", "31 October 2026 Balance")] + [(c, "f", "+%s129+%s128" % (colname(c), colname(c))) for c in range(2, 11)] + [(11, "f", "SUM(B130:J130)")]
styled(130, spec, 128)
styled(132, [(5, "s", "Xero 2-1210 at 31 October before the October journal"), (6, "n", XERO_OPENING)], 128)
styled(133, [(5, "s", "After the October journal, should equal K130"), (6, "f", "F132+C129+E129+F129+I129+J129")], 128)
ws.fix_dimension(); parts["xl/worksheets/sheet2.xml"] = ws.dump(); parts["xl/sharedStrings.xml"] = ss.dump()
wbx = parts["xl/workbook.xml"].decode("utf8")
if "fullCalcOnLoad" not in wbx: wbx = re.sub(r"<calcPr([^>]*?)/>", r'<calcPr\1 fullCalcOnLoad="1"/>', wbx, count=1)
parts["xl/workbook.xml"] = wbx.encode("utf8")
dst = os.path.join(OUT, "2026-10 Accrual Expenses 2-1210 Reconciliation.xlsx")
with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as o:
    for n in order: o.writestr(n, parts[n])

narr = "Expense Accrual October 2026"
with open(os.path.join(OUT, "2026-10_Accrual_Expenses_Journal_Xero.csv"), "w", newline="") as fh:
    w = csv.writer(fh)
    w.writerow(["*Narration", "*Date", "Description", "*AccountCode", "*TaxRate", "*Amount", "TrackingName1", "TrackingOption1", "TrackingName2", "TrackingOption2"])
    for acct, name, amt in LINES:
        for a, s in ((acct, amt), ("21210", -amt)):
            w.writerow([narr, "31/10/2026", "%s - %s" % (narr, name), a, "BAS Excluded", "%.2f" % s, "Cost Centres", "CTS", "Job Numbers", "9000 - OFFICE / ADMIN [CTS]"])
print("journal total", round(sum(a for _, _, a in LINES), 2), "| closing", round(XERO_OPENING + sum(a for _, _, a in LINES), 2))
