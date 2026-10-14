"""Bring the WIP side of the FY27 Revenue Tracker into line with Xero 11300 for September 2026.

Works on Danica's 29/09/2026 tracker. Enters the September WIP lines as posted in Xero
(journals 70539, 71186 and 71188), corrects three September amounts, sets the Aware
Appspace deferral to Xero's 12-month schedule, adds job 2609803 to the job list, and
types the WIP Schedule and Xero 11300 balances into the Month-End and WIP Summary checks.
Edits the sheet XML directly so tables, comments and formulas survive.
"""
import datetime, re, sys, zipfile
from collections import defaultdict
from lxml import etree
import openpyxl
from openpyxl.utils import column_index_from_string as CI

SRC, WIPSCHED, OUT = sys.argv[1], sys.argv[2], sys.argv[3]
NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
N = "{%s}" % NS
PATHS = {"Lists": "xl/worksheets/sheet3.xml", "Month-End": "xl/worksheets/sheet4.xml",
         "WIP Movements": "xl/worksheets/sheet12.xml", "WIP Summary": "xl/worksheets/sheet13.xml",
         "Deferrals": "xl/worksheets/sheet14.xml"}
SEP = 46295  # 30/09/2026

def colnum(c):
    n = 0
    for ch in c: n = n * 26 + ord(ch) - 64
    return n

class Sheet:
    def __init__(self, xml):
        self.root = etree.fromstring(xml)
        self.data = self.root.find(N + "sheetData")
        self.rows = {int(r.get("r")): r for r in self.data.findall(N + "row")}
    def row(self, r):
        if r not in self.rows:
            e = etree.Element(N + "row"); e.set("r", str(r))
            after = [k for k in self.rows if k < r]
            if after: self.rows[max(after)].addnext(e)
            else: self.data.insert(0, e)
            self.rows[r] = e
        return self.rows[r]
    def cell(self, ref):
        col, r = re.match(r"([A-Z]+)(\d+)", ref).groups()
        row = self.row(int(r))
        for c in row.findall(N + "c"):
            if c.get("r") == ref: return c
        c = etree.SubElement(row, N + "c"); c.set("r", ref)
        cs = sorted(row.findall(N + "c"), key=lambda e: colnum(re.match(r"[A-Z]+", e.get("r")).group()))
        for e in cs: row.remove(e)
        for e in cs: row.append(e)
        return c
    def set(self, ref, val):
        c = self.cell(ref)
        assert c.find(N + "f") is None, "formula cell %s" % ref
        for ch in list(c): c.remove(ch)
        c.attrib.pop("t", None)
        if val is None: return
        if isinstance(val, str):
            c.set("t", "inlineStr")
            t = etree.SubElement(etree.SubElement(c, N + "is"), N + "t"); t.text = val
        else:
            v = etree.SubElement(c, N + "v"); v.text = repr(round(float(val), 2))
    def bytes(self):
        return etree.tostring(self.root, xml_declaration=True, encoding="UTF-8", standalone=True)

z = zipfile.ZipFile(SRC)
sh = {k: Sheet(z.read(v)) for k, v in PATHS.items()}
ro = openpyxl.load_workbook(SRC, read_only=True, data_only=True)
wip = sh["WIP Movements"]

# ---- check the rows being changed are the ones expected
def expect(r, job, typ, amt):
    v = list(ro["WIP Movements"].iter_rows(min_row=r, max_row=r, values_only=True))[0]
    assert str(v[1]).strip() == job and v[7] == typ and abs(v[9] - amt) < 0.005, (r, v[:10])
expect(11, "25021901", "Release of deferral", 536.25)
expect(21, "2506208", "Accrual - unbilled work", 1787.5)
expect(52, "2604704", "Accrual - unbilled work", 1361.25)
expect(173, "26092101", "Deferral - invoiced in advance", -950)
assert list(ro["WIP Movements"].iter_rows(min_row=177, max_row=177, values_only=True))[0][1] is None

# ---- September amounts per the WIP Schedule and Xero journal 71188
wip.set("J11", 2516.25)
wip.set("R11", "Invoiced in advance. De-recognise based on labour hours. Sep 26: 15.25 hrs x $165 (WIP Schedule row 174).")
wip.set("J21", 6402.5)
wip.set("R21", "Based on labour hours. Sep 26: 49.25 hrs x $130 (WIP Schedule row 195).")
wip.set("H52", "Reversal of prior accrual")
wip.set("J52", -4585)
wip.set("R52", "Invoiced INV-10596 on 09/09/2026. Accrued labour reversed in Sep 26 (WIP Schedule row 377).")
wip.set("J173", -550)
wip.set("R173", "Separate journal; Auto-reverse. Labour 400 and equipment 150. Video labour 400 is on 26092101V below.")

FLIP = ("Posted in Xero as a credit to 11300 (journal %s). Not invoiced at 30/09, so it should be a debit "
        "like August's pre-production lines. Change the sign here if Xero is corrected.")
NEW = [
    # job, cost centre, invoice, description, type, rev/cost, amount, P&L, notes
    ("26092101V", "VIDEO", "INV-10645", "Bankwest - BCEC WGEA Report Launch 13.10.2026 | Video Labour Revenue (INV-10645)",
     "Deferral - invoiced in advance", "Revenue", -400, "42150", "Separate journal; Auto-reverse"),
    # Xero journal 70539, 01/09/2026: August's auto-reversing lines reverse
    ("26073107", "ONSITE", None, "DTTL | SkyBus for KK for DTTL HOB Tech & Coverage 6-9 Sep 26", "Release of deferred cost", "Cost", -40, "58110", "Auto-reversed 01/09/2026 (Xero journal 70539)"),
    ("26073107", "ONSITE", None, "DTTL | Return flights for KK ADL-HOB for DTTL HOB Tech & Coverage 6-9 Sep 26", "Release of deferred cost", "Cost", -531.82, "58110", "Auto-reversed 01/09/2026 (Xero journal 70539)"),
    ("26073107", "ONSITE", None, "DTTL | Accommodation for KK for DTTL HOB Tech & Coverage 6-9 Sep 26", "Release of deferred cost", "Cost", -1589.98, "58110", "Auto-reversed 01/09/2026 (Xero journal 70539)"),
    ("26071701", "PRODUCTION", None, "ASX Live - 9.9.26 [PRD] Goget Travel Expense", "Release of deferred cost", "Cost", -69.75, "58115", "Auto-reversed 01/09/2026 (Xero journal 70539)"),
    ("26061602", "PRODUCTION", None, "ASX CEO Connect (Extended Results) 1.9.26 [PRD] Pre-prod labour", "Reversal of prior accrual", "Revenue", -1852.5, "42100", "Auto-reversed 01/09/2026 (Xero journal 70539)"),
    ("26071701", "PRODUCTION", None, "ASX Live - 9.9.26 [PRD] Pre-prod labour", "Reversal of prior accrual", "Revenue", -1885, "42100", "Auto-reversed 01/09/2026 (Xero journal 70539)"),
    ("26072102", "PRODUCTION", None, "Ventia - RUOK? Day - 10.09.2026 [PRD] pre-prd labour", "Reversal of prior accrual", "Revenue", -585, "42100", "Auto-reversed 01/09/2026 (Xero journal 70539)"),
    ("26082502", "PRODUCTION", None, "ICC - Presenter & Speaker Preparation - September 2026 [PRD] pre-prod labour", "Reversal of prior accrual", "Revenue", -292.5, "42100", "Auto-reversed 01/09/2026 (Xero journal 70539)"),
    ("26032302", "PRODUCTION", None, "RSNSW - OGM - 2.9.2026 [PRD] pre-prod labour", "Reversal of prior accrual", "Revenue", -260, "42100", "Auto-reversed 01/09/2026 (Xero journal 70539)"),
    ("26061602V", "VIDEO", None, "ASX CEO Connect (Extended Results) 1.9.26 [VID] pre-vid labour", "Reversal of prior accrual", "Revenue", -190, "42150", "Auto-reversed 01/09/2026 (Xero journal 70539)"),
    ("2591012C", "PRODUCTION", None, "St Luke Equipment Sale | DL to advise to invoice in SEPT26 | Indicative Revenue", "Release of deferral", "Revenue", 4227.27, "42600", "Auto-reversed 01/09/2026 (Xero journal 70539)"),
    # Xero journals 71186 and 71188, 30/09/2026
    ("26082502", "PRODUCTION", "INV-10674", "ICC - September Presenter Support Presentations - ComBio - 1/10/2026 | Labour Revenue (INV-10674)", "Deferral - invoiced in advance", "Revenue", -762.5, "42100", "Separate journal; Auto-reverse"),
    ("2506208", "PRODUCTION", None, "CBA AGM 2026 - Goget hire | Travel Expense (INV10498230926)", "Deferred cost - paid in advance", "Cost", 72, "58115", "Separate journal; Auto-reverse"),
    ("26042001", "PRODUCTION", None, "ASX - AGM - 22.10.26 [PRD] pre-prod labour revenue", "Accrual - unbilled work", "Revenue", -1625, "42100", FLIP % "71186"),
    ("2603105", "PRODUCTION", None, "Cochlear - AGM - 26.10.26 [PRD] pre-prod labour revenue", "Accrual - unbilled work", "Revenue", -2697.5, "42100", FLIP % "71186"),
    ("26082402", "PRODUCTION", None, "TLC - 2026 AGM - 14.10.26 [PRD] pre-prod labour revenue", "Accrual - unbilled work", "Revenue", -715, "42100", FLIP % "71186"),
    ("2512805", "PRODUCTION", None, "RSNSW Ordinary General Meeting 7.10.2026 [PRD] pre-prod labour revenue", "Accrual - unbilled work", "Revenue", -195, "42100", FLIP % "71186"),
    ("26082702", "PRODUCTION", None, "Garvan - Breast Cancer Public Lecture - 1.10.26 [PRD] pre-prod labour revenue", "Accrual - unbilled work", "Revenue", -130, "42100", FLIP % "71186"),
    ("26081801", "PRODUCTION", None, "Microsoft - SME&C Town Hall 14.10.2026 [PRD] pre-prod labour revenue", "Accrual - unbilled work", "Revenue", -487.5, "42100", FLIP % "71186"),
    ("2608402", "PRODUCTION", None, "Garvan - Suttons Event - 15.10.26 [PRD] pre-prod labour revenue", "Accrual - unbilled work", "Revenue", -97.5, "42100", FLIP % "71186"),
    ("26011407", "PRODUCTION", None, "ASX CEO Connect 27.10.26 [PRD] pre-prod labour revenue", "Accrual - unbilled work", "Revenue", -260, "42100", FLIP % "71186"),
    ("2608101", "PRODUCTION", None, "ART - AMM Captioning - 25.11.26 [PRD] pre-prod labour revenue", "Accrual - unbilled work", "Revenue", -97.5, "42100", "Reverse in November 2026. " + FLIP % "71188"),
    ("2609803", "PRODUCTION", None, "Automic - Noumi LTD AGM - 5.11.26 [PRD] pre-prod labour revenue", "Accrual - unbilled work", "Revenue", -162.5, "42100", "Reverse in November 2026. " + FLIP % "71188"),
    ("26062901", "INTEGRATION", None, "Aware | SYD 28.06 Services [INT] Labour Revenue Pre work", "Deferral - invoiced in advance", "Revenue", -148, "42800",
     "Based on labour hours; Reverse when invoiced. Same treatment as August (row 162). No services invoice in Xero at 30/09."),
]
for i, (job, cc, inv, desc, typ, rc, amt, gl, note) in enumerate(NEW):
    r = 177 + i
    for col, v in (("A", SEP), ("B", job), ("E", cc), ("F", inv), ("G", desc), ("H", typ), ("I", rc), ("J", amt), ("K", gl), ("R", note)):
        wip.set("%s%d" % (col, r), v)
    # the Month column carries a date format in the populated rows; copy it
    wip.cell("A%d" % r).set("s", wip.cell("A176").get("s"))

# ---- Deferrals: Aware Appspace on Xero's schedule (12 months Aug 26 - Jul 27, Aug and Sep released in Sep)
d = sh["Deferrals"]
rv = list(ro["Deferrals"].iter_rows(min_row=57, max_row=58, values_only=True))
assert rv[0][1] == "INV-10635" and rv[1][1] == "WIP-474"
d.set("G57", 19430.92)
d.set("K57", rv[0][10] + ". Xero journals 71180/71192: 12 months Aug 26 - Jul 27 at 1,766.45, with Aug and Sep both released in Sep. "
      "Amount here is the invoice less August's month, so Sep releases two months as Xero did.")
d.set("G58", 17487.85)
d.set("J58", 1589.81)
d.set("K58", rv[1][10] + ". Xero journals 71180/71192: 12 months Aug 26 - Jul 27 at 1,589.81, with Aug and Sep both released in Sep. "
      "Amount here is the bill less August's month, so Sep releases two months as Xero did.")

# ---- Lists: job 2609803 is in Xero's job list but not in the tracker's
lists = list(ro["Lists"].iter_rows(min_row=5, max_row=1500, min_col=16, max_col=18, values_only=True))
jobs = [str(x[0]).strip() if x[0] is not None else None for x in lists]
assert "2609803" not in jobs
nxt = 5 + max(i for i, j in enumerate(jobs) if j)  + 1
sh["Lists"].set("P%d" % nxt, "2609803")
sh["Lists"].set("Q%d" % nxt, "Automic - Noumi LTD AGM - 5.11.26")
sh["Lists"].set("R%d" % nxt, "Production")
jobs[nxt - 5] = "2609803"

# ---- WIP Summary column G: the WIP Schedule's September closing balance per job (row n shows lst_Jobs item n-13)
ws = openpyxl.load_workbook(WIPSCHED, data_only=True)["Schedule Updated"]
per = defaultdict(float)
for r in range(6, ws.max_row + 1):
    v = ws.cell(r, CI("BZ")).value
    if isinstance(v, (int, float)) and abs(v) > 0.004:
        per[str(ws.cell(r, 1).value).split("\n")[0].strip()] += v
per = {j: v for j, v in per.items() if abs(v) >= 0.01}
missing = [j for j in per if j not in jobs]
assert not missing, missing
summ = sh["WIP Summary"]
for j, v in per.items():
    summ.set("G%d" % (14 + jobs.index(j)), round(v, 2))

# ---- Month-End: WIP Schedule and Xero 11300 closing balances at 30/09/2026
total = round(sum(per.values()), 2)
assert total == -146188.18, total
sh["Month-End"].set("B23", total)
sh["Month-End"].set("B25", -146188.18)

wbxml = z.read("xl/workbook.xml").decode()
if "fullCalcOnLoad" not in wbxml:
    wbxml = re.sub(r"<calcPr([^>]*)/>", r'<calcPr\1 fullCalcOnLoad="1"/>', wbxml)
with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as o:
    for item in z.infolist():
        name = item.filename
        if name in PATHS.values():
            data = sh[[k for k, v in PATHS.items() if v == name][0]].bytes()
        elif name == "xl/workbook.xml":
            data = wbxml.encode()
        else:
            data = z.read(name)
        o.writestr(item, data)
print("written", OUT, "jobs pasted", len(per))
