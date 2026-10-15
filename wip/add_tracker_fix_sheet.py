"""Add a "Fix on tracker" sheet to Danica's September WIP Adjustment Schedule.

Lists the tracker lines that still need to catch up with the WIP schedule and Xero, in the
same layout and narration style as the Journal Edit Schedule. Nothing is shown as done.
The sheet XML is added to her file directly so her comments and formatting survive."""
import re, sys, zipfile
from xml.sax.saxutils import escape

SRC, OUT = sys.argv[1], sys.argv[2]
NAME = "Fix on tracker"
S = dict(title=3, sect=4, label=5, money=7, blabel=55, bmoney=57, head=16, band=214, item=19, mid=20,
         wrap=21, num=22, done=24, totl=27, totn=28, note=30, find=31, fhead=32, f1=33, f2=34, f3=35, red=61,
         leg=40)

def col(n):
    s = ""
    while n: n, r = divmod(n - 1, 26); s = chr(65 + r) + s
    return s

rows = {}   # row -> list of xml cells
heights = {}
def put(r, c, v, style, formula=False):
    ref = "%s%d" % (c, r)
    if v is None:
        x = '<c r="%s" s="%d"/>' % (ref, style)
    elif formula:
        x = '<c r="%s" s="%d"><f>%s</f></c>' % (ref, style, escape(v))
    elif isinstance(v, str):
        x = '<c r="%s" s="%d" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (ref, style, escape(v))
    else:
        x = '<c r="%s" s="%d"><v>%s</v></c>' % (ref, style, repr(round(float(v), 2)))
    rows.setdefault(r, []).append((c, x))

put(1, "A", "WIP 11300 Fix on tracker, September 2026", S["title"])
put(2, "B", "Status: these fixes are not yet done on the revenue tracker. Tick Done as each line is fixed.", S["item"])
put(3, "B", "STARTING POSITION", S["sect"]); put(3, "G", "Ending  POSITION (once the fixes are done)", S["sect"])
put(4, "B", "Revenue tracker closing balance now", S["label"]); put(4, "E", -130154.73, S["money"])
put(5, "B", "WIP schedule and Xero 11300 once corrected", S["label"]); put(5, "E", -133253.18, S["money"])
put(6, "B", "Variance to clear", S["blabel"]); put(6, "E", "E4-E5", S["bmoney"], True)
put(4, "G", "Revenue tracker after the fixes", S["label"])
put(5, "G", "WIP schedule and Xero 11300 once corrected", S["label"]); put(5, "J", "E5", S["money"], True)
put(6, "G", "Variance ", S["blabel"]); put(6, "J", "J4-J5", S["bmoney"], True)
put(7, "B", "The ten pre-production lines are already positive on the tracker. The lines below are already right in "
            "Xero and on the WIP schedule. They are only missing or out of date on the tracker.", S["item"])
for c, h in zip("BCDEFGHIJ", ["Item", "Tab", "Row", "Tracker now", "Should be", "Effect on\ntracker",
                              "WIP schedule\nrow", "What to do on the tracker", "Done"]):
    put(9, c, h, S["head"])

SECTIONS = [
    ("WIP MOVEMENTS: THREE AMOUNTS STILL ON AUGUST'S FIGURES", [
        ("DAFF consulting invoiced in advance", "WIP Movements", "Row 11", 536.25, 2516.25, "174",
         "I fixed this one on the tracker. It was still carrying August's amount. In September 15.25 hours were worked "
         "at $165, so 2,516.25 of the revenue invoiced in advance was earned and released. I changed the Amount to "
         "2,516.25. WIP schedule row 174."),
        ("CBA AGM 2026 pre-production", "WIP Movements", "Row 21", 1787.5, 6402.5, "195",
         "Same reasoning as DAFF. It was still carrying August's amount. In September 49.25 hours were worked at $130, "
         "so I changed the Amount to 6,402.50. WIP schedule row 195."),
        ("Aware Seminar Room Audio Upgrade", "WIP Movements", "Row 52", 1361.25, -4585, "377",
         "This job was invoiced on 9 September as INV-10596, so the labour accrued in earlier months comes back out. The "
         "tracker was still adding another month of accrual. I changed the Type to Reversal of prior accrual and the "
         "Amount to -4,585. WIP schedule row 377."),
    ]),
    ("DEFERRALS: AWARE APPSPACE INV-10635 (XERO JOURNALS #71180 AND #71192)", [
        ("Aware Appspace licence, revenue", "Deferrals", "Row 57", -19270.34, -17664.47, "471",
         "Xero spreads this licence over 12 months, August to July, and released August and September together in "
         "September. The tracker was spreading it over 11 months from September. I typed 19,430.92 in Amount Override. "
         "That is the invoice less August's month, so September now releases two months, as Xero did. WIP schedule row 471."),
        ("Aware Appspace licence, cost", "Deferrals", "Row 58", 17343.33, 15898.04, "472",
         "Same reasoning as the revenue line. I typed 17,487.85 in Amount Override and 1,589.81 in Per Month Override, "
         "so September releases two months of cost, as Xero did. WIP schedule row 472."),
    ]),
    ("JOURNAL #70539: 1 SEPTEMBER REVERSAL OF AUGUST'S AUTO-REVERSING LINES", [
        ("DTTL SkyBus for KK, job 26073107", "WIP Movements", "New line", 0, -40, "453",
         "This line was missing from the tracker. August's line reversed automatically in Xero on 1 September and the "
         "WIP schedule has the reversal, but the tracker did not. I added it the same way August's tracker shows July's "
         "reversals in rows 90 to 104. Month 30/09/2026, cost centre ONSITE, Type Release of deferred cost, Cost, "
         "P&L 58110. WIP schedule row 453."),
        ("DTTL return flights for KK, job 26073107", "WIP Movements", "New line", 0, -531.82, "454",
         "Same reasoning as SkyBus. ONSITE, Release of deferred cost, Cost, P&L 58110. WIP schedule row 454."),
        ("DTTL accommodation for KK, job 26073107", "WIP Movements", "New line", 0, -1589.98, "455",
         "Same reasoning as SkyBus. ONSITE, Release of deferred cost, Cost, P&L 58110. WIP schedule row 455."),
        ("ASX Live GoGet, job 26071701", "WIP Movements", "New line", 0, -69.75, "456",
         "Same reasoning as SkyBus. PRODUCTION, Release of deferred cost, Cost, P&L 58115. WIP schedule row 456."),
        ("ASX CEO Connect pre-prod, job 26061602", "WIP Movements", "New line", 0, -1852.5, "459",
         "Same reasoning as SkyBus. PRODUCTION, Reversal of prior accrual, Revenue, P&L 42100. WIP schedule row 459."),
        ("ASX Live pre-prod, job 26071701", "WIP Movements", "New line", 0, -1885, "460",
         "Same reasoning as SkyBus. PRODUCTION, Reversal of prior accrual, Revenue, P&L 42100. WIP schedule row 460."),
        ("Ventia RUOK Day pre-prod, job 26072102", "WIP Movements", "New line", 0, -585, "461",
         "Same reasoning as SkyBus. PRODUCTION, Reversal of prior accrual, Revenue, P&L 42100. WIP schedule row 461."),
        ("ICC presenter prep, job 26082502", "WIP Movements", "New line", 0, -292.5, "462",
         "Same reasoning as SkyBus. PRODUCTION, Reversal of prior accrual, Revenue, P&L 42100. WIP schedule row 462."),
        ("RSNSW OGM 2.9.26 pre-prod, job 26032302", "WIP Movements", "New line", 0, -260, "463",
         "Same reasoning as SkyBus. PRODUCTION, Reversal of prior accrual, Revenue, P&L 42100. WIP schedule row 463."),
        ("ASX CEO Connect video, job 26061602V", "WIP Movements", "New line", 0, -190, "464",
         "Same reasoning as SkyBus. VIDEO, Reversal of prior accrual, Revenue, P&L 42150. WIP schedule row 464."),
        ("St Luke equipment sale, job 2591012C", "WIP Movements", "New line", 0, 4227.27, "465",
         "Same reasoning as SkyBus. This one was deferred in August, so the reversal releases it. PRODUCTION, Release of "
         "deferral, Revenue, P&L 42600. WIP schedule row 465."),
    ]),
    ("JOURNAL #71186: PRE-PROD AND INT WORK AND EXPENSE ACCRUAL FOR OCTOBER 2026", [
        ("ICC ComBio, job 26082502, INV-10674", "WIP Movements", "New line", 0, -762.5, "478",
         "This line was missing from the tracker. It was invoiced in September for work on 1 October, so the revenue is "
         "pushed out to October and WIP sits on the credit side. I added it with Month 30/09/2026, PRODUCTION, Type "
         "Deferral - invoiced in advance, Revenue, P&L 42100, invoice INV-10674. WIP schedule row 478."),
        ("CBA AGM 2026 GoGet hire, job 2506208", "WIP Movements", "New line", 0, 72, "479",
         "This line was missing from the tracker. It was paid in September for the October AGM, so the cost is held in "
         "WIP until the event. I added it with PRODUCTION, Type Deferred cost - paid in advance, Cost, P&L 58115. "
         "WIP schedule row 479."),
        ("Bankwest BCEC WGEA, job 26092101", "WIP Movements", "Row 173", -950, -550, "473, 475",
         "Xero split this invoice. Labour 400 and equipment 150 are on job 26092101. Video labour 400 is on 26092101V. "
         "I changed row 173 to -550 and added the video line below. The total does not change. WIP schedule rows 473 and 475."),
        ("Bankwest BCEC WGEA video, job 26092101V", "WIP Movements", "New line", 0, -400, "474",
         "Same invoice as the line above. I added it with VIDEO, Type Deferral - invoiced in advance, Revenue, P&L 42150, "
         "invoice INV-10645. WIP schedule row 474."),
    ]),
    ("JOURNAL #71188: WIP JOURNAL SEPTEMBER 2026", [
        ("Aware SYD 28.06 Services, job 26062901", "WIP Movements", "New line", 0, -148, "458",
         "This line was missing from the tracker. Xero and the WIP schedule both have it as a credit, the same as "
         "August's 1,450 in row 162. I added it with INTEGRATION, Type Deferral - invoiced in advance, Revenue, "
         "P&L 42800. WIP schedule row 458."),
    ]),
]
r = 10
first = None
for title, items in SECTIONS:
    r += 1
    put(r, "B", title, S["band"])
    for c in "CDEFGHIJ": put(r, c, None, S["band"])
    heights[r] = 15
    for item, tab, row, now, should, wrow, text in items:
        r += 1
        first = first or r
        put(r, "B", item, S["item"]); put(r, "C", tab, S["mid"]); put(r, "D", row, S["wrap"])
        put(r, "E", now, S["num"]); put(r, "F", should, S["num"]); put(r, "G", "F%d-E%d" % (r, r), S["num"], True)
        put(r, "H", wrow, S["mid"]); put(r, "I", text, S["wrap"]); put(r, "J", None, S["done"])
        heights[r] = 62 if len(text) > 300 else 48 if len(text) > 180 else 36
last = r
r += 1
put(r, "D", "TOTAL", S["totl"])
for c in "EFG": put(r, c, "SUM(%s%d:%s%d)" % (c, first, c, last), S["totn"], True)
tot = r
put(4, "J", "E4+G%d" % tot, S["money"], True)
r += 2
put(r, "B", "Once every line above is done, the tracker closes at a credit of 133,253.18, the same as the WIP schedule "
            "and Xero 11300 once the ten pre-production lines are corrected. On the Month-End tab, both variance lines "
            "in section 2 should then say Reconciled.", S["note"])
heights[r] = 48
widths = [(1, 1, 3), (2, 2, 36), (3, 3, 15), (4, 4, 10), (5, 7, 13), (8, 8, 12), (9, 9, 95), (10, 10, 8)]

sheet = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
         '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
         'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
         '<sheetViews><sheetView workbookViewId="0" showGridLines="0"/></sheetViews>',
         '<sheetFormatPr defaultRowHeight="15"/>', '<cols>']
sheet += ['<col min="%d" max="%d" width="%s" customWidth="1"/>' % w for w in widths]
sheet += ['</cols>', '<sheetData>']
for rr in sorted(rows):
    cells = "".join(x for _, x in sorted(rows[rr], key=lambda t: (len(t[0]), t[0])))
    h = ' ht="%s" customHeight="1"' % heights[rr] if rr in heights else ""
    sheet.append('<row r="%d"%s>%s</row>' % (rr, h, cells))
sheet += ['</sheetData>', '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>',
          '</worksheet>']
sheet_xml = "".join(sheet).encode()

z = zipfile.ZipFile(SRC)
names = z.namelist()
n = 1
while "xl/worksheets/sheet%d.xml" % n in names: n += 1
part = "xl/worksheets/sheet%d.xml" % n
wb = z.read("xl/workbook.xml").decode()
assert NAME not in wb
sid = max(int(x) for x in re.findall(r'sheetId="(\d+)"', wb)) + 1
rid = "rIdFix%d" % n
wb = re.sub(r'(<sheet name="What i did "[^>]*/>)', r'\1<sheet name="%s" sheetId="%d" r:id="%s"/>' % (NAME, sid, rid), wb)
rels = z.read("xl/_rels/workbook.xml.rels").decode().replace(
    "</Relationships>", '<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%d.xml"/></Relationships>' % (rid, n))
ct = z.read("[Content_Types].xml").decode().replace(
    "</Types>", '<Override PartName="/%s" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>' % part)
if "fullCalcOnLoad" not in wb:
    wb = re.sub(r"<calcPr([^>]*)/>", r'<calcPr\1 fullCalcOnLoad="1"/>', wb)
repl = {"xl/workbook.xml": wb.encode(), "xl/_rels/workbook.xml.rels": rels.encode(), "[Content_Types].xml": ct.encode()}
with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as o:
    for item in z.infolist():
        o.writestr(item, repl.get(item.filename, z.read(item.filename)))
    o.writestr(part, sheet_xml)
print("added", NAME, "as", part)
