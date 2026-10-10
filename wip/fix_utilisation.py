"""September 2026 utilisation report: workable hours and overtime checks.

Workable hours are business days in the person's employed window (state public
holidays taken out, as paid in Raw data) times their daily hours. People who left
or changed to full time part way through are counted only for the days they were
employed on that basis. Termination payouts (unused leave paid out) are not hours
worked, so they are kept out of overtime.
"""
import re, sys, zipfile
from lxml import etree

SRC, OUT = sys.argv[1], sys.argv[2]
N = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
SEP, Q3, Y12 = "xl/worksheets/sheet12.xml", "xl/worksheets/sheet21.xml", "xl/worksheets/sheet22.xml"
PIVOT_12M = "xl/pivotTables/pivotTable10.xml"
CACHES = ["xl/pivotCache/pivotCacheDefinition20.xml", "xl/pivotCache/pivotCacheDefinition21.xml",
          "xl/pivotCache/pivotCacheDefinition28.xml"]
PAYOUT = "-SUMPRODUCT(SUMIFS('Raw data'!$Q:$Q,'Raw data'!$I:$I,$C{r},'Raw data'!$C:$C,\"Unused*\",'Raw data'!$T:$T,{{\"Jul-26\",\"Aug-26\",\"Sep-26\"}}))"

z = zipfile.ZipFile(SRC)
docs = {p: etree.fromstring(z.read(p)) for p in (SEP, Q3, Y12)}

def cell(p, ref):
    for c in docs[p].iter(N + "c"):
        if c.get("r") == ref: return c
    raise KeyError(ref)

def formula(p, ref, new, old=None):
    c = cell(p, ref); f = c.find(N + "f")
    if old is not None: assert f.text == old, (ref, f.text)
    assert not f.get("ref") or f.get("ref") == ref, ("shared master", ref)
    for a in ("t", "si", "ref"): f.attrib.pop(a, None)
    f.text = new
    v = c.find(N + "v")
    if v is not None: c.remove(v)
    c.attrib.pop("t", None)

# ---- September tab: Chloe's leave % divided by the Support total
formula(SEP, "O25", 'GETPIVOTDATA("Sum of Leave hours",$B$5,"Name","Chloe Landayan","Person Group","Support")/H25',
        'GETPIVOTDATA("Sum of Leave hours",$B$5,"Name","Chloe Landayan","Person Group","Support")/H28')

# ---- Jul-Sep tab
formula(Q3, "H18", "33*5.076923", "65*5.076923")   # Tawonashe, WA, last worked 14/08/2026
formula(Q3, "H24", "38*8", "66*8")                 # Dylan, VIC, last worked 21/08/2026
formula(Q3, "H29", "3*8", "66*0.878787")           # Robert, ACT, last worked 03/07/2026
formula(Q3, "N23", "SUM(D23:G23)-I23", "SUM(D25:G25)-I23")
for r in (18, 24, 29):
    formula(Q3, "N%d" % r, "SUM(D{r}:G{r})-I{r}".format(r=r) + PAYOUT.format(r=r))
formula(Q3, "J32", "D32/H32", "D31/H32")

# ---- Last 12 months tab (Oct 2025 - Sep 2026)
for ref, new, old in (("I6", "164*8+1", "142*8+1"),        # Jaazaniah, FT from 09/02/2026
                      ("I18", "219*4.984064", "251*4.984064"),  # Tawonashe, WA, to 14/08/2026
                      ("I19", "252*8", "251*8"),            # Kyle, SA calendar
                      ("I23", "92*8", "70*8"),              # Sangeeta, FT from 25/05/2026
                      ("I24", "224*8", "251*8"),            # Dylan, VIC, to 21/08/2026
                      ("I26", "252*8", "251*8"),            # Jason, QLD calendar
                      ("I28", "87*8", "251*8"),             # Layla, FT from 01/06/2026
                      ("I29", "172*8", "251*8"),            # John, FT from 27/01/2026
                      ("I30", "SUM(I14:I29)", "SUM(I14:I28)"),
                      ("J30", "SUM(J14:J29)", "SUM(J14:J28)"),
                      ("K30", "E30/I30", "E31/I30")):
    formula(Y12, ref, new, old)
formula(Y12, "J29", 'I29+GETPIVOTDATA("Sum of Public Holiday",$C$5,"Name","John Abbas Rizvi","p","Support")',
        'I29+GETPIVOTDATA("Sum of Public Holiday",$C$5,"Name","Layla Phillips","p","Support")')

# untick the two termination payout pay types in the 12 month pivot's Pay Type filter
ptx = etree.fromstring(z.read(PIVOT_12M))
paytype = ptx.find(N + "pivotFields").findall(N + "pivotField")[2]
cache21 = etree.fromstring(z.read("xl/pivotCache/pivotCacheDefinition21.xml"))
shared = [i.get("v") for i in cache21.find(N + "cacheFields").findall(N + "cacheField")[2].find(N + "sharedItems")]
for it in paytype.find(N + "items"):
    if it.get("x") is not None and "nused" in str(shared[int(it.get("x"))]):
        it.set("h", "1")
assert sum(1 for it in paytype.find(N + "items") if it.get("h")) >= 2
pt = etree.tostring(ptx, xml_declaration=True, encoding="UTF-8", standalone=True).decode("utf-8")

with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as o:
    for it in z.infolist():
        n = it.filename
        if n in docs:
            data = etree.tostring(docs[n], xml_declaration=True, encoding="UTF-8", standalone=True)
        elif n == PIVOT_12M:
            data = pt.encode("utf-8")
        elif n in CACHES:
            s = z.read(n).decode("utf-8")
            if "refreshOnLoad=" not in s:
                s = s.replace("<pivotCacheDefinition ", '<pivotCacheDefinition refreshOnLoad="1" ', 1)
            data = s.encode("utf-8")
        elif n == "xl/workbook.xml":
            s = z.read(n).decode("utf-8")
            if "fullCalcOnLoad" not in s:
                s = re.sub(r"<calcPr([^>]*?)/>", r'<calcPr\1 fullCalcOnLoad="1"/>', s)
            data = s.encode("utf-8")
        else:
            data = z.read(n)
        o.writestr(it, data)
print("written", OUT)
