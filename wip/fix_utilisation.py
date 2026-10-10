"""September 2026 utilisation report: workable hours and overtime checks.

Workable hours are business days in the person's employed window (state public
holidays taken out, as paid in Raw data) times their daily hours. People who left
or changed to full time part way through are counted only for the days they were
employed on that basis. Overtime is the overtime paid in Raw data (pay types
containing "Overtime") for full and part time hours, not the gap between hours and
workable hours. Termination payouts are left out of the 12 month pivot.
"""
import re, sys, zipfile
from lxml import etree

SRC, OUT = sys.argv[1], sys.argv[2]
N = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
SEP, Q3, Y12 = "xl/worksheets/sheet12.xml", "xl/worksheets/sheet21.xml", "xl/worksheets/sheet22.xml"
PIVOT_12M = "xl/pivotTables/pivotTable10.xml"
CACHES = ["xl/pivotCache/pivotCacheDefinition20.xml", "xl/pivotCache/pivotCacheDefinition21.xml",
          "xl/pivotCache/pivotCacheDefinition28.xml"]
NAMECOL = "C"

z = zipfile.ZipFile(SRC)
docs = {p: etree.fromstring(z.read(p)) for p in (SEP, Q3, Y12)}

def cell(p, ref):
    for c in docs[p].iter(N + "c"):
        if c.get("r") == ref: return c
    raise KeyError(ref)

def formula(p, ref, new, old=None):
    c = cell(p, ref); f = c.find(N + "f")
    if old is not None: assert f.text == old, (ref, f.text)
    for a in ("t", "si", "ref"): f.attrib.pop(a, None)
    f.text = new
    v = c.find(N + "v")
    if v is not None: c.remove(v)
    c.attrib.pop("t", None)

def paid_ot(r, months):
    m = "{" + ",".join('"%s"' % x for x in months) + "}"
    return ("SUMPRODUCT(SUMIFS('Raw data'!$H:$H,'Raw data'!$I:$I,$%s%d,'Raw data'!$C:$C,\"*Overtime*\","
            "'Raw data'!$T:$T,%s,'Raw data'!$N:$N,{\"Full-time\";\"Part-time\"}))") % (NAMECOL, r, m)

def ot_column(p, col, people, totals, months):
    global NAMECOL
    for r in people: formula(p, "%s%d" % (col, r), paid_ot(r, months))
    for r, f in totals.items(): formula(p, "%s%d" % (col, r), f)

SEPT = ["Sep-26"]; QTR = ["Jul-26", "Aug-26", "Sep-26"]
YEAR = ["Oct-25", "Nov-25", "Dec-25", "Jan-26", "Feb-26", "Mar-26", "Apr-26", "May-26", "Jun-26",
        "Jul-26", "Aug-26", "Sep-26"]

# ---- September tab
formula(SEP, "O25", 'GETPIVOTDATA("Sum of Leave hours",$B$5,"Name","Chloe Landayan","Person Group","Support")/H25',
        'GETPIVOTDATA("Sum of Leave hours",$B$5,"Name","Chloe Landayan","Person Group","Support")/H28')
for ref in ("H9", "H10", "H12", "H15", "H18", "H24", "H25", "H26"):   # manual +/- hours taken out
    formula(SEP, ref, "22*8")
NAMECOL = "C"
ot_column(SEP, "N", [r for r in range(6, 28) if r not in (11, 13)],
          {11: "SUM(N6:N10)", 13: "N12", 28: "SUM(N14:N27)", 29: "N11+N13+N28"}, SEPT)

# ---- Jul-Sep tab
formula(Q3, "H18", "33*5.076923", "65*5.076923")   # Tawonashe, WA, last worked 14/08/2026
formula(Q3, "H24", "38*8", "66*8")                 # Dylan, VIC, last worked 21/08/2026
formula(Q3, "H29", "3*8", "66*0.878787")           # Robert, ACT, last worked 03/07/2026
for ref, old in (("H6", "66*8+2"), ("H7", "66*8+1.5"), ("H15", "66*8+4"), ("H17", "66*8-2")):
    formula(Q3, ref, "66*8", old)
ot_column(Q3, "N", [r for r in range(6, 31) if r not in (11, 13)],
          {11: "SUM(N6:N10)", 13: "N12", 31: "SUM(N14:N30)", 32: "N11+N13+N31"}, QTR)
formula(Q3, "J32", "D32/H32", "D31/H32")

# ---- Last 12 months tab (Oct 2025 - Sep 2026)
for ref, new, old in (("I6", "164*8", "142*8+1"),           # Jaazaniah, FT from 09/02/2026
                      ("I9", "252*8", "252*8-0.05"),        # Blake, manual adjustment out
                      ("I18", "219*4.984064", "251*4.984064"),  # Tawonashe, WA, to 14/08/2026
                      ("I19", "252*8", "251*8"),            # Kyle, SA calendar
                      ("I22", "252*4.011905", "252*4.011905-1"),  # Kent, manual adjustment out
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
NAMECOL = "D"
ot_column(Y12, "Q", [r for r in range(6, 30) if r not in (11, 13)],
          {11: "SUM(Q6:Q10)", 13: "Q12", 30: "SUM(Q14:Q29)", 31: "Q11+Q13+Q30"}, YEAR)

# every shared formula child must still have its master
for p_, root in docs.items():
    masters = {f.get("si") for f in root.iter(N + "f") if f.get("t") == "shared" and f.get("ref")}
    orphans = [c.get("r") for c in root.iter(N + "c") if c.find(N + "f") is not None
               and c.find(N + "f").get("t") == "shared" and c.find(N + "f").get("si") not in masters]
    assert not orphans, (p_, orphans)

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
