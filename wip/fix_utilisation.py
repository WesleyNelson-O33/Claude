"""September 2026 utilisation report: overtime only.

The overtime column on the Jul-Sept tab and the 12 month tab becomes a running
total of the overtime paid in Raw data (pay types containing "Overtime", full and
part time) for Jul-Sep 2026 and for Oct 2025 - Sep 2026. Nothing else changes.
"""
import re, sys, zipfile
from lxml import etree

SRC, OUT = sys.argv[1], sys.argv[2]
N = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
Q3, Y12 = "xl/worksheets/sheet21.xml", "xl/worksheets/sheet22.xml"

z = zipfile.ZipFile(SRC)
docs = {p: etree.fromstring(z.read(p)) for p in (Q3, Y12)}

def formula(p, ref, new):
    for c in docs[p].iter(N + "c"):
        if c.get("r") == ref: break
    else: raise KeyError(ref)
    f = c.find(N + "f")
    for a in ("t", "si", "ref"): f.attrib.pop(a, None)
    f.text = new
    v = c.find(N + "v")
    if v is not None: c.remove(v)
    c.attrib.pop("t", None)

def paid_ot(namecol, r, months):
    m = "{" + ",".join('"%s"' % x for x in months) + "}"
    return ("SUMPRODUCT(SUMIFS('Raw data'!$H:$H,'Raw data'!$I:$I,$%s%d,'Raw data'!$C:$C,\"*Overtime*\","
            "'Raw data'!$T:$T,%s,'Raw data'!$N:$N,{\"Full-time\";\"Part-time\"}))") % (namecol, r, m)

QTR = ["Jul-26", "Aug-26", "Sep-26"]
YEAR = ["Oct-25", "Nov-25", "Dec-25", "Jan-26", "Feb-26", "Mar-26", "Apr-26", "May-26", "Jun-26",
        "Jul-26", "Aug-26", "Sep-26"]

# Jul-Sept tab, Overtime Hours (column N)
for r in range(6, 31):
    if r not in (11, 13): formula(Q3, "N%d" % r, paid_ot("C", r, QTR))
for r, f in {11: "SUM(N6:N10)", 13: "N12", 31: "SUM(N14:N30)", 32: "N11+N13+N31"}.items():
    formula(Q3, "N%d" % r, f)

# All - last 12 months tab, Overtime Hours (column Q)
for r in range(6, 30):
    if r not in (11, 13): formula(Y12, "Q%d" % r, paid_ot("D", r, YEAR))
for r, f in {11: "SUM(Q6:Q10)", 13: "Q12", 30: "SUM(Q14:Q29)", 31: "Q11+Q13+Q30"}.items():
    formula(Y12, "Q%d" % r, f)

# every shared formula child must still have its master
for p_, root in docs.items():
    masters = {f.get("si") for f in root.iter(N + "f") if f.get("t") == "shared" and f.get("ref")}
    orphans = [c.get("r") for c in root.iter(N + "c") if c.find(N + "f") is not None
               and c.find(N + "f").get("t") == "shared" and c.find(N + "f").get("si") not in masters]
    assert not orphans, (p_, orphans)

with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as o:
    for it in z.infolist():
        n = it.filename
        if n in docs:
            data = etree.tostring(docs[n], xml_declaration=True, encoding="UTF-8", standalone=True)
        elif n == "xl/workbook.xml":   # so the new overtime formulas calculate on open
            s = z.read(n).decode("utf-8")
            if "fullCalcOnLoad" not in s:
                s = re.sub(r"<calcPr([^>]*?)/>", r'<calcPr\1 fullCalcOnLoad="1"/>', s)
            data = s.encode("utf-8")
        else:
            data = z.read(n)
        o.writestr(it, data)
print("written", OUT)
