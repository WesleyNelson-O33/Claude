"""September 2026 utilisation report fixes on the September tab.

1. O26 Layla Phillips' leave % looked up John Rizvi.
2. O25 Chloe Landayan's leave % divided by the Support total (H28), not her own hours (H25).
3. Daniel Sobkowski's row was out of date against Raw data: the pivot had not been
   refreshed after his late timesheets. His row, the Consulting total and the grand
   total are updated to the raw data, and the pivot is set to refresh when opened.
"""
import re, sys, zipfile
from lxml import etree

SRC, OUT = sys.argv[1], sys.argv[2]
N = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
SHEET = "xl/worksheets/sheet12.xml"
CACHE = "xl/pivotCache/pivotCacheDefinition20.xml"

z = zipfile.ZipFile(SRC)
root = etree.fromstring(z.read(SHEET))
cells = {c.get("r"): c for c in root.iter(N + "c")}

def setv(ref, v):
    c = cells[ref]; e = c.find(N + "v")
    if e is None: e = etree.SubElement(c, N + "v")
    e.text = repr(float(v))

def setf(ref, old, new):
    f = cells[ref].find(N + "f")
    assert f.text == old, (ref, f.text)
    f.text = new

setf("O26", 'GETPIVOTDATA("Sum of Leave hours",$B$5,"Name","John Abbas Rizvi","Person Group","Support")/H26',
     'GETPIVOTDATA("Sum of Leave hours",$B$5,"Name","Layla Phillips","Person Group","Support")/H26')
setv("O26", 0)
setf("O25", 'GETPIVOTDATA("Sum of Leave hours",$B$5,"Name","Chloe Landayan","Person Group","Support")/H28',
     'GETPIVOTDATA("Sum of Leave hours",$B$5,"Name","Chloe Landayan","Person Group","Support")/H25')
setv("O25", 12 / 176)

# Daniel Sobkowski, September per Raw data: chargeable 40.25, non-chargeable 81.5, leave 53, PH 0
H, I, L = 176, 176, 0.5
for r in (12, 13):
    D, E, F, G = 40.25, 81.5, 53, 0
    for col, v in (("D", D), ("E", E)): setv("%s%d" % (col, r), v)
    J = D / H
    setv("J%d" % r, J); setv("K%d" % r, 1 - J); setv("M%d" % r, J - L); setv("N%d" % r, D + E + F + G - I)
gD, gE, gF, gG = 2413 - 15.25 + 40.25, 326.5 - 43.25 + 81.5, 603.75, 24
gH, gI, gL = 3385.999824, 3409.999824, 0.8085000000000001
setv("D29", gD); setv("E29", gE)
gJ = gD / gH
setv("J29", gJ); setv("K29", 1 - gJ); setv("M29", gJ - gL); setv("N29", gD + gE + gF + gG - gI)

cache = z.read(CACHE).decode("utf-8")
if "refreshOnLoad=" not in cache:
    cache = cache.replace("<pivotCacheDefinition ", '<pivotCacheDefinition refreshOnLoad="1" ', 1)

with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as o:
    for it in z.infolist():
        if it.filename == SHEET:
            data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
        elif it.filename == CACHE:
            data = cache.encode("utf-8")
        else:
            data = z.read(it.filename)
        o.writestr(it, data)
print("written", OUT)
