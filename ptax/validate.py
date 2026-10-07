"""Check an xlsx package the way Excel does before it offers to repair."""
import sys, zipfile, re
from collections import Counter
from lxml import etree

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"
bad = []

f = sys.argv[1]
z = zipfile.ZipFile(f)
names = set(z.namelist())

# 1  every xml part parses
for n in sorted(names):
    if n.endswith((".xml", ".rels", ".vml")):
        try:
            etree.fromstring(z.read(n))
        except Exception as e:
            bad.append("unparsable %s: %s" % (n, str(e)[:90]))

# 2  every part has a content type, and every override points at a real part
types = etree.fromstring(z.read("[Content_Types].xml"))
defaults = {d.get("Extension").lower() for d in types.findall("{%s}Default" % CT)}
overrides = [o.get("PartName").lstrip("/") for o in types.findall("{%s}Override" % CT)]
for o in overrides:
    if o not in names:
        bad.append("content type points at a missing part: " + o)
for dup, c in Counter(overrides).items():
    if c > 1:
        bad.append("duplicate content type override: " + dup)
for n in names:
    if n == "[Content_Types].xml" or n in overrides:
        continue
    if n.rsplit(".", 1)[-1].lower() not in defaults:
        bad.append("no content type for " + n)

# 3  every relationship target resolves
for n in sorted(names):
    if not n.endswith(".rels"):
        continue
    base = n.rsplit("_rels/", 1)[0]
    for rel in etree.fromstring(z.read(n)).findall("{%s}Relationship" % PR):
        if rel.get("TargetMode") == "External":
            continue
        t = rel.get("Target")
        if t.startswith("/"):
            t = t.lstrip("/")
        else:
            t = base + t
        while "/../" in t:
            t = re.sub(r"[^/]+/\.\./", "", t, count=1)
        if t not in names:
            bad.append("%s: broken rel %s -> %s" % (n, rel.get("Id"), rel.get("Target")))

# 4  workbook sheets: unique names and ids, each rel present
wb = etree.fromstring(z.read("xl/workbook.xml"))
sheets = wb.findall(".//{%s}sheet" % NS)
for field in ("name", "sheetId"):
    for v, c in Counter(s.get(field) for s in sheets).items():
        if c > 1:
            bad.append("duplicate sheet %s: %s" % (field, v))
wrels = {r.get("Id") for r in
         etree.fromstring(z.read("xl/_rels/workbook.xml.rels")).findall("{%s}Relationship" % PR)}
for s in sheets:
    rid = s.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
    if rid not in wrels:
        bad.append("sheet %s has no relationship %s" % (s.get("name"), rid))

# 5  styles: counts, unique numFmtIds, every s= index in range
st = etree.fromstring(z.read("xl/styles.xml"))
for tag in ("numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs"):
    e = st.find("{%s}%s" % (NS, tag))
    if e is not None and e.get("count") and int(e.get("count")) != len(e):
        bad.append("styles %s count=%s but has %d children" % (tag, e.get("count"), len(e)))
nf = st.find("{%s}numFmts" % NS)
if nf is not None:
    for v, c in Counter(x.get("numFmtId") for x in nf).items():
        if c > 1:
            bad.append("duplicate numFmtId " + v)
nxf = len(st.find("{%s}cellXfs" % NS))
for n in sorted(names):
    if not n.startswith("xl/worksheets/sheet"):
        continue
    for s_ in re.findall(rb'<c r="[A-Z]+\d+" s="(\d+)"', z.read(n)):
        if int(s_) >= nxf:
            bad.append("%s uses style %s, only %d exist" % (n, s_.decode(), nxf))
            break

# 6  worksheet child order must follow the schema
ORDER = ["sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData",
         "sheetCalcPr", "sheetProtection", "protectedRanges", "scenarios", "autoFilter",
         "sortState", "dataConsolidate", "customSheetViews", "mergeCells", "phoneticPr",
         "conditionalFormatting", "dataValidations", "hyperlinks", "printOptions",
         "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks",
         "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing",
         "legacyDrawing", "legacyDrawingHF", "drawingHF", "picture", "oleObjects",
         "controls", "webPublishItems", "tableParts", "extLst"]
for n in sorted(names):
    if not n.startswith("xl/worksheets/sheet"):
        continue
    seen = [etree.QName(c).localname for c in etree.fromstring(z.read(n))
            if isinstance(c.tag, str)]
    idx = [ORDER.index(t) for t in seen if t in ORDER]
    if idx != sorted(idx):
        bad.append("%s: elements out of order %s" % (n, seen))

print(("FAILED\n  " + "\n  ".join(bad)) if bad else "OK, no problems found")
sys.exit(1 if bad else 0)
