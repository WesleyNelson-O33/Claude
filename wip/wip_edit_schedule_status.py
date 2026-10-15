"""Mark Danica's September WIP Adjustment Schedule as not yet corrected.

Keeps her narration, unticks Done, restates the ending position from the starting position
(the corrected tabs were removed), and words the status lines so nothing reads as done in Xero.
Edits the sheet XML directly so her comments and formatting survive."""
import re, sys, zipfile
from lxml import etree

SRC, OUT = sys.argv[1], sys.argv[2]
NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
N = "{%s}" % NS
RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

def colnum(c):
    n = 0
    for ch in c: n = n * 26 + ord(ch) - 64
    return n

class Sheet:
    def __init__(self, xml):
        self.root = etree.fromstring(xml)
        self.data = self.root.find(N + "sheetData")
        self.rows = {int(r.get("r")): r for r in self.data.findall(N + "row")}
    def cell(self, ref):
        col, r = re.match(r"([A-Z]+)(\d+)", ref).groups(); r = int(r)
        if r not in self.rows:
            e = etree.Element(N + "row"); e.set("r", str(r))
            before = [k for k in self.rows if k < r]
            (self.rows[max(before)].addnext(e) if before else self.data.insert(0, e)); self.rows[r] = e
        row = self.rows[r]
        for c in row.findall(N + "c"):
            if c.get("r") == ref: return c
        c = etree.SubElement(row, N + "c"); c.set("r", ref)
        cs = sorted(row.findall(N + "c"), key=lambda e: colnum(re.match(r"[A-Z]+", e.get("r")).group()))
        for e in cs: row.remove(e)
        for e in cs: row.append(e)
        return c
    def set(self, ref, val=None, formula=None):
        c = self.cell(ref)
        for ch in list(c): c.remove(ch)
        c.attrib.pop("t", None)
        if formula:
            f = etree.SubElement(c, N + "f"); f.text = formula
        elif isinstance(val, str):
            c.set("t", "inlineStr")
            t = etree.SubElement(etree.SubElement(c, N + "is"), N + "t"); t.text = val
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        elif val is not None:
            etree.SubElement(c, N + "v").text = str(val)
    def bytes(self):
        return etree.tostring(self.root, xml_declaration=True, encoding="UTF-8", standalone=True)

z = zipfile.ZipFile(SRC)
wbx = etree.fromstring(z.read("xl/workbook.xml"))
rels = etree.fromstring(z.read("xl/_rels/workbook.xml.rels"))
target = {r.get("Id"): "xl/" + r.get("Target").lstrip("/").replace("xl/", "") for r in rels}
path = {s.get("name"): target[s.get("{%s}id" % RNS)] for s in wbx.iter(N + "sheet")}
js, wd = Sheet(z.read(path["Journal Edit Schedule"])), Sheet(z.read(path["What i did "]))

# Journal Edit Schedule
js.set("B2", "Status: the corrections below are not yet done in Xero or on the WIP schedule. Tick Done as each line is corrected.")
js.set("G3", "Ending  POSITION (once the corrections are done)")
js.set("J4", formula="E4-2*SUM(E11:E21)")
js.set("J5", formula="E5-2*SUM(E11:E21)")
for r in list(range(11, 19)) + [20, 21]:
    js.set("J%d" % r, None)
js.set("B26", "The schedule and Xero 11300 agreed at a credit of 146,188.18, but both carried the same wrong sign on the ten "
              "lines above, so it did not show as a variance. Once every line above is done, both close at a credit of "
              "133,253.18. I then re-export the WIP 11300 account transactions and re-total the schedule to confirm they agree.")
for r in range(30, 40):
    j = "#71188" if r >= 38 else "#71186"
    js.set("F%d" % r, "Flip in Xero %s and fix the schedule sign" % j)

# What i did: status line, and the two statements that read as already done
wd.set("A2", "My notes on how I am correcting the September WIP journals and what I based each decision on. "
             "Status: the corrections are not yet done in Xero, on the WIP schedule or on the revenue tracker.")
wd.set("B30", "The same ten lines are to be changed to a debit on the WIP schedule (rows 480 to 489) and on the revenue tracker")
wd.set("B43", "Once the corrections are done, the WIP 11300 account transactions are re-exported and the schedule re-totalled,")
wd.set("B44", "and the two closing balances should agree to the cent at a credit of 133,253.18. This is performed before September is closed off.")

wbtxt = z.read("xl/workbook.xml").decode()
if "fullCalcOnLoad" not in wbtxt:
    wbtxt = re.sub(r"<calcPr([^>]*)/>", r'<calcPr\1 fullCalcOnLoad="1"/>', wbtxt)
edited = {path["Journal Edit Schedule"]: js.bytes(), path["What i did "]: wd.bytes(), "xl/workbook.xml": wbtxt.encode()}
with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as o:
    for item in z.infolist():
        o.writestr(item, edited.get(item.filename, z.read(item.filename)))
print("written", OUT)
