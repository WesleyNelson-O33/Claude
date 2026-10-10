"""Bring the FY27 Revenue Tracker up to date for September 2026.

Copies the V2 tracker lines that were not yet in FY27, and corrects the
Finance cream cells (invoice no, date, ex GST) to what is posted in Xero.
Edits the sheet XML directly so tables, comments and protection survive.
"""
import datetime, re, shutil, sys, zipfile
from lxml import etree

SRC, OUT = sys.argv[1], sys.argv[2]
NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
N = "{%s}" % NS
SHEETS = {"Finance": "xl/worksheets/sheet3.xml", "Onsite": "xl/worksheets/sheet8.xml",
          "Production": "xl/worksheets/sheet9.xml", "WIP Movements": "xl/worksheets/sheet11.xml"}

def serial(s):
    return (datetime.date.fromisoformat(s) - datetime.date(1899, 12, 30)).days

def colnum(c):
    n = 0
    for ch in c: n = n * 26 + ord(ch) - 64
    return n

def val_of(c):
    if c.get("t") == "s":
        return SST[int(c.find(N + "v").text)]
    if c.get("t") == "inlineStr":
        return "".join(c.find(N + "is").itertext())
    v = c.find(N + "v")
    return "" if v is None or v.text is None else v.text

class Sheet:
    def __init__(self, xml):
        self.root = etree.fromstring(xml)
        self.data = self.root.find(N + "sheetData")
        self.rows = {int(r.get("r")): r for r in self.data.findall(N + "row")}
    def cell(self, ref):
        col, r = re.match(r"([A-Z]+)(\d+)", ref).groups(); r = int(r)
        row = self.rows[r]
        for c in row.findall(N + "c"):
            if c.get("r") == ref: return c
        c = etree.SubElement(row, N + "c"); c.set("r", ref)
        cs = sorted(row.findall(N + "c"), key=lambda e: colnum(re.match(r"[A-Z]+", e.get("r")).group()))
        for e in cs: row.remove(e)
        for e in cs: row.append(e)
        return c
    def text(self, ref):
        return val_of(self.cell(ref))
    def set(self, ref, val):
        c = self.cell(ref)
        assert c.find(N + "f") is None, "formula cell %s" % ref
        for ch in list(c): c.remove(ch)
        c.attrib.pop("t", None)
        if val is None: return
        if isinstance(val, str):
            c.set("t", "inlineStr")
            is_ = etree.SubElement(c, N + "is"); t = etree.SubElement(is_, N + "t"); t.text = val
        else:
            v = etree.SubElement(c, N + "v"); v.text = repr(float(val)) if not float(val).is_integer() else str(int(val))
    def find_row(self, col, value):
        for r, row in self.rows.items():
            for c in row.findall(N + "c"):
                if c.get("r") == "%s%d" % (col, r) and val_of(c).strip() == value: return r
        raise KeyError(value)
    def bytes(self):
        return etree.tostring(self.root, xml_declaration=True, encoding="UTF-8", standalone=True)

z = zipfile.ZipFile(SRC)
SST = ["".join(si.itertext()) for si in etree.fromstring(z.read("xl/sharedStrings.xml")).findall(N + "si")]
sh = {k: Sheet(z.read(v)) for k, v in SHEETS.items()}
fin, ons, prd, wip = sh["Finance"], sh["Onsite"], sh["Production"], sh["WIP Movements"]

def finance(rid, inv="keep", date="keep", amt="keep"):
    r = fin.find_row("A", rid)
    if inv != "keep": fin.set("G%d" % r, inv)
    if date != "keep": fin.set("H%d" % r, None if date is None else serial(date))
    if amt != "keep": fin.set("I%d" % r, amt)

def dept(s, idcol, rid, **cells):
    r = s.find_row(idcol, rid)
    for col, v in cells.items(): s.set("%s%d" % (col, r), v)
    return r

S30 = "2026-09-30"
# ---- Onsite: DTTL Additionals September 2026 is INV-10665 in Xero, split by job
for rid, amt in (("ONS-0051", 7144.62), ("ONS-0053", 7734.61)):
    finance(rid, "INV-10665", S30, amt)
dept(ons, "X", "ONS-0064", O=892.5); finance("ONS-0064", "INV-10665", S30, 892.5)
dept(ons, "X", "ONS-0070", C="2608104", O=637.5); finance("ONS-0070", "INV-10665", S30, 637.5)
dept(ons, "X", "ONS-0066", C="2608105", O=765); finance("ONS-0066", "INV-10665", S30, 765)
r = dept(ons, "X", "ONS-0065", O=0)
ons.set("W%d" % r, "Line cancelled 10/10/2026: the same invoice is on the job rows ONS-0051, 0053, 0064, 0066 and 0070 (INV-10665). INV-10651 is not in Xero.")
finance("ONS-0065", None, None, None)
# ---- Onsite: other September contract lines, dates and amounts per Xero
dept(ons, "X", "ONS-0069", C="26022301", O=11000); finance("ONS-0069", "INV-10652", S30, 11000)
finance("ONS-0060", "INV-10667", S30, 11960)
finance("ONS-0055", date=S30, amt=7312)
finance("ONS-0056", date=S30)
finance("ONS-0061", date=S30, amt=4320)
finance("ONS-0057", date=S30, amt=22381.67)
finance("ONS-0058", date=S30, amt=11814.29)
finance("ONS-0059", date=S30)
finance("ONS-0054", date=S30)
finance("ONS-0062", date=S30)
dept(ons, "X", "ONS-0063", O=7650); finance("ONS-0063", "INV-10654", S30, 7650)
# ---- Onsite: new lines from V2
for rid, job, client, desc, amt, inv in (
        ("ONS-0073", "25101301B", "PWC", "PWC - Sydney - 12 months - September 2026", 11907, "INV-10655"),
        ("ONS-0074", "1668", "CBA", "85720001MEL435CSJUN27 AFTERHOURS", 85, "INV-10673")):
    dept(ons, "X", rid, C=job, F=client, G=desc, H="ONSITE", I="Contract", J="GST 10%", O=amt)
    finance(rid, inv, S30, amt)
# ---- Onsite: September CBA lines in Xero but in neither tracker, set up like August's 1668 rows
CBA_SEP = [("INV-10658", 425), ("INV-10659", 425), ("INV-10660", 425), ("INV-10661", 340), ("INV-10662", 425)]
for i, (inv, amt) in enumerate(CBA_SEP + [("INV-10657", 1080)]):
    rid = "ONS-%04d" % (75 + i)
    desc = "85720001CBPSCSJUN27 AFTERHOURS" if inv == "INV-10657" else "85720001CBPSCSJUN27 DEDICATED TECH"
    r = ons.find_row("X", rid); assert ons.text("C%d" % r) == "", rid
    dept(ons, "X", rid, C="1668", F="CBA", G=desc, H="ONSITE", I="Contract", J="GST 10%", O=amt,
         W="Added by Finance 10/10/2026 from Xero. Not in the V2 tracker.")
    finance(rid, inv, S30, amt)
# ---- Production: new lines from V2 (dept value from V2, Finance values from Xero)
PRD_ROWS = [
    # job, client, email, description, cost centre, event date, V2 value, invoice, inv date, Xero ex GST
    ("26081101", "Vega Global", None, "Payments Forum 15 - 17.9.26", "PRODUCTION", "2026-09-15", 1790, "INV-10647", "2026-09-22", 1790),
    ("26022302", "Global Philanthropic", None, "Talking Philanthropy 18.9.26", "PRODUCTION", "2026-09-18", 3949, "INV-10670", S30, 3949),
    ("26091702", "AusPayNet", None, "Event Support 21 & 22.9.26", "PRODUCTION", "2026-09-21", 680, "INV-10671", S30, 680),
    ("26092104", "Dimensional", None, "Webcast AV Support 22.9.26", "PRODUCTION", "2026-09-22", 460, "INV-10664", "2026-09-29", 460),
    ("26091601", "Automic", None, "September All Hands - 25.9.26", "PRODUCTION", "2026-09-25", 7978, "INV-10663", "2026-09-29", 7978),
    ("26092701", "ASX", None, "External Listing Test Days - 29.9.26", "PRODUCTION", "2026-10-29", 5225, "INV-10668", S30, 5225),
    ("26091002", "Global Philanthropic", None, "Talking Philanthropy - Event Videography", "VIDEO", "2026-09-18", 1249, "INV-10669", S30, 1249),
    ("26072003", "CBA", None, "Live Event Support September 2026", "PRODUCTION", S30, 4907.5, "INV-10675", S30, 4907.5),
    ("2591012C", "St Luke", None, "St Luke's Liverpool - Digital Mixer Equipment", "PRODUCTION", S30, 4535, "INV-10676", S30, 4535),
    ("26082502", "ICC", "jbell@iccsydney.com", "ICC - September Presenter Support Presentations - TTS", "PRODUCTION", "2026-09-23", 3727.25, "INV-10672", S30, 3727.5),
    ("26082502", "ICC", "jbell@iccsydney.com", "ICC - September Presenter Support Presentations - ComBio", "PRODUCTION", S30, 2226.25, "INV-10674", S30, 2226.25),
]
for i, (job, client, email, desc, cc, ev, val, inv, idate, xamt) in enumerate(PRD_ROWS):
    rid = "PRD-%04d" % (88 + i)
    r = prd.find_row("AU", rid)
    assert prd.text("C%d" % r) == "", rid
    dept(prd, "AU", rid, C=job, F=client, G=desc, H=cc, I="Project", J="GST 10%", O=val, V=serial(ev), AB="Y")
    if email: prd.set("T%d" % r, email)
    finance(rid, inv, idate, xamt)
# ---- Production: September invoices in Xero but in neither tracker (both are WIP deferrals)
for rid, job, client, desc, ev, amt, inv, idate in (
        ("PRD-0099", "26081302", "Ubank", "UBank - Dignity Volunteer Packing Day - 11.11.2026", "2026-11-11", 9879, "INV-10631", "2026-09-16"),
        ("PRD-0100", "26092101", "Bankwest", "Bankwest - BCEC WGEA Report Launch 13.10.2026", "2026-10-13", 950, "INV-10645", "2026-09-22")):
    r = prd.find_row("AU", rid); assert prd.text("C%d" % r) == "", rid
    dept(prd, "AU", rid, C=job, F=client, G=desc, H="PRODUCTION", I="Project", J="GST 10%", O=amt, V=serial(ev),
         AT="Added by Finance 10/10/2026 from Xero. Not in the V2 tracker. Deferred on WIP Movements.")
    finance(rid, inv, idate, amt)
# ---- WIP Movements: September labour-hours lines, same method as August
def wip_row(job, month, typ):
    for r in sorted(wip.rows):
        if r < 5: continue
        if wip.text("B%d" % r).strip() == job and wip.text("H%d" % r) == typ:
            c = wip.cell("A%d" % r); v = c.find(N + "v")
            if v is not None and int(float(v.text)) == serial(month): return r
    raise KeyError(job)
r = wip_row("25021901", S30, "Release of deferral"); wip.set("J%d" % r, 2516.25)
wip.set("R%d" % r, "Invoiced in advance. De-recognise based on labour hours. Sep 26: 15.25 hrs x $165 (utilisation report).")
r = wip_row("2506208", S30, "Accrual - unbilled work"); wip.set("J%d" % r, 6402.5)
wip.set("R%d" % r, "Based on labour hours. Sep 26: 49.25 hrs x $130 (utilisation report).")
r = wip_row("2604704", S30, "Accrual - unbilled work")
wip.set("H%d" % r, "Reversal of prior accrual"); wip.set("J%d" % r, -4585)
wip.set("R%d" % r, "Invoiced INV-10596 on 09/09/2026. Reverse the accrued labour (1,402.50 + 1,821.25 + 1,361.25). No hours in Sep 26.")
for job, acct in (("6824", "42500"), ("2755", "42100"), ("300", "45015")):
    pass
for r in sorted(wip.rows):
    if r >= 5 and wip.text("B%d" % r).strip() == "26081302":
        amt = abs(float(wip.cell("J%d" % r).find(N + "v").text))
        wip.set("K%d" % r, {6824: "42500", 2755: "42100", 300: "45015"}[int(amt)])
    if r >= 5 and wip.text("B%d" % r).strip() == "26092101":
        wip.set("K%d" % r, "42100,42150,42500")

# ---- write the package, full recalculation on open
wbxml = z.read("xl/workbook.xml").decode()
if "fullCalcOnLoad" not in wbxml:
    wbxml = re.sub(r"<calcPr([^>]*)/>", r'<calcPr\1 fullCalcOnLoad="1"/>', wbxml)
with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as o:
    for item in z.infolist():
        name = item.filename
        if name in SHEETS.values():
            data = sh[[k for k, v in SHEETS.items() if v == name][0]].bytes()
        elif name == "xl/workbook.xml":
            data = wbxml.encode()
        else:
            data = z.read(name)
        o.writestr(item, data)
print("written", OUT)
