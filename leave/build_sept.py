"""Roll the AL-LSL Provision workbook forward from August 2026 to September 2026,
the same way it was rolled from July to August. Edited straight in the XML so the
pivot tables, array formulas, comments and named views all survive.

Also completes August on the Balance Sheet tab, as the workbook's Instructions
say to: last month's accrual figures hard-coded when the new month is done.
"""
import datetime, os, re, shutil, sys, zipfile, collections, copy
from lxml import etree
import openpyxl

U = "/root/.claude/uploads/a91f56e6-b7da-5dad-8e0b-6d5f516f9360/"
AUG = U + "55bf3a59-2026-08_AL-LSL_Provision.xlsx"
SEP_RPT = U + "ea1b8da1-LeaveBalances_as_at_2026-09-30_1.xlsx"
OUT = "/home/user/Claude/leave/out/2026-09_AL-LSL_Provision.xlsx"

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
def q(t): return "{%s}%s" % (NS, t)

def colnum(s):
    n = 0
    for ch in s: n = n * 26 + ord(ch) - 64
    return n
def colname(i):
    s = ""
    while i:
        i, r = divmod(i - 1, 26); s = chr(65 + r) + s
    return s
def split(ref):
    m = re.match(r"([A-Z]+)(\d+)$", ref); return m.group(1), int(m.group(2))
def serial(d): return (d - datetime.date(1899, 12, 30)).days

class SST:
    def __init__(self, data):
        self.t = etree.fromstring(data); self.items = self.t.findall(q("si"))
        self.idx = {}
        for i, si in enumerate(self.items):
            self.idx.setdefault("".join(x.text or "" for x in si.iter(q("t"))), i)
    def get(self, s):
        if s in self.idx: return self.idx[s]
        si = etree.SubElement(self.t, q("si")); t = etree.SubElement(si, q("t")); t.text = s
        if s != s.strip(): t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        self.idx[s] = len(self.items); self.items.append(si); return self.idx[s]
    def text(self, i): return "".join(x.text or "" for x in self.items[i].iter(q("t")))
    def dump(self):
        self.t.set("count", str(len(self.items))); self.t.set("uniqueCount", str(len(self.items)))
        return etree.tostring(self.t, xml_declaration=True, encoding="UTF-8", standalone=True)

class Sheet:
    def __init__(self, data, sst):
        self.t = etree.fromstring(data); self.sd = self.t.find(q("sheetData")); self.sst = sst
        self.rows = {int(r.get("r")): r for r in self.sd.findall(q("row"))}
    def row(self, n):
        if n in self.rows: return self.rows[n]
        new = etree.Element(q("row")); new.set("r", str(n))
        later = [k for k in self.rows if k > n]
        if later: self.rows[min(later)].addprevious(new)
        else: self.sd.append(new)
        self.rows[n] = new; return new
    def cell(self, ref, create=True):
        col, n = split(ref); row = self.row(n) if create else self.rows.get(n)
        if row is None: return None
        for c in row.findall(q("c")):
            if c.get("r") == ref: return c
        if not create: return None
        c = etree.Element(q("c")); c.set("r", ref)
        after = None
        for x in row.findall(q("c")):
            if colnum(split(x.get("r"))[0]) < colnum(col): after = x
        if after is None: row.insert(0, c)
        else: after.addnext(c)
        if row.get("spans"): row.attrib.pop("spans")
        return c
    def value(self, ref):
        c = self.cell(ref, create=False)
        if c is None: return None
        v = c.find(q("v"))
        if v is None: return None
        if c.get("t") == "s": return self.sst.text(int(v.text))
        if c.get("t") in ("str", "inlineStr", "e"): return v.text
        return float(v.text)
    def formula(self, ref):
        c = self.cell(ref, create=False)
        if c is None: return None
        f = c.find(q("f")); return None if f is None else (f.text or "")
    def _clear(self, c):
        for ch in list(c):
            c.remove(ch)
        for a in ("t", "cm", "vm"):
            if a in c.attrib: del c.attrib[a]
    def set(self, ref, kind, val=None, style_from=None, cached=None, array=False):
        c = self.cell(ref)
        if style_from and not c.get("s"):
            s = self.cell(style_from, create=False)
            if s is not None and s.get("s"): c.set("s", s.get("s"))
        self._clear(c)
        if kind == "n":
            etree.SubElement(c, q("v")).text = repr(float(val)) if float(val) != int(float(val)) else str(int(float(val)))
        elif kind == "s":
            c.set("t", "s"); etree.SubElement(c, q("v")).text = str(self.sst.get(val))
        elif kind == "f":
            f = etree.SubElement(c, q("f")); f.text = val
            if array: f.set("t", "array"); f.set("ref", ref)
            if cached is not None:
                if isinstance(cached, bool):
                    c.set("t", "b"); etree.SubElement(c, q("v")).text = "1" if cached else "0"
                elif isinstance(cached, str):
                    c.set("t", "str"); etree.SubElement(c, q("v")).text = cached
                else:
                    etree.SubElement(c, q("v")).text = repr(round(float(cached), 10))
        elif kind == "blank":
            pass
    def hardcode(self, ref):
        """Replace a formula with the value it shows."""
        c = self.cell(ref, create=False)
        if c is None: return
        f = c.find(q("f"))
        if f is None: return
        v = c.find(q("v")); t = c.get("t")
        self._clear(c)
        if v is not None:
            if t in ("str",):
                c.set("t", "s"); etree.SubElement(c, q("v")).text = str(self.sst.get(v.text or ""))
            elif t == "b":
                c.set("t", "b"); etree.SubElement(c, q("v")).text = v.text
            else:
                etree.SubElement(c, q("v")).text = v.text
    def move(self, src, dst):
        """Move one cell's contents and style to another address (cut and paste)."""
        a = self.cell(src, create=False)
        if a is None: return
        b = self.cell(dst)
        for ch in list(b): b.remove(ch)
        for k in list(b.attrib):
            if k != "r": del b.attrib[k]
        for k, v in a.attrib.items():
            if k != "r": b.set(k, v)
        for ch in a: b.append(copy.deepcopy(ch))
        f = b.find(q("f"))
        if f is not None and f.get("ref"): f.set("ref", dst)
        a.getparent().remove(a)
    def fix_dimension(self, ref):
        d = self.t.find(q("dimension"))
        if d is not None: d.set("ref", ref)
    def dump(self): return etree.tostring(self.t, xml_declaration=True, encoding="UTF-8", standalone=True)


def sep_rows():
    ws = openpyxl.load_workbook(SEP_RPT, data_only=True)["Export"]
    rows = [list(r) for r in ws.iter_rows(min_row=2, max_col=12, values_only=True)
            if r[5] in ("Annual Leave", "Long Service Leave", "Time In Lieu Taken")]
    # same grouping as the August data tab: one leave type after another
    aug = openpyxl.load_workbook(AUG, data_only=True)["Data - EH Leave Balances EOM"]
    order = []
    for r in aug.iter_rows(min_row=2, max_col=6, values_only=True):
        if r[5] and r[5] not in order: order.append(r[5])
    rows.sort(key=lambda r: order.index(r[5]))      # stable, keeps report order within a type
    return rows, order


def full(r): return ("%s %s" % (r[2], r[3])).strip()


def build():
    rows, order = sep_rows()
    z = zipfile.ZipFile(AUG); parts = {n: z.read(n) for n in z.namelist()}; z.close()
    sst = SST(parts["xl/sharedStrings.xml"])
    S = {k: Sheet(parts["xl/worksheets/sheet%d.xml" % i], sst) for k, i in
         [("data", 3), ("al", 5), ("lsl", 6), ("til", 10), ("jnl", 11), ("bs", 12)]}
    rep = {}

    # ------------------------------------------------------------ 1  data tab
    d = S["data"]
    old_last = max(n for n in d.rows if d.value("F%d" % n)) if any(d.value("F%d" % n) for n in d.rows) else 1
    for i, r in enumerate(rows):
        n = 2 + i
        for j, v in enumerate(r):
            ref = "%s%d" % (colname(1 + j), n)
            if v is None: d.set(ref, "blank", style_from="%s2" % colname(1 + j))
            elif isinstance(v, (int, float)): d.set(ref, "n", v, style_from="%s2" % colname(1 + j))
            else: d.set(ref, "s", str(v), style_from="%s2" % colname(1 + j))
    new_last = 1 + len(rows)
    for n in range(new_last + 1, old_last + 1):
        for j in range(12): d.set("%s%d" % (colname(1 + j), n), "blank")
    rep["data_rows"] = len(rows); rep["data_order"] = order

    # sums the workbook will produce
    by = {t: collections.defaultdict(float) for t in order}
    for r in rows: by[r[5]][full(r).lower()] += r[9] or 0

    # ------------------------------------------------------------ 2  annual leave
    al = S["al"]
    bh_f = {n: al.formula("BH%d" % n) for n in range(3, 98)}
    # openpyxl spells out shared formulas cell by cell; the raw XML only has them on the first cell
    _ox = openpyxl.load_workbook(AUG)["Annual Leave"]
    bi_f = {n: str(_ox["BI%d" % n].value).lstrip("=") for n in range(3, 13)}
    aug_vals = {n: (al.value("BH%d" % n) or 0) for n in range(3, 98)}
    for n in range(13, 98): al.hardcode("BH%d" % n)
    al.hardcode("BH98")
    al.set("BI1", "f", "BH1+31", style_from="BH1")
    al.set("BI2", "s", "Sum of Annual Leave Total Value", style_from="BH2")
    sepv, diff = {}, {}
    for n in range(3, 98):
        name = (al.value("A%d" % n) or "").strip().lower()
        sepv[n] = round(by["Annual Leave"].get(name, 0.0), 2) if name else 0.0
        al.set("BI%d" % n, "f", bh_f[n], style_from="BH%d" % n, cached=sepv[n])
    al.set("BI98", "f", "SUM(BI3:BI97)", style_from="BH98", cached=round(sum(sepv.values()), 2))
    al.set("BJ2", "s", "Difference", style_from="BI2")
    for n in range(3, 13):
        f = re.sub(r"([A-Z]{2})(\d+)", lambda m: colname(colnum(m.group(1)) + 1) + m.group(2), bi_f[n])
        al.set("BJ%d" % n, "f", f, style_from="BI%d" % n, cached=0)
    for n in range(13, 98):
        diff[n] = round(sepv[n] - aug_vals[n], 2)
        al.set("BJ%d" % n, "f", "+BI%d-BH%d" % (n, n), style_from="BI%d" % n, cached=diff[n])
    tot_diff = round(sum(diff.values()), 2)
    al.set("BJ98", "f", "SUM(BJ3:BJ97)", style_from="BI98", cached=tot_diff)
    al.set("BK98", "f", "BI98=Export!P4", style_from="BJ98", cached=True)
    al.set("BL98", "s", "Check against export tab", style_from="BK98")
    al.set("BI101", "f", "BI98-BH98", style_from="BH101", cached=tot_diff)
    al.set("BJ101", "f", "SUM(BJ13:BJ97)", style_from="BJ101", cached=tot_diff)
    al.set("BK101", "f", 'BJ101=GETPIVOTDATA("Difference",$BP$98)', style_from="BJ101", cached=True)
    groups = {"BS100": ("Office/Admin", "Finance", "HR", "Management"), "BS101": ("Production",),
              "BS102": ("Onsite",), "BS103": ("Video",)}
    al_lines = {}
    for ref, ds in groups.items():
        f = al.formula(ref).replace("BI:BI", "BJ:BJ")
        v = round(sum(diff[n] for n in diff if al.value("B%d" % n) in ds), 2)
        al_lines[ref] = v
        al.set(ref, "f", f, cached=v, array=True)
    al.set("BS107", "f", "SUM(BS100:BS106)", cached=round(sum(al_lines.values()), 2))
    al.set("BQ111", "f", "SUM(BJ17:BJ20)", style_from="BQ111")
    rep["AL"] = {"aug_total": round(sum(aug_vals.values()), 2), "sep_total": round(sum(sepv.values()), 2),
                 "movement": tot_diff, "lines": al_lines,
                 "data_total": round(sum(by["Annual Leave"].values()), 2)}

    # ------------------------------------------------------------ 3  long service
    ls = S["lsl"]
    jk_f = {n: ls.formula("JK%d" % n) for n in range(2, 18)}
    jk_v = {n: (ls.value("JK%d" % n) or 0) for n in range(2, 18)}
    for n in range(2, 18): ls.hardcode("JK%d" % n)
    ls.hardcode("JK19")
    ls.set("JO1", "s", "Department", style_from="JJ1")
    ls.set("JP1", "n", serial(datetime.date(2026, 9, 1)), style_from="JK1")
    ls.set("JQ1", "s", "Difference", style_from="JL1")
    lsep, ldiff, ldept = {}, {}, {}
    for n in range(2, 18):
        name = ls.value("JI%d" % n); dept = ls.value("JJ%d" % n)
        if name is not None: ls.set("JN%d" % n, "s", name, style_from="JI%d" % n)
        if dept is not None: ls.set("JO%d" % n, "s", dept, style_from="JJ%d" % n)
        ldept[n] = dept
        if jk_f[n]:
            v = round(by["Long Service Leave"].get((name or "").strip().lower(), 0.0), 2)
            ls.set("JP%d" % n, "f", jk_f[n].replace("JI%d" % n, "JN%d" % n), style_from="JK%d" % n, cached=v)
        else:                                   # typed in by hand last month; keep it that way
            v = jk_v[n]; ls.set("JP%d" % n, "n", v, style_from="JK%d" % n)
        lsep[n] = v; ldiff[n] = round(v - jk_v[n], 2)
        ls.set("JQ%d" % n, "f", "+JP%d-JK%d" % (n, n), style_from="JL%d" % n, cached=ldiff[n])
    ls.set("JN19", "s", "Total", style_from="JI19")
    ls.set("JP19", "f", "SUM(JP2:JP18)", style_from="JK19", cached=round(sum(lsep.values()), 2))
    lmove = round(sum(ldiff.values()), 2)
    ls.set("JQ19", "f", "SUM(JQ2:JQ17)", style_from="JL19", cached=lmove)
    # the department summary moves along with the new block, as it did last month
    for n in range(29, 42):
        for a, b in (("JP", "JU"), ("JQ", "JV"), ("JR", "JW")):
            ls.move("%s%d" % (a, n), "%s%d" % (b, n))
    lsl_lines = {}
    names = {30: "OFFICE / ADMIN", 31: "PRODUCTION (excl VIDEO)", 32: "ONSITE (including PRD backup)",
             33: "VIDEO DEPT", 34: "INTEGRATION", 35: "Consulting"}
    for n in range(30, 36):
        f = ls.formula("JU%d" % n).replace("JJ:JJ", "JO:JO").replace("JL:JL", "JQ:JQ")
        v = round(sum(ldiff[k] for k in ldiff if ldept[k] == names[n]), 2)
        lsl_lines[n] = v
        ls.set("JU%d" % n, "f", f, cached=v)
    ls.set("JU36", "f", "SUM(JU30:JU35)", cached=round(sum(lsl_lines.values()), 2))
    ls.set("JU37", "f", "JU36=JQ19", cached=True)
    ls.fix_dimension("A1:JW41")
    rep["LSL"] = {"movement": lmove, "lines": lsl_lines, "sep_block": round(sum(lsep.values()), 2)}

    # ------------------------------------------------------------ 4  time in lieu tab
    til = S["til"]
    look = {}
    el = openpyxl.load_workbook(AUG, data_only=True)["Employee Lookup"]
    for r in el.iter_rows(min_row=2, values_only=True):
        if r[4] is not None: look[int(r[4])] = r[7]
    trows = [r for r in rows if r[5] == "Time In Lieu Taken"]
    # FILTER keeps the Export order, which is the data sorted by leave type
    a2 = til.cell("A2"); fa = a2.find(q("f")); old_ref = fa.get("ref")
    old_end = split(old_ref.split(":")[1])[1]
    fa.set("ref", "A2:L%d" % (1 + len(trows)))
    for i, r in enumerate(trows):
        n = 2 + i
        for j, v in enumerate(r):
            ref = "%s%d" % (colname(1 + j), n)
            if ref == "A2":
                a2.find(q("v")).text = str(v); continue
            if v is None: til.set(ref, "blank")
            elif isinstance(v, (int, float)): til.set(ref, "n", v)
            else: til.set(ref, "s", str(v))
        mc = til.cell("M%d" % n, create=False)
        if mc is not None and mc.find(q("f")) is not None:
            for ch in mc.findall(q("v")): mc.remove(ch)
            mc.set("t", "str"); etree.SubElement(mc, q("v")).text = look[int(r[0])]
    for n in range(2 + len(trows), max(old_end, 27) + 1):
        for j in range(12): til.set("%s%d" % (colname(1 + j), n), "blank")
        mc = til.cell("M%d" % n, create=False)
        if mc is not None and mc.find(q("f")) is not None:
            for ch in mc.findall(q("v")): mc.remove(ch)
            mc.set("t", "b"); etree.SubElement(mc, q("v")).text = "0"
    tdept = collections.defaultdict(float); tloc = collections.defaultdict(float)
    for r in trows:
        tdept[look[int(r[0])]] += r[9] or 0; tloc[r[4]] += r[9] or 0
    for n, code in ((2, "CONS"), (3, "HR"), (4, "INT"), (5, "OFF"), (6, "ONS"), (7, "PRD"), (8, "VID")):
        c = til.cell("Q%d" % n, create=False)
        if c is not None and c.find(q("v")) is not None: c.find(q("v")).text = repr(round(tdept.get(code, 0.0), 2))
    # Job labels typed beside the pivot (AD) follow the pivot's row order once
    # the (blank) row is filtered out and the refresh drops locations with no TIL.
    order = ["Bank West Contract [ONS]", "CBA Brisbane ONS [ONS]", "CBA Team Leader [ONS]", "CBA Tech TN [ONS]",
             "DTTL ADL Onsite [ONS]", "DTTL BNE Onsite [ONS]", "DTTL Mel Tech FT 1 [ONS]", "DTTL Mel Tech FT 2 [ONS]",
             "OFFICE / ADMIN [CTS]", "PRODUCTION [PRD]", "Ricoh - QANTAS [ONS]", "VIDEO DEPT [VID]",
             "DTTL SYD Internal [ONS]", "DTTL SYD FT Tech 1 [ONS]", "DTTL SYD NTSL [ONS]", "DTTL SYD FT Tech 2 [ONS]",
             "INTEGRATION [INT]"]
    labels = {}
    for n in range(9, 27):
        v = til.value("AD%d" % n)
        if v: labels[v.split(" - ", 1)[1]] = v
    labels["DTTL Mel Tech FT 1 [ONS]"] = "725 - DTTL Mel Tech FT 1 [ONS]"
    shown = [loc for loc in order if tloc.get(loc)]
    assert set(shown) == {k for k, v in tloc.items() if v}, "new TIL location not in the pivot order"
    for n in range(9, 27):
        i = n - 9
        if i < len(shown): til.set("AD%d" % n, "s", labels[shown[i]])
        else: til.set("AD%d" % n, "blank")
    rep["TIL"] = {"by_dept": {k: round(v, 2) for k, v in tdept.items()},
                  "by_location": {k: round(v, 2) for k, v in tloc.items()},
                  "total": round(sum(tloc.values()), 2)}

    # ------------------------------------------------------------ 5  journals tab
    j = S["jnl"]
    j.set("C1", "n", serial(datetime.date(2026, 9, 30)), style_from="C1")
    typed = {"D29": round(tloc.get("PRODUCTION [PRD]", 0), 2), "D32": round(tloc.get("VIDEO DEPT [VID]", 0), 2),
             "D35": round(tloc.get("INTEGRATION [INT]", 0), 2), "D44": round(tloc.get("CBA Brisbane ONS [ONS]", 0), 2),
             "D74": round(tloc.get("DTTL ADL Onsite [ONS]", 0), 2), "D41": 0}
    for ref, v in typed.items(): j.set(ref, "n", v)
    j.set("D65", "f", 'GETPIVOTDATA("Leave Value",\'Time in Lieu\'!$W$7,"Location","DTTL Mel Tech FT 1 [ONS]")',
          cached=round(tloc.get("DTTL Mel Tech FT 1 [ONS]", 0), 2))
    for ref in ("P11", "P22", "P30"): j.set(ref, "s", "No")
    for n, (a, b) in zip(range(16, 22), [(30, 30), (31, 31), (32, 32), (33, 33), (34, 34), (35, 35)]):
        for col in "CD":
            j.set("%s%d" % (col, n), "f", j.formula("%s%d" % (col, n)).replace("$JP%d" % a, "$JU%d" % a))
        j.set("E%d" % n, "f", j.formula("E%d" % n).replace("JQ%d" % a, "JV%d" % a))
        j.set("J%d" % n, "f", j.formula("J%d" % n).replace("JR%d" % a, "JW%d" % a))
    j.set("N22", "f", j.formula("N22").replace("JP36", "JU36"))

    # ------------------------------------------------------------ 6  balance sheet
    b = S["bs"]
    for ref, v in (("B3", -1654.68), ("B9", 920.66), ("B15", -103.02), ("B17", -214.10), ("B20", 0)):
        b.set(ref, "n", v)
    for ref, v in (("C3", 1113.61), ("C9", 1811.55), ("C15", 1208.90), ("C16", 82.85),
                   ("C17", 2566.49), ("C20", 2131.28)):
        b.set(ref, "n", v, style_from="B%s" % ref[1:])
    b.set("D3", "f", "'Annual Leave'!BI101", style_from="C3", cached=tot_diff)
    b.set("D9", "f", "'Long Service Leave'!JQ19", style_from="C9", cached=lmove)
    b.set("D15", "f", "-1208.9+'Time in Lieu'!Q7", style_from="C15", cached=round(tdept.get("PRD", 0) - 1208.9, 2))
    b.set("D16", "f", "-82.85+'Time in Lieu'!Q8", style_from="C16", cached=round(tdept.get("VID", 0) - 82.85, 2))
    b.set("D17", "f", "-3357.32+'Time in Lieu'!Q6", style_from="C17", cached=round(tdept.get("ONS", 0) - 3357.32, 2))
    b.set("D20", "f", "-2190.44+'Time in Lieu'!Q5", style_from="C20", cached=round(tdept.get("OFF", 0) - 2190.44, 2))
    for c18 in ("C18", "C19", "D18", "D19"):
        pass

    # ------------------------------------------------------------ 7  package
    for k, i in [("data", 3), ("al", 5), ("lsl", 6), ("til", 10), ("jnl", 11), ("bs", 12)]:
        parts["xl/worksheets/sheet%d.xml" % i] = S[k].dump()
    parts["xl/sharedStrings.xml"] = sst.dump()
    for i in (1, 2):
        p = "xl/pivotCache/pivotCacheDefinition%d.xml" % i
        s = parts[p].decode("utf8")
        s = s.replace('<worksheetSource ref="B2:BI97" sheet="Annual Leave"/>',
                      '<worksheetSource ref="B2:BJ97" sheet="Annual Leave"/>')
        if "refreshOnLoad=" not in s:
            s = s.replace("<pivotCacheDefinition ", '<pivotCacheDefinition refreshOnLoad="1" ', 1)
        parts[p] = s.encode("utf8")
    # Time in Lieu pivot: filter out the (blank) row the empty source rows create
    p = "xl/pivotTables/pivotTable2.xml"
    s = parts[p].decode("utf8")
    assert s.count('<item x="16"/>') == 1 and s.count('<i><x v="18"/></i>') == 1
    s = s.replace('<item x="16"/>', '<item h="1" x="16"/>')
    s = s.replace('<i><x v="18"/></i>', "").replace('<rowItems count="18">', '<rowItems count="17">')
    s = s.replace('<location ref="W7:AC26"', '<location ref="W7:AC25"')
    parts[p] = s.encode("utf8")
    w = parts["xl/workbook.xml"].decode("utf8")
    w = re.sub(r"<calcPr([^>]*)/>", lambda m: "<calcPr%s fullCalcOnLoad=\"1\"/>" % m.group(1).replace(' fullCalcOnLoad="1"', ""), w)
    parts["xl/workbook.xml"] = w.encode("utf8")
    if "xl/calcChain.xml" in parts:
        del parts["xl/calcChain.xml"]
        r = parts["xl/_rels/workbook.xml.rels"].decode("utf8")
        r = re.sub(r'<Relationship [^>]*Target="calcChain.xml"/>', "", r)
        parts["xl/_rels/workbook.xml.rels"] = r.encode("utf8")
        ct = parts["[Content_Types].xml"].decode("utf8")
        ct = re.sub(r'<Override PartName="/xl/calcChain.xml"[^>]*/>', "", ct)
        parts["[Content_Types].xml"] = ct.encode("utf8")
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    zo = zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED)
    for n, data in parts.items(): zo.writestr(n, data)
    zo.close()
    return rep


if __name__ == "__main__":
    import json
    rep = build()
    json.dump(rep, open("/home/user/Claude/leave/out/september-build.json", "w"), indent=1, default=str)
    print(json.dumps(rep, indent=1, default=str))
