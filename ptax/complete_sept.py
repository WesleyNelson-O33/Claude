"""Complete the September payroll tax workbook the way Gi does it.

Everything is written straight into the workbook XML. The file carries a pivot
table, threaded comments, drawings and external links, and openpyxl drops all
of those on a round trip.

Source of the figures: the EH Payroll Split File, PRT tab (the pivot filtered
to September, grouped on residential state) for wages and super, and the split
file's EH data (2) tab for the hours that drive the job number allocation.
"""
import collections, csv, datetime, json, os, re, shutil, zipfile
from lxml import etree

SRC = "/home/user/Claude/ptax/2026-09 Payroll Tax.xlsx"
SPLIT = "/home/user/Claude/ptax/EH Payroll Split File FY2027.xlsx"
OUT = "/home/user/Claude/ptax/out"

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
def q(t):
    return "{%s}%s" % (NS, t)

TS_SHEET = "September 2026 Timesheets"
MONTH = "September"

# ---------------------------------------------------------------- source data

def colname(i):
    s = ""
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s

def colnum(s):
    n = 0
    for ch in s:
        n = n * 26 + ord(ch) - 64
    return n

def read_september():
    """Timesheet rows and hours by state and job, from the split file."""
    import openpyxl
    wb = openpyxl.load_workbook(SPLIT, data_only=True, read_only=True)
    ws = wb["EH data (2)"]
    it = ws.iter_rows(min_row=2, values_only=True)
    next(it)
    rows = []
    hrs = collections.defaultdict(lambda: collections.defaultdict(float))
    for r in it:
        if r[0] != MONTH:
            continue
        ext, loc, st, u = r[13], r[14], r[21], r[10] or 0.0
        job = "%s - %s" % (
            ("%d" % ext) if isinstance(ext, (int, float)) and float(ext).is_integer()
            else ("" if ext is None else str(ext)),
            "" if loc is None else loc)
        rows.append((r[6], ext, loc, u, st))
        hrs[st][job] += u
    wb.close()
    return rows, {s: dict(v) for s, v in hrs.items()}

# -------------------------------------------------------------- xml machinery

class Strings:
    def __init__(self, data):
        self.t = etree.fromstring(data)
        self.items = self.t.findall(q("si"))
        self.index = {}
        for i, si in enumerate(self.items):
            self.index.setdefault(self.text(si), i)

    @staticmethod
    def text(si):
        return "".join(n.text or "" for n in si.iter(q("t")))

    def id(self, s):
        if s in self.index:
            return self.index[s]
        si = etree.SubElement(self.t, q("si"))
        t = etree.SubElement(si, q("t"))
        t.text = s
        if s != s.strip() or "\n" in s:
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        self.index[s] = len(self.items)
        self.items.append(si)
        return self.index[s]

    def dump(self):
        n = len(self.items)
        self.t.set("count", str(n))
        self.t.set("uniqueCount", str(n))
        return etree.tostring(self.t, xml_declaration=True, encoding="UTF-8", standalone=True)


class Sheet:
    def __init__(self, data, strings):
        self.t = etree.fromstring(data)
        self.sd = self.t.find(q("sheetData"))
        self.ss = strings

    def rowel(self, n):
        for r in self.sd.findall(q("row")):
            if int(r.get("r")) == n:
                return r
        prev = None
        for r in self.sd.findall(q("row")):
            if int(r.get("r")) < n:
                prev = r
        new = etree.Element(q("row"))
        new.set("r", str(n))
        if prev is None:
            self.sd.insert(0, new)
        else:
            prev.addnext(new)
        return new

    @staticmethod
    def cells(row):
        out = {}
        for c in row.findall(q("c")):
            out[colnum(re.match(r"[A-Z]+", c.get("r")).group(0))] = c
        return out

    def set_row(self, n, spec, style_from=None):
        """spec: ordered [(col, kind, value)]; kind in n/s/f/blank.

        Styles are reused from the cell that is already at that column; where
        there is none, from the column named in style_from.
        """
        row = self.rowel(n)
        old = self.cells(row)
        fallback = old.get(style_from) if style_from else None
        kept = {}
        for col, kind, val in spec:
            src = old.get(col, fallback)
            c = etree.Element(q("c"))
            c.set("r", "%s%d" % (colname(col), n))
            if src is not None and src.get("s"):
                c.set("s", src.get("s"))
            if kind == "n":
                etree.SubElement(c, q("v")).text = repr(float(val)) if val != int(val) else str(int(val))
            elif kind == "s":
                c.set("t", "s")
                etree.SubElement(c, q("v")).text = str(self.ss.id(val))
            elif kind == "f":
                etree.SubElement(c, q("f")).text = val
            kept[col] = c
        for c in row.findall(q("c")):
            row.remove(c)
        for col in sorted(kept):
            row.append(kept[col])
        if row.get("spans"):
            cols = sorted(kept)
            row.set("spans", "%d:%d" % (cols[0], cols[-1]) if cols else "1:1")
        return row

    def replace_tail(self, n, start_col, spec, style_from=None):
        """Rewrite the cells from start_col rightwards, leaving the rest."""
        row = self.rowel(n)
        old = self.cells(row)
        fallback = old.get(style_from) if style_from else None
        for col, c in old.items():
            if col >= start_col:
                row.remove(c)
        for col, kind, val in spec:
            src = old.get(col, fallback)
            c = etree.SubElement(row, q("c"))
            c.set("r", "%s%d" % (colname(col), n))
            if src is not None and src.get("s"):
                c.set("s", src.get("s"))
            if kind == "n":
                etree.SubElement(c, q("v")).text = repr(float(val)) if val != int(val) else str(int(val))
            elif kind == "s":
                c.set("t", "s")
                etree.SubElement(c, q("v")).text = str(self.ss.id(val))
            elif kind == "f":
                etree.SubElement(c, q("f")).text = val
        if row.get("spans"):
            cols = sorted(self.cells(row))
            row.set("spans", "%d:%d" % (cols[0], cols[-1]))
        return row

    def retire_rows(self, keep_upto):
        for r in list(self.sd.findall(q("row"))):
            if int(r.get("r")) > keep_upto:
                self.sd.remove(r)

    def fix_dimension(self):
        maxr = 0
        maxc = 0
        for r in self.sd.findall(q("row")):
            maxr = max(maxr, int(r.get("r")))
            for c in r.findall(q("c")):
                maxc = max(maxc, colnum(re.match(r"[A-Z]+", c.get("r")).group(0)))
        d = self.t.find(q("dimension"))
        if d is not None and maxr and maxc:
            d.set("ref", "A1:%s%d" % (colname(maxc), maxr))

    def set_autofilter(self, ref):
        a = self.t.find(q("autoFilter"))
        if a is not None:
            a.set("ref", ref)

    def dump(self):
        return etree.tostring(self.t, xml_declaration=True, encoding="UTF-8", standalone=True)

# ------------------------------------------------------------------ the build

CC = ('IF(ISNUMBER(SEARCH("PRD",{r}{cc})),"Production",'
      'IF(ISNUMBER(SEARCH("CONS",{r}{cc})),"Consulting",'
      'IF(ISNUMBER(SEARCH("ONS",{r}{cc})),"Onsite",'
      'IF(ISNUMBER(SEARCH("INT",{r}{cc})),"Integration",'
      'IF(ISNUMBER(SEARCH("VID",{r}{cc})),"Video",'
      'IF(ISNUMBER(SEARCH("CTS",{r}{cc})),"CTS",""))))))')

def cost_centre(job):
    for tag, name in (("PRD", "Production"), ("CONS", "Consulting"), ("ONS", "Onsite"),
                      ("INT", "Integration"), ("VID", "Video"), ("CTS", "CTS")):
        if tag in job:
            return name
    return ""

OTHER = ["QLD", "SA", "VIC", "WA"]


def job_order(previous, present):
    """Gi's column order: last month's order for jobs that are still there,
    then any new job appended on the right."""
    keep = [j for j in previous if j in present]
    return keep + sorted(j for j in present if j not in previous)


def build():
    os.makedirs(OUT, exist_ok=True)
    ts_rows, hrs = read_september()
    nsw = hrs["NSW"]
    other_present = set()
    for s in OTHER:
        other_present |= set(hrs.get(s, {}))

    z = zipfile.ZipFile(SRC)
    parts = {n: z.read(n) for n in z.namelist()}
    z.close()

    ss = Strings(parts["xl/sharedStrings.xml"])

    # previous column order, read off the sheets before they are rewritten
    alloc = Sheet(parts["xl/worksheets/sheet11.xml"], ss)
    prev_nsw = []
    for col, c in sorted(Sheet.cells(alloc.rowel(47)).items()):
        if col == 1:
            continue
        v = ss.text(ss.items[int(c.find(q("v")).text)]) if c.get("t") == "s" else None
        if v and v != "Grand Total":
            prev_nsw.append(v)
    osheet = Sheet(parts["xl/worksheets/sheet13.xml"], ss)
    prev_os = []
    for col, c in sorted(Sheet.cells(osheet.rowel(2)).items()):
        if col == 1:
            continue
        v = ss.text(ss.items[int(c.find(q("v")).text)]) if c.get("t") == "s" else None
        if v and v != "Grand Total":
            prev_os.append(v)

    nsw_order = job_order(prev_nsw, nsw)
    os_order = job_order(prev_os, other_present)
    assert len(nsw_order) == len(nsw)
    assert len(os_order) == len(other_present)

    report = {"nsw_jobs": len(nsw_order), "other_jobs": len(os_order),
              "nsw_hours": round(sum(nsw.values()), 5),
              "other_hours": {s: round(sum(hrs.get(s, {}).values()), 5) for s in OTHER}}

    # ---------------------------------------------------------- 1  timesheets
    ts = Sheet(parts["xl/worksheets/sheet5.xml"], ss)
    ts.retire_rows(2)
    for i, (name, ext, loc, units, state) in enumerate(ts_rows):
        r = 3 + i
        spec = [(1, "s", name or ""),
                (2, "n", ext) if isinstance(ext, (int, float)) else (2, "s", "" if ext is None else str(ext)),
                (3, "s", loc or ""),
                (4, "f", 'B%d&" - "&C%d' % (r, r)),
                (5, "n", units),
                (6, "s", state or "")]
        ts.set_row(r, spec, style_from=1)
    last = 2 + len(ts_rows)
    ts.set_autofilter("A2:S%d" % last)
    ts.fix_dimension()
    parts["xl/worksheets/sheet5.xml"] = ts.dump()
    report["timesheet_rows"] = len(ts_rows)

    # ------------------------------------------- 2  other states pivot values
    def pivot_block(sheet, hdr_row, first_state_row, order, totals_row, sum_row=None):
        gt = 2 + len(order)
        sheet.set_row(hdr_row, [(1, "s", "Row Labels")] +
                      [(2 + i, "s", j) for i, j in enumerate(order)] +
                      [(gt, "s", "Grand Total")], style_from=2)
        for k, st in enumerate(OTHER):
            spec = [(1, "s", st)]
            tot = 0.0
            for i, j in enumerate(order):
                v = hrs.get(st, {}).get(j)
                if v is not None:
                    spec.append((2 + i, "n", round(v, 5)))
                    tot += v
            spec.append((gt, "n", round(tot, 5)))
            sheet.set_row(first_state_row + k, spec, style_from=2)
        spec = [(1, "s", "Grand Total")]
        grand = 0.0
        for i, j in enumerate(order):
            v = sum(hrs.get(st, {}).get(j, 0.0) for st in OTHER)
            spec.append((2 + i, "n", round(v, 5)))
            grand += v
        spec.append((gt, "n", round(grand, 5)))
        sheet.set_row(totals_row, spec, style_from=2)
        if sum_row is not None:
            sheet.set_row(sum_row, [(2 + i, "f", "SUM(%s%d:%s%d)" % (
                colname(2 + i), first_state_row, colname(2 + i), first_state_row + 3))
                for i in range(len(order))], style_from=2)
        return gt

    os_gt = pivot_block(osheet, 2, 3, os_order, 7, sum_row=1)
    osheet.set_autofilter("A2:%s2" % colname(os_gt - 1))
    osheet.fix_dimension()
    parts["xl/worksheets/sheet13.xml"] = osheet.dump()
    report["other_grand_total_col"] = colname(os_gt)

    # the same block is pasted at the top of PTax Allocations NSW
    pivot_block(alloc, 2, 3, os_order, 7)

    # ---------------------------------------------- 3  the NSW allocation
    gt47 = pivot_block_nsw(alloc, nsw_order, nsw)
    gtc = colname(gt47)
    n = len(nsw_order)
    jobs_last = 5 + n            # last journal column
    gt_mirror = jobs_last + 1

    payable = nsw_payable()
    total_hours = sum(nsw.values())
    raw = [payable * nsw[j] / total_hours for j in nsw_order]
    plug = round(round(payable, 2) - sum(round(v, 2) for v in raw), 2)
    report["nsw_payable"] = round(payable, 2)
    report["nsw_rounding_plug"] = plug

    alloc.replace_tail(74, 6, [(c, "f", CC.format(r=colname(c), cc=79))
                               for c in range(6, gt_mirror + 1)], style_from=6)
    tail77 = [(5, "f", "ROUND(SUM(F77:%s77),2)" % colname(jobs_last))]
    for c in range(6, gt_mirror + 1):
        f = "ROUND(%s78,2)" % colname(c)
        if c == 6 and plug:
            f += "%+.2f" % plug
        tail77.append((c, "f", f))
    alloc.replace_tail(77, 5, tail77, style_from=6)
    alloc.replace_tail(78, 5, [(5, "f", "+E79-'Main & NSW'!D23")] +
                       [(c, "f", "SUBTOTAL(9,%s80:%s90)" % (colname(c), colname(c)))
                        for c in range(6, gt_mirror + 1)], style_from=6)
    alloc.set_row(79, [(1, "s", "State"), (2, "s", "Row Labels"),
                       (3, "f", "SUBTOTAL(9,C80:C88)"), (4, "f", "SUBTOTAL(9,D80:D88)"),
                       (5, "f", "SUM(E80:E82)")] +
                  [(c, "f", "%s$47" % colname(c - 4)) for c in range(6, gt_mirror + 1)],
                  style_from=6)
    alloc.set_row(80, [(1, "s", "NSW"), (2, "s", "NSW"),
                       (3, "f", "'Main & NSW'!D12"), (4, "f", "+C80/$C$79"),
                       (5, "f", "D80*'Main &amp; NSW'!D23".replace("&amp;", "&"))] +
                  [(c, "f", "$E80*(%s48/$%s48)" % (colname(c - 4), gtc))
                   for c in range(6, jobs_last + 1)] +
                  [(gt_mirror, "f", "$E80*(%s48/$%s48)" % (gtc, gtc))], style_from=6)
    # the journal tick-off block: one NSW journal this month
    alloc.replace_tail(86, 6, [(6, "n", round(payable, 2))], style_from=6)
    alloc.replace_tail(87, 6, [], style_from=6)
    alloc.set_autofilter("A47:%s49" % colname(gt47 - 1))
    alloc.fix_dimension()
    parts["xl/worksheets/sheet11.xml"] = alloc.dump()
    report["nsw_grand_total_col"] = gtc

    # ------------------------------------------ 4  the other states accrual
    acc = Sheet(parts["xl/worksheets/sheet14.xml"], ss)
    m = len(os_order)
    oj_last = 5 + m
    osgt = colname(os_gt)
    acc.replace_tail(12, 6, [(c, "f", CC.format(r=colname(c), cc=16))
                             for c in range(6, oj_last + 1)], style_from=6)
    acc.replace_tail(14, 5, [(5, "f", "ROUND(SUM(F14:%s14),0)" % colname(oj_last))] +
                     [(c, "f", "ROUND(%s15,2)" % colname(c)) for c in range(6, oj_last + 1)],
                     style_from=6)
    acc.replace_tail(15, 6, [(c, "f", "SUM(%s17:%s22)" % (colname(c), colname(c)))
                             for c in range(6, oj_last + 1)], style_from=6)
    acc.replace_tail(16, 6, [(c, "f", "'Other States PTAX Allocations'!%s$2" % colname(c - 4))
                             for c in range(6, oj_last + 1)], style_from=6)
    acc.replace_tail(17, 6, [(c, "n", 0) for c in range(6, oj_last + 1)], style_from=6)
    for k, st in enumerate(OTHER):
        r = 18 + k
        pr = 3 + k
        acc.replace_tail(r, 6, [
            (c, "f", "$E%d*('Other States PTAX Allocations'!%s%d/"
                     "'Other States PTAX Allocations'!$%s%d)"
             % (r, colname(c - 4), pr, osgt, pr)) for c in range(6, oj_last + 1)],
            style_from=6)
    acc.fix_dimension()
    parts["xl/worksheets/sheet14.xml"] = acc.dump()

    # -------------------------------------- 5  rename the timesheets sheet
    wbx = parts["xl/workbook.xml"].decode("utf8")
    wbx = wbx.replace('<sheet name="August 2026 Timesheets"',
                      '<sheet name="%s"' % TS_SHEET)
    wbx = wbx.replace("'August 2026 Timesheets'!$A$2:$S$1832",
                      "'%s'!$A$2:$S$%d" % (TS_SHEET, last))
    wbx = wbx.replace("'Other States = Accrual'!$A$15:$X$15",
                      "'Other States = Accrual'!$A$15:$%s$15" % colname(oj_last))
    wbx = wbx.replace("'Other States PTAX Allocations'!$A$2:$W$2",
                      "'Other States PTAX Allocations'!$A$2:$%s$2" % colname(os_gt - 1))
    wbx = wbx.replace("'PTax Allocations NSW'!$A$47:$CJ$49",
                      "'PTax Allocations NSW'!$A$47:$%s$49" % colname(gt47 - 1))
    parts["xl/workbook.xml"] = wbx.encode("utf8")

    pc = parts["xl/pivotCache/pivotCacheDefinition1.xml"].decode("utf8")
    pc = pc.replace('<worksheetSource ref="A2:F5236" sheet="August 2026 Timesheets"/>',
                    '<worksheetSource ref="A2:F%d" sheet="%s"/>' % (last, TS_SHEET))
    if "refreshOnLoad=" in pc:
        pc = re.sub(r'refreshOnLoad="[^"]*"', 'refreshOnLoad="1"', pc, count=1)
    else:
        pc = pc.replace("<pivotCacheDefinition ", '<pivotCacheDefinition refreshOnLoad="1" ', 1)
    parts["xl/pivotCache/pivotCacheDefinition1.xml"] = pc.encode("utf8")

    app = parts["docProps/app.xml"].decode("utf8")
    parts["docProps/app.xml"] = app.replace(
        "August 2026 Timesheets", TS_SHEET).encode("utf8")

    parts["xl/sharedStrings.xml"] = ss.dump()

    dst = os.path.join(OUT, "2026-09 Payroll Tax.xlsx")
    zo = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
    for name, data in parts.items():
        zo.writestr(name, data)
    zo.close()

    journals(dst, nsw_order, nsw, payable, plug, os_order, hrs, report)
    json.dump(report, open(os.path.join(OUT, "september-build.json"), "w"), indent=1)
    for k, v in report.items():
        print("%-24s %s" % (k, v))
    return dst


def pivot_block_nsw(sheet, order, nsw):
    gt = 2 + len(order)
    sheet.set_row(47, [(1, "s", "Row Labels")] +
                  [(2 + i, "s", j) for i, j in enumerate(order)] +
                  [(gt, "s", "Grand Total")], style_from=2)
    tot = round(sum(nsw.values()), 5)
    for r, label in ((48, "NSW"), (49, "Grand Total")):
        sheet.set_row(r, [(1, "s", label)] +
                      [(2 + i, "n", round(nsw[j], 5)) for i, j in enumerate(order)] +
                      [(gt, "n", tot)], style_from=2)
    return gt


SAFE = re.compile(r"^[0-9eE.+\-*/() ]+$")


def arith(v):
    """Evaluate the plain arithmetic Danica types into the wages lines."""
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).lstrip("=").strip()
    if not s:
        return 0.0
    if not SAFE.match(s):
        raise ValueError("not plain arithmetic: %r" % v)
    return float(eval(s, {"__builtins__": {}}, {}))


def nsw_payable():
    """The September NSW payroll tax payable, off the figures already in the
    workbook, using the Main & NSW tab's own formulas."""
    import openpyxl
    wb = openpyxl.load_workbook(SRC, data_only=False)
    main = wb["Main & NSW"]
    nsw_wages = sum(arith(main.cell(row=r, column=4).value) for r in range(3, 12))
    inter = 0.0
    for tab in ("ACT", "QLD", "SA", "VIC", "WA"):
        ws = wb[tab]
        inter += sum(arith(ws.cell(row=r, column=4).value) for r in range(3, 13))
    days = arith(main["D1"].value)
    year = sum(arith(main.cell(row=1, column=c).value) for c in range(2, 14))
    rate = arith(main["A22"].value)
    wb.close()
    total = nsw_wages + inter
    threshold = (nsw_wages / total) * (1200000.0 / year * days)
    return (total - threshold - inter) * rate


XERO_HEAD = ["*Narration", "*Date", "Description", "*AccountCode", "TaxRate", "*Amount",
             "TrackingName1", "TrackingOption1", "TrackingName2", "TrackingOption2"]
DEBIT, CREDIT = "65106", "21440"
ADMIN_JOB = "9000 - OFFICE / ADMIN [CTS]"
BAS = "BAS Excluded"


def write_journal(path, narration, date, lines):
    total = round(sum(a for _, a in lines), 2)
    with open(path, "w", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(XERO_HEAD)
        for job, amt in lines:
            if round(amt, 2) == 0:
                continue
            w.writerow([narration, date, job, DEBIT, BAS, "%.2f" % amt,
                        "Cost Centres", cost_centre(job), "Job Numbers", job])
        w.writerow([narration, date, narration, CREDIT, BAS, "%.2f" % -total,
                    "Cost Centres", "CTS", "Job Numbers", ADMIN_JOB])
    return total


def journals(dst, nsw_order, nsw, payable, plug, os_order, hrs, report):
    date = "30/09/2026"
    tot = sum(nsw.values())
    lines = []
    for i, j in enumerate(nsw_order):
        amt = round(payable * nsw[j] / tot, 2)
        if i == 0:
            amt = round(amt + plug, 2)
        lines.append((j, amt))
    report["nsw_journal_total"] = write_journal(
        os.path.join(OUT, "2026-09 Payroll Tax NSW Xero Journal.csv"),
        "Payroll Tax NSW September 2026", date, lines)

    accrual = {"QLD": 500.0, "SA": 200.0, "VIC": 1300.0, "WA": 700.0}
    olines = []
    for j in os_order:
        amt = 0.0
        for st in OTHER:
            h = hrs.get(st, {})
            if j in h and sum(h.values()):
                amt += accrual[st] * h[j] / sum(h.values())
        olines.append((j, round(amt, 2)))
    report["other_states_journal_total"] = write_journal(
        os.path.join(OUT, "2026-09 Payroll Tax Other States Xero Journal - ACCRUAL.csv"),
        "Payroll Tax ACT/QLD/SA/VIC/WA September 2026 Accrual", date, olines)


if __name__ == "__main__":
    print(build())
