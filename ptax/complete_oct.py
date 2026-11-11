"""Complete the October 2026 payroll tax workbook the way Gi does it.

Danica's October file already has Main & NSW done, the return lodged, and the
October timesheets pasted in. What was left: the pivot values were still
September's, the NSW allocation row held pasted hours instead of formulas, and
the other states tabs were still September. This rebuilds Parts 3 to 5 and
writes the two Xero journals.

Everything is written straight into the workbook XML. The file carries a pivot
table, drawings and external links, and openpyxl drops those on a round trip.
"""
import collections, csv, json, os, re, zipfile
from lxml import etree

SRC = "/root/.claude/uploads/fc902e4e-1220-55ba-999f-a6d0cad6db14/61b5b64f-2026-10_Payroll_Tax.xlsx"
OUT = "/home/user/Claude/ptax/out"
MONTH, YEAR, DATE = "October", "2026", "31/10/2026"
OLD_TS, TS_SHEET = "September 2026 Timesheets", "October 2026 Timesheets"
COL = "E"            # the October column on Main & NSW
ACCRUAL = {"QLD": 500.0, "SA": 200.0, "VIC": 1300.0, "WA": 700.0}
OTHER = ["QLD", "SA", "VIC", "WA"]
SHEET = {"ts": "xl/worksheets/sheet6.xml", "alloc": "xl/worksheets/sheet12.xml",
         "os": "xl/worksheets/sheet14.xml", "acc": "xl/worksheets/sheet15.xml"}

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
def q(t): return "{%s}%s" % (NS, t)

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
exec(open("/tmp/claude-0/-home-user-Claude/fc902e4e-1220-55ba-999f-a6d0cad6db14/scratchpad/xmlmachinery.py").read())

CC = ('IF(ISNUMBER(SEARCH("PRD",{r}{cc})),"Production",'
      'IF(ISNUMBER(SEARCH("CONS",{r}{cc})),"Consulting",'
      'IF(ISNUMBER(SEARCH("ONS",{r}{cc})),"Onsite",'
      'IF(ISNUMBER(SEARCH("INT",{r}{cc})),"Integration",'
      'IF(ISNUMBER(SEARCH("VID",{r}{cc})),"Video",'
      'IF(ISNUMBER(SEARCH("CTS",{r}{cc})),"CTS",""))))))')

def cost_centre(job):
    for tag, name in (("PRD", "Production"), ("CONS", "Consulting"), ("ONS", "Onsite"),
                      ("INT", "Integration"), ("VID", "Video"), ("CTS", "CTS")):
        if tag in job: return name
    return ""

def job_order(previous, present):
    keep = [j for j in previous if j in present]
    return keep + sorted(j for j in present if j not in previous)

def read_hours():
    """Hours by state and job off the timesheets tab already in the file."""
    import openpyxl
    wb = openpyxl.load_workbook(SRC, data_only=True, read_only=True)
    ws = wb[OLD_TS]
    hrs = collections.defaultdict(lambda: collections.defaultdict(float)); n = 0
    for r in ws.iter_rows(min_row=3, values_only=True):
        if r[0] is None: continue
        n += 1
        assert r[5] in ("NSW",) + tuple(OTHER), "stateless row: %r" % (r[:6],)
        hrs[r[5]][r[3]] += r[4] or 0.0
    wb.close()
    return n, {s: dict(v) for s, v in hrs.items()}

def nsw_payable():
    import openpyxl
    wb = openpyxl.load_workbook(SRC, data_only=True, read_only=True)
    m = wb["Main & NSW"]
    v = m[COL + "23"].value; wb.close()
    return float(v)

def pivot_block(sheet, hdr_row, first_state_row, order, totals_row, hrs, sum_row=None):
    gt = 2 + len(order)
    sheet.set_row(hdr_row, [(1, "s", "Row Labels")] + [(2 + i, "s", j) for i, j in enumerate(order)] +
                  [(gt, "s", "Grand Total")], style_from=2)
    for k, st in enumerate(OTHER):
        spec = [(1, "s", st)]; tot = 0.0
        for i, j in enumerate(order):
            v = hrs.get(st, {}).get(j)
            if v is not None: spec.append((2 + i, "n", round(v, 5))); tot += v
        spec.append((gt, "n", round(tot, 5)))
        sheet.set_row(first_state_row + k, spec, style_from=2)
    spec = [(1, "s", "Grand Total")]; grand = 0.0
    for i, j in enumerate(order):
        v = sum(hrs.get(st, {}).get(j, 0.0) for st in OTHER); spec.append((2 + i, "n", round(v, 5))); grand += v
    spec.append((gt, "n", round(grand, 5)))
    sheet.set_row(totals_row, spec, style_from=2)
    if sum_row is not None:
        sheet.set_row(sum_row, [(2 + i, "f", "SUM(%s%d:%s%d)" % (colname(2 + i), first_state_row, colname(2 + i), first_state_row + 3))
                                for i in range(len(order))], style_from=2)
    return gt

def allocate(order, weights, pot):
    tot = sum(weights.values())
    out = [(j, round(pot * weights[j] / tot, 2)) for j in order]
    diff = round(pot - sum(a for _, a in out), 2)
    if diff: out[0] = (out[0][0], round(out[0][1] + diff, 2))
    return out

XERO_HEAD = ["*Narration", "*Date", "Description", "*AccountCode", "*TaxRate", "*Amount",
             "TrackingName1", "TrackingOption1", "TrackingName2", "TrackingOption2"]
DEBIT, CREDIT, BAS = "65106", "21440", "BAS Excluded"
ADMIN_JOB = "9000 - OFFICE / ADMIN [CTS]"

def write_journal(path, narration, lines, credit_desc):
    total = round(sum(a for _, a in lines), 2)
    with open(path, "w", newline="") as fh:
        w = csv.writer(fh); w.writerow(XERO_HEAD)
        w.writerow([narration, DATE, credit_desc, CREDIT, BAS, "%.2f" % -total, "Cost Centres", "CTS", "Job Numbers", ADMIN_JOB])
        for job, amt in lines:
            if round(amt, 2) == 0: continue
            w.writerow([narration, DATE, job, DEBIT, BAS, "%.2f" % amt, "Cost Centres", cost_centre(job), "Job Numbers", job])
    return total

def build():
    os.makedirs(OUT, exist_ok=True)
    nrows, hrs = read_hours()
    nsw = hrs["NSW"]; last = 2 + nrows
    other_present = set()
    for s in OTHER: other_present |= set(hrs.get(s, {}))

    z = zipfile.ZipFile(SRC); parts = {n: z.read(n) for n in z.namelist()}; order_parts = [i.filename for i in z.infolist()]; z.close()
    ss = Strings(parts["xl/sharedStrings.xml"])

    alloc = Sheet(parts[SHEET["alloc"]], ss)
    prev_nsw = [ss.text(ss.items[int(c.find(q("v")).text)]) for col, c in sorted(Sheet.cells(alloc.rowel(47)).items())
                if col > 1 and c.get("t") == "s"]
    prev_nsw = [v for v in prev_nsw if v != "Grand Total"]
    osheet = Sheet(parts[SHEET["os"]], ss)
    prev_os = [ss.text(ss.items[int(c.find(q("v")).text)]) for col, c in sorted(Sheet.cells(osheet.rowel(2)).items())
               if col > 1 and c.get("t") == "s"]
    prev_os = [v for v in prev_os if v != "Grand Total"]
    nsw_order = job_order(prev_nsw, nsw); os_order = job_order(prev_os, other_present)
    assert len(nsw_order) == len(nsw) and len(os_order) == len(other_present)
    report = {"timesheet_rows": nrows, "nsw_jobs": len(nsw_order), "other_jobs": len(os_order),
              "nsw_hours": round(sum(nsw.values()), 5),
              "other_hours": {s: round(sum(hrs.get(s, {}).values()), 5) for s in OTHER},
              "new_nsw_jobs": [j for j in nsw_order if j not in prev_nsw],
              "gone_nsw_jobs": [j for j in prev_nsw if j not in nsw],
              "new_other_jobs": [j for j in os_order if j not in prev_os]}

    # 1  timesheets: already October. Trim anything below the data, fix filter and dimension.
    ts = Sheet(parts[SHEET["ts"]], ss)
    ts.retire_rows(last); ts.set_autofilter("A2:S%d" % last); ts.fix_dimension()
    parts[SHEET["ts"]] = ts.dump()

    # 2  other states pivot values
    os_gt = pivot_block(osheet, 2, 3, os_order, 7, hrs, sum_row=1)
    osheet.set_autofilter("A2:%s2" % colname(os_gt - 1)); osheet.fix_dimension()
    parts[SHEET["os"]] = osheet.dump()

    # 3  NSW pivot values and allocation
    gt47 = 2 + len(nsw_order)
    alloc.set_row(47, [(1, "s", "Row Labels")] + [(2 + i, "s", j) for i, j in enumerate(nsw_order)] + [(gt47, "s", "Grand Total")], style_from=2)
    tot = round(sum(nsw.values()), 5)
    for r, label in ((48, "NSW"), (49, "Grand Total")):
        alloc.set_row(r, [(1, "s", label)] + [(2 + i, "n", round(nsw[j], 5)) for i, j in enumerate(nsw_order)] + [(gt47, "n", tot)], style_from=2)
    gtc = colname(gt47); n = len(nsw_order); jobs_last = 5 + n; gt_mirror = jobs_last + 1
    payable = nsw_payable()
    raw = [payable * nsw[j] / tot for j in nsw_order]
    plug = round(round(payable, 2) - sum(round(v, 2) for v in raw), 2)
    report["nsw_payable"] = round(payable, 2); report["nsw_rounding_plug"] = plug
    M = "'Main & NSW'!"
    alloc.replace_tail(74, 6, [(c, "f", CC.format(r=colname(c), cc=79)) for c in range(6, gt_mirror + 1)], style_from=6)
    tail77 = [(5, "f", "ROUND(SUM(F77:%s77),2)" % colname(jobs_last))]
    for c in range(6, gt_mirror + 1):
        f = "ROUND(%s78,2)" % colname(c)
        if c == 6 and plug: f += "%+.2f" % plug
        tail77.append((c, "f", f))
    alloc.replace_tail(77, 5, tail77, style_from=6)
    alloc.replace_tail(78, 5, [(5, "f", "+E79-%s%s23" % (M, COL))] +
                       [(c, "f", "SUBTOTAL(9,%s80:%s90)" % (colname(c), colname(c))) for c in range(6, gt_mirror + 1)], style_from=6)
    alloc.set_row(79, [(1, "s", "State"), (2, "s", "Row Labels"), (3, "f", "SUBTOTAL(9,C80:C88)"), (4, "f", "SUBTOTAL(9,D80:D88)"),
                       (5, "f", "SUM(E80:E82)")] + [(c, "f", "%s$47" % colname(c - 4)) for c in range(6, gt_mirror + 1)], style_from=6)
    alloc.set_row(80, [(1, "s", "NSW"), (2, "s", "NSW"), (3, "f", "%s%s12" % (M, COL)), (4, "f", "+C80/$C$79"),
                       (5, "f", "D80*%s%s23" % (M, COL))] +
                  [(c, "f", "$E80*(%s48/$%s48)" % (colname(c - 4), gtc)) for c in range(6, jobs_last + 1)] +
                  [(gt_mirror, "f", "$E80*(%s48/$%s48)" % (gtc, gtc))], style_from=6)
    # the journal tick-off block: one NSW journal, and it agrees to the lodgement
    alloc.replace_tail(86, 6, [(6, "n", round(payable, 2))], style_from=6)
    alloc.replace_tail(87, 6, [(6, "n", round(payable, 2))], style_from=6)
    alloc.replace_tail(88, 7, [(7, "s", "Agrees to lodgement, %s %s" % (MONTH, YEAR))], style_from=7)
    alloc.set_autofilter("A47:%s49" % colname(gt47 - 1)); alloc.fix_dimension()
    parts[SHEET["alloc"]] = alloc.dump()
    report["nsw_grand_total_col"] = gtc

    # 4  other states accrual
    acc = Sheet(parts[SHEET["acc"]], ss)
    m = len(os_order); oj_last = 5 + m; osgt = colname(os_gt)
    acc.replace_tail(12, 6, [(c, "f", CC.format(r=colname(c), cc=16)) for c in range(6, oj_last + 1)], style_from=6)
    ow = {j: sum(ACCRUAL[st] * hrs[st][j] / sum(hrs[st].values()) for st in OTHER if j in hrs.get(st, {})) for j in os_order}
    oplug = round(sum(ACCRUAL.values()) - sum(round(v, 2) for v in ow.values()), 2)
    report["other_states_rounding_plug"] = oplug
    tail14 = [(5, "f", "ROUND(SUM(F14:%s14),0)" % colname(oj_last))]
    for c in range(6, oj_last + 1):
        f = "ROUND(%s15,2)" % colname(c)
        if c == 6 and oplug: f += "%+.2f" % oplug
        tail14.append((c, "f", f))
    acc.replace_tail(14, 5, tail14, style_from=6)
    acc.replace_tail(15, 6, [(c, "f", "SUM(%s17:%s22)" % (colname(c), colname(c))) for c in range(6, oj_last + 1)], style_from=6)
    acc.replace_tail(16, 6, [(c, "f", "'Other States PTAX Allocations'!%s$2" % colname(c - 4)) for c in range(6, oj_last + 1)], style_from=6)
    acc.replace_tail(17, 6, [(c, "n", 0) for c in range(6, oj_last + 1)], style_from=6)
    for k, st in enumerate(OTHER):
        r = 18 + k; pr = 3 + k
        acc.replace_tail(r, 6, [(c, "f", "$E%d*('Other States PTAX Allocations'!%s%d/'Other States PTAX Allocations'!$%s%d)"
                                 % (r, colname(c - 4), pr, osgt, pr)) for c in range(6, oj_last + 1)], style_from=6)
    acc.fix_dimension(); parts[SHEET["acc"]] = acc.dump()

    # 5  rename the timesheets sheet and repoint the pivot cache
    wbx = parts["xl/workbook.xml"].decode("utf8")
    wbx = wbx.replace('<sheet name="%s"' % OLD_TS, '<sheet name="%s"' % TS_SHEET)
    wbx = re.sub(r"'%s'!\$A\$2:\$S\$\d+" % OLD_TS, "'%s'!$A$2:$S$%d" % (TS_SHEET, last), wbx)
    wbx = wbx.replace("'%s'" % OLD_TS, "'%s'" % TS_SHEET)
    wbx = re.sub(r"'Other States = Accrual'!\$A\$15:\$[A-Z]+\$15", "'Other States = Accrual'!$A$15:$%s$15" % colname(oj_last), wbx)
    wbx = re.sub(r"'Other States PTAX Allocations'!\$A\$2:\$[A-Z]+\$2", "'Other States PTAX Allocations'!$A$2:$%s$2" % colname(os_gt - 1), wbx)
    wbx = re.sub(r"'PTax Allocations NSW'!\$A\$47:\$[A-Z]+\$49", "'PTax Allocations NSW'!$A$47:$%s$49" % colname(gt47 - 1), wbx)
    if "fullCalcOnLoad" not in wbx:
        wbx = re.sub(r'<calcPr([^>]*?)/>', r'<calcPr\1 fullCalcOnLoad="1"/>', wbx, count=1)
    parts["xl/workbook.xml"] = wbx.encode("utf8")
    pc = parts["xl/pivotCache/pivotCacheDefinition1.xml"].decode("utf8")
    pc = re.sub(r'<worksheetSource ref="A2:F\d+" sheet="%s"/>' % OLD_TS, '<worksheetSource ref="A2:F%d" sheet="%s"/>' % (last, TS_SHEET), pc)
    pc = re.sub(r'refreshOnLoad="[^"]*"', 'refreshOnLoad="1"', pc, count=1) if "refreshOnLoad=" in pc else pc.replace("<pivotCacheDefinition ", '<pivotCacheDefinition refreshOnLoad="1" ', 1)
    parts["xl/pivotCache/pivotCacheDefinition1.xml"] = pc.encode("utf8")
    for p in ("docProps/app.xml",):
        parts[p] = parts[p].decode("utf8").replace(OLD_TS, TS_SHEET).encode("utf8")
    # a stale calc chain after cells change kind is worse than none; Excel rebuilds it
    cc = "xl/calcChain.xml"
    if cc in parts:
        rels = parts["xl/_rels/workbook.xml.rels"].decode("utf8")
        parts["xl/_rels/workbook.xml.rels"] = re.sub(r'<Relationship Id="rId\d+"[^>]*Target="calcChain\.xml"[^>]*/>', "", rels).encode("utf8")
        parts["[Content_Types].xml"] = re.sub(r'<Override PartName="/xl/calcChain\.xml"[^>]*/>', "", parts["[Content_Types].xml"].decode("utf8")).encode("utf8")
        del parts[cc]; order_parts = [x for x in order_parts if x != cc]
    parts["xl/sharedStrings.xml"] = ss.dump()

    dst = os.path.join(OUT, "2026-10 Payroll Tax.xlsx")
    zo = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
    for name in order_parts: zo.writestr(name, parts[name])
    zo.close()

    report["nsw_journal_total"] = write_journal(os.path.join(OUT, "2026-10 Payroll Tax NSW Xero Journal.csv"),
        "Payroll Tax NSW %s %s" % (MONTH, YEAR), allocate(nsw_order, nsw, round(payable, 2)), "OFFICE / ADMIN")
    report["other_states_journal_total"] = write_journal(os.path.join(OUT, "2026-10 Payroll Tax Alloc Other States Xero Journal - ACCRUAL.csv"),
        "Payroll Tax ACT/QLD/SA/VIC/WA %s %s Accrual" % (MONTH, YEAR), allocate(os_order, ow, sum(ACCRUAL.values())), "OFFICE / ADMIN [CTS]")
    json.dump(report, open(os.path.join(OUT, "october-build.json"), "w"), indent=1)
    for k, v in report.items(): print("%-26s %s" % (k, v))
    return dst

if __name__ == "__main__":
    print(build())
