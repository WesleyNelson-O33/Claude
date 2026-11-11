"""Roll the prepayments schedule to October 2026 and write the Xero journal.

Edits the sheet XML directly: the workbook carries drawings and comments that
openpyxl would drop. Method is the one in the September file: every row has a
SEQUENCE spill across its months, the Released column sums to the current
month, and the journal is one credit to 11510 and one debit per line.
"""
import csv, datetime, os, re, zipfile
from lxml import etree
import openpyxl
from openpyxl.utils import column_index_from_string as CI, get_column_letter as CL

SRC = "/root/.claude/uploads/fc902e4e-1220-55ba-999f-a6d0cad6db14/0516ff6b-2026-09_Prepayments.xlsx"
OUT = "/home/user/Claude/prepay/out"
PREV, CUR = "BG", "BH"            # September and October month columns
CUR_DATE_SERIAL = 46296           # 1/10/2026
NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
def q(t): return "{%s}%s" % (NS, t)
def colname(i): return CL(i)
def colnum(s): return CI(s)
exec(open("/tmp/claude-0/-home-user-Claude/fc902e4e-1220-55ba-999f-a6d0cad6db14/scratchpad/xmlmachinery.py").read())

JOBNAME = {"9105": "ONSITE (including PROD backup)", "9200": "INTEGRATION [INT]", "9400": "CONSULTING [CONS]",
           "9000": "OFFICE / ADMIN [CTS]", "9300": "HUMAN RESOURCES [CTS]",
           "585": "CBA Team Leader [ONS]", "587": "CBA Tech TN [ONS]", "1350": "CBA Brisbane ONS [ONS]"}
CC = {"9105": "Onsite", "9200": "Integration", "9400": "Consulting", "9000": "CTS", "9300": "CTS",
      "585": "Onsite", "587": "Onsite", "1350": "Onsite"}
CCF = ('IF(ISNUMBER(SEARCH("PRD",D%(r)d)),"Production",IF(ISNUMBER(SEARCH("CONS",D%(r)d)),"Consulting",'
       'IF(ISNUMBER(SEARCH("ONS",D%(r)d)),"Onsite",IF(ISNUMBER(SEARCH("INT",D%(r)d)),"Integration",'
       'IF(ISNUMBER(SEARCH("VID",D%(r)d)),"Video",IF(ISNUMBER(SEARCH("CTS",D%(r)d)),"CTS",""))))))')
GLNAME = {61400: "Dues & Subscriptions", 62300: "Other Insurance"}

# New October prepayments, from the Xero bills. (gl, job, inv no, date, supplier, details, cost ex GST, months, first month column)
NEW = [
    (61400, "9105", "50BB48DA0013", datetime.date(2026, 10, 6), "Qwilr Pty Ltd", "Starter plan 6/10/26 to 6/10/27, USD 1,260", 909.09, 12, "BH"),
    (61400, "9200", "50BB48DA0013", datetime.date(2026, 10, 6), "Qwilr Pty Ltd", "Starter plan 6/10/26 to 6/10/27, USD 1,260", 454.55, 12, "BH"),
    (61400, "9400", "50BB48DA0013", datetime.date(2026, 10, 6), "Qwilr Pty Ltd", "Starter plan 6/10/26 to 6/10/27, USD 1,260", 454.54, 12, "BH"),
    (61400, "9000", "INV137607", datetime.date(2026, 10, 1), "The CEO Institute (DL)", "Syndicate 41 membership Oct to Dec 2026", 2100.00, 3, "BH"),
    (61400, "9300", "INV51110778741", datetime.date(2026, 10, 1), "Linkedin Singapore", "Job Slots 3 of 3, 15/08/2026 to 14/08/2027. Oct release includes Aug and Sep catch up", 6360.00, 12, "BF"),
    (62300, "585", "INV02434590", datetime.date(2026, 10, 23), "Avetta", "Advantage membership 23/10/26 to 22/10/27", 297.34, 12, "BI"),
    (62300, "587", "INV02434590", datetime.date(2026, 10, 23), "Avetta", "Advantage membership 23/10/26 to 22/10/27", 297.33, 12, "BI"),
    (62300, "1350", "INV02434590", datetime.date(2026, 10, 23), "Avetta", "Advantage membership 23/10/26 to 22/10/27", 297.33, 12, "BI"),
]

def serial(d): return (d - datetime.date(1899, 12, 30)).days

def existing_lines():
    """October lines already on the schedule, off the cached values."""
    wb = openpyxl.load_workbook(SRC, data_only=True); ws = wb["Prepayment"]
    cur = colnum(CUR); out = []
    for r in range(8, ws.max_row + 1):
        v = ws.cell(r, cur).value
        if isinstance(v, (int, float)) and abs(v) > 0.004:
            out.append({"row": r, "amount": round(float(v), 2), "desc": ws.cell(r, 80).value, "job": ws.cell(r, 5).value,
                        "gl": ws.cell(r, 1).value, "jobname": ws.cell(r, 4).value, "cc": ws.cell(r, 6).value, "supplier": ws.cell(r, 9).value})
    total_cached = ws.cell(6, cur).value
    wb.close(); return out, total_cached

def build():
    os.makedirs(OUT, exist_ok=True)
    lines, cached_total = existing_lines()
    assert abs(sum(l["amount"] for l in lines) - cached_total) < 0.05, (sum(l["amount"] for l in lines), cached_total)
    z = zipfile.ZipFile(SRC); parts = {n: z.read(n) for n in z.namelist()}; order = [i.filename for i in z.infolist()]; z.close()
    wbx = parts["xl/workbook.xml"].decode("utf8"); rels = parts["xl/_rels/workbook.xml.rels"].decode("utf8")
    rid = {m.group(1): m.group(2) for m in re.finditer(r'Id="(rId\d+)"[^>]*Target="(worksheets/sheet\d+\.xml)"', rels)}
    rid.update({m.group(2): m.group(1) for m in re.finditer(r'Target="(worksheets/sheet\d+\.xml)"[^>]*Id="(rId\d+)"', rels)})
    sheets = {m.group(1): "xl/" + rid[m.group(3)] for m in re.finditer(r'<sheet name="([^"]+)" sheetId="(\d+)"[^>]*r:id="(rId\d+)"', wbx)}
    ss = Strings(parts["xl/sharedStrings.xml"])
    ws = Sheet(parts[sheets["Prepayment"]], ss)

    # 1  the month
    c1 = Sheet.cells(ws.rowel(1))[3]; c1.attrib.pop("t", None)
    for ch in list(c1): c1.remove(ch)
    etree.SubElement(c1, q("v")).text = str(CUR_DATE_SERIAL)

    # 2  Released column: sum to October on every row that has one
    bx = colnum("BX"); rolled = 0
    for row in ws.sd.findall(q("row")):
        c = Sheet.cells(row).get(bx)
        if c is None: continue
        f = c.find(q("f"))
        if f is not None and f.text and re.search(r":%s\d+\)" % PREV, f.text):
            f.text = re.sub(r":%s(\d+)\)" % PREV, r":%s\1)" % CUR, f.text); rolled += 1
            v = c.find(q("v"))
            if v is not None: c.remove(v)

    # 3  new rows, styled on row 219
    r0 = 220; new_lines = []
    for i, (gl, job, inv, d, sup, det, cost, months, first) in enumerate(NEW):
        r = r0 + i; fc = colnum(first); lc = fc + months - 1
        spec = [(1, "n", gl), (2, "s", GLNAME[gl]), (3, "n", int(job)), (4, "s", JOBNAME[job]), (5, "f", 'C%d&" - "&D%d' % (r, r)),
                (6, "f", CCF % dict(r=r)), (7, "s", inv), (8, "n", serial(d)), (9, "s", sup), (10, "s", det), (11, "n", cost), (12, "n", months),
                (13, "f", "ROUND(K%d/L%d,2)" % (r, r)),
                (bx, "f", "SUM(%s%d:%s%d)" % (first if colnum(first) <= colnum(CUR) else CUR, r, CUR, r)), (bx + 1, "f", "K%d" % r), (bx + 2, "f", "BY%d-BX%d" % (r, r)),
                (bx + 3, "f", "BY%d-K%d" % (r, r)), (bx + 4, "f", '"Inv"&G%d&"|"&TEXT(H%d,"dd/mm")&"|"&I%d&"|"&J%d' % (r, r, r, r))]
        row = ws.set_row(r, spec, style_from=None)
        # copy styles from row 219 column by column, then add the SEQUENCE spill
        s219 = Sheet.cells(ws.rowel(219))
        for col, c in Sheet.cells(row).items():
            src = s219.get(col)
            if src is not None and src.get("s"): c.set("s", src.get("s"))
        c = etree.SubElement(row, q("c")); c.set("r", "%s%d" % (first, r))
        if s219.get(colnum("BG")) is not None and s219[colnum("BG")].get("s"): c.set("s", s219[colnum("BG")].get("s"))
        f = etree.SubElement(c, q("f")); f.set("t", "array"); f.set("ref", "%s%d:%s%d" % (first, r, colname(lc), r))
        f.text = "_xlfn.SEQUENCE(1,L%d,M%d,0)" % (r, r)
        cells = sorted(Sheet.cells(row).items())
        for _, x in cells: row.remove(x)
        for _, x in cells: row.append(x)
        row.set("spans", "1:%d" % max(col for col, _ in cells))
        if colnum(first) <= colnum(CUR):
            n_months = colnum(CUR) - colnum(first) + 1
            new_lines.append({"row": r, "amount": round(round(cost / months, 2) * n_months, 2), "desc": "Inv%s|%s|%s|%s" % (inv, d.strftime("%d/%m"), sup, det),
                              "job": "%s - %s" % (job, JOBNAME[job]), "gl": gl, "jobname": JOBNAME[job], "cc": CC[job], "supplier": sup})
    ws.fix_dimension()
    parts[sheets["Prepayment"]] = ws.dump()

    # 4  the journal dump sheet, in her layout: A amount, B description, C job, D account, E job name, H cost centre
    all_lines = lines + new_lines
    js = Sheet(parts[sheets["2025-07 Prepayments Journal Xer"]], ss)
    js.retire_rows(0)
    for i, l in enumerate(all_lines, start=1):
        job = str(l["job"]).replace("9000 - CTS", "9000 - OFFICE / ADMIN [CTS]")
        js.set_row(i, [(1, "n", l["amount"]), (2, "s", l["desc"] or ""), (3, "s", job), (4, "n", int(l["gl"])),
                       (5, "s", job.split(" - ", 1)[1] if " - " in job else ""), (8, "s", l["cc"] or "")])
    js.fix_dimension(); parts[sheets["2025-07 Prepayments Journal Xer"]] = js.dump()

    if "fullCalcOnLoad" not in wbx:
        wbx = re.sub(r'<calcPr([^>]*?)/>', r'<calcPr\1 fullCalcOnLoad="1"/>', wbx, count=1)
    parts["xl/workbook.xml"] = wbx.encode("utf8"); parts["xl/sharedStrings.xml"] = ss.dump()
    dst = os.path.join(OUT, "2026-10 Prepayments.xlsx")
    zo = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
    for n in order: zo.writestr(n, parts[n])
    zo.close()

    # 5  the Xero journal
    H = ["*Narration", "*Date", "Description", "*AccountCode", "*TaxRate", "*Amount", "TrackingName1", "TrackingOption1", "TrackingName2", "TrackingOption2"]
    narr, date = "Prepayments October 2026", "31/10/2026"
    rows = []
    for l in all_lines:
        job = str(l["job"]).replace("9000 - CTS", "9000 - OFFICE / ADMIN [CTS]")
        rows.append([narr, date, l["desc"], 11510, "BAS Excluded", "%.2f" % -l["amount"], "Cost Centres", l["cc"], "Job Numbers", job])
    for l in all_lines:
        job = str(l["job"]).replace("9000 - CTS", "9000 - OFFICE / ADMIN [CTS]")
        rows.append([narr, date, l["desc"], int(l["gl"]), "BAS Excluded", "%.2f" % l["amount"], "Cost Centres", l["cc"], "Job Numbers", job])
    with open(os.path.join(OUT, "2026-10_Prepayments_Journal_Xero.csv"), "w", newline="", encoding="cp1252", errors="replace") as fh:
        w = csv.writer(fh); w.writerow(H); w.writerows(rows)
    total = round(sum(l["amount"] for l in all_lines), 2)
    print("rows rolled:", rolled, "| existing lines:", len(lines), round(sum(l["amount"] for l in lines), 2),
          "| new lines:", len(new_lines), round(sum(l["amount"] for l in new_lines), 2), "| journal total:", total)
    return dst

if __name__ == "__main__":
    print(build())
