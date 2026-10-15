"""Clear the 12,249 Month-End revenue variance on Danica's FY27 tracker (her _2 copy).

Adds the two September department rows that were missing and types their Xero invoices on Finance:
  Onsite ONS-0073: Aware Super Contract September 2026, job 26022301, INV-10652, 11,000.
  Production PRD-0102: Talking Philanthropy - Event Videography, job 26091002, VIDEO, INV-10669, 1,249.
Also sets Onsite ONS-0051 (CN-10666) expected to -1186.15 so the credit note agrees.
Same XML editing as fill_fy27_sept_wip.py so tables, comments and formulas survive."""
import re, sys, zipfile
import openpyxl
exec(open(__file__.replace("fix_tracker_sept_v3.py", "fill_fy27_sept_wip.py")).read().split("z = zipfile.ZipFile(SRC)")[0]
     .replace("SRC, WIPSCHED, OUT = sys.argv[1], sys.argv[2], sys.argv[3]", "SRC, OUT = sys.argv[1], sys.argv[2]"))
PATHS = {"Finance": "xl/worksheets/sheet2.xml", "Onsite": "xl/worksheets/sheet8.xml",
         "Production": "xl/worksheets/sheet10.xml"}
EVENT = 46283  # 18/09/2026

z = zipfile.ZipFile(SRC)
sh = {k: Sheet(z.read(v)) for k, v in PATHS.items()}
ro = openpyxl.load_workbook(SRC, read_only=True, data_only=True)
def val(sheet, ref): return ro[sheet][ref].value
fin_ids = {r[0]: i + 7 for i, r in enumerate(ro["Finance"].iter_rows(min_row=7, max_row=4506, max_col=1, values_only=True))}

# the rows being filled are empty and carry the expected Row IDs
assert val("Onsite", "X77") == "ONS-0073" and val("Onsite", "C77") is None and val("Onsite", "G77") is None
assert val("Production", "AU106") == "PRD-0102" and val("Production", "C106") is None
assert val("Onsite", "X55") == "ONS-0051" and val("Onsite", "O55") == 1186.15
fo, fp = fin_ids["ONS-0073"], fin_ids["PRD-0102"]
assert all(val("Finance", "%s%d" % (c, r)) is None for c in "GHI" for r in (fo, fp))

on = sh["Onsite"]
for c, v in (("C", "26022301"), ("F", "Aware Super"), ("G", "Aware Super Contract September 2026"),
             ("H", "ONSITE"), ("I", "Contract"), ("J", "GST 10%"), ("O", 11000)):
    on.set("%s77" % c, v)
on.set("O55", -1186.15)

pr = sh["Production"]
for c, v in (("C", "26091002"), ("F", "Global Philanthropic"), ("G", "Talking Philanthropy - Event Videography"),
             ("H", "VIDEO"), ("I", "Project"), ("J", "GST 10%"), ("O", 1249), ("V", EVENT)):
    pr.set("%s106" % c, v)

fin = sh["Finance"]
for r, inv, amt in ((fo, "INV-10652", 11000), (fp, "INV-10669", 1249)):
    fin.set("G%d" % r, inv); fin.set("H%d" % r, SEP); fin.set("I%d" % r, amt)

wbxml = z.read("xl/workbook.xml").decode()
if "fullCalcOnLoad" not in wbxml:
    wbxml = re.sub(r"<calcPr([^>]*)/>", r'<calcPr\1 fullCalcOnLoad="1"/>', wbxml)
with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as o:
    for item in z.infolist():
        name = item.filename
        if name in PATHS.values():
            data = sh[[k for k, v in PATHS.items() if v == name][0]].bytes()
        elif name == "xl/workbook.xml":
            data = wbxml.encode()
        else:
            data = z.read(name)
        o.writestr(item, data)
print("written", OUT, "Finance rows", fo, fp)
