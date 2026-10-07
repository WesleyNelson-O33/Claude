"""September payroll tax fixes, written straight into the workbook XML.

Both workbooks carry pivot tables, so they are edited in place rather than
round-tripped through openpyxl, which would drop them.
"""
import os, re, shutil, zipfile

UP = "/root/.claude/uploads/a91f56e6-b7da-5dad-8e0b-6d5f516f9360/"
OUT = "/home/user/Claude/ptax/"

# Employee default location basis, the same basis the split file uses.
PTAX = {
    "xl/worksheets/sheet3.xml": {            # Main & NSW
        "D3": ("v", 160958.61),              # wages, was 160,129.12
        "D5": ("v", 21293.84),               # NSW super only, was all states 27,816.03
    },
    "xl/worksheets/sheet9.xml": {            # VIC
        "D3": ("v", 25982.64),               # was 26,418.15
        "D5": ("v", 3016.17),                # was 3,068.43
        "C14": ("f", "(C13/C13)*(1200000/$N$1*C1)"),   # threshold, was =SUM(C4:C13)
        "D14": ("f", "(D13/D13)*(1200000/$N$1*D1)"),
    },
    "xl/worksheets/sheet8.xml": {            # SA
        "D3": ("v", 6511.87),                # was 7,036.87, Kyle's 525 per diem out
        "C14": ("f", "(C13/C13)*(1200000/$N$1*C1)"),
        "D14": ("f", "(D13/D13)*(1200000/$N$1*D1)"),
    },
    "xl/worksheets/sheet10.xml": {           # WA
        "D3": ("v", 9608.58),                # was =9608.58-300
        "C14": ("f", "(C13/C13)*(1200000/$N$1*C1)"),
        "D14": ("f", "(D13/D13)*(1200000/$N$1*D1)"),
    },
    "xl/worksheets/sheet7.xml": {            # QLD, figures unchanged
        "C14": ("f", "(C13/C13)*(1200000/$N$1*C1)"),
        "D14": ("f", "(D13/D13)*(1200000/$N$1*D1)"),
    },
}


def set_cell(xml, ref, kind, val):
    """Replace one cell's contents, keeping its style."""
    m = re.search(r'<c r="%s"([^>/]*)(?:/>|>(.*?)</c>)' % ref, xml, re.S)
    assert m, "cell %s not found" % ref
    attrs = re.sub(r'\st="[^"]*"', "", m.group(1))          # drop any type, these are numeric
    body = ("<v>%s</v>" % val) if kind == "v" else ("<f>%s</f>" % val.replace("&", "&amp;"))
    return xml[:m.start()] + '<c r="%s"%s>%s</c>' % (ref, attrs, body) + xml[m.end():]


def rewrite(src, dst, edits=(), drop_rows=None, drop_part=None):
    z = zipfile.ZipFile(src)
    parts = {n: z.read(n) for n in z.namelist()}
    order = [i.filename for i in z.infolist()]
    z.close()

    for part, cells in edits:
        xml = parts[part].decode("utf-8")
        for ref, (kind, val) in cells.items():
            xml = set_cell(xml, ref, kind, val)
        parts[part] = xml.encode("utf-8")

    if drop_rows:
        part, lo, hi = drop_rows
        xml = parts[part].decode("utf-8")
        for r in range(lo, hi + 1):
            m = re.search(r'<row r="%d"[^>]*?(?:/>|>.*?</row>)' % r, xml, re.S)
            if m:
                xml = xml[:m.start()] + xml[m.end():]
        parts[part] = xml.encode("utf-8")

    # a stale calc cache after edits is worse than none; let Excel rebuild it
    cc = "xl/calcChain.xml"
    if cc in parts:
        rels = parts["xl/_rels/workbook.xml.rels"].decode("utf-8")
        rid = re.search(r'<Relationship Id="rId\d+"[^>]*Target="calcChain\.xml"[^>]*/>', rels)
        if rid:
            parts["xl/_rels/workbook.xml.rels"] = rels.replace(rid.group(0), "").encode("utf-8")
        ct = parts["[Content_Types].xml"].decode("utf-8")
        parts["[Content_Types].xml"] = re.sub(
            r'<Override PartName="/xl/calcChain\.xml"[^>]*/>', "", ct).encode("utf-8")
        del parts[cc]
        order = [n for n in order if n != cc]

    wbx = parts["xl/workbook.xml"].decode("utf-8")
    if "fullCalcOnLoad" not in wbx:
        wbx = re.sub(r'<calcPr([^>]*?)/>', r'<calcPr\1 fullCalcOnLoad="1"/>', wbx, count=1)
        if "fullCalcOnLoad" not in wbx:
            wbx = wbx.replace("</workbook>", '<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>')
    parts["xl/workbook.xml"] = wbx.encode("utf-8")

    zo = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
    for n in order:
        zo.writestr(n, parts[n])
    zo.close()
    print("wrote", dst, len(order), "parts")


os.makedirs(OUT, exist_ok=True)
rewrite(UP + "82522077-2026-09_Payroll_Tax.xlsx",
        OUT + "2026-09 Payroll Tax.xlsx",
        edits=list(PTAX.items()))
rewrite(UP + "ae7509ab-EH_Payroll_Split_File_FY2027.xlsx",
        OUT + "EH Payroll Split File FY2027.xlsx",
        drop_rows=("xl/worksheets/sheet8.xml", 15, 24))
