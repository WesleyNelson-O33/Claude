"""Add two tabs to the FY27 budget so it can feed the Controller Pack.

Pack Map        the mapping engine: one row per budget line, a split basis,
                a department weight and a pack subcategory.
Pack Budget FY27  the same 227 row skeleton as the pack's Budget FY27 sheet,
                filled by formula. Copy B7:M232 and paste values into the pack.

Written straight into the workbook's XML so the pivot tables, charts, threaded
comments and external links in the budget are left untouched.
"""
import json, re, zipfile
from lxml import etree

SRC = ("/root/.claude/uploads/a91f56e6-b7da-5dad-8e0b-6d5f516f9360/"
       "a1424a99-CTS_Budget_FY27_Final.xlsx")
OUT = "/home/user/Claude/pack/CTS Budget FY27 Final - 3 Way fixed.xlsx"
SKEL = ("/tmp/claude-0/-home-user-Claude/a91f56e6-b7da-5dad-8e0b-6d5f516f9360/"
        "scratchpad/skeleton.json")

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
Q = lambda t: "{%s}%s" % (NS, t)
DETAIL = "'FY27 Budget - detailed'"

DEPTS = ["ONSITE", "PRODUCTION", "VIDEO", "INTEGRATION", "CONSULTING", "ADMIN"]
WCOL = dict(zip(DEPTS, "TUVWXY"))          # weight column on Pack Map
MAP0, MAPN = 21, 79                        # first and last mapping row

# ---------------------------------------------------------------- source lines
GA = "General & Administrative Exp"
OPEX = "Operating Expenses"
EMP = "Employment Expenses"
COS = "Cost of Sales"
# (budget row, department rule, pack category, pack subcategory, rows to net off)
LINES = [
    (3, "Consulting", "Income", "Non Contract Income", None),
    (4, "Production", "Income", "Non Contract Income", None),
    (5, "Onsite", "Income", "Contract Income", None),
    (8, "Consulting", COS, "Direct Salaries", None),
    (9, "Consulting", COS, "Equipment - Purchase", None),
    (10, "Consulting", COS, "Other Direct Expenses", None),
    (7, "Consulting", COS, "Direct Superannuation", (8, 9, 10)),
    (12, "Production", COS, "Direct Salaries", None),
    (13, "Production", COS, "Equipment - Hires", None),
    (14, "Production", COS, "Other Direct Expenses", None),
    (11, "Production", COS, "Direct Superannuation", (12, 13, 14)),
    (16, "Onsite", COS, "Direct Salaries", None),
    (17, "Onsite", COS, "Subscriptions & Licences", None),
    (18, "Onsite", COS, "Other Direct Expenses", None),
    (15, "Onsite", COS, "Direct Superannuation", (16, 17, 18)),
]
for r in list(range(23, 37)) + [38, 39, 41] + list(range(43, 56)) + list(range(57, 68)):
    sub = (GA if r <= 36 else "Business Insurances" if r in (38, 39)
           else "Advertising & Promotion Exp" if r == 41
           else OPEX if r <= 55 else EMP)
    if r == 32:
        sub = "Credit Card fees paid"
    LINES.append((r, "Split", "Expenses", sub, None))
# the three departmental overhead blocks, already department specific in the budget
LINES += [(70, "Consulting", "Expenses", EMP, None),
          (71, "Onsite", "Expenses", EMP, None),
          (72, "Production", "Expenses", EMP, None)]
assert len(LINES) == MAPN - MAP0 + 1, len(LINES)

MONTHS = ["Jul-26", "Aug-26", "Sep-26", "Oct-26", "Nov-26", "Dec-26",
          "Jan-27", "Feb-27", "Mar-27", "Apr-27", "May-27", "Jun-27"]


# ------------------------------------------------------------------- xml helpers
def esc(s):
    return (str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def col(n):                                   # 1 -> A
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


class Sheet:
    def __init__(self):
        self.rows = {}

    def _c(self, r, c, s):
        return self.rows.setdefault(r, {}).setdefault(c, {"s": s})

    def text(self, r, c, v, s=0):
        self.rows.setdefault(r, {})[c] = {"s": s, "t": esc(v)}

    def num(self, r, c, v, s=0):
        self.rows.setdefault(r, {})[c] = {"s": s, "n": v}

    def f(self, r, c, v, s=0):
        self.rows.setdefault(r, {})[c] = {"s": s, "f": esc(v.lstrip("="))}

    def blank(self, r, c, s=0):
        self.rows.setdefault(r, {})[c] = {"s": s}

    def xml(self, cols, freeze=None):
        maxr = max(self.rows) if self.rows else 1
        maxc = max((max(cs) for cs in self.rows.values()), default=1)
        out = [('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<worksheet xmlns="%s" xmlns:r="%s">' % (NS, R)),
               '<dimension ref="A1:%s%d"/>' % (col(maxc), maxr)]
        pane = ('<pane ySplit="%d" topLeftCell="A%d" activePane="bottomLeft" state="frozen"/>'
                '<selection pane="bottomLeft"/>' % (freeze, freeze + 1)) if freeze else ""
        out.append('<sheetViews><sheetView showGridLines="0" workbookViewId="0">'
                   + pane + '</sheetView></sheetViews>')
        out.append('<sheetFormatPr defaultRowHeight="15"/>')
        if cols:
            out.append("<cols>" + "".join(
                '<col min="%d" max="%d" width="%s" customWidth="1"/>' % (a, b, w)
                for a, b, w in cols) + "</cols>")
        out.append("<sheetData>")
        for r in sorted(self.rows):
            out.append('<row r="%d">' % r)
            for c in sorted(self.rows[r]):
                d = self.rows[r][c]
                ref = '<c r="%s%d" s="%d"' % (col(c), r, d["s"])
                if "t" in d:
                    out.append(ref + ' t="inlineStr"><is><t xml:space="preserve">'
                               + d["t"] + "</t></is></c>")
                elif "f" in d:
                    out.append(ref + "><f>" + d["f"] + "</f></c>")
                elif "n" in d:
                    out.append(ref + "><v>" + repr(d["n"]) + "</v></c>")
                else:
                    out.append(ref + "/>")
            out.append("</row>")
        out.append("</sheetData></worksheet>")
        return "".join(out).encode("utf-8")


# ------------------------------------------------------------------- new styles
def add_styles(root):
    """Append the fonts, fills, formats and cell styles the two tabs need."""
    fonts, fills, numfmts, xfs = (root.find(Q(n)) for n in
                                  ("fonts", "fills", "numFmts", "cellXfs"))
    if numfmts is None:
        numfmts = etree.Element(Q("numFmts"))
        root.insert(list(root).index(fonts), numfmts)

    def font(**kw):
        e = etree.SubElement(fonts, Q("font"))
        if kw.get("b"):
            etree.SubElement(e, Q("b"))
        if kw.get("i"):
            etree.SubElement(e, Q("i"))
        etree.SubElement(e, Q("sz")).set("val", str(kw.get("sz", 10)))
        etree.SubElement(e, Q("color")).set("rgb", "FF" + kw.get("c", "000000"))
        etree.SubElement(e, Q("name")).set("val", "Arial")
        fonts.set("count", str(len(fonts)))
        return len(fonts) - 1

    def fill(rgb):
        e = etree.SubElement(fills, Q("fill"))
        p = etree.SubElement(e, Q("patternFill"))
        p.set("patternType", "solid")
        etree.SubElement(p, Q("fgColor")).set("rgb", "FF" + rgb)
        etree.SubElement(p, Q("bgColor")).set("indexed", "64")
        fills.set("count", str(len(fills)))
        return len(fills) - 1

    used = {int(e.get("numFmtId")) for e in numfmts}

    def numfmt(code, fid):
        while fid in used:
            fid += 1
        used.add(fid)
        e = etree.SubElement(numfmts, Q("numFmt"))
        e.set("numFmtId", str(fid))
        e.set("formatCode", code)
        numfmts.set("count", str(len(numfmts)))
        return fid

    def xf(fo=0, fi=0, nf=0, center=False, indent=0):
        e = etree.SubElement(xfs, Q("xf"))
        for k, v in (("numFmtId", nf), ("fontId", fo), ("fillId", fi), ("borderId", 0)):
            e.set(k, str(v))
        e.set("xfId", "0")
        e.set("applyFont", "1")
        e.set("applyFill", "1")
        e.set("applyNumberFormat", "1")
        if center or indent:
            e.set("applyAlignment", "1")
            a = etree.SubElement(e, Q("alignment"))
            if center:
                a.set("horizontal", "center")
            if indent:
                a.set("indent", str(indent))
        xfs.set("count", str(len(xfs)))
        return len(xfs) - 1

    MONEY = numfmt('$#,##0;($#,##0);"-"', 300)
    PCT = numfmt("0.0%", 301)
    MON = numfmt("mmm-yy", 302)

    f_body = font()
    f_bold = font(b=True)
    f_white = font(b=True, c="FFFFFF")
    f_title = font(b=True, sz=14, c="1F3864")
    f_note = font(i=True, sz=9, c="595959")
    f_blue = font(c="0000FF")
    fl_navy, fl_grey, fl_yellow = fill("1F3864"), fill("F2F2F2"), fill("FFFF00")

    return {
        "title": xf(f_title),
        "note": xf(f_note),
        "head": xf(f_white, fl_navy, center=True),
        "headmon": xf(f_white, fl_navy, MON, center=True),
        "dept": xf(f_white, fl_navy),
        "deptm": xf(f_white, fl_navy, MONEY),
        "cat": xf(f_bold, fl_grey, indent=1),
        "catm": xf(f_bold, fl_grey, MONEY),
        "sub": xf(f_body, indent=2),
        "subm": xf(f_body, 0, MONEY),
        "body": xf(f_body),
        "bold": xf(f_bold),
        "money": xf(f_body, 0, MONEY),
        "pct": xf(f_body, 0, PCT),
        "inpct": xf(f_blue, fl_yellow, PCT),
        "input": xf(f_blue, fl_yellow),
        "white": xf(f_body, 0),
    }


# -------------------------------------------------------------------- Pack Map
def build_map(S):
    ws = Sheet()
    ws.text(1, 1, "Pack Map  -  how each budget line reaches the Controller Pack", S["title"])
    ws.text(2, 1, "Yellow is yours. Everything else is a formula off the FY27 Budget - "
                  "detailed tab, so it moves when the budget moves.", S["note"])

    # the three split bases, read live off her matrix
    ws.text(4, 1, "Split bases", S["bold"])
    for i, g in enumerate(("PRD/VID", "ONS", "CONS/INT")):
        ws.text(4, 2 + i, g, S["head"])
    for i, (name, src) in enumerate((("3 Way", 17), ("Staff", 19), ("Office Dept", 21))):
        r = 5 + i
        ws.text(r, 1, name, S["body"])
        for j in range(3):
            ws.f(r, 2 + j, "=%s!%s%d" % (DETAIL, "ABC"[j], src), S["pct"])

    # the two sub splits she owns
    ws.text(9, 9, "Production share of PRD/VID", S["body"])
    ws.num(9, 10, 0.79, S["inpct"])
    ws.text(10, 9, "Integration share of CONS/INT", S["body"])
    ws.num(10, 10, 0.72, S["inpct"])
    ws.text(11, 9, "Video and Consulting take the rest. Defaults are your own FY23 "
                   "actuals on the PRD and CONS tabs.", S["note"])

    # rule table: a weight per pack department
    ws.text(9, 1, "Rule", S["head"])
    for i, d in enumerate(DEPTS):
        ws.text(9, 2 + i, d, S["head"])
    rules = [("Split-3 Way", 5), ("Split-Staff", 6), ("Split-Office Dept", 7),
             ("Onsite", None), ("Production", None), ("Consulting", None)]
    for i, (name, src) in enumerate(rules):
        r = 10 + i
        ws.text(r, 1, name, S["body"])
        if src:
            cells = ["=$C$%d" % src, "=$B$%d*$J$9" % src, "=$B$%d*(1-$J$9)" % src,
                     "=$D$%d*$J$10" % src, "=$D$%d*(1-$J$10)" % src, "=0"]
        elif name == "Onsite":
            cells = ["=1", "=0", "=0", "=0", "=0", "=0"]
        elif name == "Production":
            cells = ["=0", "=$J$9", "=1-$J$9", "=0", "=0", "=0"]
        else:
            cells = ["=0", "=0", "=0", "=$J$10", "=1-$J$10", "=0"]
        for j, f in enumerate(cells):
            ws.f(r, 2 + j, f, S["pct"])

    # the mapping itself
    heads = (["Budget row", "Line", "Basis", "Rule", "Category", "Subcategory", "Sign"]
             + MONTHS + [d + " wt" for d in DEPTS])
    for i, h in enumerate(heads):
        ws.text(MAP0 - 1, 1 + i, h, S["head"])
    for i, (src, rule, cat, sub, net) in enumerate(LINES):
        r = MAP0 + i
        ws.num(r, 1, src, S["body"])
        ws.f(r, 2, '=%s!G%d%s' % (DETAIL, src, '&" - the rest"' if net else ""), S["body"])
        ws.f(r, 3, '=%s!E%d&""' % (DETAIL, src), S["body"])
        ws.text(r, 4, rule, S["input"])
        ws.text(r, 5, cat, S["input"])
        ws.text(r, 6, sub, S["input"])
        ws.f(r, 7, '=IF(OR($E%d="Income",$E%d="Other Income"),1,-1)' % (r, r), S["body"])
        for j in range(12):
            mc = col(8 + j)
            term = "%s!%s%d" % (DETAIL, mc, src)
            if net:
                term = "(%s-%s)" % (term, "-".join("%s!%s%d" % (DETAIL, mc, x) for x in net))
            ws.f(r, 8 + j, "=$G%d*%s" % (r, term), S["money"])
        key = 'IF($D%d="Split","Split-"&$C%d,$D%d)' % (r, r, r)
        for j in range(6):
            ws.f(r, 20 + j,
                 "=IFERROR(INDEX(%s$10:%s$15,MATCH(%s,$A$10:$A$15,0)),0)"
                 % (col(2 + j), col(2 + j), key), S["pct"])
    cols = [(1, 1, 11), (2, 2, 38), (3, 3, 13), (4, 4, 12), (5, 5, 15), (6, 6, 30),
            (7, 7, 6), (8, 19, 12), (20, 25, 11)]
    return ws.xml(cols, freeze=MAP0 - 1)


# ------------------------------------------------------- Pack Budget FY27 sheet
def build_budget(S):
    skel = json.load(open(SKEL))              # [kind, label, row] from the pack
    ws = Sheet()
    ws.text(1, 1, "Budget FY27 for the Controller Pack", S["title"])
    ws.text(2, 1, "Calculated. This is the pack's Budget FY27 sheet, filled from the "
                  "budget. Costs are negative, the way the pack reads them.", S["note"])
    ws.text(3, 1, "To load it: copy B7 to M232, then Paste Special Values into cell B7 "
                  "of Budget FY27 in the Controller Pack.", S["note"])
    ws.text(5, 1, "Line", S["head"])
    from datetime import date
    epoch = date(1899, 12, 30)
    for j, m in enumerate(MONTHS):
        d = date(2026 + (j >= 6), (7 + j - 1) % 12 + 1, 1)
        ws.num(5, 2 + j, (d - epoch).days, S["headmon"])
    ws.text(5, 14, "FY total", S["head"])

    dept = None
    order = []
    for kind, label, r in skel:
        if kind == "dept":
            dept = label
            ws.text(r, 1, label, S["dept"])
            for c in range(2, 15):
                ws.blank(r, c, S["deptm"])
            order.append(("dept", r, []))
        elif kind == "cat":
            ws.text(r, 1, label, S["cat"])
            for c in range(2, 15):
                ws.blank(r, c, S["catm"])
            for k in range(len(order) - 1, -1, -1):
                if order[k][0] == "dept":
                    order[k][2].append(r)
                    break
            order.append(("cat", r, []))
        elif kind == "sub":
            ws.text(r, 1, label, S["sub"])
            ws.text(r, 16, dept, S["body"])
            ws.text(r, 17, label, S["body"])
            w = WCOL[dept]
            for j in range(12):
                ws.f(r, 2 + j,
                     "=SUMPRODUCT(('Pack Map'!$F$%d:$F$%d=$Q%d)*('Pack Map'!$%s$%d:$%s$%d)"
                     "*('Pack Map'!%s$%d:%s$%d))"
                     % (MAP0, MAPN, r, w, MAP0, w, MAPN,
                        col(8 + j), MAP0, col(8 + j), MAPN), S["subm"])
            ws.f(r, 14, "=SUM(B%d:M%d)" % (r, r), S["subm"])
            for k in range(len(order) - 1, -1, -1):
                if order[k][0] == "cat":
                    order[k][2].append(r)
                    break
        else:
            order.append(("blank", r, []))

    for kind, r, kids in order:                # roll subs into cats, cats into depts
        if kind in ("cat", "dept") and kids:
            for j in range(13):
                c = col(2 + j)
                ws.f(r, 2 + j, "=" + "+".join("%s%d" % (c, k) for k in kids),
                     S["catm"] if kind == "cat" else S["deptm"])
    ws.text(5, 16, "Dept", S["head"])
    ws.text(5, 17, "Subcategory", S["head"])
    ws.text(4, 16, "Helper columns, ignore them", S["note"])
    cols = [(1, 1, 38), (2, 13, 12), (14, 14, 13), (15, 15, 3), (16, 16, 14), (17, 17, 30)]
    return ws.xml(cols, freeze=5)


# ---------------------------------------------------- the 3 Way fix and the orange
TOTALS = {26: 500.0, 28: 0.0, 49: 200.0, 50: 275.0, 53: 12430.0}
THREE = {"A": 0.34, "B": 0.33, "C": 0.33}


def three_way(xml, styles_root):
    """Point five overhead lines at the 3 Way row, then band them orange."""
    for r, total in TOTALS.items():
        for c in "ABC":
            m = re.search(r'<c r="%s%d"[^>]*>.*?</c>' % (c, r), xml, re.S)
            new = re.sub(r"<f>([^<]*)</f>",
                         lambda x: "<f>" + re.sub(r"\b([ABC])(19|21)\b", r"\g<1>17",
                                                 x.group(1)) + "</f>", m.group(0))
            new = re.sub(r"<v>[^<]*</v>", "<v>%s</v>" % round(total * THREE[c], 10), new)
            assert new != m.group(0), (c, r)
            xml = xml[:m.start()] + new + xml[m.end():]
        me = re.search(r'<c r="E%d"[^>]*t="s"[^>]*>.*?</c>' % r, xml, re.S)
        xml = xml[:me.start()] + re.sub(r"<v>\d+</v>", "<v>42</v>", me.group(0)) + xml[me.end():]
        mf = re.search(r'<c r="F%d"[^>]*>.*?</c>' % r, xml, re.S)
        if mf:
            xml = xml[:mf.start()] + xml[mf.end():]

    fills = styles_root.find(Q("fills"))
    e = etree.SubElement(fills, Q("fill"))
    pf = etree.SubElement(e, Q("patternFill"))
    pf.set("patternType", "solid")
    etree.SubElement(pf, Q("fgColor")).set("rgb", "FFFFA500")
    etree.SubElement(pf, Q("bgColor")).set("indexed", "64")
    fills.set("count", str(len(fills)))
    orange = len(fills) - 1

    xfs = styles_root.find(Q("cellXfs"))
    items = list(xfs)
    newxf = {}
    for sid in (46, 44, 202, 203, 699):
        clone = etree.fromstring(etree.tostring(items[sid]))
        clone.set("fillId", str(orange))
        clone.set("applyFill", "1")
        xfs.append(clone)
        newxf[sid] = len(xfs) - 1
    xfs.set("count", str(len(xfs)))

    blank = {"D": newxf[46], "F": newxf[44]}
    for r in TOTALS:
        m = re.search(r'(<row r="%d"[^>]*>)(.*?)(</row>)' % r, xml, re.S)
        body = m.group(2)
        for cell, old in re.findall(r'<c r="([A-G]%d)" s="(\d+)"' % r, body):
            body = body.replace('<c r="%s" s="%s"' % (cell, old),
                                '<c r="%s" s="%d"' % (cell, newxf[int(old)]), 1)
        for c, sid in blank.items():
            if '<c r="%s%d"' % (c, r) not in body:
                nxt = re.search(r'<c r="[%s]%d"' % ({"D": "EFG", "F": "G"}[c], r), body)
                tag = '<c r="%s%d" s="%d"/>' % (c, r, sid)
                body = (body[:nxt.start()] + tag + body[nxt.start():]) if nxt else body + tag
        xml = xml[:m.start()] + m.group(1) + body + m.group(3) + xml[m.end():]
    return xml


# ------------------------------------------------------------------------ main
def main():
    import os
    z = zipfile.ZipFile(SRC)                      # always from her original
    parts = {n: z.read(n) for n in z.namelist()}
    order = [i.filename for i in z.infolist()]
    z.close()

    root = etree.fromstring(parts["xl/styles.xml"])
    parts["xl/worksheets/sheet6.xml"] = three_way(
        parts["xl/worksheets/sheet6.xml"].decode("utf-8"), root).encode("utf-8")
    S = add_styles(root)
    parts["xl/styles.xml"] = etree.tostring(root, xml_declaration=True,
                                            encoding="UTF-8", standalone=True)

    wbx = parts["xl/workbook.xml"].decode("utf-8")
    rels = parts["xl/_rels/workbook.xml.rels"].decode("utf-8")
    ct = parts["[Content_Types].xml"].decode("utf-8")

    # calcChain is a cache of formula order. Adding sheets invalidates it, so drop it
    # and let Excel rebuild on open rather than ship a stale one.
    cc = "xl/calcChain.xml"
    if cc in parts:
        rid = re.search(r'<Relationship Id="(rId\d+)"[^>]*Target="calcChain.xml"[^>]*/>', rels)
        if rid:
            rels = rels.replace(rid.group(0), "")
        ct = re.sub(r'<Override PartName="/xl/calcChain\.xml"[^>]*/>', "", ct)
        del parts[cc]
        order = [n for n in order if n != cc]

    sid = max(int(x) for x in re.findall(r'sheetId="(\d+)"', wbx))
    rid = max(int(x) for x in re.findall(r'Id="rId(\d+)"', rels))
    nsheet = max(int(m) for m in re.findall(r"worksheets/sheet(\d+)\.xml", rels))
    add = []
    for title, xml in (("Pack Map", build_map(S)),
                       ("Pack Budget FY27", build_budget(S))):
        sid, rid, nsheet = sid + 1, rid + 1, nsheet + 1
        name = "worksheets/sheet%d.xml" % nsheet
        parts["xl/" + name] = xml
        add.append("xl/" + name)
        wbx = wbx.replace("</sheets>", '<sheet name="%s" sheetId="%d" r:id="rId%d"/></sheets>'
                          % (title, sid, rid))
        rels = rels.replace("</Relationships>",
                            '<Relationship Id="rId%d" Type="%s/worksheet" Target="%s"/>'
                            "</Relationships>" % (rid, R, name))
        ct = ct.replace("</Types>", '<Override PartName="/xl/%s" ContentType="application/vnd.'
                        'openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'
                        % name)
    wbx = wbx.replace('<calcPr calcId="191028"/>',
                      '<calcPr calcId="191028" fullCalcOnLoad="1"/>')
    parts["xl/workbook.xml"] = wbx.encode("utf-8")
    parts["xl/_rels/workbook.xml.rels"] = rels.encode("utf-8")
    parts["[Content_Types].xml"] = ct.encode("utf-8")

    zo = zipfile.ZipFile(OUT + ".tmp", "w", zipfile.ZIP_DEFLATED)
    for n in order + add:
        zo.writestr(n, parts[n])
    zo.close()
    os.replace(OUT + ".tmp", OUT)
    print("built", OUT, len(order) + len(add), "parts")


if __name__ == "__main__":
    main()
