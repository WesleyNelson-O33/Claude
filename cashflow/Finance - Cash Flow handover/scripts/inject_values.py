"""Write cached results next to every formula so viewers that don't recalculate still show numbers.

LibreOffice is unavailable in the build sandbox, so results are computed with the
`formulas` package and written into each sheet's XML as <v> values.
"""
import re, zipfile, shutil, os, html
import formulas

def _values(path):
    sol = formulas.ExcelModel().loads(path).finish().calculate()
    out = {}
    for k, v in sol.items():
        if not hasattr(v, "value"): continue
        sheet = k.split("]")[-1].split("!")[0].strip("'").upper()
        ref = k.split("!")[-1]
        val = v.value[0][0]
        out[(sheet, ref)] = val
    return out

def _fmt(val):
    """Return (t_attr, text) for the <v> element."""
    if val is None or (isinstance(val, str) and val == ""):
        return "str", ""
    if isinstance(val, bool):
        return "b", "1" if val else "0"
    if isinstance(val, str):
        if val.startswith("#"):
            return "e", val
        return "str", html.escape(val, quote=False)
    try:
        f = float(val)
    except Exception:
        return "str", html.escape(str(val), quote=False)
    if f == int(f) and abs(f) < 1e15:
        return None, str(int(f))
    return None, repr(f)

def inject(path):
    vals = _values(path)
    tmp = path + ".tmp"
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        wbxml = zin.read("xl/workbook.xml").decode()
        rels = zin.read("xl/_rels/workbook.xml.rels").decode()
        rid_to_target = dict(re.findall(r'Target="/?([^"]+)"\s+Id="(rId\d+)"', rels)[i][::-1] for i in range(len(re.findall(r'Target="/?([^"]+)"\s+Id="(rId\d+)"', rels))))
        name_of = {}
        for name, rid in re.findall(r'<sheet name="([^"]+)" sheetId="\d+" state="\w+" r:id="(rId\d+)"', wbxml):
            name_of[rid_to_target[rid].lstrip("/")] = html.unescape(name).upper()
        patched = 0
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename in name_of:
                sheet = name_of[item.filename]
                xml = data.decode()
                def sub(m):
                    nonlocal patched
                    attrs, formula = m.group(1), m.group(2)
                    ref = re.search(r'r="([A-Z]+\d+)"', attrs).group(1)
                    if (sheet, ref) not in vals:
                        return m.group(0)
                    t, text = _fmt(vals[(sheet, ref)])
                    attrs = re.sub(r'\s+t="[^"]*"', "", attrs)
                    if t: attrs += f' t="{t}"'
                    patched += 1
                    return f"<c{attrs}><f>{formula}</f><v>{text}</v></c>"
                xml = re.sub(r'<c((?:\s+[a-z]+="[^"]*")+)><f>(.*?)</f>(?:<v\s*/>|<v></v>)?</c>', sub, xml, flags=re.S)
                data = xml.encode()
            zout.writestr(item, data)
    shutil.move(tmp, path)
    print(f"inject_values: wrote cached results for {patched} formula cells")
    return patched
