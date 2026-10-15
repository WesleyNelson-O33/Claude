"""
Re-apply <outlinePr summaryBelow="0"> after recalculation.

LibreOffice rewrites the workbook when it recalculates, and drops the outline
properties on the way through. Without summaryBelow="0" Excel assumes the
summary row sits BELOW its detail, so the + and - buttons render against the
wrong rows on every collapsible sheet.

Re-saving with openpyxl would fix the attribute but discard the cached formula
values, so this edits the sheet XML inside the zip and copies every other part
across untouched.
"""
import re, shutil, sys, zipfile
from pathlib import Path

OUTLINE = '<outlinePr applyStyles="0" summaryBelow="0" summaryRight="0"/>'

def apply(path, sheet_names):
    path = Path(path)
    z = zipfile.ZipFile(path)
    order = re.findall(r'<sheet name="([^"]+)"', z.read("xl/workbook.xml").decode())
    targets = {}
    for i, n in enumerate(order, 1):
        if n in sheet_names:
            targets[f"xl/worksheets/sheet{i}.xml"] = n
    missing = set(sheet_names) - set(targets.values())
    if missing:
        raise SystemExit(f"{path.name}: sheets not found -> {sorted(missing)}")

    parts, done = {}, []
    for item in z.infolist():
        data = z.read(item.filename)
        if item.filename in targets:
            x = data.decode()
            if "<outlinePr" in x:
                pass                                        # already there, leave it
            elif re.search(r"<sheetPr[^>]*/>", x):          # self-closing: expand it
                x = re.sub(r"<sheetPr([^>]*)/>", rf"<sheetPr\1>{OUTLINE}</sheetPr>", x, count=1)
                done.append(targets[item.filename])
            elif "<sheetPr" in x:                           # insert as the first child
                x = re.sub(r"(<sheetPr[^>]*>)", rf"\1{OUTLINE}", x, count=1)
                done.append(targets[item.filename])
            else:                                           # no sheetPr at all
                x = re.sub(r"(<worksheet[^>]*>)", rf"\1<sheetPr>{OUTLINE}</sheetPr>", x, count=1)
                done.append(targets[item.filename])
            data = x.encode()
        parts[item.filename] = (item, data)

    tmp = path.with_suffix(".outlinefix.tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        for name, (info, data) in parts.items():
            zi = zipfile.ZipInfo(name, date_time=info.date_time)
            zi.compress_type = zipfile.ZIP_DEFLATED
            zi.external_attr = info.external_attr
            out.writestr(zi, data)
    shutil.move(tmp, path)
    return done

def verify(path, sheet_names):
    z = zipfile.ZipFile(path)
    order = re.findall(r'<sheet name="([^"]+)"', z.read("xl/workbook.xml").decode())
    ok = []
    for i, n in enumerate(order, 1):
        if n not in sheet_names: continue
        x = z.read(f"xl/worksheets/sheet{i}.xml").decode()
        m = re.search(r'<outlinePr[^>]*/>', x)
        ok.append((n, bool(m) and 'summaryBelow="0"' in m.group(0)))
    return ok

if __name__ == "__main__":
    target, sheets = sys.argv[1], sys.argv[2].split(",")
    changed = apply(target, sheets)
    print(f"  patched: {changed or '(already present)'}")
    for n, good in verify(target, sheets):
        print(f"  {'OK ' if good else 'BAD'} {n}")
