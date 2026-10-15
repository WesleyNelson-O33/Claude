"""
Put back the row-outline settings that LibreOffice strips when it recalculates.

    python3 fix_outline.py "Workbook.xlsx" "Sheet One,Sheet Two"

Our summary rows sit ABOVE their detail, so the sheet needs
<outlinePr summaryBelow="0">. LibreOffice drops the whole element on save, and
without it Excel draws the collapse buttons against the wrong rows. Run this
AFTER recalc, every time, as the last step of the build.

Every other part of the zip is copied through byte for byte.
"""
import re, shutil, sys, zipfile

TAG = '<outlinePr applyStyles="0" summaryBelow="0" summaryRight="0"/>'

def sheet_parts(z):
    """Sheet display name -> part name inside the zip."""
    wbx = z.read('xl/workbook.xml').decode('utf-8')
    rels = z.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    tgt = dict(re.findall(r'Id="(rId\d+)"[^>]*?Target="([^"]+)"', rels))
    tgt.update(dict(re.findall(r'Target="([^"]+)"[^>]*?Id="(rId\d+)"', rels))and{})
    out = {}
    for name, rid in re.findall(r'<sheet[^>]*name="([^"]+)"[^>]*r:id="(rId\d+)"', wbx):
        t = tgt.get(rid)
        if not t: continue
        t = t.lstrip('/')
        out[name.replace('&amp;', '&').replace('&lt;', '<').replace('&gt;', '>')] = (
            t if t.startswith('xl/') else 'xl/' + t)
    return out

def patch(xml):
    if '<outlinePr' in xml:
        return re.sub(r'<outlinePr[^>]*/>|<outlinePr[^>]*>.*?</outlinePr>', TAG, xml, count=1)
    m = re.search(r'<sheetPr\b([^>]*?)(/?)>', xml)
    if m:
        if m.group(2) == '/':
            return xml[:m.start()] + f'<sheetPr{m.group(1)}>{TAG}</sheetPr>' + xml[m.end():]
        return xml[:m.end()] + TAG + xml[m.end():]
    m = re.search(r'(<worksheet\b[^>]*>)', xml)
    if not m: return None
    return xml[:m.end()] + f'<sheetPr>{TAG}</sheetPr>' + xml[m.end():]

def main(path, sheets):
    want = [s.strip() for s in sheets.split(',') if s.strip()]
    with zipfile.ZipFile(path) as z:
        parts = sheet_parts(z)
        missing = [s for s in want if s not in parts]
        if missing:
            print('  not in this workbook:', missing); return 1
        items = [(i, z.read(i.filename)) for i in z.infolist()]
    targets = {parts[s]: s for s in want}
    done = []
    shutil.copy(path, path + '.bak')
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as out:
        for info, data in items:
            if info.filename in targets:
                new = patch(data.decode('utf-8'))
                if new is None:
                    print('  could not patch', info.filename)
                else:
                    data = new.encode('utf-8'); done.append(targets[info.filename])
            out.writestr(info, data)
    print('  patched:', done)
    with zipfile.ZipFile(path) as z:
        for part, name in targets.items():
            ok = TAG in z.read(part).decode('utf-8')
            print(f'  {"OK " if ok else "BAD"} {name}')
    return 0

if __name__ == '__main__':
    if len(sys.argv) != 3:
        print(__doc__); sys.exit(2)
    sys.exit(main(sys.argv[1], sys.argv[2]))
