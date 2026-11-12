"""Put a sheet's row grouping buttons above the detail instead of below.

Excel decides where a group's plus and minus button goes from summaryBelow.
Left on, it expects the total row under its detail. Danica's layout puts the
heading above its detail, so the button lands on the wrong row and the
bottom group gets no button at all. That is why the grouping would not
collapse.

openpyxl can set this, but the LibreOffice recalc step drops it again, so it
has to be written straight into the file afterwards, after every other step.

    tools/outline_above.py <workbook> <sheet name> [<sheet name> ...]
"""
import re
import shutil
import sys
import zipfile

OUTLINE = '<outlinePr summaryBelow="false" summaryRight="true"/>'


def fix(xml):
    if '<outlinePr' in xml:
        def setit(m):
            tag = m.group(0)
            if 'summaryBelow' in tag:
                return re.sub(r'summaryBelow="[^"]*"', 'summaryBelow="false"', tag)
            return tag[:-2] + ' summaryBelow="false"/>' if tag.endswith('/>') else tag
        return re.sub(r'<outlinePr[^>]*/?>', setit, xml, count=1)
    m = re.search(r'<sheetPr\b[^>]*?/>', xml)
    if m:                                   # empty sheetPr, give it a body
        return xml[:m.start()] + m.group(0)[:-2] + '>' + OUTLINE + '</sheetPr>' + xml[m.end():]
    m = re.search(r'<sheetPr\b[^>]*?>', xml)
    if m:
        at = m.end()
        tab = re.compile(r'<tabColor[^>]*/?>').match(xml, at)
        if tab:
            at = tab.end()
        return xml[:at] + OUTLINE + xml[at:]
    m = re.search(r'<worksheet\b[^>]*?>', xml)
    if not m:
        return None
    return xml[:m.end()] + '<sheetPr>' + OUTLINE + '</sheetPr>' + xml[m.end():]


def main():
    if len(sys.argv) < 3:
        print(__doc__.strip())
        return 1
    path, wanted = sys.argv[1], sys.argv[2:]
    z = zipfile.ZipFile(path)
    names = z.namelist()
    book = z.read('xl/workbook.xml').decode('utf-8')
    rels = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"',
                           z.read('xl/_rels/workbook.xml.rels').decode('utf-8')))
    sheets = re.findall(r'<sheet name="([^"]*)"[^>]*r:id="(rId\d+)"', book)

    targets = {}
    for raw, rid in sheets:
        name = raw.replace('&amp;', '&').replace('&lt;', '<').replace('&gt;', '>')
        if name in wanted:
            targets['xl/' + rels[rid].lstrip('/').replace('xl/', '')] = name
    missing = set(wanted) - set(targets.values())
    if missing:
        print('STOPPED: no such sheet: %s' % ', '.join(sorted(missing)))
        return 1

    out = path + '.tmp'
    zo = zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED)
    done = []
    for n in names:
        data = z.read(n)
        if n in targets:
            fixed = fix(data.decode('utf-8'))
            if fixed is None:
                print('STOPPED: %s does not look like a worksheet' % n)
                zo.close()
                return 1
            data = fixed.encode('utf-8')
            done.append(targets[n])
        zo.writestr(n, data)
    zo.close()
    z.close()
    shutil.move(out, path)
    for name in done:
        print('  %r: grouping buttons set above the detail' % name)
    return 0


if __name__ == '__main__':
    sys.exit(main())
