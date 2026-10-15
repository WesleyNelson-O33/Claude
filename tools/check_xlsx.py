"""
Check an xlsx for the things Excel refuses to open, which LibreOffice and
openpyxl will both happily let through.

    python3 check_xlsx.py "Workbook.xlsx"

Every finding here is something that produces Excel's "We found a problem with
some content" dialog. A clean run is not a guarantee, but it catches the
schema-order and dangling-reference faults that cause it in practice.
"""
import re, sys, zipfile
import xml.etree.ElementTree as ET

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
RNS = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
WS_ORDER = ['sheetPr', 'dimension', 'sheetViews', 'sheetFormatPr', 'cols', 'sheetData',
            'sheetCalcPr', 'sheetProtection', 'protectedRanges', 'scenarios', 'autoFilter',
            'sortState', 'dataConsolidate', 'customSheetViews', 'mergeCells', 'phoneticPr',
            'conditionalFormatting', 'dataValidations', 'hyperlinks', 'printOptions',
            'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks',
            'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing',
            'legacyDrawing', 'legacyDrawingHF', 'drawingHF', 'picture', 'oleObjects',
            'controls', 'webPublishItems', 'tableParts', 'extLst']
SP_ORDER = ['tabColor', 'outlinePr', 'pageSetUpPr']

def ordered(seq, order, where, out):
    idx = []
    for s in seq:
        if s not in order:
            out.append(f'{where}: unknown element <{s}>'); return
        idx.append(order.index(s))
    for i in range(len(idx) - 1):
        if idx[i] >= idx[i + 1]:
            out.append(f'{where}: <{seq[i+1]}> must come before <{seq[i]}>')

def main(path):
    out = []
    z = zipfile.ZipFile(path)
    names = set(z.namelist())
    wbx = z.read('xl/workbook.xml').decode()
    rels = z.read('xl/_rels/workbook.xml.rels').decode()
    # Excel and openpyxl write these attributes in a different order
    tgt = dict(re.findall(r'Id="(rId\d+)"[^>]*?Target="([^"]+)"', rels))
    tgt.update({i: t for t, i in re.findall(r'Target="([^"]+)"[^>]*?Id="(rId\d+)"', rels)})
    sheets = re.findall(r'<sheet[^>]*name="([^"]+)"[^>]*r:id="(rId\d+)"', wbx)
    def unesc(x):
        return (x.replace('&amp;', '&').replace('&apos;', "'").replace('&lt;', '<')
                 .replace('&gt;', '>').replace('&quot;', '"'))
    sheetnames = [unesc(n) for n, _ in sheets]
    # styles: how many dxf records exist
    styles = z.read('xl/styles.xml').decode()
    m = re.search(r'<dxfs count="(\d+)"', styles)
    ndxf = int(m.group(1)) if m else 0
    nfmt = len(re.findall(r'<numFmt ', styles))
    for name, rid in sheets:
        t = tgt.get(rid)
        if not t:
            out.append(f'{name}: no relationship for {rid}'); continue
        t = t.lstrip('/')
        part = t if t.startswith('xl/') else 'xl/' + t
        if part not in names:
            out.append(f'{name}: part {part} missing from the zip'); continue
        root = ET.fromstring(z.read(part))
        where = unesc(name)
        ordered([c.tag.replace(NS, '') for c in root], WS_ORDER, where, out)
        sp = root.find(NS + 'sheetPr')
        if sp is not None:
            ordered([c.tag.replace(NS, '') for c in sp], SP_ORDER, where + ' sheetPr', out)
        for cf in root.iter(NS + 'conditionalFormatting'):
            if not cf.get('sqref'):
                out.append(f'{where}: conditionalFormatting with no sqref')
            for rule in cf:
                d = rule.get('dxfId')
                if d is not None and int(d) >= ndxf:
                    out.append(f'{where}: conditional format points at dxf {d}, only {ndxf} exist')
        af = root.find(NS + 'autoFilter')
        if af is not None and not af.get('ref'):
            out.append(f'{where}: autoFilter with no ref')
    # defined names must point at a sheet that exists
    for nm, ref in re.findall(r'<definedName name="([^"]+)"[^>]*>([^<]*)</definedName>', wbx):
        for s in re.findall(r"(?:'([^']+)'|([A-Za-z_][\w.]*))!", unesc(ref)):
            sn = (s[0] or s[1]).replace("''", "'")
            if sn not in sheetnames:
                out.append(f'defined name {nm} points at sheet "{sn}", which does not exist')
        if '#REF' in ref:
            out.append(f'defined name {nm} contains #REF: {ref}')
    if 'xl/calcChain.xml' in names:
        out.append('calcChain.xml is present and is almost certainly stale; delete it')
    ct = z.read('[Content_Types].xml').decode()
    for p in names:
        if p.startswith('xl/worksheets/sheet') and f'/{p}' not in ct and 'Default Extension="xml"' not in ct:
            out.append(f'{p} has no content type')
    print(f'{len(out)} problem(s)')
    for o in out: print('  -', o)
    return 1 if out else 0

if __name__ == '__main__':
    sys.exit(main(sys.argv[1]))
