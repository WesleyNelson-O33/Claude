"""Put the monthly control onto the PRT tab of the payroll split file.

    split total, less any pay run the report does not have,
    should equal the Employment Hero residential address report

October 2026 is the month it was built for. The split file carries an ad hoc
pay run of 9 October that the report leaves out, worth 4,146.15, and that is
196.72 of New South Wales payroll tax. The control makes that visible before
the return goes in rather than after.

It goes in rows 36 to 42, clear of everything already on the tab, and is
written as XML so the five pivot tables, the comments and the drawings are
left exactly as they are.

Run with the split file to read and the file to write.
"""
import re
import sys
import zipfile

SHEET = 'xl/worksheets/sheet8.xml'          # PRT
NUM, TXT = '102', '103'                     # styles copied from B21 and C21
DATE = '87'                                 # a date format, as used for the pay dates

ADHOC = ("IF($B$37=\"\",0,-(SUMIFS('EH data (2)'!S:S,'EH data (2)'!C:C,"
         "$B$37)+SUMIFS('EH data (2)'!U:U,'EH data (2)'!C:C,$B$37)))")

BLOCK = [
    ('A36', TXT, 'text', 'THE MONTHLY CONTROL'),
    ('B37', DATE, 'formula', 'DATE(2026,10,9)'),
    ('C37', TXT, 'text',
     'Pay date the report does not have. Clear this if there is none.'),
    ('B38', NUM, 'formula', 'B21'),
    ('C38', TXT, 'text',
     'Split total. Gross, plus SG super, plus the employer contribution.'),
    ('B39', NUM, 'formula', ADHOC),
    ('C39', TXT, 'text', 'Less that pay run, both its wages and its super'),
    ('B40', NUM, 'formula', 'B38+B39'),
    ('C40', TXT, 'text', 'Should now equal the report'),
    ('B41', NUM, 'number', '236101.92'),
    ('C41', TXT, 'text',
     'EH residential address report grand total. Retype each month.'),
    ('B42', NUM, 'formula', 'B40-B41'),
    ('C42', TXT, 'text',
     'Must be nil. If it is not, name every dollar before you lodge.'),
]


def colnum(c):
    n = 0
    for ch in c:
        n = n * 26 + ord(ch) - 64
    return n


def put(xml, ref, style, kind, value):
    col, row = re.match(r'([A-Z]+)(\d+)', ref).groups()
    if kind == 'formula':
        cell = f'<c r="{ref}" s="{style}"><f>{value}</f></c>'
    elif kind == 'number':
        cell = f'<c r="{ref}" s="{style}"><v>{value}</v></c>'
    else:
        esc = (value.replace('&', '&amp;').replace('<', '&lt;')
               .replace('>', '&gt;'))
        cell = (f'<c r="{ref}" s="{style}" t="inlineStr">'
                f'<is><t xml:space="preserve">{esc}</t></is></c>')

    old = re.search(rf'<c r="{ref}"(?:[^>/]*)(?:/>|>.*?</c>)', xml, re.S)
    if old:
        return xml[:old.start()] + cell + xml[old.end():], 'replaced'

    rm = re.search(rf'(<row r="{row}"[^>]*?)(?:/>|>(.*?)</row>)', xml, re.S)
    if not rm:
        return xml, 'row missing'
    body = rm.group(2) or ''
    out, placed = [], False
    for cm in re.finditer(r'<c r="([A-Z]+)\d+".*?(?:/>|</c>)', body, re.S):
        if not placed and colnum(cm.group(1)) > colnum(col):
            out.append(cell)
            placed = True
        out.append(cm.group(0))
    if not placed:
        out.append(cell)
    attrs = re.sub(r'\s*/?$', '', rm.group(1))
    return xml[:rm.start()] + f'{attrs}>{"".join(out)}</row>' + xml[rm.end():], \
        'added'


def main():
    if len(sys.argv) != 3:
        print(__doc__)
        return 1
    src, dst = sys.argv[1], sys.argv[2]
    zin = zipfile.ZipFile(src)
    parts = {n: zin.read(n) for n in zin.namelist()}
    xml = parts[SHEET].decode('utf-8')

    for ref, style, kind, value in BLOCK:
        xml, how = put(xml, ref, style, kind, value)
        if how == 'row missing':
            print(f'  STOPPED, row for {ref} is not in the sheet')
            return 1
        print(f'  {ref:<5} {how:<9} {str(value)[:62]}')
    parts[SHEET] = xml.encode('utf-8')

    wb = parts['xl/workbook.xml'].decode('utf-8')
    if 'fullCalcOnLoad' not in wb:
        wb = re.sub(r'<calcPr ', '<calcPr fullCalcOnLoad="1" ', wb, count=1)
        parts['xl/workbook.xml'] = wb.encode('utf-8')

    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zout:
        for n in zin.namelist():
            o = zin.getinfo(n)
            i = zipfile.ZipInfo(n, date_time=o.date_time)
            i.compress_type = o.compress_type
            i.external_attr = o.external_attr
            i.internal_attr = o.internal_attr
            i.create_system = o.create_system
            zout.writestr(i, parts[n])

    zc = zipfile.ZipFile(dst)
    kept = [n for n in zin.namelist()
            if re.search(r'pivot|drawing|comment|vml|table', n, re.I)]
    same = all(zc.read(n) == zin.read(n) for n in kept)
    touched = sorted(n for n in zin.namelist() if zc.read(n) != zin.read(n))
    print(f'\n  parts changed: {touched}')
    print(f'  {len(kept)} pivot, table, drawing and comment parts, '
          f'byte for byte identical: {same}')
    return 0 if same else 1


if __name__ == '__main__':
    sys.exit(main())
