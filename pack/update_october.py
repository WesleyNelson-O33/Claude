"""Load the October exports into the controller pack and move it to October.

The pack cannot be round tripped by openpyxl, so each paste tab's sheetData is
rebuilt as raw XML and swapped into the original zip. Every other part is
copied byte for byte. Styles are taken from the cells already on each tab, so
styles.xml is untouched.
"""
import datetime
import re
import zipfile
from xml.sax.saxutils import escape

import openpyxl

SRC = 'pack/CTS Financial Controller Pack - v3.xlsx'
OUT = 'pack/CTS Financial Controller Pack - October 2026.xlsx'
OCT = '/tmp/claude-0/-home-user-Claude/a00ffe59-433c-507b-9006-917d8c60083a/scratchpad/oct'
EPOCH = datetime.datetime(1899, 12, 30)


def serial(dt):
    return (dt - EPOCH).days + (dt - EPOCH).seconds / 86400.0


def read(path, first_row=1, first_col=1, last_col=None):
    ws = openpyxl.load_workbook(path, data_only=True, read_only=True).worksheets[0]
    out = []
    for row in ws.iter_rows(min_row=first_row, values_only=True):
        r = list(row[first_col - 1:last_col])
        out.append(r)
    while out and all(v in (None, '') for v in out[-1]):
        out.pop()
    return out


# ---------------------------------------------------------------- the pastes
pla = read(OCT + '/PLA.xlsx', first_row=6, last_col=12)
pl = read(OCT + '/PL.xlsx')
bs = read(OCT + '/BS2.xlsx')

ar_raw = read(OCT + '/AR.xlsx')
ar = []
contact = None
for row in ar_raw:
    a = row[0]
    rest = row[1:]
    if isinstance(a, str) and a.strip():
        t = a.strip()
        if t.startswith('Total') or t in ('Aged Receivables Detail', 'Ageing by due date',
                                          'Invoice Date', 'Percentage of total') \
                or t.startswith('Corporate Technology') or t.startswith('As at'):
            continue
        if all(v in (None, '') for v in rest):
            contact = t
        continue
    if isinstance(a, datetime.datetime) and contact:
        ar.append([contact, a, row[1], row[2], row[3],
                   row[4], row[5], row[6], row[7], row[8], row[9], None])
print('rows: P&L Analysis %d | PL %d | BS %d | AR %d' % (len(pla), len(pl), len(bs), len(ar)))

z = zipfile.ZipFile(SRC)
wbxml = z.read('xl/workbook.xml').decode('utf-8')
rels = dict(re.findall(r'<Relationship Id="([^"]+)"[^>]*?Target="([^"]+)"',
                       z.read('xl/_rels/workbook.xml.rels').decode('utf-8')))
sheetfile = {}
for name, sid, rid in re.findall(
        r'<sheet[^>]*?name="([^"]*)"[^>]*?sheetId="(\d+)"[^>]*?r:id="([^"]+)"[^>]*?/>', wbxml):
    sheetfile[name.replace('&amp;', '&')] = 'xl/' + rels[rid].lstrip('/')


def styles_of(part, data_row):
    """column letter -> style index, taken from one existing data row."""
    xml = z.read(part).decode('utf-8')
    m = re.search(r'<row r="%d"[^>]*>(.*?)</row>' % data_row, xml, re.S)
    out = {}
    if m:
        for c in re.finditer(r'<c r="([A-Z]+)\d+"(?:\s+s="(\d+)")?', m.group(1)):
            out[c.group(1)] = c.group(2) or '0'
    return out


def colname(i):
    s = ''
    while i > 0:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def build(part, grid, first_row, style_row, formulas=None):
    """Replace sheetData with the grid, starting at first_row."""
    xml = z.read(part).decode('utf-8')
    st = styles_of(part, style_row)
    head = re.search(r'<sheetData>(.*?)</sheetData>', xml, re.S)
    keep = []
    if first_row > 1:
        for m in re.finditer(r'<row r="(\d+)"[^>]*>.*?</row>|<row r="(\d+)"[^>]*/>',
                             head.group(1), re.S):
            n = int(m.group(1) or m.group(2))
            if n < first_row:
                keep.append(m.group(0))
    rows = []
    for i, line in enumerate(grid):
        rn = first_row + i
        cells = []
        for j, v in enumerate(line, 1):
            col = colname(j)
            s = st.get(col, '0')
            ref = '%s%d' % (col, rn)
            if formulas and col in formulas:
                cells.append('<c r="%s" s="%s"><f>%s</f></c>'
                             % (ref, s, escape(formulas[col].format(r=rn))))
            elif v is None or v == '':
                continue
            elif isinstance(v, datetime.datetime):
                cells.append('<c r="%s" s="%s"><v>%s</v></c>' % (ref, s, serial(v)))
            elif isinstance(v, bool):
                cells.append('<c r="%s" s="%s" t="b"><v>%d</v></c>' % (ref, s, v))
            elif isinstance(v, (int, float)):
                cells.append('<c r="%s" s="%s"><v>%s</v></c>' % (ref, s, repr(v)))
            else:
                cells.append('<c r="%s" s="%s" t="inlineStr"><is><t xml:space="preserve">'
                             '%s</t></is></c>' % (ref, s, escape(str(v))))
        if cells:
            rows.append('<row r="%d">%s</row>' % (rn, ''.join(cells)))
    body = ''.join(keep) + ''.join(rows)
    xml = xml[:head.start()] + '<sheetData>' + body + '</sheetData>' + xml[head.end():]
    xml = re.sub(r'<dimension ref="[^"]*"/>',
                 '<dimension ref="A1:%s%d"/>' % (colname(max(len(l) for l in grid)),
                                                 first_row + len(grid) - 1), xml)
    return xml


new = {
    sheetfile['P&L Analysis']: build(sheetfile['P&L Analysis'], pla, 2, 2),
    sheetfile['PL_Paste']: build(sheetfile['PL_Paste'], pl, 1, 8),
    sheetfile['BS_Paste']: build(sheetfile['BS_Paste'], bs, 1, 9),
    sheetfile['AR_Paste']: build(sheetfile['AR_Paste'], ar, 2, 2,
                                 formulas={'L': 'SUM(F{r}:K{r})'}),
}

# move the pack to October: Summary B3 is the only typed input
sx = z.read(sheetfile['Summary']).decode('utf-8')
oct_serial = serial(datetime.datetime(2026, 10, 1))
sx2 = re.sub(r'<c r="B3"([^>]*?)>.*?</c>',
             lambda m: '<c r="B3"%s><v>%s</v></c>'
             % (re.sub(r'\s+t="[^"]*"', '', m.group(1)), oct_serial), sx, count=1, flags=re.S)
assert sx2 != sx, 'Summary B3 not found'
new[sheetfile['Summary']] = sx2

if '<calcPr' in wbxml:
    wbxml = re.sub(r'<calcPr([^>]*?)/>',
                   lambda m: '<calcPr%s fullCalcOnLoad="1"/>'
                   % re.sub(r'\s*fullCalcOnLoad="[^"]*"', '', m.group(1)), wbxml)

with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as out:
    for item in z.infolist():
        if item.filename == 'xl/calcChain.xml':
            continue
        data = new.get(item.filename)
        if data is not None:
            data = data.encode('utf-8')
        elif item.filename == 'xl/workbook.xml':
            data = wbxml.encode('utf-8')
        else:
            data = z.read(item.filename)
        out.writestr(item, data)
print('wrote', OUT)
