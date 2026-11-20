"""Add the Dept P&L Monthly tab without rewriting the rest of the pack.

openpyxl cannot round trip this workbook: it drops customXml parts, a sheet
rels file and chart detail, and Excel then refuses to open it. So the new sheet
is written as raw XML and injected into the original zip. Every other part is
copied across byte for byte, charts and conditional formatting included.
Styles are borrowed from existing cells on Q1 Dept P&L, so styles.xml is not
touched either.
"""
import re
import shutil
import zipfile
from xml.sax.saxutils import escape

SRC = ('/tmp/claude-0/-home-user-Claude/a00ffe59-433c-507b-9006-917d8c60083a'
       '/scratchpad/pack2/w.xlsx')
OUT = 'pack/CTS Financial Controller Pack - v3.xlsx'
NEW_NAME = 'Dept P&L Monthly'

BLOCKS = [('Support', [('Managed Services', 'C')]),
          ('Production', [('Events', 'E'), ('Video', 'G')]),
          ('Consulting', [('Integration', 'I'), ('Consulting', 'K')]),
          ('CTS', [('CTS total', 'M')])]
LINES = [(9, 'Total Sales Revenue'), (13, 'Labour'),
         (14, 'Equipment & Services'), (15, 'Transportation')]
MONTHS = [('Jul-26', 'B'), ('Aug-26', 'C'), ('Sep-26', 'D')]
COLS, YTD, PY = ['C', 'D', 'E'], 'F', 'H'
RANGE = re.compile(r"SUM\('([^']+)'!\$?([A-Z]{1,3})\$?(\d+):\$?([A-Z]{1,3})\$?(\d+)\)")


def month_of(f, mcol):
    return RANGE.sub(lambda m: m.group(0) if m.group(3) != m.group(5)
                     else "'%s'!%s%s" % (m.group(1), mcol, m.group(3)), f)


def ci(col):
    n = 0
    for ch in col:
        n = n * 26 + ord(ch) - 64
    return n


z = zipfile.ZipFile(SRC)
wbxml = z.read('xl/workbook.xml').decode('utf-8')
rels = z.read('xl/_rels/workbook.xml.rels').decode('utf-8')
ct = z.read('[Content_Types].xml').decode('utf-8')

sheets = re.findall(r'<sheet[^>]*?name="([^"]*)"[^>]*?sheetId="(\d+)"[^>]*?r:id="([^"]+)"[^>]*?/>', wbxml)
relmap = dict(re.findall(r'<Relationship Id="([^"]+)"[^>]*?Target="([^"]+)"', rels))
q1file = None
for name, sid, rid in sheets:
    if name.replace('&amp;', '&') == 'Q1 Dept P&L':
        q1file = 'xl/' + relmap[rid].lstrip('/').replace('worksheets/', 'worksheets/')
assert q1file, 'Q1 Dept P&L not found'
q1 = z.read(q1file).decode('utf-8')

# style index of each Q1 cell, so the new tab inherits the pack's look
STYLE = {}
for m in re.finditer(r'<c r="([A-Z]+\d+)"(?:\s+s="(\d+)")?', q1):
    STYLE[m.group(1)] = m.group(2) or '0'
S_TITLE, S_SUB = STYLE.get('B1', '0'), STYLE.get('B2', '0')
S_GRP, S_SUBH = STYLE.get('C4', '0'), STYLE.get('C5', '0')
S_HDR, S_LINEHDR = STYLE.get('C6', '0'), STYLE.get('B6', '0')
S_BOLDLAB, S_INDLAB = STYLE.get('B9', '0'), STYLE.get('B13', '0')
S_MONEY, S_PCT = STYLE.get('C9', '0'), STYLE.get('C19', '0')
S_BOLDMON = STYLE.get('C16', S_MONEY)

# the Q1 formulas, read straight out of its XML
F = {}
for mr in re.finditer(r'<row r="(\d+)"[^>]*>(.*?)</row>', q1, re.S):
    rn = int(mr.group(1))
    for mc in re.finditer(r'<c r="([A-Z]+)%d"[^>]*>(.*?)</c>' % rn, mr.group(2), re.S):
        fm = re.search(r'<f[^>]*>(.*?)</f>', mc.group(2), re.S)
        if fm:
            F[(rn, mc.group(1))] = '=' + fm.group(1).replace('&amp;', '&') \
                .replace('&lt;', '<').replace('&gt;', '>')

rows_by_num = {}


def cell(ref, s, value=None, formula=None):
    if formula is not None:
        return '<c r="%s" s="%s"><f>%s</f></c>' % (ref, s, escape(formula.lstrip('=')))
    if value is None:
        return '<c r="%s" s="%s"/>' % (ref, s)
    return ('<c r="%s" s="%s" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>'
            % (ref, s, escape(str(value))))


def row(n, items, ht=15, level=0):
    # Excel requires rows in ascending order, so collect and sort at the end.
    lv = ' outlineLevel="%d"' % level if level else ''
    rows_by_num[n] = ('<row r="%d" ht="%s" customHeight="1"%s>%s</row>'
                      % (n, ht, lv, ''.join(items)))


row(1, [cell('B1', S_TITLE, 'Dept P&L Monthly')], 21)
row(2, [cell('B2', S_SUB, 'FY27 Q1 by month, by department. Same lines and the '
                          'same cost groupings as the Q1 Dept P&L tab.')])
row(3, [cell('B3', S_SUB, 'Click the minus in the margin to collapse a department. '
                          'Overhead is the residual and proves Labour, Equipment '
                          'and Transport add back to Total COGS.')])
r = 5
for grp, subs in BLOCKS:
    row(r, [cell('B%d' % r, S_GRP, grp)], 20)
    r += 1
    hdr = [cell('B%d' % r, S_LINEHDR, 'Line')]
    for i, (m, _) in enumerate(MONTHS):
        hdr.append(cell('%s%d' % (COLS[i], r), S_HDR, m))
    hdr.append(cell('%s%d' % (YTD, r), S_HDR, 'Q1 FY27'))
    hdr.append(cell('%s%d' % (PY, r), S_HDR, 'Q1 FY26'))
    row(r, hdr, 18)
    r += 1

    marks = {}
    for q1row, label in LINES:
        head = r
        subrows = []
        r += 1
        if len(subs) > 1:
            for sublab, col in subs:
                it = [cell('B%d' % r, S_INDLAB, '    ' + sublab)]
                for i, (_, mcol) in enumerate(MONTHS):
                    it.append(cell('%s%d' % (COLS[i], r), S_MONEY,
                                   formula=month_of(F[(q1row, col)], mcol)))
                it.append(cell('%s%d' % (YTD, r), S_MONEY,
                               formula='SUM(%s%d:%s%d)' % (COLS[0], r, COLS[2], r)))
                pycol = chr(ord(col) + 1)
                it.append(cell('%s%d' % (PY, r), S_MONEY, formula=F[(q1row, pycol)]))
                row(r, it, 15, level=1)
                subrows.append(r)
                r += 1
        it = [cell('B%d' % head, S_BOLDLAB, label)]
        for i in range(3):
            f = ('+'.join('%s%d' % (COLS[i], x) for x in subrows) if subrows
                 else month_of(F[(q1row, subs[0][1])], MONTHS[i][1]).lstrip('='))
            it.append(cell('%s%d' % (COLS[i], head), S_BOLDMON, formula=f))
        it.append(cell('%s%d' % (YTD, head), S_BOLDMON,
                       formula='SUM(%s%d:%s%d)' % (COLS[0], head, COLS[2], head)))
        pyf = ('+'.join('%s%d' % (PY, x) for x in subrows) if subrows
               else F[(q1row, chr(ord(subs[0][1]) + 1))].lstrip('='))
        it.append(cell('%s%d' % (PY, head), S_BOLDMON, formula=pyf))
        row(head, it, 16)
        marks[label] = head

    cogs = r
    it = [cell('B%d' % cogs, S_BOLDLAB, 'Total COGS')]
    for i, (_, mcol) in enumerate(MONTHS):
        it.append(cell('%s%d' % (COLS[i], cogs), S_BOLDMON,
                       formula='+'.join(month_of(F[(16, c)], mcol).lstrip('=')
                                        for _, c in subs)))
    it.append(cell('%s%d' % (YTD, cogs), S_BOLDMON,
                   formula='SUM(%s%d:%s%d)' % (COLS[0], cogs, COLS[2], cogs)))
    it.append(cell('%s%d' % (PY, cogs), S_BOLDMON,
                   formula='+'.join(F[(16, chr(ord(c) + 1))].lstrip('=') for _, c in subs)))
    row(cogs, it, 16)
    r += 1

    it = [cell('B%d' % r, S_INDLAB, '    Overhead')]
    for c2 in COLS + [YTD, PY]:
        it.append(cell('%s%d' % (c2, r), S_MONEY,
                       formula='ROUND(%s%d-%s%d-%s%d-%s%d,2)'
                       % (c2, cogs, c2, marks['Labour'],
                          c2, marks['Equipment & Services'],
                          c2, marks['Transportation'])))
    row(r, it, 15)
    r += 1

    gp = r
    it = [cell('B%d' % gp, S_BOLDLAB, 'Gross Profit')]
    for i, (_, mcol) in enumerate(MONTHS):
        it.append(cell('%s%d' % (COLS[i], gp), S_BOLDMON,
                       formula='+'.join(month_of(F[(18, c)], mcol).lstrip('=')
                                        for _, c in subs)))
    it.append(cell('%s%d' % (YTD, gp), S_BOLDMON,
                   formula='SUM(%s%d:%s%d)' % (COLS[0], gp, COLS[2], gp)))
    it.append(cell('%s%d' % (PY, gp), S_BOLDMON,
                   formula='+'.join(F[(18, chr(ord(c) + 1))].lstrip('=') for _, c in subs)))
    row(gp, it, 16)
    r += 1

    it = [cell('B%d' % r, S_BOLDLAB, 'GPM')]
    for c2 in COLS + [YTD, PY]:
        it.append(cell('%s%d' % (c2, r), S_PCT,
                       formula='IFERROR(%s%d/%s%d,"")'
                       % (c2, gp, c2, marks['Total Sales Revenue'])))
    row(r, it, 16)
    r += 3

row(r, [cell('B%d' % r, S_SUB, 'Overhead is the residual: Total COGS less Labour, '
                               'Equipment and Services and Transportation. It reads '
                               'nil, which proves the three groupings add back to '
                               'the cost of sales in the ledger.')])
r += 1
row(r, [cell('B%d' % r, S_SUB, 'Q1 FY26 is the comparative from the Q1 Dept P&L '
                               'tab. FY26 was never split by month in the pack, so '
                               'it is shown as the quarter only.')])
LAST = r

sheet_xml = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    '<sheetPr><outlinePr summaryBelow="0"/></sheetPr>'
    '<dimension ref="A1:H%d"/>' % LAST +
    '<sheetViews><sheetView showGridLines="0" workbookViewId="0">'
    '<pane xSplit="2" topLeftCell="C1" activePane="topRight" state="frozen"/>'
    '</sheetView></sheetViews>'
    '<sheetFormatPr defaultRowHeight="15" outlineLevelRow="1"/>'
    '<cols><col min="1" max="1" width="2" customWidth="1"/>'
    '<col min="2" max="2" width="30" customWidth="1"/>'
    '<col min="3" max="6" width="14" customWidth="1"/>'
    '<col min="7" max="7" width="2" customWidth="1"/>'
    '<col min="8" max="8" width="14" customWidth="1"/></cols>'
    '<sheetData>' + ''.join(rows_by_num[k] for k in sorted(rows_by_num)) + '</sheetData>'
    '<pageMargins left="0.5" right="0.5" top="0.6" bottom="0.6" header="0.3" footer="0.3"/>'
    '</worksheet>')

nums = [int(m) for m in re.findall(r'xl/worksheets/sheet(\d+)\.xml', ' '.join(z.namelist()))]
newfile = 'sheet%d.xml' % (max(nums) + 1)
rid_nums = [int(m) for m in re.findall(r'Id="rId(\d+)"', rels)]
newrid = 'rId%d' % (max(rid_nums) + 1)
sheet_ids = [int(s[1]) for s in sheets]
newsid = max(sheet_ids) + 1

q1pos = wbxml.index('<sheet ', wbxml.index('name="Q1 Dept P&amp;L"') - 200)
endpos = wbxml.index('/>', wbxml.index('name="Q1 Dept P&amp;L"')) + 2
wbxml = (wbxml[:endpos] +
         '<sheet name="%s" sheetId="%d" r:id="%s"/>' % (escape(NEW_NAME), newsid, newrid) +
         wbxml[endpos:])
# make Excel calculate the new sheet on open, since it carries no cached values
if '<calcPr' in wbxml:
    wbxml = re.sub(r'<calcPr([^>]*?)/>',
                   lambda m: '<calcPr%s fullCalcOnLoad="1"/>'
                   % re.sub(r'\s*fullCalcOnLoad="[^"]*"', '', m.group(1)), wbxml)
else:
    wbxml = wbxml.replace('</workbook>', '<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>')

rels = rels.replace('</Relationships>',
                    '<Relationship Id="%s" Type="http://schemas.openxmlformats.org/'
                    'officeDocument/2006/relationships/worksheet" Target="worksheets/%s"/>'
                    '</Relationships>' % (newrid, newfile))
ct = ct.replace('</Types>',
                '<Override PartName="/xl/worksheets/%s" ContentType="application/'
                'vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                '</Types>' % newfile)

with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as out:
    for item in z.infolist():
        if item.filename == 'xl/calcChain.xml':
            continue
        data = z.read(item.filename)
        if item.filename == 'xl/workbook.xml':
            data = wbxml.encode('utf-8')
        elif item.filename == 'xl/_rels/workbook.xml.rels':
            data = rels.encode('utf-8')
        elif item.filename == '[Content_Types].xml':
            data = ct.encode('utf-8')
        out.writestr(item, data)
    out.writestr('xl/worksheets/' + newfile, sheet_xml)
print('injected %s as %s, last row %d' % (NEW_NAME, newfile, LAST))
