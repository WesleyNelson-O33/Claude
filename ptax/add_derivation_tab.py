"""Add a tab to the payroll tax workbook showing how the lodged October
figure was arrived at.

The return lodged 5,996.99. The Employment Hero residential address report
gives 5,800.27. The difference is one ad hoc pay run, on 9 October, that the
report leaves out, and anyone looking at the file a year from now needs to
be able to see that without being told.

The tab carries the report as it stands, the pay run added back, the wages
used in the return, the tax worked out from them, and the cross check to
the split file. The report figures are typed because they are the source.
Everything derived from them is a formula.

The workbook holds a pivot table, four pictures and eight sets of cell
notes, so the sheet is added as XML and every other part is copied through
untouched.

Run with the workbook to read and the file to write.
"""
import re
import sys
import zipfile

TITLE, HEAD, LABEL = '3', '9', '4'      # styles already in the workbook
MONEY, MONEYB, PLAIN = '1', '10', '0'
TAB = 'Oct 2026 derivation'

# the Employment Hero residential address report, 1 to 31 October 2026
REPORT = [
    ('NSW', 162504.03, 21675.13, 2725.88, 357.24),
    ('QLD', 12936.70, 1552.40, 0.0, 0.0),
    ('SA', 5397.65, 647.72, 0.0, 0.0),
    ('VIC', 17732.10, 2067.21, 0.0, 0.0),
    ('WA', 7594.52, 911.34, 0.0, 0.0),
]
# the 9 October pay run, from the split file, wages and super
ADHOC = {'NSW': (3434.06, 412.09), 'SA': (267.86, 32.14)}

NOTES = [
    'The split file has three October pay dates: the 6th, the 9th and the '
    '20th. The 9th was an ad hoc run.',
    'The payroll tax report picks up the 6th and the 20th and leaves the 9th '
    'out. Re-running the report on 6 November gave the same figures, so it '
    'is not a timing problem.',
    'The 9 October run is ordinary pay, not an extra. One employee is paid a '
    'flat 80 hours and 3,434.06 every fortnight all year, and the period '
    'ending 2 October was paid on the 9th. The other is an 8 hour top up to '
    'the run of the 6th.',
    'Wages paid in October are October wages, so they belong in the October '
    'return. Lodging on the report alone would have understated New South '
    'Wales wages by 3,846.15 and the tax by 196.72.',
    'This is not a rule about ad hoc runs. August had one on the 26th, two '
    'termination payments of 10,931.19, and the August report included it. '
    'October is the exception and the reason is not in the data.',
    'The control that catches this is on the PRT tab of the split file. '
    'Split total, less any pay run the report does not have, must equal the '
    'report.',
]


def esc(t):
    return (t.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;'))


class Sheet:
    def __init__(self):
        self.rows = {}

    def put(self, ref, style, kind, value):
        col, row = re.match(r'([A-Z]+)(\d+)', ref).groups()
        self.rows.setdefault(int(row), []).append((col, style, kind, value))

    def text(self, ref, value, style=PLAIN):
        self.put(ref, style, 'text', value)

    def num(self, ref, value, style=MONEY):
        self.put(ref, style, 'number', value)

    def formula(self, ref, value, style=MONEY):
        self.put(ref, style, 'formula', value)

    def xml(self, widths):
        def colnum(c):
            n = 0
            for ch in c:
                n = n * 26 + ord(ch) - 64
            return n
        out = []
        for r in sorted(self.rows):
            cells = sorted(self.rows[r], key=lambda x: colnum(x[0]))
            body = []
            for col, style, kind, v in cells:
                ref = f'{col}{r}'
                if kind == 'formula':
                    # a sheet name such as Main & NSW carries an ampersand,
                    # which has to be escaped or the whole part stops parsing
                    body.append(f'<c r="{ref}" s="{style}">'
                                f'<f>{esc(str(v))}</f></c>')
                elif kind == 'number':
                    body.append(f'<c r="{ref}" s="{style}"><v>{v}</v></c>')
                else:
                    body.append(f'<c r="{ref}" s="{style}" t="inlineStr">'
                                f'<is><t xml:space="preserve">{esc(str(v))}'
                                f'</t></is></c>')
            out.append(f'<row r="{r}">{"".join(body)}</row>')
        cols = ''.join(f'<col min="{i}" max="{i}" width="{w}" customWidth="1"/>'
                       for i, w in enumerate(widths, start=1))
        last = max(self.rows)
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<worksheet xmlns="http://schemas.openxmlformats.org/'
                'spreadsheetml/2006/main">'
                f'<dimension ref="A1:F{last}"/>'
                '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
                '<sheetFormatPr defaultRowHeight="15"/>'
                f'<cols>{cols}</cols>'
                f'<sheetData>{"".join(out)}</sheetData>'
                '<pageMargins left="0.7" right="0.7" top="0.75" '
                'bottom="0.75" header="0.3" footer="0.3"/></worksheet>')


def build():
    s = Sheet()
    s.text('A1', 'October 2026 payroll tax, how the lodged figure was '
                 'arrived at', TITLE)
    s.text('A2', 'Corporate Technology Services Pty Ltd')
    s.text('A3', 'Lodged 5,996.99. The Employment Hero residential address '
                 'report gives 5,800.27. This sets out the difference.')

    s.text('A5', '1. THE EH RESIDENTIAL ADDRESS REPORT, 1 TO 31 OCTOBER 2026',
           LABEL)
    for i, h in enumerate(('State', 'Total wages', 'Total super',
                           'Salary sacrifice', 'Pre-tax deductions',
                           'Grand total')):
        s.text(f'{chr(65+i)}6', h, HEAD)
    for i, (st, w, su, ss, pt) in enumerate(REPORT):
        r = 7 + i
        s.text(f'A{r}', st)
        for c, v in zip('BCDE', (w, su, ss, pt)):
            s.num(f'{c}{r}', v)
        s.formula(f'F{r}', f'SUM(B{r}:E{r})')
    r = 7 + len(REPORT)
    s.text(f'A{r}', 'Total', LABEL)
    for c in 'BCDEF':
        s.formula(f'{c}{r}', f'SUM({c}7:{c}{r-1})', MONEYB)

    s.text('A14', '2. THE AD HOC PAY RUN OF 9 OCTOBER, WHICH THE REPORT DOES '
                  'NOT HAVE', LABEL)
    for i, h in enumerate(('State', 'Wages', 'Super', 'Total')):
        s.text(f'{chr(65+i)}15', h, HEAD)
    for i, (st, (w, su)) in enumerate(ADHOC.items()):
        r = 16 + i
        s.text(f'A{r}', st)
        s.num(f'B{r}', w)
        s.num(f'C{r}', su)
        s.formula(f'D{r}', f'B{r}+C{r}')
    s.text('A18', 'Total', LABEL)
    for c in 'BCD':
        s.formula(f'{c}18', f'SUM({c}16:{c}17)', MONEYB)

    s.text('A20', '3. THE WAGES USED IN THE RETURN', LABEL)
    for i, h in enumerate(('State', 'Per the report', 'The ad hoc run',
                           'Used in the return')):
        s.text(f'{chr(65+i)}21', h, HEAD)
    for i, (st, *_rest) in enumerate(REPORT):
        r, src = 22 + i, 7 + i
        s.text(f'A{r}', st)
        s.formula(f'B{r}', f'F{src}')
        if st == 'NSW':
            s.formula(f'C{r}', 'D16')
        elif st == 'SA':
            s.formula(f'C{r}', 'D17')
        else:
            s.num(f'C{r}', 0)
        s.formula(f'D{r}', f'B{r}+C{r}')
    s.text('A27', 'Total', LABEL)
    for c in 'BCD':
        s.formula(f'{c}27', f'SUM({c}22:{c}26)', MONEYB)

    s.text('A29', '4. THE TAX', LABEL)
    lines = [
        ('New South Wales wages', 'D22'),
        ('Interstate wages', 'D23+D24+D25+D26'),
        ('Total Australian wages', 'C30+C31'),
        ('Days in October', '31'),
        ('Threshold entitlement', '(C30/C32)*(1200000/365*C33)'),
        ('Taxable after the threshold', 'C32-C34-C31'),
    ]
    for i, (lbl, f) in enumerate(lines):
        r = 30 + i
        s.text(f'A{r}', lbl)
        s.formula(f'C{r}', f, PLAIN if lbl.startswith('Days') else MONEY)
    s.text('A36', 'Rate')
    s.put('C36', PLAIN, 'number', '0.0545')
    s.text('A37', 'PAYROLL TAX PAYABLE', LABEL)
    s.formula('C37', 'C35*C36', MONEYB)
    s.text('A38', 'Agrees with Main & NSW cell E23')
    s.formula('C38', "'Main & NSW'!E23")
    s.formula('D38', 'IF(ROUND(C37-C38,2)=0,"yes","CHECK")', PLAIN)

    s.text('A40', '5. WHY THE RETURN DIFFERS FROM THE REPORT', LABEL)
    for i, n in enumerate(NOTES):
        s.text(f'A{41+i}', n)

    r = 41 + len(NOTES) + 1
    s.text(f'A{r}', '6. THE CROSS CHECK TO THE SPLIT FILE', LABEL)
    s.text(f'A{r+1}', 'Split gross, plus SG super, plus the employer '
                      'contribution')
    s.num(f'C{r+1}', 240248.07)
    s.text(f'A{r+2}', 'Less the 9 October pay run')
    s.formula(f'C{r+2}', '-D18')
    s.text(f'A{r+3}', 'Equals the report grand total')
    s.formula(f'C{r+3}', f'C{r+1}+C{r+2}')
    s.text(f'A{r+4}', 'The report says')
    s.formula(f'C{r+4}', 'F12')
    s.text(f'A{r+5}', 'Difference, must be nil', LABEL)
    s.formula(f'C{r+5}', f'C{r+3}-C{r+4}', MONEYB)
    s.text(f'A{r+7}', 'And the wages used in the return above, '
                      'which is the split file total')
    s.formula(f'C{r+7}', 'D27', MONEYB)
    return s


def main():
    if len(sys.argv) != 3:
        print(__doc__)
        return 1
    src, dst = sys.argv[1], sys.argv[2]
    zin = zipfile.ZipFile(src)
    parts = {n: zin.read(n) for n in zin.namelist()}

    nums = [int(m.group(1)) for n in zin.namelist()
            if (m := re.match(r'xl/worksheets/sheet(\d+)\.xml$', n))]
    part = f'xl/worksheets/sheet{max(nums)+1}.xml'
    wb = parts['xl/workbook.xml'].decode('utf-8')
    rels = parts['xl/_rels/workbook.xml.rels'].decode('utf-8')
    if TAB in wb:
        print(f'  a tab called {TAB} is already there, nothing done')
        return 1
    rid = 'rId%d' % (max(int(i) for i in re.findall(
        'Id="rId([0-9]+)"', rels)) + 1)
    sid = max(int(i) for i in re.findall(r'sheetId="(\d+)"', wb)) + 1

    parts[part] = build().xml([44, 17, 17, 17, 17, 17]).encode('utf-8')
    parts['xl/_rels/workbook.xml.rels'] = rels.replace(
        '</Relationships>',
        f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/'
        f'officeDocument/2006/relationships/worksheet" '
        f'Target="worksheets/sheet{max(nums)+1}.xml"/></Relationships>'
    ).encode('utf-8')
    parts['xl/workbook.xml'] = re.sub(
        r'</sheets>',
        f'<sheet name="{TAB}" sheetId="{sid}" r:id="{rid}"/></sheets>',
        wb, count=1).encode('utf-8')
    ct = parts['[Content_Types].xml'].decode('utf-8')
    parts['[Content_Types].xml'] = ct.replace(
        '</Types>',
        f'<Override PartName="/{part}" ContentType="application/vnd.'
        f'openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '</Types>').encode('utf-8')

    order = list(zin.namelist()) + [part]
    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zout:
        for n in order:
            if n in zin.namelist():
                o = zin.getinfo(n)
                i = zipfile.ZipInfo(n, date_time=o.date_time)
                i.compress_type = o.compress_type
                i.external_attr = o.external_attr
                i.internal_attr = o.internal_attr
                i.create_system = o.create_system
            else:
                i = zipfile.ZipInfo(n, date_time=(2026, 11, 6, 12, 0, 0))
                i.compress_type = zipfile.ZIP_DEFLATED
            zout.writestr(i, parts[n])

    zc = zipfile.ZipFile(dst)
    kept = [n for n in zin.namelist()
            if re.search(r'pivot|drawing|comment|vml', n, re.I)]
    same = all(zc.read(n) == zin.read(n) for n in kept)
    print(f'  tab "{TAB}" added as {part}, sheetId {sid}, {rid}')
    print(f'  {len(zin.namelist())} parts in, {len(zc.namelist())} out')
    print(f'  {len(kept)} pivot, drawing and note parts, '
          f'byte for byte identical: {same}')
    return 0 if same else 1


if __name__ == '__main__':
    sys.exit(main())
