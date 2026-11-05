"""Correct the October 2026 payroll tax workbook by editing its sheet XML.

The workbook holds a pivot table, three drawings and eight sets of cell notes.
openpyxl drops all of those when it saves, so the file is edited inside the
zip and every other part is copied through byte for byte.

Only formulas are changed. No figure is typed over, and nothing that needs the
October Employment Hero report or the October timesheets is touched.

One mistake explains nearly all of it. A total was dragged one row too far, in
two places:

  Main & NSW   the Total Wages formula on row 19 was dragged down over rows
               20, 21 and 22 in the June column. Excel recorded it as one
               shared formula covering E19 to M22, so June's threshold,
               taxable wages and tax are all the Total Wages formula shifted
               down, and read +M13+M19, +M14+M20 and +M15+M21.

  QLD SA VIC WA  the Total Wages formula on row 13 was dragged down onto row
               14 from October onwards. Excel recorded it as one shared
               formula covering rows 13 and 14, so row 14 reads SUM(E4:E13)
               and counts the super twice.

What is changed:

  Main & NSW   B20:M20   the threshold, wrapped in IFERROR so a month with no
                         data reads nil instead of #DIV/0!, which is what
                         breaks the year to date totals in column N
  Main & NSW   M21, M22  taken out of the dragged group and given June's own
                         formulas
  Main & NSW   N20, N23  the two annual totals, widened to run to June
  QLD SA VIC WA  E14:L14 taken out of the dragged group and given the same
                         formula the other months use, wrapped in IFERROR so
                         a month with no wages reads nil rather than #DIV/0!

Left alone on purpose:

  Main & NSW   E3        September's 2,001.04, needs the October EH report
  Main & NSW   E3 etc    the exempt per diems, needs the October EH report
  Main & NSW   K23       the stray 1.1276, source unknown
  September 2026 Timesheets          needs the October export

Row 14 on the state tabs is left in place and made consistent, but it computes
nothing. The 1.2 million threshold is New South Wales', the other states are
accrued at a fixed monthly figure, and no cell reads row 14. It would be
better off the sheet, which is Danica's call, not mine.

Run with the workbook to read and the file to write.
"""
import re
import sys
import zipfile

MONTHS = 'BCDEFGHIJKLM'
NSW = 'xl/worksheets/sheet3.xml'
STATES = {'QLD': 'xl/worksheets/sheet7.xml', 'SA': 'xl/worksheets/sheet8.xml',
          'VIC': 'xl/worksheets/sheet9.xml', 'WA': 'xl/worksheets/sheet10.xml'}

# ---------------------------------------------------------------- the edits
# shared ranges to narrow, so a dragged total stops covering the row below
NARROW = {
    NSW: [('E19:M22', 'E19:M19')],
}
for part in STATES.values():
    NARROW[part] = [('C13:L14', 'C13:L13'), ('E13:L14', 'E13:L13'),
                    ('G13:J14', 'G13:J13')]

# formulas to set explicitly, cell by cell
SET = {NSW: {}, **{p: {} for p in STATES.values()}}
for col in MONTHS:
    SET[NSW][f'{col}20'] = f'IFERROR(({col}12/{col}19)*(1200000/$N$1*{col}1),0)'
SET[NSW]['M21'] = '+M19-M20-M18'
SET[NSW]['M22'] = '+M21*$A$22'
SET[NSW]['N20'] = 'SUM(B20:M20)'
SET[NSW]['N23'] = 'SUM(B23:M23)'
for part in STATES.values():
    for col in 'EFGHIJKL':
        SET[part][f'{col}14'] = (
            f'IFERROR(({col}13/{col}13)*(1200000/$N$1*{col}1),0)')

A1 = re.compile(r'([A-Z]+)(\d+)')


def inside(ref, cell):
    """Is a cell reference inside an A1 range."""
    def xy(s):
        c, r = A1.match(s).groups()
        n = 0
        for ch in c:
            n = n * 26 + ord(ch) - 64
        return n, int(r)
    a, b = (ref.split(':') + [ref])[:2]
    (x1, y1), (x2, y2), (x, y) = xy(a), xy(b), xy(cell)
    return min(x1, x2) <= x <= max(x1, x2) and min(y1, y2) <= y <= max(y1, y2)


def set_formula(xml, ref, formula):
    """Give one cell an explicit formula. Any shared-formula link, stale
    cached value and error flag go, and the cell keeps its style."""
    pat = rf'<c r="{re.escape(ref)}"(?P<a>[^>/]*)(?:/>|>(?P<b>.*?)</c>)'
    m = re.search(pat, xml, re.S)
    if not m:
        return xml, None
    was = 'empty'
    if m.group('b'):
        f = re.search(r'<f[^>]*?(?:/>|>(.*?)</f>)', m.group('b'), re.S)
        if f:
            was = f.group(1) or '(shared, no text of its own)'
    attrs = re.sub(r'\s*t="[^"]*"', '', m.group('a'))
    new = f'<c r="{ref}"{attrs}><f>{formula}</f></c>'
    return xml[:m.start()] + new + xml[m.end():], was


def orphans(xml):
    """Every cell that points at a shared formula must sit inside the range
    its master declares. A dragged or narrowed range leaves orphans, and
    Excel repairs or mis-reads those, so the file must have none."""
    refs = {si: r for r, si in re.findall(
        r'<f t="shared" ref="([^"]+)" si="(\d+)"', xml)}
    bad = []
    for cell, si in re.findall(
            r'<c r="([A-Z]+\d+)"[^>]*><f t="shared"[^>]*si="(\d+)"', xml):
        if si not in refs:
            bad.append(f'{cell} points at si={si}, which has no master')
        elif not inside(refs[si], cell):
            bad.append(f'{cell} is outside si={si} range {refs[si]}')
    return bad


def main():
    if len(sys.argv) != 3:
        print(__doc__)
        return 1
    src, dst = sys.argv[1], sys.argv[2]
    zin = zipfile.ZipFile(src)
    parts = {n: zin.read(n) for n in zin.namelist()}

    for part in [NSW] + list(STATES.values()):
        tab = 'Main & NSW' if part == NSW else \
            next(k for k, v in STATES.items() if v == part)
        xml = parts[part].decode('utf-8')
        print(f'[{tab}]')
        for old, new in NARROW[part]:
            if f'ref="{old}"' in xml:
                xml = xml.replace(f'ref="{old}"', f'ref="{new}"')
                print(f'  shared range {old} narrowed to {new}')
        for ref, formula in sorted(SET[part].items()):
            xml, was = set_formula(xml, ref, formula)
            if was is None:
                print(f'  STOPPED, {ref} not found')
                return 1
            print(f'  {ref:<5} {was[:38]:<40}->  {formula}')
        bad = orphans(xml)
        if bad:
            print('  STOPPED, shared formulas left dangling:')
            for b in bad:
                print(f'    {b}')
            return 1
        print('  no dangling shared formulas')
        parts[part] = xml.encode('utf-8')

    wb = parts['xl/workbook.xml'].decode('utf-8')
    if 'fullCalcOnLoad' not in wb:
        wb = re.sub(r'<calcPr ', '<calcPr fullCalcOnLoad="1" ', wb, count=1)
        parts['xl/workbook.xml'] = wb.encode('utf-8')
        print('\n  workbook set to work itself out when it opens, because '
              'every changed cell had its stored answer removed')

    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zout:
        for n in zin.namelist():
            old = zin.getinfo(n)
            # a fresh entry, because the original carries the old checksum
            # and sizes and the changed parts no longer match them
            new = zipfile.ZipInfo(n, date_time=old.date_time)
            new.compress_type = old.compress_type
            new.external_attr = old.external_attr
            new.internal_attr = old.internal_attr
            new.create_system = old.create_system
            zout.writestr(new, parts[n])

    zc = zipfile.ZipFile(dst)
    touched = sorted(n for n in zin.namelist() if zc.read(n) != zin.read(n))
    kept = [n for n in zin.namelist()
            if re.search(r'pivot|drawing|comment|vml', n, re.I)]
    same = all(zc.read(n) == zin.read(n) for n in kept)
    print(f'\n  {len(zin.namelist())} parts in, {len(zc.namelist())} out')
    print(f'  parts changed: {len(touched)}')
    for n in touched:
        print(f'    {n}')
    print(f'  {len(kept)} pivot, drawing and note parts, byte for byte '
          f'identical: {same}')
    return 0 if same else 1


if __name__ == '__main__':
    sys.exit(main())
