"""Correct the October rows in the Employment Hero payroll split file.

Three things are wrong with October on the 'EH data (2)' tab, and all three
are fixed with the file's own formulas rather than anything new.

1. Pay run 7 is in the file twice. 751 rows carry a state and the same 751
   rows are there again with the state blank, identical in every other
   column. They double count 101,245.94 of gross and 2,379.25 hours. The
   blank copy is cleared.

2. The State column is a lookup, =VLOOKUP(E,'Employee Details Report'!A:Q,
   9,FALSE), and it was never dragged past row 6065. That is why pay runs 8,
   9, 10 and 11 have no state. The same lookup is written down the rest of
   the October rows. All 61 employees involved are in that report and none
   of the states it returns contradicts a state already in the file.

3. Pay Dte Mth is typed in for October, with a trailing space, and pay run
   11 is labelled October although it is paid on 3 November. The column is a
   formula elsewhere in the file, =IF(C="","",TEXT(C,"mmmm")), which reads
   the month off the pay date. That formula is written down the October
   rows, so run 11 moves to November by itself and the trailing space goes.

Rows are cleared, not deleted, so nothing below them shifts and no formula,
pivot or comment anywhere in the file loses its place. The tab holds no
table and no pivot of its own, and the pivots that read it pick the change
up when they are refreshed.

Run with the split file to read and the file to write.
"""
import re
import sys
import zipfile

import openpyxl

SHEET = 'xl/worksheets/sheet5.xml'          # EH data (2)
TAB = 'EH data (2)'
MONTH = 'October'
STATE_COL, MONTH_COL = 22, 1                # V and A
RUN_COL, EMP_COL, STATE_IDX = 2, 5, 22

STATE_F = "VLOOKUP(E{r},'Employee Details Report'!A:Q,9,FALSE)"
MONTH_F = 'IF(C{r}="","",TEXT(C{r},"mmmm"))'


def survey(path):
    """Which rows belong to October, which carry a state, and which are the
    duplicate copy of pay run 7."""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb[TAB]
    oct_rows, run7_stated, run7_blank = [], [], []
    for i, row in enumerate(ws.iter_rows(values_only=True), 1):
        if i <= 2 or row[EMP_COL - 1] is None:
            continue
        if str(row[0]).strip() != MONTH:
            continue
        oct_rows.append(i)
        if row[RUN_COL - 1] == 7:
            (run7_stated if row[STATE_IDX - 1] else run7_blank).append(i)
    wb.close()
    return oct_rows, run7_stated, run7_blank


def clear_rows(xml, rows):
    """Empty every cell on these rows, keeping the row elements so nothing
    below them moves."""
    done = 0
    for r in sorted(rows):
        pat = rf'(<row r="{r}"[^>]*?)(?:/>|>.*?</row>)'
        m = re.search(pat, xml, re.S)
        if not m:
            return xml, f'row {r} not found'
        xml = xml[:m.start()] + m.group(1) + '/>' + xml[m.end():]
        done += 1
    return xml, done


def set_formula(xml, row, col, formula):
    ref = f'{chr(64 + col) if col <= 26 else "?"}{row}'
    pat = rf'(<c r="{ref}")(?P<a>[^>/]*)(?:/>|>(?P<b>.*?)</c>)'
    m = re.search(pat, xml, re.S)
    if m:
        attrs = re.sub(r'\s*t="[^"]*"', '', m.group('a'))
        return xml[:m.start()] + f'<c r="{ref}"{attrs}><f>{formula}</f></c>' \
            + xml[m.end():], True
    # no cell there yet, so add one in the right place on the row
    rm = re.search(rf'(<row r="{row}"[^>]*?)(/>|>(.*?)</row>)', xml, re.S)
    if not rm:
        return xml, False
    body = rm.group(3) or ''
    cell = f'<c r="{ref}"><f>{formula}</f></c>'
    out = []
    placed = False
    for cm in re.finditer(r'<c r="([A-Z]+)(\d+)".*?(?:/>|</c>)', body, re.S):
        c = cm.group(1)
        n = 0
        for ch in c:
            n = n * 26 + ord(ch) - 64
        if not placed and n > col:
            out.append(cell)
            placed = True
        out.append(cm.group(0))
    if not placed:
        out.append(cell)
    new = f'{rm.group(1)}>{"".join(out)}</row>'
    return xml[:rm.start()] + new + xml[rm.end():], True


def main():
    if len(sys.argv) != 3:
        print(__doc__)
        return 1
    src, dst = sys.argv[1], sys.argv[2]

    oct_rows, stated, blank = survey(src)
    print(f'  October rows on {TAB}: {len(oct_rows)}')
    print(f'  pay run 7: {len(stated)} with a state, {len(blank)} without')
    if len(stated) != len(blank):
        print('  those two should match if it is a straight duplicate, '
              'stopping')
        return 1

    zin = zipfile.ZipFile(src)
    parts = {n: zin.read(n) for n in zin.namelist()}
    xml = parts[SHEET].decode('utf-8')

    xml, n = clear_rows(xml, blank)
    if isinstance(n, str):
        print(f'  STOPPED, {n}')
        return 1
    print(f'  {n} duplicate rows cleared')

    keep = [r for r in oct_rows if r not in set(blank)]
    ok = 0
    for r in keep:
        xml, a = set_formula(xml, r, STATE_COL, STATE_F.format(r=r))
        xml, b = set_formula(xml, r, MONTH_COL, MONTH_F.format(r=r))
        ok += a and b
    print(f'  {ok} rows given the state lookup and the month formula')

    # stored answers on this tab are stale now, so they go and the file is
    # told to work itself out when it opens
    xml, cleared = re.subn(
        r'(<c r="[A-Z]+\d+"[^>]*?>)(<f[^>]*?(?:/>|>.*?</f>))<v>[^<]*</v>(</c>)',
        r'\1\2\3', xml, flags=re.S)
    xml = re.sub(r'(<c r="[A-Z]+\d+"[^>]*?)\s*t="(?:e|str)"([^>]*>)',
                 r'\1\2', xml)
    print(f'  {cleared} stored answers cleared on this tab')
    parts[SHEET] = xml.encode('utf-8')

    wb = parts['xl/workbook.xml'].decode('utf-8')
    if 'fullCalcOnLoad' not in wb:
        wb = re.sub(r'<calcPr ', '<calcPr fullCalcOnLoad="1" ', wb, count=1)
        parts['xl/workbook.xml'] = wb.encode('utf-8')

    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zout:
        for n2 in zin.namelist():
            old = zin.getinfo(n2)
            new = zipfile.ZipInfo(n2, date_time=old.date_time)
            new.compress_type = old.compress_type
            new.external_attr = old.external_attr
            new.internal_attr = old.internal_attr
            new.create_system = old.create_system
            zout.writestr(new, parts[n2])

    zc = zipfile.ZipFile(dst)
    touched = sorted(n2 for n2 in zin.namelist() if zc.read(n2) != zin.read(n2))
    kept = [n2 for n2 in zin.namelist()
            if re.search(r'pivot|drawing|comment|vml|table', n2, re.I)]
    same = all(zc.read(n2) == zin.read(n2) for n2 in kept)
    print(f'\n  parts changed: {touched}')
    print(f'  {len(kept)} pivot, table, drawing and comment parts, '
          f'byte for byte identical: {same}')
    return 0 if same else 1


if __name__ == '__main__':
    sys.exit(main())
