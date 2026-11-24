"""Complete the November payroll tax workbook, the way October's was done.

What it changes, and nothing else:

  Main & NSW      column F, the three New South Wales lines, on to the
                  Employment Hero basis Danica asked for
  VIC             column F, wages and super, likewise
  Timesheets      renamed for November and repasted from the split file
  PTax Allocations NSW
                  both pivots and the allocation block rebuilt at November's
                  width, pointing at column F of Main & NSW
  Other States PTAX Allocations
                  the pivot rebuilt for November
  Nov 2026 derivation
                  a new tab, the same shape as October's, setting out how the
                  lodged figure was arrived at

The workbook holds a pivot table, drawings and eight sets of cell notes, so
each sheet part is rebuilt as raw XML and every other part is copied through
byte for byte. The sheets that carry notes keep their legacyDrawing, so the
notes stay anchored.

Styles are read off the cells already in the workbook, so styles.xml is never
touched and the look is hers.

    ptax/complete_nov_workbook.py <November workbook> <split file> <out>
"""
import os
import re
import sys
import zipfile
from collections import defaultdict

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

MONTH, MONTH_SHORT, YEAR = 'November', 'Nov', 2026
# Danica asked for no tab renaming, so the timesheets tab keeps its name and
# only its contents change. Renaming it is Excel's job: do it on the tab and
# Excel moves the pivot source and the hidden filter name with it. Doing it
# from outside means changing three parts by hand and Excel refuses to open
# the pivot if any one of them is missed.
TIMESHEETS = 'October 2026 Timesheets'
DERIV = 'Nov 2026 derivation'
COL = 'F'                      # November is column F on Main & NSW
NSW_TAX = 6768.53

# Main & NSW column F, the Employment Hero payroll tax report basis
NSW_LINES = {3: 177335.13, 5: 23281.90, 6: 982.69}
VIC_LINES = {3: 20672.04, 5: 2383.45}

# the employee the report puts in another state, same override as the journals
MOVED = {'Nicol Schildknecht': 'VIC'}

STATES = ['NSW', 'QLD', 'SA', 'VIC', 'WA']


# ------------------------------------------------------------------ helpers
def esc(t):
    return str(t).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


def colletter(n):
    s = ''
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def colnum(s):
    n = 0
    for ch in s:
        n = n * 26 + ord(ch) - 64
    return n


def styles_of(xml):
    """ref -> style index, for every cell already on the sheet."""
    out = {}
    for m in re.finditer(r'<c r="([A-Z]+\d+)"(?:\s+s="(\d+)")?', xml):
        out[m.group(1)] = m.group(2) or '0'
    return out


def cell(ref, style, kind, value):
    if kind == 'f':
        return f'<c r="{ref}" s="{style}"><f>{esc(value)}</f></c>'
    if kind == 'n':
        return f'<c r="{ref}" s="{style}"><v>{value}</v></c>'
    return (f'<c r="{ref}" s="{style}" t="inlineStr">'
            f'<is><t xml:space="preserve">{esc(value)}</t></is></c>')


def sheet_data(rows):
    """rows: {rownum: [(colletter, style, kind, value), ...]}. Excel needs the
    rows in ascending order and the cells inside each row in column order."""
    out = []
    for r in sorted(rows):
        body = ''.join(cell(f'{c}{r}', s, k, v)
                       for c, s, k, v in sorted(rows[r],
                                                key=lambda x: colnum(x[0])))
        out.append(f'<row r="{r}">{body}</row>')
    return '<sheetData>' + ''.join(out) + '</sheetData>'


def swap(xml, data, dim=None, autofilter=None):
    out = re.sub(r'<sheetData>.*?</sheetData>|<sheetData/>', data, xml,
                 count=1, flags=re.S)
    if dim:
        out = re.sub(r'<dimension ref="[^"]*"/>',
                     f'<dimension ref="{dim}"/>', out, count=1)
    if autofilter is not None:
        out = re.sub(r'<autoFilter ref="[^"]*"',
                     f'<autoFilter ref="{autofilter}"', out, count=1)
    return out


def set_cell(xml, ref, value):
    """Replace one numeric cell, keeping its style."""
    pat = re.compile(r'<c r="%s"((?:\s+[a-z:]+="[^"]*")*)\s*(?:/>|>.*?</c>)'
                     % ref, re.S)
    m = pat.search(xml)
    if not m:
        raise SystemExit('cannot find %s' % ref)
    attrs = re.sub(r'\s+t="[^"]*"', '', m.group(1))
    return xml[:m.start()] + f'<c r="{ref}"{attrs}><v>{value}</v></c>' \
        + xml[m.end():]


# ------------------------------------------------------------- the source
def november(split):
    """Timesheet rows and units per state per job, for November."""
    wb = openpyxl.load_workbook(split, data_only=True, read_only=True)
    it = wb['EH data (2)'].iter_rows(values_only=True)
    next(it)
    next(it)
    lines, units = [], defaultdict(float)
    for r in it:
        if all(v is None for v in r) or r[10] is None:
            continue
        if str(r[0]).strip() != MONTH:
            continue
        state = MOVED.get(str(r[6]).strip(), str(r[21]).strip())
        job = f'{r[13]} - {r[14]}' if r[13] is not None else str(r[14])
        lines.append((str(r[6]).strip(), r[13], str(r[14]), job,
                      float(r[10]), state))
        units[(state, job)] += float(r[10])
    wb.close()
    lines.sort(key=lambda x: (x[0], x[3]))
    return lines, units


def main():
    if len(sys.argv) != 4:
        print(__doc__)
        return 1
    src, split, dst = sys.argv[1], sys.argv[2], sys.argv[3]
    lines, units = november(split)
    if not lines:
        print('STOPPED: no %s rows in the split file' % MONTH)
        return 1

    all_jobs = sorted({j for _s, j in units})
    nsw_jobs = sorted({j for s, j in units if s == 'NSW'})
    oth_jobs = sorted({j for s, j in units if s != 'NSW'})
    by_state = defaultdict(float)
    for (s, _j), u in units.items():
        by_state[s] += u
    total_units = sum(by_state.values())

    z = zipfile.ZipFile(src)
    parts = {n: z.read(n) for n in z.namelist()}
    S = {n: parts[f'xl/worksheets/{n}.xml'].decode('utf-8')
         for n in ('sheet3', 'sheet6', 'sheet10', 'sheet12', 'sheet14')}
    st = {n: styles_of(x) for n, x in S.items()}

    def sty(sheet, ref, fallback='0'):
        return st[sheet].get(ref, fallback)

    # ------------------------------------------------ Main & NSW and VIC
    for row, value in NSW_LINES.items():
        S['sheet3'] = set_cell(S['sheet3'], f'{COL}{row}', value)
    for row, value in VIC_LINES.items():
        S['sheet10'] = set_cell(S['sheet10'], f'{COL}{row}', value)

    # ----------------------------------------------------- the timesheets
    head = ['Employee Name', 'Location External Id', 'Location Name',
            'Job Name and Job Number', 'Units', 'State']
    rows = {1: [('D', sty('sheet6', 'D1'), 't', '*formula based'),
                ('E', sty('sheet6', 'E1'), 'n', round(total_units, 2)),
                ('K', sty('sheet6', 'K1'), 't',
                 '**Complete EH Split File, %s %d' % (MONTH, YEAR))],
            2: [(colletter(i + 1), sty('sheet6', f'{colletter(i + 1)}2'), 't', h)
                for i, h in enumerate(head)]}
    for i, (emp, ext, name, job, u, state) in enumerate(lines, start=3):
        rows[i] = [
            ('A', sty('sheet6', 'A3'), 't', emp),
            ('B', sty('sheet6', 'B3'), 't' if not isinstance(ext, (int, float))
             else 'n', ext if ext is not None else ''),
            ('C', sty('sheet6', 'C3'), 't', name),
            ('D', sty('sheet6', 'D3'), 't', job),
            ('E', sty('sheet6', 'E3'), 'n', u),
            ('F', sty('sheet6', 'F3'), 't', state),
        ]
    last = 2 + len(lines)
    S['sheet6'] = swap(S['sheet6'], sheet_data(rows), f'A1:L{last}',
                       f'A2:S{last}')

    # ------------------------------------------- PTax Allocations NSW
    g = lambda r, c: units.get((r, c), 0)
    rows = {}
    rows[1] = [('A', sty('sheet12', 'A1'), 't', 'Sum of Units'),
               ('B', sty('sheet12', 'B1'), 't', 'Column Labels')]
    rows[2] = [('A', sty('sheet12', 'A2'), 't', 'Row Labels')] + [
        (colletter(2 + i), sty('sheet12', 'B2'), 't', j)
        for i, j in enumerate(all_jobs)] + [
        (colletter(2 + len(all_jobs)), sty('sheet12', 'B2'), 't', 'Grand Total')]
    for k, state in enumerate(STATES):
        r = 3 + k
        cs = [('A', sty('sheet12', 'A3'), 't', state)]
        for i, j in enumerate(all_jobs):
            if g(state, j):
                cs.append((colletter(2 + i), sty('sheet12', 'D3'), 'n',
                           round(g(state, j), 2)))
        cs.append((colletter(2 + len(all_jobs)), sty('sheet12', 'D3'), 'n',
                   round(by_state[state], 2)))
        rows[r] = cs
    cs = [('A', sty('sheet12', 'A8'), 't', 'Grand Total')]
    for i, j in enumerate(all_jobs):
        cs.append((colletter(2 + i), sty('sheet12', 'B8'), 'n',
                   round(sum(g(s, j) for s in STATES), 2)))
    cs.append((colletter(2 + len(all_jobs)), sty('sheet12', 'B8'), 'n',
               round(total_units, 2)))
    rows[8] = cs

    # the New South Wales only pivot the allocation reads
    nsw_total = by_state['NSW']
    gt = colletter(2 + len(nsw_jobs))          # the Grand Total column
    rows[47] = [('A', sty('sheet12', 'A47'), 't', 'Row Labels')] + [
        (colletter(2 + i), sty('sheet12', 'B47'), 't', j)
        for i, j in enumerate(nsw_jobs)] + [
        (gt, sty('sheet12', 'B47'), 't', 'Grand Total')]
    rows[48] = [('A', sty('sheet12', 'A48'), 't', 'NSW')] + [
        (colletter(2 + i), sty('sheet12', 'B48'), 'n', round(g('NSW', j), 2))
        for i, j in enumerate(nsw_jobs)] + [
        (gt, sty('sheet12', 'B48'), 'n', round(nsw_total, 2))]
    rows[49] = [('A', sty('sheet12', 'A49'), 't', 'Grand Total')] + [
        (colletter(2 + i), sty('sheet12', 'B49'), 'n', round(g('NSW', j), 2))
        for i, j in enumerate(nsw_jobs)] + [
        (gt, sty('sheet12', 'B49'), 'n', round(nsw_total, 2))]

    # the allocation block. Column F lines up with pivot column B, so the
    # job in allocation column X is the one in pivot column X minus four.
    first, lastcol = 6, 6 + len(nsw_jobs)      # F .. the Grand Total column
    def pivot_of(c):
        return colletter(colnum(c) - 4)

    cost = ('=IF(ISNUMBER(SEARCH("PRD",{c}79)),"Production",'
            'IF(ISNUMBER(SEARCH("VID",{c}79)),"Video",'
            'IF(ISNUMBER(SEARCH("INT",{c}79)),"Integration",'
            'IF(ISNUMBER(SEARCH("CONS",{c}79)),"Consulting",'
            'IF(ISNUMBER(SEARCH("CTS",{c}79)),"CTS","Onsite")))))')

    cols = [colletter(c) for c in range(first, lastcol + 1)]
    rows[74] = [('E', sty('sheet12', 'E74'), 't', 'Cost Centre')] + [
        (c, sty('sheet12', 'F74'), 'f', cost.format(c=c)) for c in cols]
    lastjob = colletter(lastcol - 1)
    rows[76] = [('E', sty('sheet12', 'E76'), 'f', f'=+E77-E79')]
    # The rounding difference goes on the largest line, the same line the
    # journal puts it on, so the tab and the journal agree to the cent.
    share = {j: round(NSW_TAX * g('NSW', j) / nsw_total, 2) for j in nsw_jobs}
    plug = round(NSW_TAX - sum(v for v in share.values() if v), 2)
    biggest = max((j for j in share if share[j]), key=lambda j: share[j])
    plug_col = colletter(first + nsw_jobs.index(biggest))
    rows[77] = [('E', sty('sheet12', 'E77'), 'f',
                 f'=ROUND(SUM({cols[0]}77:{lastjob}77),2)')] + [
        (c, sty('sheet12', 'G77'), 'f',
         f'=ROUND({c}78,2){plug:+.2f}' if c == plug_col
         else f'=ROUND({c}78,2)') for c in cols]
    rows[78] = [('D', sty('sheet12', 'D78'), 't', 'CHECK'),
                ('E', sty('sheet12', 'E78'), 'f',
                 f"=+E79-'Main & NSW'!{COL}23")] + [
        (c, sty('sheet12', 'G78'), 'f', f'=SUBTOTAL(9,{c}80:{c}90)')
        for c in cols]
    rows[79] = [('A', sty('sheet12', 'A79'), 't', 'State'),
                ('B', sty('sheet12', 'B79'), 't', 'Row Labels'),
                ('C', sty('sheet12', 'C79'), 'f', '=SUBTOTAL(9,C80:C88)'),
                ('D', sty('sheet12', 'D79'), 'f', '=SUBTOTAL(9,D80:D88)'),
                ('E', sty('sheet12', 'E79'), 'f', '=SUM(E80:E82)')] + [
        (c, sty('sheet12', 'F79'), 'f', f'={pivot_of(c)}$47') for c in cols]
    rows[80] = [('A', sty('sheet12', 'A80'), 't', 'NSW'),
                ('B', sty('sheet12', 'B80'), 't', 'NSW'),
                ('C', sty('sheet12', 'C80'), 'f',
                 f"='Main & NSW'!{COL}12"),
                ('D', sty('sheet12', 'D80'), 'f', '=+C80/$C$79'),
                ('E', sty('sheet12', 'E80'), 'f',
                 f"=D80*'Main & NSW'!{COL}23")] + [
        (c, sty('sheet12', 'F80'), 'f',
         f'=$E80*({pivot_of(c)}48/${gt}48)') for c in cols]
    rows[92] = [('E', sty('sheet12', 'E86'), 't', 'Journal 1'),
                ('F', sty('sheet12', 'F86'), 'f', '=E77')]
    rows[93] = [('E', sty('sheet12', 'E87'), 't', 'Lodgement to ATO'),
                ('F', sty('sheet12', 'F87'), 'f', f"='Main & NSW'!{COL}23")]
    rows[94] = [('E', sty('sheet12', 'E88'), 't', 'Check '),
                ('F', sty('sheet12', 'F88'), 'f', '=F92-F93'),
                ('G', sty('sheet12', 'G88'), 't',
                 f'Agrees to lodgement, {MONTH} {YEAR}')]
    end = colletter(2 + len(all_jobs))
    S['sheet12'] = swap(S['sheet12'], sheet_data(rows), f'A1:{end}94',
                        f'A47:{gt}49')

    # ------------------------------- Other States PTAX Allocations
    oth_states = [s for s in STATES if s != 'NSW']
    gto = colletter(2 + len(oth_jobs))
    rows = {}
    rows[1] = [(colletter(2 + i), sty('sheet14', 'B1'), 'n',
                round(sum(g(s, j) for s in oth_states), 2))
               for i, j in enumerate(oth_jobs)]
    rows[2] = [('A', sty('sheet14', 'A2'), 't', 'Row Labels')] + [
        (colletter(2 + i), sty('sheet14', 'B2'), 't', j)
        for i, j in enumerate(oth_jobs)]
    for k, state in enumerate(oth_states):
        r = 3 + k
        cs = [('A', sty('sheet14', 'A3'), 't', state)]
        for i, j in enumerate(oth_jobs):
            if g(state, j):
                cs.append((colletter(2 + i), sty('sheet14', 'C3'), 'n',
                           round(g(state, j), 2)))
        cs.append((gto, sty('sheet14', 'Q3'), 'n', round(by_state[state], 2)))
        rows[r] = cs
    cs = [('A', sty('sheet14', 'A7'), 't', 'Grand Total')]
    for i, j in enumerate(oth_jobs):
        cs.append((colletter(2 + i), sty('sheet14', 'B7'), 'n',
                   round(sum(g(s, j) for s in oth_states), 2)))
    cs.append((gto, sty('sheet14', 'Q7'), 'n',
               round(sum(by_state[s] for s in oth_states), 2)))
    rows[7] = cs
    S['sheet14'] = swap(S['sheet14'], sheet_data(rows), f'A1:{gto}7',
                        f'A2:{gto}2')
    S['sheet14'] = re.sub(r'<sortState .*?</sortState>', '', S['sheet14'],
                          flags=re.S)

    for n, x in S.items():
        parts[f'xl/worksheets/{n}.xml'] = x.encode('utf-8')

    # ------------------------------------------------ the derivation tab
    from nov_derivation import build as build_deriv
    nums = [int(m.group(1)) for n in z.namelist()
            if (m := re.match(r'xl/worksheets/sheet(\d+)\.xml$', n))]
    part = f'xl/worksheets/sheet{max(nums) + 1}.xml'
    wbx = parts['xl/workbook.xml'].decode('utf-8')
    rels = parts['xl/_rels/workbook.xml.rels'].decode('utf-8')
    if DERIV not in wbx:
        rid = 'rId%d' % (max(int(i) for i in re.findall(
            'Id="rId([0-9]+)"', rels)) + 1)
        sid = max(int(i) for i in re.findall(r'sheetId="(\d+)"', wbx)) + 1
        parts[part] = build_deriv().encode('utf-8')
        parts['xl/_rels/workbook.xml.rels'] = rels.replace(
            '</Relationships>',
            f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org'
            f'/officeDocument/2006/relationships/worksheet" '
            f'Target="worksheets/sheet{max(nums) + 1}.xml"/></Relationships>'
        ).encode('utf-8')
        wbx = re.sub(r'</sheets>',
                     f'<sheet name="{DERIV}" sheetId="{sid}" r:id="{rid}"/>'
                     '</sheets>', wbx, count=1)
        ct = parts['[Content_Types].xml'].decode('utf-8')
        parts['[Content_Types].xml'] = ct.replace(
            '</Types>',
            f'<Override PartName="/{part}" ContentType="application/vnd.'
            f'openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            '</Types>').encode('utf-8')

    # The pivot on the allocations tab reads the timesheets tab by name, so
    # renaming the tab without telling the pivot cache leaves Excel refusing
    # to open the pivot source. The range grows with the month's row count.
    cache = 'xl/pivotCache/pivotCacheDefinition1.xml'
    if cache in parts:
        c = parts[cache].decode('utf-8')
        before = c
        c = re.sub(r'<worksheetSource ref="[^"]*" sheet="[^"]*"/>',
                   f'<worksheetSource ref="A2:F{last}" '
                   f'sheet="{TIMESHEETS}"/>', c, count=1)
        if c == before:
            raise SystemExit('cannot find the pivot worksheetSource')
        # the cache still holds last month's rows. refreshOnLoad is already
        # set, so Excel rebuilds it the moment the file opens.
        c = re.sub(r'recordCount="\d+"', f'recordCount="{len(lines)}"', c,
                   count=1)
        if 'refreshOnLoad' not in c:
            c = c.replace('<cacheSource', ' refreshOnLoad="1"><cacheSource', 1)
        parts[cache] = c.encode('utf-8')

    # the hidden filter name has to cover the month's rows, or the filter on
    # the timesheets tab stops at last month's last row
    wbx = wbx.replace(f"'{TIMESHEETS}'!$A$2:$S$1610",
                      f"'{TIMESHEETS}'!$A$2:$S${last}")
    if '<calcPr' in wbx:
        wbx = re.sub(r'<calcPr[^>]*/>',
                     '<calcPr calcId="191029" fullCalcOnLoad="1"/>', wbx)
    else:
        wbx = wbx.replace('</workbook>',
                          '<calcPr calcId="191029" fullCalcOnLoad="1"/>'
                          '</workbook>')
    parts['xl/workbook.xml'] = wbx.encode('utf-8')
    parts.pop('xl/calcChain.xml', None)

    order = [n for n in z.namelist() if n in parts]
    if part in parts and part not in order:
        order.append(part)
    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as out:
        for n in order:
            if n in z.namelist():
                o = z.getinfo(n)
                i = zipfile.ZipInfo(n, date_time=o.date_time)
                i.compress_type = o.compress_type
                i.external_attr = o.external_attr
                i.internal_attr = o.internal_attr
                i.create_system = o.create_system
            else:
                i = zipfile.ZipInfo(n, date_time=(YEAR, 11, 24, 12, 0, 0))
                i.compress_type = zipfile.ZIP_DEFLATED
            out.writestr(i, parts[n])

    zc = zipfile.ZipFile(dst)
    # the pivot cache definition is the one such part that has to change,
    # because it names the sheet the pivot reads
    kept = [n for n in z.namelist()
            if re.search(r'pivot|drawing|comment|vml', n, re.I)
            and n != 'xl/pivotCache/pivotCacheDefinition1.xml']
    same = all(zc.read(n) == z.read(n) for n in kept)
    print(f'  {len(lines):,} {MONTH} timesheet rows, {total_units:,.2f} units')
    print(f'  {len(all_jobs)} jobs in all, {len(nsw_jobs)} NSW, '
          f'{len(oth_jobs)} other states')
    print(f'  NSW units {by_state["NSW"]:,.2f}, other states '
          f'{total_units - by_state["NSW"]:,.2f}')
    print(f'  {len(kept)} pivot, drawing and note parts carried through '
          f'byte for byte: {same}')
    print(f'  written to {dst}')
    return 0 if same else 1


if __name__ == '__main__':
    sys.exit(main())
