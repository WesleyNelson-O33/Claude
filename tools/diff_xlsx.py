"""
Compare two workbooks on everything, not just the numbers.

    python3 diff_xlsx.py OLD.xlsx NEW.xlsx [--skip Cleanup,Engine]

Checks sheet list and order, then for every sheet: values and formulas, fill,
font, border, number format, alignment, column widths, row heights, outline
level and hidden state, freeze panes, tab colour, gridlines, conditional
formatting, autofilter and print setup. Prints what differs and where.

Written because comparing only cell values let a stray colour through twice.
"""
import sys, collections
import openpyxl
from openpyxl.utils import get_column_letter as CL

def colour(c):
    if c is None: return None
    if getattr(c, 'type', None) == 'rgb': return c.rgb
    if getattr(c, 'type', None) == 'theme': return f'theme{c.theme}/{round(c.tint or 0, 3)}'
    if getattr(c, 'type', None) == 'indexed': return f'idx{c.indexed}'
    return None

def norm_align(v):   return None if v in (None, 'general') else v
def norm_wrap(v):    return bool(v)
def norm_height(v):  return None if v in (None, 15.0, 15) else round(v, 2)

def look(cell):
    f, b = cell.font, cell.border
    fl = cell.fill
    return (
        colour(fl.fgColor) if fl is not None and fl.fill_type == 'solid' else None,
        f.name, f.size, f.bold, f.italic, colour(f.color),
        cell.number_format,
        norm_align(cell.alignment.horizontal), norm_wrap(cell.alignment.wrap_text),
        tuple((getattr(b, s).style, colour(getattr(b, s).color))
              for s in ('left', 'right', 'top', 'bottom')),
    )

def sheet_props(ws):
    tc = ws.sheet_properties.tabColor
    op = ws.sheet_properties.outlinePr
    return {
        'freeze': ws.freeze_panes,
        'tabColor': colour(tc),
        'gridlines': ws.sheet_view.showGridLines,
        'summaryBelow': op.summaryBelow if op else None,
        'autofilter': str(ws.auto_filter.ref) if ws.auto_filter.ref else None,
        'orientation': ws.page_setup.orientation,
        'fitToWidth': ws.page_setup.fitToWidth,
        'printTitles': ws.print_title_rows,
    }

def main(fa, fb, skip):
    A = openpyxl.load_workbook(fa)
    B = openpyxl.load_workbook(fb)
    out = []
    if A.sheetnames != B.sheetnames:
        out.append(f'SHEET ORDER\n   old: {A.sheetnames}\n   new: {B.sheetnames}')
    for nm in A.sheetnames:
        if nm in skip or nm not in B.sheetnames: continue
        a, b = A[nm], B[nm]
        pa, pb = sheet_props(a), sheet_props(b)
        for k in pa:
            if k == 'fitToWidth' and (pa[k] or 1) == (pb[k] or 1): continue
        if pa[k] != pb[k]:
                out.append(f'{nm}: {k}  old={pa[k]!r}  new={pb[k]!r}')
        ca = {str(r.sqref) for r in a.conditional_formatting}
        cb = {str(r.sqref) for r in b.conditional_formatting}
        if ca != cb:
            out.append(f'{nm}: conditional formatting  old={sorted(ca)}  new={sorted(cb)}')
        rows = min(max(a.max_row, b.max_row), 500)
        cols = min(max(a.max_column, b.max_column), 40)
        buckets = collections.Counter(); examples = collections.defaultdict(list)
        for r in range(1, rows + 1):
            ra, rb = a.row_dimensions[r], b.row_dimensions[r]
            for k, x, y in (('height', norm_height(ra.height), norm_height(rb.height)),
                            ('outline', ra.outlineLevel or 0, rb.outlineLevel or 0),
                            ('hidden', bool(ra.hidden), bool(rb.hidden))):
                if x != y:
                    buckets[f'row {k}'] += 1
                    if len(examples[f'row {k}']) < 4:
                        examples[f'row {k}'].append(f'r{r} old={x!r} new={y!r}')
            for c in range(1, cols + 1):
                x, y = a.cell(r, c), b.cell(r, c)
                if x.value != y.value:
                    buckets['value'] += 1
                    if len(examples['value']) < 6:
                        examples['value'].append(
                            f'{CL(c)}{r} old={str(x.value)[:44]!r} new={str(y.value)[:44]!r}')
                la, lb = look(x), look(y)
                if la != lb:
                    for i, what in enumerate(('fill', 'font name', 'font size', 'bold', 'italic',
                                              'font colour', 'number format', 'align',
                                              'wrap', 'border')):
                        if la[i] != lb[i]:
                            buckets[what] += 1
                            if len(examples[what]) < 4:
                                examples[what].append(f'{CL(c)}{r} old={la[i]!r} new={lb[i]!r}')
        for c in range(1, cols + 1):
            L = CL(c)
            wa = a.column_dimensions[L].width if L in a.column_dimensions else None
            wrb = b.column_dimensions[L].width if L in b.column_dimensions else None
            if (round(wa, 2) if wa else None) != (round(wrb, 2) if wrb else None):
                buckets['column width'] += 1
                if len(examples['column width']) < 4:
                    examples['column width'].append(f'{L} old={wa} new={wrb}')
        if buckets:
            out.append(f'{nm}: ' + ', '.join(f'{k} x{v}' for k, v in buckets.most_common()))
            for k in buckets:
                for e in examples[k]: out.append(f'      {k}: {e}')
    print(f'{len(out)} difference group(s)')
    for o in out: print('  -' if not o.startswith('   ') else '', o)
    return 0

if __name__ == '__main__':
    skip = set()
    if '--skip' in sys.argv:
        skip = set(sys.argv[sys.argv.index('--skip') + 1].split(','))
    sys.exit(main(sys.argv[1], sys.argv[2], skip))
