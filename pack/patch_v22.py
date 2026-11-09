"""Take the P&L off the green, and put every number back on black.

The rule, in order:
  A row of dark teal that carries no money is chrome. It becomes charcoal with
    white text and a bright teal hairline under it.
  A row of dark teal that does carry money is a band she has to read numbers
    off. It becomes light grey with charcoal text, so black numbers and red
    negatives are legible on it.
  Dark teal used as a font colour becomes charcoal.

Charcoal is the one colour in her palette that is not a department, so the
chrome can never be mistaken for the department colour coding.

The department headings on P&L by Dept are left exactly as they are. Those
are the colour coding.
"""
import os, json, copy
import openpyxl
from openpyxl.styles import PatternFill, Border, Side

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = f'{HERE}/CTS Financial Controller Pack - v20.xlsx'
B = json.load(open(f'{HERE}/data/brand.json'))
TEAL, AQUA, INK, WHITE = (B['dark_teal'].upper(), B['bright_teal'],
                          B['charcoal'], B['white'])
GREY, GREYL = B['grey'], B['grey_lt']

# the department colour coding, which this sweep must not touch
KEEP = {('P&L by Dept', 6), ('P&L by Dept', 58)}

wb = openpyxl.load_workbook(OUT)

def is_money(fmt):
    if not fmt or fmt == 'General':
        return False
    low = fmt.lower()
    if 'mmm' in low or 'yy' in low or low.startswith('d'):
        return False
    return '#' in fmt or '0' in fmt

def money_row(ws, r):
    for c in range(2, min(ws.max_column, 40) + 1):
        if is_money(ws.cell(r, c).number_format):
            return True
    return False

hair = Side(style='medium', color=AQUA)
chrome = banded = refont = 0
for ws in wb.worksheets:
    if ws.max_row > 5000:
        continue
    teal_rows = {}
    for row in ws.iter_rows():
        for c in row:
            f = c.fill
            rgb = getattr(f.start_color, 'rgb', None) if f and f.fill_type == 'solid' else None
            if isinstance(rgb, str) and rgb[-6:].upper() == TEAL:
                teal_rows.setdefault(c.row, []).append(c)
    for r, cells in teal_rows.items():
        if (ws.title, r) in KEEP:
            continue
        if money_row(ws, r):
            for c in cells:                      # a band she reads numbers off
                c.fill = PatternFill('solid', fgColor=GREY)
                nf = copy.copy(c.font)
                nf.color = openpyxl.styles.Color(rgb='FF' + INK)
                c.font = nf
            banded += 1
        else:                                    # chrome
            for c in cells:
                c.fill = PatternFill('solid', fgColor=INK)
                nf = copy.copy(c.font)
                nf.color = openpyxl.styles.Color(rgb='FF' + WHITE)
                c.font = nf
                bd = copy.copy(c.border)
                bd.bottom = hair
                c.border = bd
            chrome += 1
    # dark teal as a font colour anywhere else becomes charcoal
    for row in ws.iter_rows():
        for c in row:
            fc = getattr(c.font, 'color', None)
            rgb = getattr(fc, 'rgb', None) if fc is not None else None
            if isinstance(rgb, str) and rgb[-6:].upper() == TEAL:
                nf = copy.copy(c.font)
                nf.color = openpyxl.styles.Color(rgb='FF' + INK)
                c.font = nf
                refont += 1

# tab colours too, so the strip along the bottom stops being green
for ws in wb.worksheets:
    tc = ws.sheet_properties.tabColor
    rgb = getattr(tc, 'rgb', None) if tc is not None else None
    if isinstance(rgb, str) and rgb[-6:].upper() == TEAL:
        ws.sheet_properties.tabColor = INK

print(f'chrome rows moved to charcoal: {chrome}')
print(f'number bands moved to light grey: {banded}')
print(f'cells whose text was dark teal, now charcoal: {refont}')
wb.save(OUT)
print('saved')
