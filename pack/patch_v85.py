"""Split Depreciation & amortisation into two lines, and sit them between
EBITDA and EBIT.

Before, every P&L block read:

    Gross profit / Gross profit % / Depreciation & amortisation /
    EBITDA / EBITDA % / EBIT / EBIT % / Net profit / Net profit %

which put the add-back above the subtotal it is added back to. Now it reads:

    Gross profit / Gross profit % / EBITDA / EBITDA % /
    Depreciation & amortisation / Instant asset write off /
    EBIT / EBIT % / Net profit / Net profit %

The two lines are the pack's own split. The Engine already tags each GL row
D&A or IAWO, and the old single line was the sum of both, with the write off
switched on by Setup B28. Nothing is reclassified: the two lines add back to
the old figure to the cent, in every month, on every tab.

The block grows by one row and takes the blank row under it, so nothing below
row 182 moves and no other tab needs touching apart from the row numbers it
quotes.

It also adds Setup B40, a Yes/No that blanks the Net profit lines. It blanks
them with a conditional number format, so the figures are still there for the
tabs that read them: Recode Check, Xero History and Cost Analysis all compare
net profit back to Xero and would break if the cells were emptied.
"""
import re
import sys
from copy import copy

import openpyxl
from openpyxl.formatting.rule import Rule
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.styles.numbers import NumberFormat
from openpyxl.utils import get_column_letter as gl

PACK = 'pack/CTS Financial Controller Pack - v20.xlsx'

MAIN = ['Summary', 'P&L FY27', 'Budget FY27', 'Actual FY26', 'Budget FY26']
PLTABS = ['P&L FY27', 'Budget FY27', 'Actual FY26', 'Budget FY26']

# old row -> new row, on the five sheets that carry the main block
MAIN_MAP = {176: 175,   # EBITDA
            177: 176,   # EBITDA %
            175: 177,   # Depreciation & amortisation
            178: 179,   # EBIT
            179: 180,   # EBIT %
            180: 181,   # Net profit
            181: 182}   # Net profit %
DA_ROW, IAWO_ROW = 177, 178
EBITDA_ROW, EBIT_ROW = 175, 179
NET_ROWS = (181, 182)

# the six departments' Engine rows, D&A then IAWO
ENG_DA = [192, 195, 198, 201, 204, 207]
ENG_IAWO = [193, 196, 199, 202, 205, 208]

SETUP_SECTION, SETUP_TOGGLE = 39, 40

# A qualified range writes the sheet once: 'Tab'!$B$180:$M$180. The second
# endpoint has no sheet of its own but belongs to the same sheet, so it has to
# move with the first. Missing that is what half-remapped the ranges first time.
CELLREF = re.compile(
    r"(?:'([^']+)'|([A-Za-z_][A-Za-z0-9_.]*))!"
    r"(\$?)([A-Z]{1,2})(\$?)(\d+)"
    r"(?::(\$?)([A-Z]{1,2})(\$?)(\d+))?")
BARE = re.compile(r"(?<![A-Za-z0-9_])(\$?)([A-Z]{1,2})(\$?)(\d+)")


def remap(text, rowmap, sheets, own_sheet):
    """Point row references at their new rows.

    Cross-sheet references are rewritten when they name one of `sheets`. Bare
    references are rewritten when the formula itself sits on one of `sheets`.
    The cross-sheet ones are masked out first so the bare pass cannot see them.
    """
    if not isinstance(text, str) or not text.startswith('='):
        return text
    held = []

    def mask(m):
        name = m.group(1) or m.group(2)
        hit = name in sheets

        def row_of(n):
            r = int(n)
            return rowmap[r] if hit and r in rowmap else r

        q = f"'{m.group(1)}'" if m.group(1) else m.group(2)
        ref = f'{q}!{m.group(3)}{m.group(4)}{m.group(5)}{row_of(m.group(6))}'
        if m.group(10) is not None:
            ref += (f':{m.group(7)}{m.group(8)}{m.group(9)}'
                    f'{row_of(m.group(10))}')
        held.append(ref)
        return f'\x00{len(held) - 1}\x00'

    text = CELLREF.sub(mask, text)
    if own_sheet in sheets:
        def bare(m):
            row = int(m.group(4))
            row = rowmap.get(row, row)
            return f'{m.group(1)}{m.group(2)}{m.group(3)}{row}'
        text = BARE.sub(bare, text)
    return re.sub(r'\x00(\d+)\x00', lambda m: held[int(m.group(1))], text)


def snapshot(ws, rows):
    """Formula/value plus style for every cell on these rows."""
    out = {}
    for r in rows:
        out[r] = [(c, ws.cell(r, c).value, copy(ws.cell(r, c)._style))
                  for c in range(1, ws.max_column + 1)]
    return out


def put(ws, row, cells, rowmap, sheets, name):
    for c, v, style in cells:
        cell = ws.cell(row, c)
        cell.value = remap(v, rowmap, sheets, name)
        cell._style = copy(style)


def main():
    wb = openpyxl.load_workbook(PACK, data_only=False)
    touched = {'remapped': 0, 'written': 0}

    # ---------------------------------------------------------------- 1
    # Every formula in the workbook that quotes a row of the five sheets,
    # except the block rows themselves, which move with their own remap.
    block = set(MAIN_MAP) | {182}
    for name in wb.sheetnames:
        ws = wb[name]
        for row in ws.iter_rows():
            for cell in row:
                if name in MAIN and cell.row in block:
                    continue
                v = cell.value
                if not isinstance(v, str) or not v.startswith('='):
                    continue
                new = remap(v, MAIN_MAP, MAIN, name)
                if new != v:
                    cell.value = new
                    touched['remapped'] += 1
    print(f"  references repointed outside the block: {touched['remapped']}")

    # ---------------------------------------------------------------- 2
    # Move the block on the five sheets.
    for name in MAIN:
        ws = wb[name]
        snap = snapshot(ws, range(175, 182))
        for c in range(1, ws.max_column + 1):
            for r in range(175, 183):
                ws.cell(r, c).value = None
        for old, new in MAIN_MAP.items():
            put(ws, new, snap[old], MAIN_MAP, MAIN, name)
            touched['written'] += 1
        # the second depreciation line is the first one, one row down
        put(ws, IAWO_ROW,
            [(c, remap(v, {DA_ROW: IAWO_ROW}, MAIN, name), s)
             for c, v, s in snapshot(ws, [DA_ROW])[DA_ROW]],
            {}, MAIN, name)
        ws.cell(DA_ROW, 1).value = '    Depreciation & amortisation'
        ws.cell(IAWO_ROW, 1).value = '    Instant asset write off'
        ws.row_dimensions[IAWO_ROW].height = ws.row_dimensions[DA_ROW].height

    # ---------------------------------------------------------------- 3
    # P&L FY27 drives the two lines off the Engine. The other three tabs carry
    # no monthly depreciation, exactly as the single line carried none.
    ws = wb['P&L FY27']
    for c in range(2, 14):
        L, ec = gl(c), gl(c + 3)
        da = '+'.join(f'Engine!${ec}{r}' for r in ENG_DA)
        ia = '+'.join(f'Engine!${ec}{r}' for r in ENG_IAWO)
        ws.cell(DA_ROW, c).value = (
            f'=IF({L}$5<=Setup!$B$5,({da}),\'Budget FY27\'!{L}{DA_ROW})')
        ws.cell(IAWO_ROW, c).value = (
            f'=IF({L}$5<=Setup!$B$5,IF(Setup!$B$28="Yes",({ia}),0),'
            f'\'Budget FY27\'!{L}{IAWO_ROW})')

    # ---------------------------------------------------------------- 4
    # EBITDA now takes both lines out, EBIT puts both back.
    for name in PLTABS:
        ws = wb[name]
        for c in range(1, ws.max_column + 1):
            L = gl(c)
            for row, sign in ((EBITDA_ROW, '-'), (EBIT_ROW, '+')):
                cell = ws.cell(row, c)
                v = cell.value
                if not isinstance(v, str) or f'{sign}{L}{DA_ROW}' not in v:
                    continue
                cell.value = v + f'{sign}{L}{IAWO_ROW}'
                touched['written'] += 1
    print(f"  block rows rewritten: {touched['written']}")

    # ---------------------------------------------------------------- 5
    # The same move on each of the four departmental blocks.
    ws = wb['P&L by Dept']
    BLOCKS = [
        (70, 6, [195, 198], [196, 199],
         "('Overhead Split'!$Q$52+'Overhead Split'!$R$52)", '+'),
        (147, 83, [201, 204], [202, 205],
         "('Overhead Split'!$S$52+'Overhead Split'!$T$52)", '+'),
        (197, 160, [192], [193], "'Overhead Split'!$P$52", '+'),
        (247, 210, [207], [208], None, '-'),
    ]
    for b, d, dep, iawo, key, sign in BLOCKS:
        bmap = {b + 3: b + 2, b + 4: b + 3, b + 2: b + 4,
                b + 5: b + 6, b + 6: b + 7, b + 7: b + 8, b + 8: b + 9}
        snap = snapshot(ws, range(b + 2, b + 9))
        for c in range(1, ws.max_column + 1):
            for r in range(b + 2, b + 10):
                ws.cell(r, c).value = None
        for old, new in bmap.items():
            put(ws, new, snap[old], bmap, ['P&L by Dept'], 'P&L by Dept')
        put(ws, b + 5,
            [(c, remap(v, {b + 4: b + 5}, ['P&L by Dept'], 'P&L by Dept'), s)
             for c, v, s in snapshot(ws, [b + 4])[b + 4]],
            {}, ['P&L by Dept'], 'P&L by Dept')
        ws.cell(b + 4, 1).value = 'Depreciation & amortisation'
        ws.cell(b + 5, 1).value = 'Instant asset write off'
        ws.row_dimensions[b + 5].height = ws.row_dimensions[b + 4].height
        for c in range(2, 14):
            L, ec = gl(c), gl(c + 3)
            idx = (f"INDEX('Overhead Split'!$C52:$N52,1,"
                   f"MATCH({L}${d},'Overhead Split'!$C$22:$N$22,0))")
            admin = ('' if key is None and sign == '+' else
                     f'{sign}IF(Setup!$B$28="Yes",'
                     + (f'{key}*{idx}' if key else idx) + ',0)')
            ws.cell(b + 4, c).value = (
                f'=IF({L}${d}<=Setup!$B$5,'
                + '+'.join(f'Engine!${ec}{r}' for r in dep) + ',0)')
            ws.cell(b + 5, c).value = (
                f'=IF({L}${d}<=Setup!$B$5,'
                + '+'.join(f'IF(Setup!$B$28="Yes",Engine!${ec}{r},0)'
                           for r in iawo)
                + f'{admin},0)')
        for c in range(1, ws.max_column + 1):
            L = gl(c)
            for row, s in ((b + 2, '-'), (b + 6, '+')):
                cell = ws.cell(row, c)
                v = cell.value
                if isinstance(v, str) and f'{s}{L}{b + 4}' in v:
                    cell.value = v + f'{s}{L}{b + 5}'
    print('  departmental blocks rebuilt: 4')

    # ---------------------------------------------------------------- 6
    # The switch, and the format that blanks the Net profit lines.
    st = wb['Setup']
    st.cell(SETUP_SECTION, 1).value = '5. WHAT TO SHOW'
    st.cell(SETUP_SECTION, 1)._style = copy(st.cell(27, 1)._style)
    st.cell(SETUP_TOGGLE, 1).value = 'Show the Net profit lines'
    st.cell(SETUP_TOGGLE, 1)._style = copy(st.cell(28, 1)._style)
    st.cell(SETUP_TOGGLE, 2).value = 'Yes'
    st.cell(SETUP_TOGGLE, 2)._style = copy(st.cell(28, 2)._style)

    blank = Rule(type='expression', stopIfTrue=False,
                 formula=[f'Setup!$B${SETUP_TOGGLE}="No"'],
                 dxf=DifferentialStyle(
                     numFmt=NumberFormat(numFmtId=183, formatCode=';;;')))
    for name in MAIN:
        ws = wb[name]
        last = gl(ws.max_column)
        ws.conditional_formatting.add(
            f'A{NET_ROWS[0]}:{last}{NET_ROWS[1]}', blank)
    ws = wb['P&L by Dept']
    last = gl(ws.max_column)
    for b, *_ in BLOCKS:
        ws.conditional_formatting.add(f'A{b + 8}:{last}{b + 9}', blank)
    print(f'  Setup B{SETUP_TOGGLE} added, Net profit lines made hideable')

    wb.save(PACK)
    print('saved', PACK)
    return 0


if __name__ == '__main__':
    sys.exit(main())
