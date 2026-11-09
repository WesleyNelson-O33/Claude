"""The Aware Super equipment split, brought into line with Xero.

INV-10596 was raised on the quote, which had the whole of Phase 2 as one
figure, so the pack's GL paste carries 144,049.00 on Installation Labour and
the 819.28 lighting programming line on Equipment Sales. Xero has since been
recoded to the invoice as it reads now:

    Phase 2 - Installation and Integration   93,421.00   42600 Equipment Sales
    Phase 2 - Installation and Integration   50,628.00   42300 Installation Labour
    Lighting Integration Programming            819.28   42300 Installation Labour

That moves 92,601.72 from Installation Labour to Equipment Sales, on
INTEGRATION, in September 2026. Nothing else changes: the invoice total, the
job number, the contact, the month and the GST are all as they were.

The split is applied to the 'P&L Analysis' tab, because that is the GL paste
and Cleanup is formula driven off it row for row. Cleanup's formulas address
'P&L Analysis' positionally and already run to row 20000, and Setup's three
range checks are to $D$20000 as well, so the inserted line is picked up with
nothing to widen.

This edit is a bridge for September only. The recode is in Xero now, so the
next GL export carries it and this patch is not needed again.
"""
import sys
import copy
import openpyxl

PACK = 'pack/CTS Financial Controller Pack - v20.xlsx'
SHEET = 'P&L Analysis'

LABOUR, EQUIP = '42300', '42600'
LABOUR_NAME, EQUIP_NAME = 'Installation Labour', 'Equipment Sales'
MOVE = 92601.72


def snapshot(ws, last):
    return [tuple(ws.cell(r, c).value for c in range(1, 14))
            for r in range(1, last + 1)]


def main():
    wb = openpyxl.load_workbook(PACK)
    ws = wb[SHEET]

    last = max(r for r in range(ws.max_row, 1, -1)
               if ws.cell(r, 1).value not in (None, ''))
    before = snapshot(ws, last)

    # locate the two lines by what they are, not by row number
    lab = eq = None
    for r in range(2, last + 1):
        if ws.cell(r, 9).value != 'INV-10596':
            continue
        if ws.cell(r, 1).value == LABOUR and ws.cell(r, 7).value == 144049:
            lab = r
        if ws.cell(r, 1).value == EQUIP and ws.cell(r, 7).value == 819.28:
            eq = r
    if lab is None or eq is None:
        print('  the two INV-10596 lines are not as expected, nothing done')
        return 1
    if eq != lab - 1:
        print(f'  expected the 819.28 line directly above {lab}, found {eq}')
        return 1
    print(f'  found the 819.28 line at r{eq} and the 144,049.00 line at r{lab}')

    # the 819.28 lighting programming line is labour, not equipment
    ws.cell(eq, 1).value = LABOUR
    ws.cell(eq, 2).value = LABOUR_NAME

    # Phase 2 splits: the equipment half stays on the existing row
    ws.cell(lab, 1).value = EQUIP
    ws.cell(lab, 2).value = EQUIP_NAME
    ws.cell(lab, 7).value = 93421.00
    ws.cell(lab, 11).value = ('Aware Super - Phase 2 - Installation and '
                              'Integration. Equipment')

    # the labour half goes on a new line directly under it
    new = lab + 1
    ws.insert_rows(new)
    for c in range(1, 14):
        src, dst = ws.cell(lab, c), ws.cell(new, c)
        dst.value = src.value
        dst.number_format = src.number_format
        dst.font = copy.copy(src.font)
        dst.alignment = copy.copy(src.alignment)
    ws.cell(new, 1).value = LABOUR
    ws.cell(new, 2).value = LABOUR_NAME
    ws.cell(new, 7).value = 50628.00
    ws.cell(new, 11).value = ('Aware Super - Phase 2 - Installation and '
                              'Integration. Labour')
    print(f'  labour half written to r{new}')

    # ---- nothing else may have moved
    after = snapshot(ws, last + 1)
    EQ_DESC = 'Aware Super - Phase 2 - Installation and Integration. Equipment'
    LB_DESC = 'Aware Super - Phase 2 - Installation and Integration. Labour'

    def recode(row, code, name, credit=None, desc=None):
        """The same GL line with the account, and optionally the credit and
        the description, replaced. Everything else is carried over."""
        out = list(row)
        out[0], out[1] = code, name
        if credit is not None:
            out[6] = credit
        if desc is not None:
            out[10] = desc
        return tuple(out)

    want = (before[:eq - 1]
            + [recode(before[eq - 1], LABOUR, LABOUR_NAME)]
            + [recode(before[lab - 1], EQUIP, EQUIP_NAME, 93421.00, EQ_DESC)]
            + [recode(before[lab - 1], LABOUR, LABOUR_NAME, 50628.00, LB_DESC)]
            + before[lab:])
    bad = [i + 1 for i, (a, b) in enumerate(zip(after, want)) if a != b]
    if bad or len(after) != len(want):
        print(f'  rows not as intended: {bad[:10]} '
              f'(after {len(after)}, want {len(want)})')
        return 1
    print(f'  {len(after)} rows checked, only the three intended lines differ')

    # ---- the move is the figure already agreed with Xero
    def total(rows, code):
        return round(sum((r[6] or 0) - (r[5] or 0)
                         for r in rows if r[0] == code), 2)
    for code, name in ((EQUIP, EQUIP_NAME), (LABOUR, LABOUR_NAME)):
        b, a = total(before, code), total(after, code)
        print(f'  {name:<20} {b:>14,.2f} -> {a:>14,.2f}  '
              f'move {a - b:>+13,.2f}')
    if round(total(after, EQUIP) - total(before, EQUIP), 2) != MOVE:
        print('  the move is not 92,601.72')
        return 1
    if round(total(before, LABOUR) - total(after, LABOUR), 2) != MOVE:
        print('  the labour side does not mirror it')
        return 1
    b2 = round(total(before, EQUIP) + total(before, LABOUR), 2)
    a2 = round(total(after, EQUIP) + total(after, LABOUR), 2)
    if b2 != a2:
        print(f'  the two together have changed: {b2} -> {a2}')
        return 1
    print(f'  the two together are unchanged at {a2:,.2f}')

    wb.save(PACK)
    print('  saved')
    return 0


if __name__ == '__main__':
    sys.exit(main())
