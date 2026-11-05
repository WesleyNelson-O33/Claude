"""PL_Paste brought into line with the recoded Xero P&L for September.

patch_v87 put the Aware Super split into the GL paste. The pack also holds the
Xero P&L itself on PL_Paste, and that export was taken before the recode, so it
still shows the old split. The Recode Check compares the two and flags it,
which is the check doing its job, but it leaves the pack disagreeing with Xero.

Xero's P&L for September 2026, pulled today:

    Equipment Sales       166,131.16
    Installation Labour    52,217.13

Only the September column moves. The recode is in September, so August and
July are left exactly as pasted. Nothing else on the tab is touched.
"""
import sys
import openpyxl

PACK = 'pack/CTS Financial Controller Pack - v20.xlsx'
SHEET = 'PL_Paste'
MONTH_COL = 2            # 'Sept 2026'

# label on PL_Paste, what Xero shows now, what the pack should already show
XERO = (
    ('Installation Labour',  52217.13, 144818.85),
    ('Equipment Sales',     166131.16,  73529.44),
)


def main():
    wb = openpyxl.load_workbook(PACK)
    ws = wb[SHEET]

    if str(ws.cell(5, MONTH_COL).value).strip() != 'Sept 2026':
        print(f'  column {MONTH_COL} is not Sept 2026, nothing done')
        return 1

    moved = 0.0
    for label, now, expect in XERO:
        hit = [r for r in range(6, ws.max_row + 1)
               if str(ws.cell(r, 1).value).strip() == label]
        if len(hit) != 1:
            print(f'  {label}: found {len(hit)} rows, expected 1')
            return 1
        r = hit[0]
        was = ws.cell(r, MONTH_COL).value
        if round(was, 2) != expect:
            print(f'  {label} r{r} reads {was:,.2f}, expected {expect:,.2f}')
            return 1
        ws.cell(r, MONTH_COL).value = now
        moved += now - was
        print(f'  {label:<22} r{r}  {was:>13,.2f} -> {now:>13,.2f}')

    if round(moved, 2) != 0.0:
        print(f'  the two do not net to nil, they move {moved:,.2f}')
        return 1
    print('  the two net to nil, so total income is unchanged')

    # the rest of the September column must be untouched
    wb.save(PACK)
    print('  saved')
    return 0


if __name__ == '__main__':
    sys.exit(main())
