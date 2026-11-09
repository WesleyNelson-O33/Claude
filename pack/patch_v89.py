"""The Aware Super Seminar Room job put on the cost and sell breakdown.

Job 2604704, invoice INV-10596. The breakdown prepared on 9 November gives
the supplier cost and splits both the sell and the cost between equipment
and services. Two things follow from it.

The sell split moves. Services are labour and equipment is equipment, so
the breakdown's 161,969.52 equipment and 43,382.65 services go to Equipment
Sales and Installation Labour. The ledger has 153,904.89 and 51,447.28, so
8,064.63 moves from labour to equipment. Total income does not change.

The cost is short. Complete AV's total for the job is 184,475.80 and the
ledger carries 161,926.60 of it, because the second purchase order has not
been billed yet. The missing 22,549.20 is accrued, 18,520.00 to equipment
and 4,029.20 to sub-contract labour.

This is a manual adjustment in the pack only. Nothing has been posted in
Xero.
"""
import sys
import openpyxl

PACK = 'pack/CTS Financial Controller Pack - v20.xlsx'
SHEET = 'PL_Paste'
MONTH_COL = 2            # 'Sept 2026'

# label on PL_Paste, what it reads now, what it should read
MOVES = (
    ('Installation Labour',               52217.13,  44152.50),
    ('Equipment Sales',                  166131.16, 174195.79),
    ('Equipment - Purchase',             142536.76, 161056.76),
    ('Sub-Contract Labour - INTEGRATION',  45490.30,  49519.50),
)
INCOME = ('Installation Labour', 'Equipment Sales')


def main():
    wb = openpyxl.load_workbook(PACK)
    ws = wb[SHEET]

    if str(ws.cell(5, MONTH_COL).value).strip() != 'Sept 2026':
        print(f'  column {MONTH_COL} is not Sept 2026, nothing done')
        return 1

    inc = cost = 0.0
    for label, expect, now in MOVES:
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
        if label in INCOME:
            inc += now - was
        else:
            cost += now - was
        print(f'  {label:<34} r{r}  {was:>13,.2f} -> {now:>13,.2f}')

    if round(inc, 2) != 0.0:
        print(f'  the two income lines do not net to nil, {inc:,.2f}')
        return 1
    print(f'  income lines net to nil, so total income is unchanged')
    print(f'  cost goes up {cost:,.2f}, so September profit falls by that')

    wb.save(PACK)
    print('  saved')
    return 0


if __name__ == '__main__':
    sys.exit(main())
