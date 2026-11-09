"""The Aware Super Seminar Room job on the GL paste, so it reaches the report.

patch_v89 put the cost and sell breakdown into PL_Paste, which is the Xero
P&L the pack holds as a control. The Summary and every reporting tab read
the GL paste on 'P&L Analysis' instead, so the figures on the report did
not move. This puts the same adjustment there.

Two parts.

The sell split. Services are labour and equipment is equipment, so the
breakdown's 161,969.52 equipment and 43,382.65 services go to Equipment
Sales and Installation Labour. The GL paste has 153,904.89 and 51,447.28,
so 8,064.63 moves off the Phase 2 labour line and onto the Phase 2
equipment line. Total income does not change.

The cost. Complete AV's total for the job is 184,475.80 and the paste
carries 161,926.60, because the second purchase order has not been billed.
Two lines are appended for the missing 22,549.20, 18,520.00 to Equipment
Purchase and 4,029.20 to Sub-Contract Labour, on the same date and job as
the invoice.

Nothing is posted in Xero. The pack is deliberately ahead of the ledger by
this amount, and the P&L control and the recode check say so.
"""
import sys
import openpyxl

PACK = 'pack/CTS Financial Controller Pack - v20.xlsx'
SHEET = 'P&L Analysis'
INV = 'INV-10596'
MOVE = 8064.63
ACCRUALS = (
    ('51100', 'Equipment - Purchase', 18520.00,
     'Aware Super - PO 2 of 2 not yet billed, accrued to the cost and '
     'sell breakdown of 9 Nov 2026'),
    ('51325', 'Sub-Contract Labour - INTEGRATION', 4029.20,
     'Aware Super - PO 2 of 2 not yet billed, accrued to the cost and '
     'sell breakdown of 9 Nov 2026'),
)


def main():
    wb = openpyxl.load_workbook(PACK)
    ws = wb[SHEET]
    last = max(r for r in range(ws.max_row, 1, -1)
               if ws.cell(r, 1).value not in (None, ''))

    # the two Phase 2 lines, found by what they are
    lab = eq = None
    for r in range(2, last + 1):
        if ws.cell(r, 9).value != INV:
            continue
        d = str(ws.cell(r, 11).value or '')
        if not d.startswith('Aware Super - Phase 2'):
            continue
        if str(ws.cell(r, 1).value) == '42300':
            lab = r
        elif str(ws.cell(r, 1).value) == '42600':
            eq = r
    if lab is None or eq is None:
        print('  the two Phase 2 lines are not as expected, nothing done')
        return 1
    if round(ws.cell(lab, 7).value, 2) != 50628.00 or \
       round(ws.cell(eq, 7).value, 2) != 93421.00:
        print(f'  labour reads {ws.cell(lab, 7).value}, '
              f'equipment reads {ws.cell(eq, 7).value}, not as expected')
        return 1
    ws.cell(lab, 7).value = round(50628.00 - MOVE, 2)
    ws.cell(eq, 7).value = round(93421.00 + MOVE, 2)
    print(f'  r{lab} 42300 Installation Labour  50,628.00 -> '
          f'{ws.cell(lab, 7).value:,.2f}')
    print(f'  r{eq} 42600 Equipment Sales      93,421.00 -> '
          f'{ws.cell(eq, 7).value:,.2f}')

    # the invoice lines give the date, job and cost centre for the accrual
    src = eq
    date, job, cc = (ws.cell(src, 4).value, ws.cell(src, 8).value,
                     ws.cell(src, 12).value)
    r = last + 1
    for code, name, amt, desc in ACCRUALS:
        ws.cell(r, 1).value = code
        ws.cell(r, 2).value = name
        ws.cell(r, 3).value = 'Manual Journal'
        ws.cell(r, 4).value = date
        ws.cell(r, 5).value = 'Complete AV Solutions Pty Ltd'
        ws.cell(r, 6).value = amt            # debit
        ws.cell(r, 7).value = 0
        ws.cell(r, 8).value = job
        ws.cell(r, 10).value = 'Pack adjustment'
        ws.cell(r, 11).value = desc
        ws.cell(r, 12).value = cc
        print(f'  r{r} {code} {name:<34} dr {amt:>11,.2f}  appended')
        r += 1

    total = sum(a[2] for a in ACCRUALS)
    print(f'  cost accrued {total:,.2f}, so September profit falls by that')
    wb.save(PACK)
    print('  saved')
    return 0


if __name__ == '__main__':
    sys.exit(main())
