"""The Aware Super accrual cut back to the breakdown's own services cost.

patch_v89 and patch_v90 accrued 22,549.20, which took the job's total cost
to 186,542.09 and the margin to 9.2 per cent. The breakdown's services cost
of 41,320.05 covers every service on the job, our own labour, super and
payroll tax included, not only what Complete AV charge. That 2,066.29 is
already on the job, so accruing on top of it double counted it.

The sub-contract labour accrual comes down from 4,029.20 to 1,962.91. Total
cost on the job is then 184,475.80 and the margin 20,876.37, which is the
10.2 per cent the breakdown states.

Both places are corrected: the appended line on the GL paste and the
September column on PL_Paste.
"""
import sys
import openpyxl

PACK = 'pack/CTS Financial Controller Pack - v20.xlsx'
WAS, NOW = 4029.20, 1962.91
BACK = round(WAS - NOW, 2)
MARK = 'PO 2 of 2 not yet billed'


def main():
    wb = openpyxl.load_workbook(PACK)

    # the appended accrual line on the GL paste
    ws = wb['P&L Analysis']
    hit = [r for r in range(2, ws.max_row + 1)
           if str(ws.cell(r, 1).value) == '51325'
           and MARK in str(ws.cell(r, 11).value or '')]
    if len(hit) != 1:
        print(f'  found {len(hit)} accrual lines on 51325, expected 1')
        return 1
    r = hit[0]
    if round(ws.cell(r, 6).value, 2) != WAS:
        print(f'  r{r} reads {ws.cell(r, 6).value}, expected {WAS}')
        return 1
    ws.cell(r, 6).value = NOW
    print(f'  P&L Analysis r{r}  {WAS:,.2f} -> {NOW:,.2f}')

    # the same on the Xero P&L the pack holds as a control
    ws = wb['PL_Paste']
    if str(ws.cell(5, 2).value).strip() != 'Sept 2026':
        print('  PL_Paste column 2 is not Sept 2026, nothing done')
        return 1
    hit = [r for r in range(6, ws.max_row + 1)
           if str(ws.cell(r, 1).value).strip()
           == 'Sub-Contract Labour - INTEGRATION']
    if len(hit) != 1:
        print(f'  found {len(hit)} sub-contract rows, expected 1')
        return 1
    r = hit[0]
    was = ws.cell(r, 2).value
    if round(was, 2) != 49519.50:
        print(f'  PL_Paste r{r} reads {was:,.2f}, expected 49,519.50')
        return 1
    ws.cell(r, 2).value = round(was - BACK, 2)
    print(f'  PL_Paste     r{r}  {was:,.2f} -> {ws.cell(r, 2).value:,.2f}')

    print(f'  cost comes back {BACK:,.2f}, so September profit rises by that')
    print(f'  the job then costs 184,475.80 and makes 20,876.37, 10.2 per cent')
    wb.save(PACK)
    print('  saved')
    return 0


if __name__ == '__main__':
    sys.exit(main())
