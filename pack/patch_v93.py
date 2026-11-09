"""Services is labour, so 41,320.05 is all the labour on the job.

Danica's coding rule is that services are labour and equipment is
equipment. On that basis the breakdown's services cost of 41,320.05 is
every labour cost on the job, not only what Complete AV charge. Our own
labour, super and payroll tax of 2,066.29 is already there, so the
accrual needed on sub-contract labour is 1,962.91, not 4,029.20.

patch_v91 made this change, patch_v92 reversed it while I checked where
the 41,320.05 came from. Every line behind it is a Complete AV line, so
on the supplier's own basis the accrual would be 4,029.20. Danica has
decided it goes on her coding basis, which is what this applies.

Job 2604704 then costs 184,475.80 and makes 20,876.37, which is the 10.2
per cent the breakdown states.
"""
import sys
import openpyxl

PACK = 'pack/CTS Financial Controller Pack - v20.xlsx'
WAS, NOW = 4029.20, 1962.91
BACK = round(WAS - NOW, 2)
MARK = 'PO 2 of 2 not yet billed'


def main():
    wb = openpyxl.load_workbook(PACK)

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
    ws.cell(r, 11).value = (
        'Aware Super - PO 2 of 2 not yet billed, accrued to the cost and '
        'sell breakdown of 9 Nov 2026. Services is labour, so the '
        '41,320.05 covers our own labour on the job as well.')
    print(f'  P&L Analysis r{r}  {WAS:,.2f} -> {NOW:,.2f}')

    ws = wb['PL_Paste']
    if str(ws.cell(5, 2).value).strip() != 'Sept 2026':
        print('  PL_Paste column 2 is not Sept 2026, nothing done')
        return 1
    hit = [r for r in range(6, ws.max_row + 1)
           if str(ws.cell(r, 1).value).strip()
           == 'Sub-Contract Labour - INTEGRATION']
    r = hit[0]
    was = ws.cell(r, 2).value
    if round(was, 2) != 49519.50:
        print(f'  PL_Paste r{r} reads {was:,.2f}, expected 49,519.50')
        return 1
    ws.cell(r, 2).value = round(was - BACK, 2)
    print(f'  PL_Paste     r{r}  {was:,.2f} -> {ws.cell(r, 2).value:,.2f}')

    print(f'  cost comes back {BACK:,.2f}')
    print('  the job costs 184,475.80 and makes 20,876.37, 10.2 per cent')
    wb.save(PACK)
    print('  saved')
    return 0


if __name__ == '__main__':
    sys.exit(main())
