"""patch_v91 reversed out. The breakdown's cost is Complete AV's only.

patch_v91 cut the sub-contract accrual from 4,029.20 to 1,962.91 on the
reading that the breakdown's services cost of 41,320.05 covered our own
labour as well. It does not. Every line behind that figure is a Complete
AV line: labour and programming 23,320.00, labour 7,196.00 and 6,561.80,
the return trip 1,380.00, commissioning 1,440.00, lighting programming
1,064.00, and the freight and waste lines. Not one is CTS labour.

So the full 4,029.20 is needed to bring Complete AV's services cost to
41,320.05, and our own labour, super and payroll tax of 2,066.29 sits on
the job on top of it.

The job then costs 186,542.09 and makes 18,810.08, which is 9.2 per cent.
The breakdown's 10.2 per cent is the margin on supplier cost alone.
"""
import sys
import openpyxl

PACK = 'pack/CTS Financial Controller Pack - v20.xlsx'
WAS, NOW = 1962.91, 4029.20
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
    print(f'  P&L Analysis r{r}  {WAS:,.2f} -> {NOW:,.2f}')

    ws = wb['PL_Paste']
    hit = [r for r in range(6, ws.max_row + 1)
           if str(ws.cell(r, 1).value).strip()
           == 'Sub-Contract Labour - INTEGRATION']
    r = hit[0]
    was = ws.cell(r, 2).value
    if round(was, 2) != 47453.21:
        print(f'  PL_Paste r{r} reads {was:,.2f}, expected 47,453.21')
        return 1
    ws.cell(r, 2).value = 49519.50
    print(f'  PL_Paste     r{r}  {was:,.2f} -> 49,519.50')

    print('  the job costs 186,542.09 and makes 18,810.08, 9.2 per cent')
    wb.save(PACK)
    print('  saved')
    return 0


if __name__ == '__main__':
    sys.exit(main())
