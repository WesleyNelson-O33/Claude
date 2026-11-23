"""The November 2026 payroll tax return, laid out the way Revenue NSW shows it.

October's return is the template, so the layout, the fonts and the order of
the fields are hers and not mine. Only the figures, the period, the due date
and the note at the bottom change.

Every figure the portal works out for itself stays a formula, so the four
lines at the bottom can be read straight off the page and checked against the
screen before submitting.

The figures come from the Main & NSW tab of the November payroll tax workbook,
which is built from the Employment Hero split file. The comparative column is
October as it was lodged.

    ptax/build_nov_return.py <October return.xlsx>
"""
import sys

import openpyxl

OUT = 'ptax/CTS payroll tax return - November 2026.xlsx'

# Main & NSW column F, November 2026
NSW = {
    'Salaries and wages': 178154.87,
    'Fringe benefits': 0,
    'Employer superannuation contribution': 23337.35,
    'Termination payments': 875.00,
    'Contractor payments': 0,
    'Employment agency payments': 0,
    'Allowances': 0,
    'Bonuses/Commissions': 0,
    'Directors fees': 0,
    'Shares and options': 0,
    'NSW apprentice and trainee wages': 0,
}
INTERSTATE = [('QLD', 14222.05), ('SA', 6289.93), ('VIC', 22537.99),
              ('WA', 11707.87), ('ACT', 0)]

# October 2026 as lodged, for the movement column
OCT = {17: 169021.21, 18: 0, 19: 22087.22, 20: 0, 21: 0, 22: 0, 23: 0,
       24: 0, 25: 0, 26: 0, 27: 0, 28: 191108.43, 29: 49139.64,
       30: 240248.07, 31: 81071.84, 32: 0, 33: 0, 34: 5996.99}

DAYS = 30                      # days in November

WHY = [
    'The Employment Hero payroll tax report gives 201,849.72 of New South '
    'Wales wages. This return uses 202,367.22, which is 517.50 more.',
    'One employee, Nicol Schildknecht, is New South Wales in the split file '
    'and Victoria on the report, which groups by residential address. '
    'That is 462.05 of wages and 55.45 of super.',
    'The two sources agree on total Australian wages of 257,125.06, so this '
    'is a state split, not a missing pay run. Last month the difference was '
    'a missing pay run.',
    'Per diems of 250.00 are exempt and are still in the New South Wales '
    'gross on the Main & NSW tab. Taking them out would reduce the tax to '
    '6,785.90.',
    'The full working is on the Nov 2026 derivation tab of the payroll tax '
    'workbook.',
]


def main():
    if len(sys.argv) != 2:
        print(__doc__)
        return 1
    wb = openpyxl.load_workbook(sys.argv[1])
    ws = wb[wb.sheetnames[0]]
    if ws['A1'].value != 'Payroll tax return, October 2026':
        print('STOPPED: %r is not the October return' % ws['A1'].value)
        return 1
    ws.title = 'November 2026 return'

    ws['A1'] = 'Payroll tax return, November 2026'
    ws['B6'] = '1 Nov 2026 to 30 Nov 2026'
    ws['B8'] = '7 Dec 2026'
    ws['B16'] = 'November 2026'
    ws['D16'] = 'October, as it was lodged'
    ws['B37'] = 'November 2026'

    for i, (label, value) in enumerate(NSW.items()):
        row = 17 + i
        if str(ws.cell(row, 1).value).strip() != label:
            print('STOPPED: row %d reads %r, expected %r'
                  % (row, ws.cell(row, 1).value, label))
            return 1
        ws.cell(row, 2).value = value
        ws.cell(row, 3).value = 'Type this in' if value else 'Leave at nil'

    for i, (state, value) in enumerate(INTERSTATE):
        row = 38 + i
        if str(ws.cell(row, 1).value).strip() != state:
            print('STOPPED: row %d reads %r, expected %r'
                  % (row, ws.cell(row, 1).value, state))
            return 1
        ws.cell(row, 2).value = value
        ws.cell(row, 3).value = 'nil this month' if not value else None

    # the threshold and the tax carry the month's day count
    ws['B31'] = '=ROUND((B28/B30)*(1200000/365*%d),2)' % DAYS
    ws['B34'] = '=ROUND((B30-B31-B29)*0.0545,2)'

    for row, value in OCT.items():
        ws.cell(row, 4).value = value

    for i, line in enumerate(WHY):
        ws.cell(52 + i, 1).value = line
    for row in range(52 + len(WHY), 57):
        ws.cell(row, 1).value = None

    wb.save(OUT)
    print('  written to %s' % OUT)

    # read it back and show the four lines that get checked on screen
    v = openpyxl.load_workbook(OUT, data_only=False)[ws.title]
    print('  Salaries and wages          %12.2f' % NSW['Salaries and wages'])
    print('  Employer super              %12.2f'
          % NSW['Employer superannuation contribution'])
    print('  Termination payments        %12.2f' % NSW['Termination payments'])
    nsw_total = sum(NSW.values())
    inter = sum(v for _s, v in INTERSTATE)
    total = nsw_total + inter
    thr = round((nsw_total / total) * (1200000 / 365 * DAYS), 2)
    tax = round((total - thr - inter) * 0.0545, 2)
    print('  Total NSW wages             %12.2f' % nsw_total)
    print('  Interstate wages            %12.2f' % inter)
    print('  Total Australian wages      %12.2f' % total)
    print('  Threshold entitlement       %12.2f' % thr)
    print('  TAX PAYABLE                 %12.2f' % tax)
    return 0


if __name__ == '__main__':
    sys.exit(main())
