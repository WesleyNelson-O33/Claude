"""
Section 6 of Overhead Split: the office staff split now runs on FTE, worked out
the way the recording does it.

Gi at 1:31:00: the employment expenses cell is the grey one, "this one here
changes every single month based on this here... that's labour hours. So update
based on utilisation." At 1:37:32 to 1:38:28 she gives the FTE method: take the
department total plus its Other CC total, filtered over twelve months, and
divide by the working hours in a year, which is 253 working days by eight hours
for FY27.

So it is utilisation, and FTE is how utilisation gets turned into a share. Her
ALL DEPT labour report is that twelve month window to July, which is why its
Support total is 40,887.94 and not a single month.

Two paste rows per department are gone. Instead the tab now holds the rolling
twelve month hours by cost centre, which is what the Labour Hours Report Grand
Total rows give her, and the annual hours divisor. FTE and the share follow.

Seeded with the windows to July, August and September off the Raw data tab of
her Utilisation Report.
"""
import openpyxl, sys

WB = sys.argv[1] if len(sys.argv) > 1 else 'CTS Financial Controller Pack - v20.xlsx'

ROWS = (113, 114, 115, 116, 117)        # Support, Events, Video, Integration, Consulting
COLS = [chr(ord('C') + i) for i in range(12)]

# Rolling twelve month hours by cost centre, per the Labour Hours Report.
HOURS = {
    'C': (40892.44, 15107.65, 1293.33, 181.42, 3713.74),    # window to July 2026
    'D': (40890.31, 14989.65, 1251.08, 248.08, 3550.43),    # window to August 2026
    'E': (37779.06, 13544.65, 1217.33, 254.08, 3513.93),    # window to September 2026
}
ANNUAL = 253 * 8                        # 2,024 hours, FY27, off the Days tab

wb = openpyxl.load_workbook(WB)
ws = wb['Overhead Split']

ws['A111'] = ("Allocated on each department's FTE. Rolling twelve month hours on that cost centre, "
              "the Grand Total row of the Labour Hours Report, divided by the working hours in a year.")
ws['A112'] = 'Labour hours, rolling twelve months'
ws['A118'] = '    Total hours'
ws['A119'] = f'    Working hours in a year ({253} days by 8)'

carry = HOURS['E']
for col in COLS:
    for row, value in zip(ROWS, HOURS.get(col, carry)):
        cell = ws[f'{col}{row}']
        cell.value = value
        cell.number_format = '#,##0.00'
    ws[f'{col}118'].number_format = '#,##0.00'
    ws[f'{col}119'] = ANNUAL
    ws[f'{col}119'].number_format = '#,##0'

# Share of FTE is the same ratio as share of hours, so rows 121-125 stand, but
# show the FTE itself so the number on the page matches the one she quotes.
ws['A120'] = 'FTE and share of FTE'
for col in COLS:
    for row in ROWS:
        ws[f'{col}{row + 8}'] = f'=IFERROR({col}{row}/{col}$118,0)'
        ws[f'{col}{row + 8}'].number_format = '0.00%'

wb.save(WB)
print('section 6 now on FTE from rolling twelve month hours')
for col, lab in (('C', 'to Jul'), ('D', 'to Aug'), ('E', 'to Sep')):
    h = HOURS[col]
    print(f'  {lab:<8} FTE ' + '  '.join(f'{v / ANNUAL:0.2f}' for v in h) +
          f'   total {sum(h) / ANNUAL:0.2f}')
