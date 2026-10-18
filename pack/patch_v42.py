"""
Make the office staff split calculate instead of being typed.

The FTE split needs worked hours by cost centre over a rolling twelve months.
The Utilisation tab only held hours by the person's own department, and only for
FY27, so the rolling window had nowhere to come from and section 6 was carrying
typed constants.

This adds one paste block to the Utilisation tab: worked hours by cost centre,
a prior year strip and a current year strip, both on the same month columns D to
O as everything else above it. Five numbers a month, the Grand Total row off the
three Labour Hours Reports. The prior year strip is a once a year paste.

Section 6 of Overhead Split then builds the rolling twelve months itself, so
nothing in the chain is typed any more.

Seeded from the Raw data tab of the August Utilisation Report for July 2025 to
August 2026, and from the September Labour Hours Reports for September 2026,
because the August file only holds part of September.
"""
import openpyxl, sys
from openpyxl.styles import Font, Alignment

WB = sys.argv[1] if len(sys.argv) > 1 else 'CTS Financial Controller Pack - v20.xlsx'

DEPTS = ['Support', 'Events', 'Video', 'Integration', 'Consulting']
PRIOR_ROWS = (50, 51, 52, 53, 54)      # Jul 2025 to Jun 2026
CURR_ROWS = (58, 59, 60, 61, 62)       # Jul 2026 to Jun 2027
MCOLS = [chr(ord('D') + i) for i in range(12)]

PRIOR = [            # Support, Events, Video, Integration, Consulting
    (3707.00, 1316.75,  41.25,  0.00, 327.25),   # Jul 2025
    (3429.75, 1353.00, 127.75,  2.25, 354.56),
    (3730.75, 1650.25,  47.75,  1.25,  67.75),
    (3587.50, 1918.00, 126.00,  0.00, 391.00),
    (3392.50, 1575.25, 167.00,  0.50, 346.50),
    (3198.25, 1177.00,  18.25,  0.25, 349.25),
    (3027.66,  886.50,  85.00, 14.75, 317.50),
    (3168.25, 1239.90, 201.00, 26.50, 301.00),
    (3672.00, 1074.50, 103.25, 33.50, 330.50),
    (3445.45,  882.75, 116.75,  5.25, 292.25),
    (3520.05, 1023.08, 131.50,  6.75, 356.33),
    (3200.31, 1056.83, 108.08, 11.42, 374.70),   # Jun 2026
]
CURR = [
    (3519.98, 1270.58, 61.00, 79.00, 232.40),    # Jul 2026
    (3427.62, 1235.00, 85.50, 68.92, 191.25),    # Aug 2026
    (3151.65, 1249.00, 37.25,  9.50, 286.25),    # Sep 2026
]

BOLD = Font(bold=True, size=10)
wb = openpyxl.load_workbook(WB)
u = wb['Utilisation']

u['A47'] = 'Worked hours by cost centre, for the office staff split'
u['A47'].font = BOLD
u['A48'] = ("The Grand Total row of each Labour Hours Report, which is chargeable plus "
            "non-chargeable plus leave on that cost centre. Five numbers a month.")

u['A49'] = 'Prior year (paste once a year)'
u['A49'].font = BOLD
u['A57'] = 'This year (paste each month)'
u['A57'].font = BOLD

# Month headers. The prior year strip is the same twelve months a year earlier.
for i, col in enumerate(MCOLS):
    u[f'{col}49'] = f'=EDATE({col}$4,-12)'
    u[f'{col}57'] = f'={col}$4'
    for row in (49, 57):
        u[f'{col}{row}'].number_format = 'mmm yy'
        u[f'{col}{row}'].alignment = Alignment(horizontal='center')

for rows, label in ((PRIOR_ROWS, 'prior'), (CURR_ROWS, 'curr')):
    for row, dept in zip(rows, DEPTS):
        u.cell(row, 2).value = dept
        u.cell(row, 3).value = 'Worked hours'

for i, col in enumerate(MCOLS):
    for row, value in zip(PRIOR_ROWS, PRIOR[i]):
        u[f'{col}{row}'] = value
        u[f'{col}{row}'].number_format = '#,##0.00'
    for row in CURR_ROWS:
        u[f'{col}{row}'].number_format = '#,##0.00'
for i, month in enumerate(CURR):
    for row, value in zip(CURR_ROWS, month):
        u[f'{MCOLS[i]}{row}'] = value

u['A64'] = ('Worked hours here are by cost centre, not by the person\'s own department, so they '
            'do not match the blocks above. That is deliberate: the split follows the cost centre.')

# ---- Overhead Split section 6: build the rolling twelve months from that block ----
o = wb['Overhead Split']
o['A111'] = ("Allocated on each department's FTE. Worked hours on that cost centre over the "
             "rolling twelve months, off the Utilisation tab, divided by the working hours in a year.")
o['A112'] = 'Labour hours, rolling twelve months'

SRC = {113: (58, 50), 114: (59, 51), 115: (60, 52), 116: (61, 53), 117: (62, 54)}
for i in range(12):                      # i = 0 for July, 11 for June
    col = chr(ord('C') + i)              # Overhead Split month columns C..N
    last = MCOLS[i]                      # last current-year column in the window
    for row, (cur, pri) in SRC.items():
        part = f"SUM(Utilisation!$D{cur}:{last}{cur})"
        if i < 11:                       # months still needing the tail of last year
            first_prior = MCOLS[i + 1]
            part += f"+SUM(Utilisation!{first_prior}{pri}:$O{pri})"
        o[f'{col}{row}'] = '=' + part
        o[f'{col}{row}'].number_format = '#,##0.00'

wb.save(WB)
print('Utilisation: worked hours by cost centre added, 24 month columns')
print('Overhead Split: rolling twelve months now a formula, no typed hours left')
