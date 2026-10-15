"""Type the utilisation hours onto the Utilisation tab.

Straight off the Person Group total rows of her own utilisation reports.
Nothing is derived, nothing is apportioned, nothing is estimated.

  August    - August 2026 Utilisation Report, August sheet, rows 11, 13, 30
  September - September 2026 Utilisation Report, September sheet, rows 11, 13, 28

Onsite on this tab is the Support group in her report. Video, Integration and
CTS are left empty: her reports do not split those out, and the three labour
hour reports are PRD, CONS and ONS only.
"""
import os, shutil
import openpyxl
from openpyxl.styles import Font, Alignment

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = ('/tmp/claude-0/-home-user-Claude/a00ffe59-433c-507b-9006-917d8c60083a/'
       'scratchpad/v83/her11.xlsx')
OUT = f'{HERE}/CTS Financial Controller Pack - v12.xlsx'
MID = '5F6670'

ROWS = {'Onsite': (9, 10, 11), 'Production': (15, 16, 17), 'Consulting': (33, 34, 35)}

#                        chargeable, non-chargeable, leave
HOURS = {
    'E': {                                              # August 2026
        'Production': (429.5,      316.25,    98.25),
        'Consulting': (73.91666,   94.33333,   0.0),
        'Onsite':     (2269.33333, 359.28203, 157.5),   # her Support total
    },
    'F': {                                              # September 2026
        'Production': (266.0,      283.25,   332.75),
        'Consulting': (40.25,       81.5,     53.0),
        'Onsite':     (2131.75,      0.0,    218.0),    # her Support total
    },
}

shutil.copy(SRC, OUT)
wb = openpyxl.load_workbook(OUT)
ws = wb['Utilisation']

for dept, (rc, rn, rl) in ROWS.items():
    assert str(ws.cell(rc, 2).value).strip() == dept, (dept, ws.cell(rc, 2).value)
    assert str(ws.cell(rc, 3).value).strip() == 'Chargeable', dept
    assert str(ws.cell(rn, 3).value).strip() == 'Non-chargeable', dept
    assert str(ws.cell(rl, 3).value).strip() == 'Leave', dept

n = 0
for col, by_dept in HOURS.items():
    for dept, (ch, nc, lv) in by_dept.items():
        rc, rn, rl = ROWS[dept]
        for row, val in ((rc, ch), (rn, nc), (rl, lv)):
            ws[f'{col}{row}'] = val
            n += 1
print('hour cells typed in:', n)

ws['A45'] = ('Hours are typed in from the Person Group total rows of the monthly '
             'utilisation report. Video, Integration and CTS are empty because '
             'those reports do not split them out.')
ws['A45'].font = Font(name='Calibri', sz=9, italic=True, color=MID)
ws['A45'].alignment = Alignment(horizontal='left', vertical='center')

wb.save(OUT)
print('saved', OUT)
