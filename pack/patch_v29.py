"""The Summary's three months roll with the dropdown.

They used to be the calendar quarter the chosen month falls in, so picking
November gave October, November and December, and December was a forecast.
Now they are the three months ending on the month she picks: November gives
September, October and November, all actual.

Where the window reaches back past the start of the financial year it reads
Actual FY26, which is where those months live, so July and August work too.

The Quarterly column further along is left alone. That one is quarter to date
and is meant to be a quarter.
"""
import os, re
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = f'{HERE}/CTS Financial Controller Pack - v20.xlsx'
wb = openpyxl.load_workbook(OUT)
s = wb['Summary']

FYSTART = 'DATE(Setup!$B$7-1,Setup!$B$6,1)'
for j, col in enumerate(('B', 'C', 'D')):
    s[f'{col}5'] = f'=EDATE(Setup!$B$5,{j - 2})'
    s[f'{col}5'].number_format = 'mmm\\-yy'
print('month headers now roll:', [s[f'{c}5'].value for c in 'BCD'])

OLD = re.compile(r"^=SUMIF\('P&L FY27'!\$B\$5:\$M\$5,([BCD])\$5,"
                 r"'P&L FY27'!\$B(\d+):\$M\d+\)$")
n = 0
for r in range(6, s.max_row + 1):
    for col in (2, 3, 4):
        v = s.cell(r, col).value
        if not isinstance(v, str):
            continue
        m = OLD.match(v)
        if not m:
            continue
        c, row = m.group(1), m.group(2)
        s.cell(r, col, f'=IF({c}$5>={FYSTART},'
                       f"SUMIF('P&L FY27'!$B$5:$M$5,{c}$5,'P&L FY27'!$B{row}:$M{row}),"
                       f"SUMIF('Actual FY26'!$B$5:$M$5,EDATE({c}$5,12),"
                       f"'Actual FY26'!$B{row}:$M{row}))")
        n += 1
print('month cells repointed:', n)

if 'Rolling 3 Months' in wb.sheetnames:
    del wb['Rolling 3 Months']
    print('Rolling 3 Months removed, the Summary does the job now')

wb.save(OUT)
print('saved')
