"""Paste a Xero-shaped P&L export over the top and check the control reconciles."""
import json, subprocess, sys, datetime as dt
import openpyxl
HERE = __file__.rsplit('/', 1)[0]
TST = f'{HERE}/_real.xlsx'
wb = openpyxl.load_workbook(TST)
pp = wb['PL_Paste']
# the shape Xero gives you: title block, then sections, then a Total row each
LAYOUT = [
    ('Profit and Loss', None), ('Corporate Technology Services Pty Ltd', None),
    ('For the month ended 31 August 2026', None), (None, None),
    ('Account', 'Aug-26'),
    ('Income', None), ('Contract Support Staff', 208213.93),
    ('Equipment Hires', 100557.14), ('Production Labour', 80676.50),
    ('Total Income', 603466.54), (None, None),
    ('Cost of Sales', None), ('Direct Salaries - ONS', 92093.64),
    ('Equipment - Purchase', 51937.76),
    ('Total Cost of Sales', 335516.00), ('Gross Profit', 267950.54), (None, None),
    ('Other Income', None), ('Interest Received', 820.38), ('Income Other', 40.97),
    ('Total Other Income', 861.35), (None, None),
    ('Operating Expenses', None), ('Indirect Salaries - ADMIN', 46342.91),
    ('Rent', 20681.16),
    ('Total Operating Expenses', 208876.19), (None, None),
    ('Net Profit', 59935.70),
]
for i, (lab, val) in enumerate(LAYOUT, start=1):
    if lab is not None: pp.cell(i, 1, lab)
    if val is not None: pp.cell(i, 2, val)
wb.save(TST)
j = json.loads(subprocess.run([sys.executable, '/mnt/skills/public/xlsx/scripts/recalc.py',
                               TST, '1800'], capture_output=True, text=True).stdout)
print('recalc:', j['status'], j['total_errors'], 'errors')
v = openpyxl.load_workbook(TST, data_only=True)
pc = v['PL_Check']
print()
print(f'{"":<46}{"per the GL":>14}{"per the Xero P&L":>18}{"":>4}{"status":<12}')
for r in range(5, 10):
    lab = str(pc.cell(r, 1).value or '')[:44]
    print(f'   {lab:<44}{(pc.cell(r,2).value or 0):>14,.2f}'
          f'{(pc.cell(r,9).value if pc.cell(r,9).value is not None else 0):>18,.2f}'
          f'    {pc.cell(r,11).value}   found={pc.cell(r,10).value}')
print(f'   {"Net profit":<44}{(pc.cell(10,2).value or 0):>14,.2f}'
      f'{(pc.cell(10,9).value or 0):>18,.2f}    {pc.cell(10,11).value}')
print()
print('   Setup B29, the column it read:', v['Setup']['B29'].value,
      '| P&L lines pasted:', v['Setup']['B34'].value)
