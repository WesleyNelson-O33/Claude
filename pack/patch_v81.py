"""Match the departmental report Danica actually presents.

The analysis holds four departmental views and they do not agree with each
other.  '4. PL Per Department' uses a Staff key of 39/24/3/1/0 over 67 and has
an off by one fault in three of its five columns.  '5. PL Per Department Excl
OI', '6. YTD PL Per Department' and the ONS, PRD and CONS department pages all
use 49/18/3/1/1 over 72, with no fault.  The ONS page is the one she presents,
so tab 5 and the 72 base are the source.

Three things follow.

The headcount goes back to 49/18/3/1/1.  The Legal Fees row picks up the GL
account, which is named plainly "Legal Fees", so the 800 a month stops falling
into the catch all and spreads on the Staff key as tab 5 has it.  And other
income comes out of the overhead spread, because tab 5 and the department pages
exclude it, leaving Interest Received and Income Other to sit on their own line.

August then ties to tab 5 for all five departments, the worst being 5.35 out on
Support.
"""
import openpyxl

WB = 'pack/CTS Financial Controller Pack - v20.xlsx'
HEADCOUNT = {'Support': 49, 'Events': 18, 'Video': 3, 'Integration': 1,
             'Consulting': 1, 'Admin': 7}
SUB = '$U$23:$U$86'
NOT_OI = f'({SUB}<>"Other Income")*({SUB}<>"Other Expenses")'

wb = openpyxl.load_workbook(WB)
ws = wb['Overhead Split']

# 1. the headcount behind the Staff key
assert ws['A5'].value == 'Department' and ws['B5'].value == 'Headcount'
for r in range(6, 12):
    name = ws.cell(r, 1).value
    print(f'  headcount {name:<12}{ws.cell(r, 2).value} -> {HEADCOUNT[name]}')
    ws.cell(r, 2, HEADCOUNT[name])

# 2. the legal fees row, which never matched the GL account name
LEGAL = 59
assert ws.cell(LEGAL, 1).value == 'Legal Fees - Other (Employsure)'
assert ws.cell(LEGAL, 2).value == 'Staff'
for c in range(3, 15):
    ws.cell(LEGAL, c).value = ws.cell(LEGAL, c).value.replace(f'$A{LEGAL}', '"Legal Fees"')
print(f'  row {LEGAL} now reads the GL account "Legal Fees"')

# 3. the spread, and the pot it is checked against, both without other income
for i, r in enumerate(range(90, 95)):
    key = openpyxl.utils.get_column_letter(16 + i)
    for c in range(3, 15):
        col = openpyxl.utils.get_column_letter(c)
        ws.cell(r, c).value = (f'=SUMPRODUCT({NOT_OI}*{col}$23:{col}$86'
                               f'*${key}$23:${key}$86)')
for c in range(3, 15):
    col = openpyxl.utils.get_column_letter(c)
    ws.cell(96, c).value = f'=SUMPRODUCT({NOT_OI}*{col}$23:{col}$86)'
print('  rows 90 to 94 and row 96 now leave other income out')

# the result table totals to the company figure before other income
chk = ws.cell(150, 2)
assert "'P&L FY27'!$B180:$M180" in chk.value, chk.value
chk.value = chk.value.replace("'P&L FY27'!$B180:$M180", "'P&L FY27'!$B178:$M178")
print('  the result check now reads the company EBIT, which is before other income')

wb.save(WB)
print('saved')
