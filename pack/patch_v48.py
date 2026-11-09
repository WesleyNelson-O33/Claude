"""
The hours by department chart could not be read.

Three faults. White label text sat on the bright teal and the pale grey
segments, where white is invisible. The segments with no hours printed a
"0.00" that ran off the end of the bar. And nothing on the picture said what
the colours were, so the reader could not tell billable hours from
non-billable.

This script does the two parts that live in the sheet: plainer column
headings, and a number format that leaves a zero blank instead of printing
it. The label colours and the key are set in build_charts.py.
"""
import openpyxl, sys

WB = sys.argv[1] if len(sys.argv) > 1 else 'CTS Financial Controller Pack - v20.xlsx'
HEAD = {'B67': 'Billable hours', 'C67': 'Non-billable hours', 'D67': 'Leave'}
FMT = '#,##0.00;;'

wb = openpyxl.load_workbook(WB)
u = wb['Utilisation']
for ref, label in HEAD.items():
    u[ref] = label
for r in range(68, 74):
    for col in range(2, 5):
        u.cell(r, col).number_format = FMT
wb.save(WB)
print('Utilisation hours block: headings in plain English, zeros left blank')
