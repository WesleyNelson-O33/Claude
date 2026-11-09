"""
Take the rest of the narration out of the workbook.

patch_v45 cleared the explanatory paragraphs. This removes what was left:

  the README sheet, which is nothing but narration and is referenced by nothing
  the helper notes still sitting beside typed cells on Setup and DSO
  headings and row labels that carried a clause, shortened to a plain label

Every edit is listed explicitly rather than matched by a rule, because what is
left is a short list and a wrong guess here deletes a working label.
"""
import openpyxl, sys

WB = sys.argv[1] if len(sys.argv) > 1 else 'CTS Financial Controller Pack - v20.xlsx'

CLEAR = [                       # helper notes beside typed cells
    ('Setup', 'C11'), ('DSO', 'C6'), ('DSO', 'C10'), ('DSO', 'C11'), ('DSO', 'C12'),
]
RENAME = {                      # clause in the label, cut back to the label
    ('Setup', 'A4'): '1.  REPORTING MONTH',
    ('Setup', 'A27'): '4.  PASTED REPORTS AND THE EBITDA BRIDGE',
    ('Setup', 'A37'): 'In General Admin, no cost centre',
    ('Budget FY27', 'A195'): 'Expenses by department',
    ('PL_Check', 'A9'): 'of which Other Expenses',
    ('Utilisation', 'A47'): 'WORKED HOURS BY COST CENTRE',
    ('Overhead Split', 'A4'): '1.  HEADCOUNT, WHICH DRIVES THE STAFF KEY',
    ('Overhead Split', 'A20'): '3.  ADMIN OVERHEAD BY ACCOUNT AND KEY',
    ('Overhead Split', 'A99'): "5.  EACH DEPARTMENT'S OWN OVERHEAD",
    ('Overhead Split', 'A119'): 'Working hours in a year',
    ('Overheads MoM', 'A88'): 'CHART DATA',
    ('Overheads MoM', 'A153'): 'The admin pot',
    ('Cost Analysis', 'A79'): 'Employment Expenses, not labour',
}

wb = openpyxl.load_workbook(WB)

if 'README' in wb.sheetnames:
    del wb['README']
    print('README sheet removed')

for sheet, cell in CLEAR:
    if sheet in wb.sheetnames and wb[sheet][cell].value is not None:
        print(f'  cleared  {sheet:<16} {cell:<6} {str(wb[sheet][cell].value)[:70]}')
        wb[sheet][cell].value = None

for (sheet, cell), new in RENAME.items():
    if sheet not in wb.sheetnames:
        continue
    old = wb[sheet][cell].value
    if old is not None and str(old).strip() != new:
        print(f'  renamed  {sheet:<16} {cell:<6} {str(old)[:56]}  ->  {new}')
        wb[sheet][cell].value = new

wb.save(WB)
print('sheets now:', len(wb.sheetnames))
