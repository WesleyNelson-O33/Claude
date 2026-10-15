"""Make the balance sheet agree with her Xero export.

After the lookup was repointed at the right columns the balance sheet picked
up, but it was still out. Three reasons, all of them the pack's account list
disagreeing with her Xero chart:

  PL Credit Card sat under Current Assets. Xero has it as a Current
    Liability, which it is. Moved, and its Cash flag dropped, because a
    credit card balance is not cash.
  Conversion clearing account had no row at all. Added to Current
    Liabilities, where Xero puts it.
  Current Year Earnings had no row at all. Added to Equity.

Rows are inserted inside the existing section totals, so the sums pick them
up. Every account row formula is written out again afterwards, because
inserting a row does not repoint a formula that names its own row.
"""
import os, shutil
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = f'{HERE}/CTS Financial Controller Pack - v13.xlsx'
OUT = f'{HERE}/CTS Financial Controller Pack - v14.xlsx'

shutil.copy(SRC, OUT)
wb = openpyxl.load_workbook(OUT)
b = wb['Balance_Sheet']

assert b['A10'].value.strip() == 'PL Credit Card', b['A10'].value
b.delete_rows(10)                       # out of Current Assets

# Current Liabilities now runs 47..77. Two new rows go in before the blank.
assert b['A46'].value.strip() == 'Current Liabilities', b['A46'].value
b.insert_rows(78, 2)
for off, (name, code) in enumerate([('PL Credit Card', '21131  '),
                                    ('Conversion clearing account', '21190  ')]):
    r = 78 + off
    b.cell(r, 1, '    ' + name)
    b.cell(r, 8, code)

# Equity now runs 88..93, with the two inserts above it.
assert b['A88'].value.strip() == 'Equity', b['A88'].value
b.insert_rows(95, 1)
b.cell(95, 1, '    Current Year Earnings')
b.cell(95, 8, '33200  ')

# Carry the look of the row above onto each new row.
for r in (78, 79, 95):
    for c in range(1, 9):
        src = b.cell(r - 1, c)
        tgt = b.cell(r, c)
        tgt.font = Font(name=src.font.name, sz=src.font.sz, b=src.font.b,
                        i=src.font.i, color=src.font.color)
        if src.fill.fill_type == 'solid':
            tgt.fill = PatternFill('solid', fgColor=src.fill.start_color.rgb)
        tgt.alignment = Alignment(horizontal=src.alignment.horizontal,
                                  vertical=src.alignment.vertical,
                                  indent=src.alignment.indent)
        tgt.number_format = src.number_format
        tgt.border = Border(left=src.border.left, right=src.border.right,
                            top=src.border.top, bottom=src.border.bottom)
    b.row_dimensions[r].height = b.row_dimensions[r - 1].height

# Every account row gets its formulas written out again against its own row.
SECTIONS = {}
cur = None
for r in range(6, 100):
    a = b.cell(r, 1).value
    if not isinstance(a, str) or not a.strip():
        continue
    if a.startswith('Total ') or a.startswith('Check') or a.startswith('Key'):
        break
    if not a.startswith('    '):
        cur = r
        SECTIONS[cur] = []
    elif cur:
        SECTIONS[cur].append(r)

IDX = ('IFERROR(INDEX(BS_Paste!$A$1:$P$400,'
       'MATCH(TRIM($A{r}),BS_Paste!$B$1:$B$400,0),Setup!$B${s}),0)')
for head, rows in SECTIONS.items():
    for r in rows:
        for col, s in ((2, 30), (3, 31), (4, 32)):
            b.cell(r, col, f'=IF(Setup!$B${s}=0,0,{IDX.format(r=r, s=s)})')
        b.cell(r, 5, f'=$B{r}-$C{r}')
        b.cell(r, 6, f'=$B{r}-$D{r}')
    for col in range(2, 7):
        cl = b.cell(head, col).column_letter
        b.cell(head, col, f'=SUM({cl}{rows[0]}:{cl}{rows[-1]})')
    print(f'  {str(b.cell(head,1).value)[:26]:<26} rows {rows[0]}..{rows[-1]}')

heads = list(SECTIONS)
ca, nca, cl, ncl, eq = heads
tot = max(SECTIONS[eq]) + 2
for col in range(2, 7):
    c = b.cell(tot, col).column_letter
    b.cell(tot,     col, f'={c}{ca}+{c}{nca}')
    b.cell(tot + 1, col, f'={c}{cl}+{c}{ncl}')
    b.cell(tot + 2, col, f'={c}{tot}-{c}{tot+1}')
    b.cell(tot + 3, col, f'={c}{eq}')
for col in range(2, 5):
    c = b.cell(tot, col).column_letter
    b.cell(tot + 5, col, f'=ROUND({c}{tot+2}-{c}{tot+3},2)')
last = SECTIONS[eq][-1]
for off, tag in enumerate(['Cash', 'Trade debtors', 'Trade creditors', 'WIP', 'Borrowings']):
    for col in range(2, 5):
        c = b.cell(tot, col).column_letter
        b.cell(tot + 8 + off, col,
               f'=SUMIF($H$6:$H${last},"*"&"{tag}",{c}$6:{c}${last})')
wc = tot + 15
b.cell(wc,     2, f'=$B${ca}-$B${cl}')
b.cell(wc + 1, 2, f'=IF($B${cl}=0,"",$B${ca}/$B${cl})')
b.cell(wc + 2, 2, f'=IF($B${cl}=0,"",($B${ca}-$B${tot+11})/$B${cl})')
b.cell(wc + 3, 2, f'=$B${tot+12}-$B${tot+8}')

wb.save(OUT)
print('saved', OUT)
